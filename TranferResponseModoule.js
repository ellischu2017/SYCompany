/**
 * TransferResponse.gs
 * 負責處理資料轉換邏輯
 */

// 定義試算表 URL (來自 Transfer.md)
const SS_URL_RAW_RESPONSES =
  "https://docs.google.com/spreadsheets/d/1Lg57GtSMtZRi3dTWX2-Xr-3slc1__UT7H1kExwAh_eM/edit?usp=drive_link";
// const SS_URL_CASE_REPORTS = "https://docs.google.com/spreadsheets/d/1ib8q-lKJgLEhRVrwncnRqOyKNauMqaV2wtYEpGlmRlk";

// 假設本腳本綁定在 SYTemp 所在的試算表，或者透過 getTargetsheet 獲取
// 這裡為了獨立運作，假設 MainSpreadsheet 是 SYTemp (或包含 SYTemp sheet 的檔案)
// 如果不是，請使用 SpreadsheetApp.openByUrl(...)

/**
 * 取得 Raw Responses 試算表中的所有 Sheet 名稱 (即案主名單)
 * 用於前端下拉選單
 */
function getRawResponseSheetNames() {
  try {
    const ss = SpreadsheetApp.openByUrl(SS_URL_RAW_RESPONSES);
    const sheets = ss.getSheets();
    // 過濾掉可能的系統隱藏表或說明表，假設大部分是案主名
    // 這裡回傳所有 Sheet Name
    return sheets.map((s) => s.getName());
  } catch (e) {
    throw new Error("無法存取 Raw Responses 試算表: " + e.message);
  }
}

/**
 * 主要轉換函式
 * @param {string} custName 案主名稱 (對應 Sheet 名稱)
 * @param {boolean} isUpdateOnly 是否僅更新 (true: > TDate, false: <= TDate)
 */
function processTransferData(custName, isUpdateOnly) {
  console.log(`processTransferData 開始處理: ${custName}, UpdateOnly: ${isUpdateOnly}`);
  // 初始化結果物件
  const globalResult = { success: true, count: 0, log: "", message: "" };

  try {
    let customersToProcess = [];

    if (custName === "all") {
      // 如果選全部，取得所有 Sheet 名稱
      customersToProcess = getRawResponseSheetNames();
      globalResult.log += `[Info] 開始批次處理所有案主 (共 ${customersToProcess.length} 位)...\n`;
    } else {
      // 單一案主
      customersToProcess = [custName];
    }

    // 為了效能，這裡先開啟目標 Spreadsheet (避免在迴圈內重複開啟)
    // 假設 Utilities.getTargetsheet 可用，若無請自行替換為 SpreadsheetApp.openByUrl(...)
    const ssCurrent = getTargetsheet("SYTemp", "SYTemp");
    const sheetSRData = ssCurrent.getSheetByName("SR_Data");
    const sheetTransResp = ssCurrent.getSheetByName("Transfer_Response");

    if (!sheetSRData || !sheetTransResp) {
      throw new Error("找不到目標工作表 (SR_Data 或 Transfer_Response)");
    }

    // 快取現有資料 Key，避免重複讀取 (Performance Optimization)
    // 注意：若資料量極大，可能需要在迴圈內分批處理，但目前先維持全讀取
    const existingData = sheetSRData.getDataRange().getValues();
    const existingKeys = new Set();
    // 建立索引 Key: 日期(0)_案主(2)_居服員(3)_SR_ID(5)
    for (let r = 1; r < existingData.length; r++) {
      const exDate = formatDate(existingData[r][0]);
      const exSRTimes = existingData[r][1];
      const exCust = existingData[r][2];
      const exUser = existingData[r][3];
      const exId = existingData[r][5];
      existingKeys.add(`${exDate}_${exSRTimes}_${exCust}_${exUser}_${exId}`);
    }

    // --- 迴圈處理案主 ---
    for (const customer of customersToProcess) {
      // 呼叫單一處理邏輯
      const result = processSingleCustomerInternal(
        customer,
        isUpdateOnly,
        sheetSRData,
        sheetTransResp,
        existingKeys,
      );

      // 累加結果
      globalResult.count += result.count;
      globalResult.log += result.log;

      // 若單一處理發生嚴重錯誤(非資料面)，可選擇是否中斷，這裡選擇記錄錯誤但繼續執行
      if (!result.success) {
        globalResult.log += `[Error] ${customer} 處理失敗: ${result.message}\n`;
      }
    }

    globalResult.log += `[Done] 所有作業完成。總新增筆數: ${globalResult.count}`;
  } catch (err) {
    globalResult.success = false;
    globalResult.message = err.message;
    globalResult.log += `[Critical Error] ${err.message}\n${err.stack}`;
  }

  return globalResult;
}

/**
 * 內部核心邏輯：處理單一案主
 * (將原本 processTransferData 的邏輯搬移至此，並接受共用的 Sheet 物件與 Key Set 以提升效能)
 */
function processSingleCustomerInternal(
  custName,
  isUpdateOnly,
  sheetSRData,
  sheetTransResp,
  existingKeys,
) {
  const result = { success: false, count: 0, log: "", message: "" };

  try {
    const ssRaw = SpreadsheetApp.openByUrl(SS_URL_RAW_RESPONSES);
    const sheetSource = ssRaw.getSheetByName(custName);

    // 如果來源表單不存在 (例如系統隱藏表單)，跳過並回傳成功
    if (!sheetSource) {
      result.log = `[Skip] 找不到來源工作表: ${custName} (略過)\n`;
      result.success = true;
      return result;
    }

    // 2. 取得 TDate
    const tDate = getTDate(sheetTransResp, custName);
    result.log += `--- 處理案主: ${custName} (上次更新: ${formatDate(tDate)}) ---\n`;

    // 3. 讀取來源資料
    const sourceData = sheetSource.getDataRange().getValues();
    if (sourceData.length < 2) {
      result.log += `[Info] 無資料可處理。\n`;
      result.success = true;
      return result;
    }

    const headers = sourceData[0];
    const idxDate = headers.indexOf("日期");
    const idxUser = headers.indexOf("居服員姓名");
    const idxLoc = headers.indexOf("意識狀況");
    const idxMood = headers.indexOf("身心狀況");
    const idxSpCons = headers.indexOf("有特殊狀況，請說明及處理");
    const idxHasTemp = headers.indexOf("是否有其他臨時服務項目");
    const idxTempItem = headers.indexOf(
      "臨時服務項目(請填寫服務代碼+項目名稱)",
    );
    const idxTempDesc = headers.indexOf("說明");

    if (idxDate === -1) {
      result.log += `[Warn] 缺少 '日期' 欄位，跳過此案主。\n`;
      result.success = true;
      return result;
    }

    // 4. 過濾與轉換資料
    const rowsToAdd = [];
    let maxDateProcessed = null;

    for (let i = 1; i < sourceData.length; i++) {
      const row = sourceData[i];
      const rowDateRaw = row[idxDate];
      if (!rowDateRaw) continue;

      const rowDate = new Date(rowDateRaw);
      const isDateValid = isUpdateOnly ? rowDate > tDate : rowDate <= tDate;

      if (isDateValid) {
        if (!maxDateProcessed || rowDate > maxDateProcessed) {
          maxDateProcessed = rowDate;
        }

        let rawSpCons = idxSpCons > -1 ? row[idxSpCons] : "";
        const finalSpCons =
          rawSpCons && String(rawSpCons).trim() !== "" ? rawSpCons : "無";

        const baseData = {
          date: formatDate(rowDate),
          SRTimes: "1",
          custName: custName,
          userName: idxUser > -1 ? row[idxUser] : "",
          payType: "補助",
          loc: idxLoc > -1 ? row[idxLoc] : "",
          mood: idxMood > -1 ? row[idxMood] : "",
          spCons: finalSpCons,
        };

        // 5. 處理 BA 碼
        for (let c = 0; c < headers.length; c++) {
          const header = headers[c];
          const cellValue = row[c];
          if (
            header.match(/^[A-Za-z][0-9a-zA-Z]+/) &&
            cellValue &&
            String(cellValue).trim() !== "" &&
            String(cellValue).trim() !== "無"
          ) {
            const srIdMatch = header.match(/^([A-Za-z][0-9a-zA-Z\-]+)/);
            const srId = srIdMatch ? srIdMatch[0] : header;

            rowsToAdd.push([
              baseData.date,
              baseData.SRTimes,
              baseData.custName,
              baseData.userName,
              baseData.payType,
              srId,
              cellValue,
              baseData.loc,
              baseData.mood,
              baseData.spCons,
            ]);
          }
        }

        // 6. 處理 臨時服務
        if (idxHasTemp > -1 && String(row[idxHasTemp]).trim() === "是") {
          const tempIdRaw = idxTempItem > -1 ? String(row[idxTempItem]) : "";
          const tempRec = idxTempDesc > -1 ? String(row[idxTempDesc]) : "";
          let tempId = "TEMP";
          const tempMatch = tempIdRaw.match(/^([A-Za-z][0-9a-zA-Z\-]+)/);
          if (tempMatch) tempId = tempMatch[0];
          else if (tempIdRaw !== "") tempId = tempIdRaw;

          if (tempRec !== "") {
            rowsToAdd.push([
              baseData.date,
              baseData.SRTimes,
              baseData.custName,
              baseData.userName,
              baseData.payType,
              tempId,
              tempRec,
              baseData.loc,
              baseData.mood,
              baseData.spCons,
            ]);
          }
        }
      }
    }

    // 7. 寫入資料 (使用傳入的 existingKeys 檢查重複)
    if (rowsToAdd.length > 0) {
      const uniqueRows = rowsToAdd.filter((row) => {
        const key = `${row[0]}_${row[1]}_${row[2]}_${row[3]}_${row[5]}`;
        if (existingKeys.has(key)) return false;
        existingKeys.add(key);
        return true;
      });

      if (uniqueRows.length > 0) {
        const lastRow = sheetSRData.getLastRow();
        sheetSRData
          .getRange(lastRow + 1, 1, uniqueRows.length, uniqueRows[0].length)
          .setValues(uniqueRows);

        result.count = uniqueRows.length;
        result.log += `> 新增 ${uniqueRows.length} 筆資料。\n`;

        if (isUpdateOnly && maxDateProcessed) {
          updateTDate(sheetTransResp, custName, maxDateProcessed);
        }
      } else {
        result.log += `> 資料皆已存在，無新增。\n`;
      }
    } else {
      // 沒有符合日期的資料，但也許需要更新 TDate?
      // 邏輯上如果是 UpdateOnly 且沒資料，就不更新 TDate
      result.log += `> 無需更新資料。\n`;
    }

    result.success = true;
  } catch (err) {
    result.success = false;
    result.message = err.message;
    result.log += `> 錯誤: ${err.message}\n`;
  }

  return result;
}

/**
 * 輔助：從 Transfer_Response 取得特定案主的上次更新日期
 * 若無資料，回傳一個極早的日期 (如 2000-01-01)
 */
function getTDate(sheet, custName) {
  const data = sheet.getDataRange().getValues();
  // 假設 Column A 是 Cust_N, Column B 是 TDate
  // 從第 2 列開始搜尋
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == custName) {
      const d = new Date(data[i][1]);
      return isNaN(d.getTime()) ? new Date("2000-01-01") : d;
    }
  }
  return new Date("2000-01-01");
}

/**
 * 輔助：更新 Transfer_Response 的 TDate
 * 若無該案主則新增一列
 */
function updateTDate(sheet, custName, newDate) {
  const data = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == custName) {
      // 更新日期 (Column B -> Index 1)
      // 使用 yyyy-MM-dd 格式字串寫入，避免時區問題
      sheet.getRange(i + 1, 2).setValue(formatDate(newDate));
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([custName, formatDate(newDate)]);
  }
}

/**
 * 格式化日期為 yyyy-MM-dd
 */
function formatDate(date) {
  return Utilities.formatDate(
    new Date(date),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd",
  );
}

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
    // const ss = SpreadsheetApp.openByUrl(SS_URL_RAW_RESPONSES);
    const ss = getTargetsheet("SYTemp", "Raw_Responses").Spreadsheet;
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
    const ssCurrent = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    const sheetSRData = ssCurrent.getSheetByName("SR_Data");
    const sheetTransResp = ssCurrent.getSheetByName("Transfer_Response");

    if (!sheetSRData || !sheetTransResp) {
      throw new Error("找不到目標工作表 (SR_Data 或 Transfer_Response)");
    }

    // 來源試算表 (Raw Responses) 也只開啟一次
    const ssRaw = SpreadsheetApp.openByUrl(SS_URL_RAW_RESPONSES);

    // [Batch Optimization 1] 預先讀取所有案主的 TDate (Transfer_Response)
    const tDateData = sheetTransResp.getDataRange().getValues();
    const tDateMap = new Map(); // Key: CustName, Value: Date
    if (tDateData.length > 1) {
      const h = tDateData[0];
      const iName = getColIndex(h, "Cust_N") > -1 ? getColIndex(h, "Cust_N") : 0;
      const iDate = getColIndex(h, "TDate") > -1 ? getColIndex(h, "TDate") : 1;
      for (let i = 1; i < tDateData.length; i++) {
        const d = new Date(tDateData[i][iDate]);
        // 若日期無效則設為很久以前
        tDateMap.set(tDateData[i][iName], isNaN(d.getTime()) ? new Date("2000-01-01") : d);
      }
    }

    // 快取現有資料 Key，避免重複讀取 (Performance Optimization)
    // 注意：若資料量極大，可能需要在迴圈內分批處理，但目前先維持全讀取
    const existingData = sheetSRData.getDataRange().getValues();
    const existingKeys = new Set();
    // 建立索引 Key: 日期(0)_案主(2)_居服員(3)_SR_ID(5)
    for (let r = 1; r < existingData.length; r++) {
      const exDate = formatDate(existingData[r][0]);
      // const exSRTimes = existingData[r][1];
      const exCust = existingData[r][2];
      const exUser = existingData[r][3];
      const exId = existingData[r][5];
      // Key 規則需與下方 processSingleCustomerInternal 一致
      existingKeys.add(`${exDate}_${existingData[r][1]}_${exCust}_${exUser}_${exId}`);

      // existingKeys.add(`${exDate}_${exSRTimes}_${exCust}_${exUser}_${exId}`);
    }

    // [Batch Optimization 2] 準備批次容器
    const batchRowsToAdd = [];
    const batchTDateUpdates = new Map(); // CustName -> NewDate
    const batchMonthsToClear = new Set();

    // --- 迴圈處理案主 ---
    for (const customer of customersToProcess) {
      // 防呆機制：檢查執行時間是否即將超時 (引用 Utilities.js 中的 isNearTimeout)
      // 預設 EXECUTION_TIMEOUT_MINUTES 為 5 分鐘，保留 1 分鐘緩衝進行收尾
      if (typeof isNearTimeout === 'function' && isNearTimeout()) {
        console.warn("系統執行時間即將逾時，為避免 Quota Exceeded 錯誤已中止後續處理。");
        globalResult.log += `[System] 執行時間即將逾時，已中止於案主: ${customer} (後續未處理)\n`;
        globalResult.message += " (因時間限制部分中斷)";
        break;
      }

      // 呼叫單一處理邏輯
      const result = processSingleCustomerInternal(
        ssRaw,          // 傳入已開啟的 Spreadsheet 物件
        customer,
        isUpdateOnly,
        // sheetSRData,
        // sheetTransResp,
        tDateMap.get(customer) || new Date("2000-01-01"), // 直接傳入 TDate
        existingKeys,
        batchRowsToAdd,    // 收集新增資料
        batchTDateUpdates, // 收集 TDate 更新
        batchMonthsToClear // 收集受影響月份
      );

      // 累加結果
      globalResult.count += result.count;
      globalResult.log += result.log;

      // 若單一處理發生嚴重錯誤(非資料面)，可選擇是否中斷，這裡選擇記錄錯誤但繼續執行
      if (!result.success) {
        globalResult.log += `[Error] ${customer} 處理失敗: ${result.message}\n`;
      }
    }

    // --- [Batch Optimization 3] 迴圈結束後，一次性寫入 ---

    // 1. 批次寫入 SR_Data
    if (batchRowsToAdd.length > 0) {
      const lastRow = sheetSRData.getLastRow();
      sheetSRData.getRange(lastRow + 1, 1, batchRowsToAdd.length, batchRowsToAdd[0].length)
        .setValues(batchRowsToAdd);
      globalResult.log += `> [Batch] 成功寫入 ${batchRowsToAdd.length} 筆資料至 SR_Data。\n`;
    }

    // 2. 批次更新 Transfer_Response (TDate)
    if (batchTDateUpdates.size > 0) {
      batchUpdateTDate(sheetTransResp, batchTDateUpdates);
      globalResult.log += `> [Batch] 更新 ${batchTDateUpdates.size} 筆案主的 TDate。\n`;
    }

    // 3. 批次清除 Cache
    if (batchMonthsToClear.size > 0) {
      batchMonthsToClear.forEach(ym => {
        CacheService.getScriptCache().remove("CustN_" + ym);
      });
      globalResult.log += `> [Batch] 清除快取月份: ${Array.from(batchMonthsToClear).join(", ")}\n`;
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
  * (修改為純內存處理，不直接寫入 Sheet)
 */
function processSingleCustomerInternal(
  ssRaw,
  custName,
  isUpdateOnly,
  // sheetSRData,
  // sheetTransResp,
  tDate,          // 直接傳入 Date 物件
  existingKeys,
  batchRowsToAdd,    // 接收陣列
  batchTDateUpdates, // 接收 Map
  batchMonthsToClear // 接收 Set
) {
  const result = { success: false, count: 0, log: "", message: "" };

  try {
    // const ssRaw = SpreadsheetApp.openByUrl(SS_URL_RAW_RESPONSES);
    // 使用傳入的 ssRaw，不重複開啟
    const sheetSource = ssRaw.getSheetByName(custName);

    // 如果來源表單不存在 (例如系統隱藏表單)，跳過並回傳成功
    if (!sheetSource) {
      result.log = `[Skip] 找不到來源工作表: ${custName} (略過)\n`;
      result.success = true;
      return result;
    }

    // 2. 取得 TDate
    // const tDate = getTDate(sheetTransResp, custName);
    // TDate 已由外部傳入
    result.log += `--- 處理案主: ${custName} (上次更新: ${formatDate(tDate)}) ---\n`;

    // 3. 讀取來源資料
    const sourceData = sheetSource.getDataRange().getValues();
    if (sourceData.length < 2) {
      result.log += `[Info] 無資料可處理。\n`;
      result.success = true;
      return result;
    }

    const headers = sourceData[0];
    const idxDate = getColIndex(headers, "日期");
    const idxUser = getColIndex(headers, "居服員姓名");
    const idxLoc = getColIndex(headers, "意識狀況");
    const idxMood = getColIndex(headers, "身心狀況");
    const idxSpCons = getColIndex(headers, "有特殊狀況，請說明及處理");
    const idxHasTemp = getColIndex(headers, "是否有其他臨時服務項目");
    const idxTempItem = getColIndex(headers, "臨時服務項目(請填寫服務代碼+項目名稱)");
    const idxTempDesc = getColIndex(headers, "說明");

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
              String(baseData.userName).trim(),
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
              String(baseData.userName).trim(),
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
    // 7. 將過濾後的資料加入 Batch 容器
    if (rowsToAdd.length > 0) {
      const uniqueRows = rowsToAdd.filter((row) => {
        const key = `${row[0]}_${row[1]}_${row[2]}_${row[3]}_${row[5]}`;
        if (existingKeys.has(key)) return false;
        existingKeys.add(key);
        return true;
      });

      if (uniqueRows.length > 0) {
        // const lastRow = sheetSRData.getLastRow();
        // sheetSRData
        // .getRange(lastRow + 1, 1, uniqueRows.length, uniqueRows[0].length)
        // .setValues(uniqueRows);
        // 加入 Batch 佇列
        batchRowsToAdd.push(...uniqueRows);

        // --- 清除相關月份的快取 ---
        // const addedMonths = new Set();
        // 收集受影響月份
        uniqueRows.forEach(row => {
          // 日期在第一欄 (index 0)
          const d = new Date(row[0]);
          const yyyyMM = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyyMM");
          // addedMonths.add(yyyyMM);
          batchMonthsToClear.add(yyyyMM);
        });

        // addedMonths.forEach(ym => {
        // CacheService.getScriptCache().remove("CustN_" + ym);
        // });

        // if (addedMonths.size > 0) {
        // result.log += `> 已清除受影響月份的快取: ${Array.from(addedMonths).join(", ")}\n`;
        // }

        result.count = uniqueRows.length;
        result.log += `> 新增 ${uniqueRows.length} 筆資料。\n`;

        // 紀錄 TDate 更新
        if (isUpdateOnly && maxDateProcessed) {
          // updateTDate(sheetTransResp, custName, maxDateProcessed);
          batchTDateUpdates.set(custName, maxDateProcessed);

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
 * 輔助：批次更新 Transfer_Response 的 TDate
 */
function batchUpdateTDate(sheet, updatesMap) {
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    // 初始化標題
    sheet.appendRow(["Cust_N", "TDate"]);
    data.push(["Cust_N", "TDate"]);
  }
  
  const headers = data[0];
  let idxName = getColIndex(headers, "Cust_N");
  let idxDate = getColIndex(headers, "TDate");
  if (idxName === -1) idxName = 0;
  if (idxDate === -1) idxDate = 1;

  // 1. 更新現有列
  const updatedRows = data.map((row, i) => {
    if (i === 0) return row; // 標題
    const name = row[idxName];
    if (updatesMap.has(name)) {
      row[idxDate] = formatDate(updatesMap.get(name));
      updatesMap.delete(name); // 標記已處理
    }
    return row;
  });

  // 2. 新增未存在的案主
  updatesMap.forEach((newDate, name) => {
    const newRow = new Array(headers.length).fill("");
    newRow[idxName] = name;
    newRow[idxDate] = formatDate(newDate);
    updatedRows.push(newRow);
  });

  // 3. 一次性寫回
  sheet.getRange(1, 1, updatedRows.length, headers.length).setValues(updatedRows);
}

/**
 * 更新 Raw Responses 
 * @param {*} formObj 
 * @returns 
 */
function UpdateRawResponse(formObj) {
  //Raw Response::  https://docs.google.com/spreadsheets/d/1Lg57GtSMtZRi3dTWX2-Xr-3slc1__UT7H1kExwAh_eM/edit?usp=drive_link
  //Tset:: https://docs.google.com/spreadsheets/d/1Srk5AVQ2mFmHsHIsdEbSQEFZaVhEMCPLdhpnJ6Y2AhM/edit?usp=drive_link
  const RAW_RESPONSES_URL = "https://docs.google.com/spreadsheets/d/1Srk5AVQ2mFmHsHIsdEbSQEFZaVhEMCPLdhpnJ6Y2AhM/edit?usp=drive_link";
  const targetSs = SpreadsheetApp.openByUrl(RAW_RESPONSES_URL);
  const sheet = targetSs.getSheetByName(formObj.custName);
  Logger.log("UpdateRawResponse Running");
  if (!sheet) return;
  try {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    //先讀取所有欄位名稱   
    const idxstamp = getColIndex(headers, "時間戳記");
    const idxDate = getColIndex(headers, "日期");
    const idxUser = getColIndex(headers, "居服員姓名");
    const idxLoc = getColIndex(headers, "意識狀況");
    const idxMood = getColIndex(headers, "身心狀況");
    const idxSpCons = getColIndex(headers, "特殊狀況");
    const idxSpConsDeal = getColIndex(headers, "有特殊狀況，請說明及處理");
    const idxHasTemp = getColIndex(headers, "是否有其他臨時服務項目");
    const idxTempItem = getColIndex(headers, "臨時服務項目(請填寫服務代碼+項目名稱)");
    const idxTempDesc = getColIndex(headers, "說明");
    var update = false;
    //篩選 data Date === formObj.date 及 User === formObj.userName 的資料
    for (let i = 1; i < data.length; i++) {
      var rowDate = data[i][idxDate];
      var sheetDate = (rowDate instanceof Date) ? Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy-MM-dd") : String(rowDate);
      if (sheetDate === formObj.date && String(data[i][idxUser]).trim() === formObj.userName) {
        update = true;
        if (idxstamp !== -1) sheet.getRange(i + 1, idxstamp + 1).setValue(Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/M/d hh:mm:ss"));
        if (idxDate !== -1) sheet.getRange(i + 1, idxDate + 1).setValue(Utilities.formatDate(rowDate, "Asia/Taipei", "yyyy/M/d"));
        if (idxLoc !== -1) sheet.getRange(i + 1, idxLoc + 1).setValue(formObj.loc || "清醒");
        if (idxMood !== -1) sheet.getRange(i + 1, idxMood + 1).setValue(formObj.mood || "穩定");
        if (formObj.spcons !== "無") {
          if (idxSpCons !== -1) sheet.getRange(i + 1, idxSpCons + 1).setValue("有");
          if (idxSpConsDeal !== -1) sheet.getRange(i + 1, idxSpConsDeal + 1).setValue(formObj.spcons);
        }
        var found = false;
        for (let c = 0; c < headers.length; c++) {
          const header = headers[c];
          // The header in the Raw Response sheet might be "BA01a" or "BA01a - Description".
          // We check if the header starts with the service ID from the form.
          if (formObj.srId && header.trim().startsWith(formObj.srId.substring(0, 4))) {
            // If a match is found, update the cell in that column with the service record.
            sheet.getRange(i + 1, c + 1).setValue(formObj.srRec || "");
            if (idxHasTemp !== -1 && sheet.getRange(i + 1, idxHasTemp + 1).getValue() !== "是") {
              sheet.getRange(i + 1, idxHasTemp + 1).setValue("否");
            }
            found = true;
            break;
          }
        }
        if (!found) {
          sheet.getRange(i + 1, idxHasTemp + 1).setValue("是");
          sheet.getRange(i + 1, idxTempItem + 1).setValue(formObj.srId);
          sheet.getRange(i + 1, idxTempDesc + 1).setValue(formObj.srRec);
          // Assuming one srId corresponds to only one column, we can stop searching.            
        }
      }
    }

    if (!update) {
      // 建立一個與標題等長的空陣列，確保欄位對應正確
      var newRow = new Array(headers.length).fill("");
      if (idxstamp !== -1) newRow[idxstamp] = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/M/d hh:mm:ss");
      if (idxDate !== -1) newRow[idxDate] = Utilities.formatDate(new Date(formObj.date), "Asia/Taipei", "yyyy/M/d");
      if (idxUser !== -1) newRow[idxUser] = formObj.userName;
      if (idxLoc !== -1) newRow[idxLoc] = formObj.loc || "清醒";
      if (idxMood !== -1) newRow[idxMood] = formObj.mood || "穩定";
      if (formObj.spcons !== "無") {
        if (idxSpCons !== -1) newRow[idxSpCons] = "有";
        if (idxSpConsDeal !== -1) newRow[idxSpConsDeal] = formObj.spcons;
      } else {
        if (idxSpCons !== -1) newRow[idxSpCons] = "無";
        if (idxSpConsDeal !== -1) newRow[idxSpConsDeal] = "";
      }
      var found = false;
      for (let c = 0; c < headers.length; c++) {
        const header = headers[c];
        if (header.trim().startsWith(formObj.srId.substring(0, 4))) {
          newRow[c] = formObj.srRec || "";
          found = true;
          break;
        }
      }
      if (!found) {
        if (idxHasTemp !== -1) newRow[idxHasTemp] = "是";
        if (idxTempItem !== -1) newRow[idxTempItem] = formObj.srId;
        if (idxTempDesc !== -1) newRow[idxTempDesc] = formObj.srRec;
      }
      sheet.appendRow(newRow);
    }

  } catch (e) {
    Logger.log("更新 Raw Response 失敗: " + e.toString());
  }
}



/**
 * 輔助：從 Transfer_Response 取得特定案主的上次更新日期
 * 若無資料，回傳一個極早的日期 (如 2000-01-01)
 */
function getTDate(sheet, custName) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return new Date("2000-01-01");

  const headers = data[0];
  let idxName = getColIndex(headers, "Cust_N");
  let idxDate = getColIndex(headers, "TDate");

  // 防呆：若找不到標題，預設為第 1, 2 欄
  if (idxName === -1) idxName = 0;
  if (idxDate === -1) idxDate = 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idxName] == custName) {
      const d = new Date(data[i][idxDate]);
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

  // 若是空表，初始化標題
  if (data.length === 0) {
    sheet.appendRow(["Cust_N", "TDate"]);
    sheet.appendRow([custName, formatDate(newDate)]);
    return;
  }

  const headers = data[0];
  let idxName = getColIndex(headers, "Cust_N");
  let idxDate = getColIndex(headers, "TDate");

  if (idxName === -1) idxName = 0;
  if (idxDate === -1) idxDate = 1;

  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idxName] == custName) {
      sheet.getRange(i + 1, idxDate + 1).setValue(formatDate(newDate));
      found = true;
      break;
    }
  }

  if (!found) {
    // 若有標題結構，依照結構寫入；否則直接附加
    const newRow = new Array(headers.length).fill("");
    newRow[idxName] = custName;
    newRow[idxDate] = formatDate(newDate);
    sheet.appendRow(newRow);
  }
}

/**
 * SRServer.gs - 服務紀錄單管理模組
 * 提供服務紀錄的查詢、新增、修改、刪除操作
 */

/**
 * 初始化 SR_server 頁面所需的所有資料
 * 需求 1, 2, 3: 一次性抓取 SYCompany 中的三个表
 */
function getSRServerInitData() {
  // 1. 取得 User 資料 (包含 User_N, User_Email, Cust_N)
  const userSheet = MainSpreadsheet.getSheetByName("User");
  var userData = [];
  if (userSheet) {
    const rawUsers = userSheet.getDataRange().getValues();
    if (rawUsers.length > 0) {
      const headers = rawUsers[0];
      const userFields = ["User_N", "User_Email", "Email", "Cust_N"];
      const userColMap = getColIndicesMap(headers, userFields);

      const idxName = userColMap["User_N"];
      const idxEmail = userColMap["User_Email"] !== -1 ? userColMap["User_Email"] : userColMap["Email"];
      const idxCust = userColMap["Cust_N"];

      if (idxName !== -1) {
        for (let i = 1; i < rawUsers.length; i++) {
          const row = rawUsers[i];
          if (row[idxName]) {
            userData.push({
              name: row[idxName].toString(),
              email: idxEmail !== -1 ? row[idxEmail].toString() : "",
              custStr: idxCust !== -1 ? row[idxCust].toString() : ""
            });
          }
        }
      }
    }
  }

  // 2. 取得 Cust 資料 (包含 Cust_N, Cust_LTC_Code)
  const custSheet = MainSpreadsheet.getSheetByName("Cust");
  var custData = [];
  if (custSheet) {
    const rawCusts = custSheet.getDataRange().getValues();
    if (rawCusts.length > 0) {
      const cHeaders = rawCusts[0];
      const custFields = ["Cust_N", "LTC_Code"];
      const custColMap = getColIndicesMap(cHeaders, custFields);
      const idxCName = custColMap["Cust_N"];
      const idxCLTC = custColMap["LTC_Code"];

      if (idxCName !== -1) {
        for (let j = 1; j < rawCusts.length; j++) {
          const row = rawCusts[j];
          if (row[idxCName]) {
            custData.push({
              name: row[idxCName].toString(),
              ltcStr: idxCLTC !== -1 ? row[idxCLTC].toString() : ""
            });
          }
        }
      }
    }
  }

  // 3. 取得 LTC_Code 資料 (SR_ID)
  const ltcSheet = MainSpreadsheet.getSheetByName("LTC_Code");
  var ltcIds = [];
  if (ltcSheet) {
    const rawLtc = ltcSheet.getDataRange().getValues();
    if (rawLtc.length > 0) {
      const lHeaders = rawLtc[0];
      const ltcFields = ["SR_ID", "SR_Cont"];
      const ltcColMap = getColIndicesMap(lHeaders, ltcFields);
      const idxSRID = ltcColMap["SR_ID"];
      const idxCont = ltcColMap["SR_Cont"];
      const targetIdx = idxSRID !== -1 ? idxSRID : 0; // 若找不到標題，預設第一欄
      const seen = {};

      for (let k = 1; k < rawLtc.length; k++) {
        const code = rawLtc[k][targetIdx].toString().trim();
        const cont = idxCont !== -1 ? rawLtc[k][idxCont].toString().trim() : "";
        if (code && !seen[code]) {
          seen[code] = true;
          ltcIds.push({ id: code, cont: cont });
        }
      }
    }
  }

  return {
    users: userData,
    custs: custData,
    ltcCodes: ltcIds
  };
}

/**
 * 檢查 SYTemp 的 User 工作表並同步回 SYCompany
 */
function processUserSync01() {
  try {
    // 1. 取得來源表 SYTemp > User
    var ssTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    var tempUserSheet = ssTemp.getSheetByName("User");
    if (!tempUserSheet) {
      console.log("SYTemp 中不存在 User 工作表，無法同步。");
      return;
    }

    var tempValues = tempUserSheet.getDataRange().getValues();
    if (tempValues.length <= 1) {
      console.log("SYTemp 的 User 表無資料可同步。");
      return;
    }

    // 2. 取得目標表 SYCompany > User
    var companyUserSheet = MainSpreadsheet.getSheetByName("User");
    if (!companyUserSheet) {
      companyUserSheet = MainSpreadsheet.insertSheet("User");
      companyUserSheet.appendRow(tempValues[0]);
    }

    // 3. 抓取 SYCompany 目前已有的 Email 清單
    const companyData = companyUserSheet.getDataRange().getValues();
    const companyHeaders = companyData.length > 0 ? companyData[0] : [];
    const companyEmailIdx = getColIndex(companyHeaders, "User_Email");
    const existingEmails = new Set();

    if (companyEmailIdx !== -1) {
      for (let i = 1; i < companyData.length; i++) {
        const email = companyData[i][companyEmailIdx];
        if (email) {
          existingEmails.add(String(email).trim().toLowerCase());
        }
      }
    }

    // 4. 過濾出 SYTemp 中「不在」SYCompany 裡的資料
    const tempHeaders = tempValues[0];
    const tempEmailIdx = getColIndex(tempHeaders, "User_Email");
    const rowsToSync = [];

    if (tempEmailIdx !== -1) {
      for (let j = 1; j < tempValues.length; j++) {
        const tempEmail = tempValues[j][tempEmailIdx] ? String(tempValues[j][tempEmailIdx]).trim().toLowerCase() : "";
        if (tempEmail !== "" && !existingEmails.has(tempEmail)) {
          rowsToSync.push(tempValues[j]);
          existingEmails.add(tempEmail); // 確保在同一次執行中不會重複加入
        }
      }
    }

    // 5. 執行寫入與清空
    if (rowsToSync.length > 0) {
      companyUserSheet
        .getRange(
          companyUserSheet.getLastRow() + 1,
          1,
          rowsToSync.length,
          rowsToSync[0].length,
        )
        .setValues(rowsToSync);
      console.log("同步成功，新增了 " + rowsToSync.length + " 筆資料。");
    } else {
      console.log("無新資料需要同步。");
    }

    // 清空 SYTemp
    if (tempValues.length > 1) {
      tempUserSheet.deleteRows(2, tempValues.length - 1);
    }
  } catch (e) {
    console.log("processUserSync01 執行出錯: " + e.toString());
  }
}

/**
 * 處理服務紀錄：根據日期判斷查詢位置
 * 1. <= 7天：SYTemp > SR_Data
 * 2. > 7天 ：SY+yyyy > yyyyMM (工作表名稱)
 */
function processSRData(formObj, actionType) {
  // 根據日期預先產生快取 key，以便後續清除
  let yearmonth = "";
  if (formObj && formObj.date) {
    const d = new Date(formObj.date);
    yearmonth = Utilities.formatDate(d, "Asia/Taipei", "yyyyMM");
  }

  // 優化 1: 查詢請求先檢查快取，避免讀取試算表 (Cache First)
  if (actionType === "query") {
    let cacheKey = `SRDataQuery_${formObj.date}_${formObj.SRTimes}_${formObj.custName}_${formObj.userName}_${formObj.payType}_${formObj.srId}`;
    let cache = CacheService.getScriptCache();
    let cachedData = cache.get(cacheKey);
    if (cachedData) {
      console.log(`processSRData query 從快取讀取, key: ${cacheKey}`);
      return JSON.parse(cachedData);
    }
  }

  try {
    // --- 1. 計算日期與判斷目標 ---
    var inputDate = new Date(formObj.date);
    inputDate.setHours(0, 0, 0, 0);

    var today = new Date();
    // 設定判斷基準日：若今日在10號(含)以前，則上個月資料仍保留在 SYTemp
    var cutoffDate = new Date(today.getFullYear(), today.getMonth(), 1);
    if (today.getDate() <= 10) {
      cutoffDate.setMonth(cutoffDate.getMonth() - 1);
    }

    var targetSS, sheetName;

    if (inputDate >= cutoffDate) {
      targetSS = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
      sheetName = "SR_Data";
    } else {
      var yearStr = inputDate.getFullYear().toString();
      var monthStr = Utilities.formatDate(inputDate, "Asia/Taipei", "yyyyMM");

      targetSS = getTargetsheet("RecUrl", "SY" + yearStr).Spreadsheet;
      sheetName = monthStr;
    }

    if (!targetSS) throw new Error("找不到對應年份的試算表或 SYTemp。");

    var targetSheet = targetSS.getSheetByName(sheetName);

    // 如果是新增(add)且月份工作表不存在，則自動建立並加入標題
    if (!targetSheet && actionType === "add") {
      targetSheet = targetSS.insertSheet(sheetName);
      var headers = [
        "Date",
        "SRTimes",
        "CUST_N",
        "USER_N",
        "Pay_Type",
        "SR_ID",
        "SR_REC",
        "LOC",
        "MOOD",
        "SPCONS",
      ];
      targetSheet.appendRow(headers);
    }

    if (!targetSheet) {
      return { found: false, message: "工作表 " + sheetName + " 尚無資料。" };
    }

    // --- 2. 執行 新增 (add) ---
    var rowData = [
      formObj.date,
      formObj.SRTimes,
      formObj.custName,
      formObj.userName,
      formObj.payType,
      formObj.srId,
      formObj.srRec,
      formObj.loc,
      formObj.mood,
      formObj.spcons,
    ];

    if (actionType === "add") {
      targetSheet.appendRow(rowData);
      UpdateRawResponse(formObj)
      // 清除對應月份的案主列表快取
      if (yearmonth) {
        CacheService.getScriptCache().remove("CustN_" + yearmonth);
      }
      return {
        success: true,
        message: "資料已成功存入 " + targetSS.getName() + " > " + sheetName,
      };
    }

    // --- 3. 執行 查詢 (query) / 更新 (update) / 刪除 (delete) ---
    // 優化策略：避免讀取整張表 (getValues)，改用 TextFinder 搜尋關鍵欄位 (CUST_N)
    var lastRow = targetSheet.getLastRow();
    var lastCol = targetSheet.getLastColumn();

    if (lastRow < 2) {
      return { success: false, found: false, message: "無資料可供查詢。" };
    }

    // 先只讀取標題列以取得欄位索引
    var headers = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const targetFields = ["Date", "SRTimes", "CUST_N", "USER_N", "Pay_Type", "SR_ID", "SR_REC", "LOC", "MOOD", "SPCONS"];
    const colMap = getColIndicesMap(headers, targetFields);

    // 準備候選列號列表 (Candidate Rows)
    var candidateRows = [];

    // 若 CUST_N 欄位存在，使用 TextFinder 快速篩選
    if (colMap["CUST_N"] !== -1) {
      var searchRange = targetSheet.getRange(2, colMap["CUST_N"] + 1, lastRow - 1, 1);
      // matchEntireCell(true) 確保完全匹配，避免部分字串誤判
      var ranges = searchRange.createTextFinder(formObj.custName).matchEntireCell(true).findAll();

      // 將搜尋結果倒序放入候選清單 (模擬從新到舊的邏輯)
      for (var k = ranges.length - 1; k >= 0; k--) {
        candidateRows.push(ranges[k].getRow());
      }
    } else {
      // 降級處理：若找不到 CUST_N 欄位，則不得不遍歷所有列 (倒序)
      for (let r = lastRow; r >= 2; r--) {
        candidateRows.push(r);
      }
    }

    // 遍歷候選列 (倒序)
    for (let rIdx of candidateRows) {
      // 讀取該單列的完整資料
      // 這裡權衡了 API 呼叫次數與資料傳輸量，針對「特定個案」查詢，這種方式通常快很多
      var row = targetSheet.getRange(rIdx, 1, 1, lastCol).getValues()[0];


      // 優化 3: 避免在迴圈內使用 Utilities.formatDate，改用原生 Date 比對
      let rawDate = row[colMap["Date"]];
      let isDateMatch = false;
      if (rawDate instanceof Date) {
        // 比對年月日 (注意 getMonth 從 0 開始)
        let d = rawDate;
        let dateStr = d.getFullYear() + "-" +
          ("0" + (d.getMonth() + 1)).slice(-2) + "-" +
          ("0" + d.getDate()).slice(-2);
        if (dateStr === formObj.date) isDateMatch = true;
      } else {
        if (String(rawDate) === formObj.date) isDateMatch = true;
      }

      if (!isDateMatch) continue; // 日期不符則跳過

      // 比對其他關鍵欄位 (CUST_N 已由 TextFinder 篩選，但保留比對邏輯也無妨)
      if (
        String(row[colMap["SRTimes"]]).trim() === formObj.SRTimes.toString().trim() &&
        // String(row[colMap["CUST_N"]]).trim() === formObj.custName.toString().trim() && // TextFinder 已精確匹配
        String(row[colMap["USER_N"]]).trim() === formObj.userName.toString().trim() &&
        String(row[colMap["Pay_Type"]]).trim() === formObj.payType.toString().trim() &&
        String(row[colMap["SR_ID"]]).trim() === formObj.srId.toString().trim()
      ) {
        if (actionType === "query") {
          let result = {
            found: true,
            success: true,
            data: {
              date: formObj.date, // 直接回傳查詢日期，節省格式化成本
              SRTimes: row[colMap["SRTimes"]],
              custName: row[colMap["CUST_N"]],
              userName: row[colMap["USER_N"]],
              payType: row[colMap["Pay_Type"]],
              srId: row[colMap["SR_ID"]],
              srRec: row[colMap["SR_REC"]],
              loc: row[colMap["LOC"]],
              mood: row[colMap["MOOD"]],
              spcons: row[colMap["SPCONS"]],
            },
          };

          // 寫入快取 (需重新產生 cacheKey 或確保變數作用域可及，此處簡單重組 key)
          try {
            let cacheKey = `SRDataQuery_${formObj.date}_${formObj.SRTimes}_${formObj.custName}_${formObj.userName}_${formObj.payType}_${formObj.srId}`;
            CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 300); // 快取 5 分鐘
          } catch (e) {
            console.error("快取 SRDataQuery 失敗: " + e.toString());
          }

          return result;
        } else if (actionType === "update") {
          // 使用欄位映射進行穩健更新
          targetFields.forEach((field, idx) => {
            let colIdx = colMap[field];
            if (colIdx !== -1) {
              targetSheet.getRange(i + 1, colIdx + 1).setValue(rowData[idx]);
            }
          });

          UpdateRawResponse(formObj);
          // 清除對應月份的案主列表快取
          if (yearmonth) {
            CacheService.getScriptCache().remove("CustN_" + yearmonth);
            // 同步清除報表資料來源快取
            CacheService.getScriptCache().remove("DataMap_" + yearmonth);
          }
          return { success: true, message: "資料已於 " + sheetName + " 更新完成。", };
        } else if (actionType === "delete") {

          targetSheet.deleteRow(rIdx);
          // 清除對應月份的案主列表快取
          // 清除對應月份的案主列表快取
          if (yearmonth) {
            CacheService.getScriptCache().remove("CustN_" + yearmonth);
          }
          return {
            success: true,
            message: "資料已從 " + sheetName + " 刪除。",
          };
        }
      }
    }

    return {
      success: false,
      found: false,
      message: actionType === "delete" ? "刪除失敗：找不到該筆資料，可能已被刪除或日期不符。" : "在此區間找不到符合條件的紀錄。"
    };
  } catch (e) {
    console.error("processSRData 發生錯誤: " + e.toString());
    return { success: false, message: "錯誤: " + e.toString() };
  }
}
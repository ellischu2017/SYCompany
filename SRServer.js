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
  var userSheet = MainSpreadsheet.getSheetByName("User");
  var userData = [];
  if (userSheet) {
    var rawUsers = userSheet.getDataRange().getValues();
    var headers = rawUsers[0];
    
    // 動態抓取欄位索引，確保名稱對應 (引用 Utilities.js 的 getColIndex 邏輯，或直接重寫以防萬一)
    var idxName = getColIndexSafe(headers, "User_N");
    var idxEmail = getColIndexSafe(headers, "User_Email");
    if (idxEmail === -1) idxEmail = getColIndexSafe(headers, "Email"); // 容錯
    var idxCust = getColIndexSafe(headers, "Cust_N");
    
    if (idxName !== -1) {
      for (var i = 1; i < rawUsers.length; i++) {
        var row = rawUsers[i];
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

  // 2. 取得 Cust 資料 (包含 Cust_N, Cust_LTC_Code)
  var custSheet = MainSpreadsheet.getSheetByName("Cust");
  var custData = [];
  if (custSheet) {
    var rawCusts = custSheet.getDataRange().getValues();
    var cHeaders = rawCusts[0];
    var idxCName = getColIndexSafe(cHeaders, "Cust_N");
    var idxCLTC = getColIndexSafe(cHeaders, "LTC_Code");
    
    if (idxCName !== -1) {
      for (var j = 1; j < rawCusts.length; j++) {
        var row = rawCusts[j];
        if (row[idxCName]) {
          custData.push({
            name: row[idxCName].toString(),
            ltcStr: idxCLTC !== -1 ? row[idxCLTC].toString() : ""
          });
        }
      }
    }
  }

  // 3. 取得 LTC_Code 資料 (SR_ID)
  var ltcSheet = MainSpreadsheet.getSheetByName("LTC_Code");
  var ltcIds = [];
  if (ltcSheet) {
    var rawLtc = ltcSheet.getDataRange().getValues();
    var lHeaders = rawLtc[0];
    var idxSRID = getColIndexSafe(lHeaders, "SR_ID");
    var idxCont = getColIndexSafe(lHeaders, "SR_Cont");
    var targetIdx = idxSRID !== -1 ? idxSRID : 0; // 若找不到標題，預設第一欄
    var seen = {};
    
    for (var k = 1; k < rawLtc.length; k++) {
       var code = rawLtc[k][targetIdx].toString().trim();
       var cont = idxCont !== -1 ? rawLtc[k][idxCont].toString().trim() : "";
       if(code && !seen[code]) {
         seen[code] = true;
         ltcIds.push({ id: code, cont: cont });
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
 * 輔助函數：安全的取得欄位索引
 */
function getColIndexSafe(headers, name) {
  if (!headers) return -1;
  var idx = headers.indexOf(name);
  if (idx !== -1) return idx;
  // 嘗試不區分大小寫
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase() === name.toLowerCase()) return i;
  }
  return -1;
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
    var companyData = companyUserSheet.getDataRange().getValues();
    var existingEmails = new Set();
    for (var i = 1; i < companyData.length; i++) {
      var email = companyData[i][1]; // User_Email
      if (email) {
        existingEmails.add(email.toString().trim().toLowerCase());
      }
    }

    // 4. 過濾出 SYTemp 中「不在」SYCompany 裡的資料
    var rowsToSync = [];
    for (var j = 1; j < tempValues.length; j++) {
      var tempEmail = tempValues[j][1]
        ? tempValues[j][1].toString().trim().toLowerCase()
        : "";

      if (tempEmail !== "" && !existingEmails.has(tempEmail)) {
        rowsToSync.push(tempValues[j]);
        existingEmails.add(tempEmail);
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
  try {
    // --- 1. 計算日期與判斷目標 ---
    var inputDate = new Date(formObj.date);
    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var diffDays =
      (today.getTime() - inputDate.getTime()) / (1000 * 60 * 60 * 24);

    var targetSS, sheetName;

    if (diffDays <= 7) {
      // 條件 1: 少於或等於 7 天
      targetSS = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
      sheetName = "SR_Data";
    } else {
      // 條件 2: 多於 7 天
      var yearStr = inputDate.getFullYear().toString();
      var monthStr = Utilities.formatDate(inputDate, "GMT+8", "yyyyMM"); 

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
      return {
        success: true,
        message: "資料已成功存入 " + targetSS.getName() + " > " + sheetName,
      };
    }

    // --- 3. 執行 查詢 (query) / 更新 (update) / 刪除 (delete) ---
    var data = targetSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var sheetDate =
        data[i][0] instanceof Date
          ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy-MM-dd")
          : data[i][0].toString();

      // 比對五個關鍵欄位：Date, Cust_N, User_N, Pay_Type, SR_ID
      // 注意: SR_ID 和 SRTimes 轉字串比對
      if (
        sheetDate === formObj.date &&
        data[i][1].toString().trim() === formObj.SRTimes.toString().trim() &&
        data[i][2].toString().trim() === formObj.custName.toString().trim() &&
        data[i][3].toString().trim() === formObj.userName.toString().trim() &&
        data[i][4].toString().trim() === formObj.payType.toString().trim() &&
        data[i][5].toString().trim() === formObj.srId.toString().trim()
      ) {
        if (actionType === "query") {
          return {
            found: true,
            success: true,
            data: {
              date: sheetDate,
              SRTimes: data[i][1],
              custName: data[i][2],
              userName: data[i][3],
              payType: data[i][4],
              srId: data[i][5],
              srRec: data[i][6],
              loc: data[i][7],
              mood: data[i][8],
              spcons: data[i][9],
            },
          };
        } else if (actionType === "update") {
          targetSheet.getRange(i + 1, 1, 1, 10).setValues([rowData]);
          return {
            success: true,
            message: "資料已於 " + sheetName + " 更新完成。",
          };
        } else if (actionType === "delete") {
          targetSheet.deleteRow(i + 1);
          return {
            success: true,
            message: "資料已從 " + sheetName + " 刪除。",
          };
        }
      }
    }

    return { found: false, message: "在此區間找不到符合條件的紀錄。" };
  } catch (e) {
    console.error("processSRData 發生錯誤: " + e.toString());
    return { success: false, message: "錯誤: " + e.toString() };
  }
}
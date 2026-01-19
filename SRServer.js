/**
 * SRServer.gs - 服務紀錄單管理模組
 * 提供服務紀錄的查詢、新增、修改、刪除操作
 */

function getSRServerInitData() {
  var userMapping = {};
  var localUserSheet = MainSpreadsheet.getSheetByName("User");
  if (localUserSheet) {
    var localData = localUserSheet.getDataRange().getValues();
    for (var i = 1; i < localData.length; i++) {
      if (localData[i][0]) userMapping[localData[i][0]] = localData[i][1];
    }
  }

  return {
    userMapping: userMapping,
    userList: Object.keys(userMapping),
    custNames: getCustList(), // 需存在於 Utilities.gs
    srIds: getLtcCodeList(), // 需存在於 Utilities.gs
  };
}

/**
 * 檢查 SYTemp 的 User 工作表並同步回 SYCompany
 */
/**
 * 同步使用者資料並檢查重復
 * 邏輯：以 Email 為基準，若 SYCompany 已存在該 Email，則跳過不匯入。
 */
function processUserSync01() {
  try {
    // 1. 取得來源表 SYTemp > User
    var ssTemp = getTargetsheet("SYTemp", "SYTemp");
    var tempUserSheet = ssTemp.getSheetByName("User");
    if (!tempUserSheet) {
      console.log("SYTemp 中不存在 User 工作表，無法同步。");
      return;
    }

    var tempValues = tempUserSheet.getDataRange().getValues();
    if (tempValues.length <= 1) {
      console.log("SYTemp 的 User 表無資料可同步。");
      return;
    } // 若除了標題外無資料，直接結束

    // 2. 取得目標表 SYCompany > User
    var companyUserSheet = MainSpreadsheet.getSheetByName("User");
    if (!companyUserSheet) {
      // 若工作表不存在則建立，並寫入標題列
      companyUserSheet = MainSpreadsheet.insertSheet("User");
      companyUserSheet.appendRow(tempValues[0]);
    }

    // 3. 抓取 SYCompany 目前已有的 Email 清單 (假設 Email 在第 2 欄，索引為 1)
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

      // 檢查是否已存在，且避免 SYTemp 內部有重復資料同時被加入
      if (tempEmail !== "" && !existingEmails.has(tempEmail)) {
        rowsToSync.push(tempValues[j]);
        existingEmails.add(tempEmail); // 暫時加入 Set 防止本次迴圈重復處理
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

    // 無論有無新資料，處理完後皆清空 SYTemp (保留標題列)
    // 這樣可以確保 SYTemp 始終只存放「待處理」的新申請
    if (tempValues.length > 1) {
      tempUserSheet.deleteRows(2, tempValues.length - 1);
    }
  } catch (e) {
    console.log("processUserSync01 執行出錯: " + e.toString());
  }
}

/**
 * 處理服務紀錄：根據日期判斷查詢位置
 * 1. 7天內：存取 SYTemp > SR_Data
 * 2. 7天以前：存取 SY + yyyy > SR_Data
 */
/**
 * 重新定義：根據日期動態分流存取路徑
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
      targetSS = getTargetsheet("SYTemp", "SYTemp");
      sheetName = "SR_Data";
    } else {
      // 條件 2: 多於 7 天
      var yearStr = inputDate.getFullYear().toString();
      var monthStr = Utilities.formatDate(inputDate, "GMT+8", "yyyyMM"); // 例如 202501

      targetSS = getTargetsheet("RecUrl", "SY" + yearStr);
      sheetName = monthStr;
    }

    if (!targetSS) throw new Error("找不到對應年份的試算表或 SYTemp。");

    var targetSheet = targetSS.getSheetByName(sheetName);

    // 如果是新增(add)且月份工作表不存在，則自動建立並加入標題
    if (!targetSheet && actionType === "add") {
      targetSheet = targetSS.insertSheet(sheetName);
      var headers = [
        "Date",
        "E-mail",
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
      formObj.email,
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
      if (
        sheetDate === formObj.date &&
        data[i][2].toString().trim() === formObj.custName.trim() &&
        data[i][3].toString().trim() === formObj.userName.trim() &&
        data[i][4].toString().trim() === formObj.payType.trim() &&
        data[i][5].toString().trim() === formObj.srId.trim()
      ) {
        if (actionType === "query") {
          return {
            found: true,
            success: true,
            data: {
              date: sheetDate,
              email: data[i][1],
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

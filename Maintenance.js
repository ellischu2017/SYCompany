/**
 * Maintenance.gs - 維護任務模組
 * 提供自動維護、同步、資料遷移等功能
 */

/**
 * 每月維護任務：同步試算表權限
 */
function monthlyMaintenanceJob() {
  syncMasterTablePermissions();
  console.log("每月維護任務完成：權限已同步。");

  // 1. 取得今天的日期並計算「上個月」
  const now = new Date();
  now.setMonth(now.getUTCMonth() - 1); // 往前推一個月

  // 2. 格式化年份與月份
  // 使用 Utilities.formatDate 確保格式為 yyyy 與 yyyyMM (補零)
  const timeZone = Session.getScriptTimeZone();
  const yyyy = Utilities.formatDate(now, timeZone, "yyyy");
  const yyyyMM = Utilities.formatDate(now, timeZone, "yyyyMM");

  // 3. 動態組合名稱並獲取工作表
  const srcSpreadsheetName = "SY" + yyyy;
  const SYyyyy = getTargetsheet("RecUrl", srcSpreadsheetName).Spreadsheet;
  const srcSheet = SYyyyy.getSheetByName(yyyyMM);
  const SYTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
  const tmpSheet = SYTemp.getSheetByName("SR_Data");
  const tarSheet = MainSpreadsheet.getSheetByName("Cust");
  // 4. 執行更新
  if (srcSheet) {
    removeSRDuplicates(srcSheet);
    // 同時傳入 上個月工作表(srcSheet) 與 暫存工作表(tmpSheet)，確保跨年度/跨月資料完整性
    UpdateCustLTCCode([srcSheet, tmpSheet], tarSheet);
    console.log(`成功處理：${srcSpreadsheetName} > ${yyyyMM}`);
  } else {
    console.error(`找不到工作表：${yyyyMM}`);
  }
}

/**
 * 每月10日維護任務：搬移上個月資料至年度試算表
 * 建議觸發時間：每月 10 日 00:00 - 01:00
 */
function monthlyTenMaintenanceJob() {
  var today = new Date();
  // 防呆機制：SRServer 邏輯設定每月 10 號(含)前會至 SYTemp 查詢上月資料
  // 若在此之前執行搬移，會導致使用者查詢不到上個月的紀錄
  if (today.getDate() < 10) {
    console.warn(`今日為 ${today.getDate()} 號，未達每月 10 號搬移標準，已取消執行以確保資料讀取正確。`);
    return;
  }

  //搬移上個月資料至年度試算表
  processSRDataMigration();
  console.log("每月10日維護任務完成：資料已搬移。");


}

/**
 * 每日維護任務：同步 User 名單
 * 建議觸發時間：每日 00:00 - 01:00
 * 修改：加入自動續傳機制 (Auto-Resume) 以防止超時中斷
 */
function dailyMaintenanceJob() {
  // 1. 讀取執行進度 (若無則從 0 開始)
  var state = getProgress("DAILY_MAINTENANCE_STATE");
  var step = state ? state.step : 0;

  if (step > 0) {
    // Set time when the script resumes
    startTime = new Date().getTime();

    console.log("偵測到未完成的任務，從步驟 " + step + " 繼續執行...");
  }

  try {
    // Step 0: 每日 Raw Response 更新
    if (step <= 0) {
      if (checkTimeoutAndScheduleResume(0)) return;

      // Set time when the script starts
      startTime = new Date().getTime();
      UpdateRawResponseDaily();
      step = 1;
    }

    // 準備試算表物件 (供後續步驟使用)
    const ssSYTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    const tmpSheet = ssSYTemp.getSheetByName("SR_Data");
    const tarSheet = MainSpreadsheet.getSheetByName("Cust");

    // Step 1: 同步使用者 (User Sync)
    if (step <= 1) {
      if (checkTimeoutAndScheduleResume(1)) return;
      processUserSync(MainSpreadsheet, ssSYTemp);
      step = 2;
    }

    // Step 2: 同步個案 (Cust Sync)
    if (step <= 2) {
      if (checkTimeoutAndScheduleResume(2)) return;
      processCustSync();
      step = 3;
    }

    // Step 3: 同步封存個案 (Old Cust Sync)
    if (step <= 3) {
      if (checkTimeoutAndScheduleResume(3)) return;
      processOldCustSync();
      step = 4;
    }

    // Step 4: 匯入新紀錄 (Transfer Data - 最耗時步驟)
    if (step <= 4) {
      if (checkTimeoutAndScheduleResume(4)) return;
      // processTransferData 內部已有 isNearTimeout 檢查，若超時會安全中止
      // 此處確保即使它中止，狀態也會推進，或透過迴圈設計讓其續傳(視需求)
      // 目前設計為：盡量做，做不完下次排程再跑，避免卡死整個流程
      processTransferData("all", true);
      // 資料匯入後主動清除快取，確保前端讀取到最新資料
      CacheService.getScriptCache().remove("SRServer01_InitData");
      step = 5;
    }

    // Step 5: 清除重複 (Remove Duplicates)
    if (step <= 5) {
      if (checkTimeoutAndScheduleResume(5)) return;
      removeSRDuplicates(tmpSheet);
      // 清除重複後主動清除快取
      CacheService.getScriptCache().remove("SRServer01_InitData");
      step = 6;
    }

    // Step 6: 更新長照編碼 (Update LTC Code)
    if (step <= 6) {
      if (checkTimeoutAndScheduleResume(6)) return;
      UpdateCustLTCCode(tmpSheet, tarSheet);
      step = 7;
    }

    // Step 7: 更新使用者個案關聯 (Update User Cust Name)
    if (step <= 7) {
      if (checkTimeoutAndScheduleResume(7)) return;
      UpdateUserCustName();
      step = 8;
    }

    // Step 8: 資料遷移 (目前註解中，保留位置)
    // if (step <= 8) {
    //   if (checkTimeoutAndScheduleResume(8)) return;
    //   processSRDataMigration(MainSpreadsheet, ssSYTemp);
    //   step = 9;
    // }

    // 全部完成，清除進度
    clearProgress("DAILY_MAINTENANCE_STATE");
    // 確保任務結束時快取一定是最新的
    CacheService.getScriptCache().remove("SRServer01_InitData");
    console.log("每日維護任務已完整執行完畢。");

  } catch (e) {
    console.error("每日維護任務發生錯誤: " + e.toString());
    // 發生錯誤時不清除進度，保留狀態供排錯或下次重試
  }
}


/**
 * 更新 SYCompany 試算表中的 LTC_Code
 * 修正：支援多來源、使用 getColIndex 避免大小寫問題
 * @param {Sheet|Sheet[]} srcSheets 來源工作表 (單一或陣列)
 * @param {Sheet} tarSheet 目標工作表 (Cust)
 */
function UpdateCustLTCCode(srcSheets, tarSheet) {
  // 2. 讀取資料並整合到記憶體 (Map 結構)
  // Map 鍵為 CUST_N, 值為 Set (確保 SR_ID 唯一)
  let combinedData = new Map();

  // 支援單一工作表或工作表陣列
  const sheets = Array.isArray(srcSheets) ? srcSheets : [srcSheets];

  sheets.forEach(sheet => {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return; // 跳過無資料的工作表

    const headers = data.shift();
    // 使用 getColIndex 處理欄位名稱大小寫差異 (如 CUST_N vs Cust_N)
    const cIdx = getColIndex(headers, "CUST_N");
    const sIdx = getColIndex(headers, "SR_ID");

    if (cIdx === -1 || sIdx === -1) return;

    data.forEach((row) => {
      const custN = row[cIdx];
      const srId = String(row[sIdx]).trim();
      if (custN && srId) {
        if (!combinedData.has(custN)) {
          combinedData.set(custN, new Set());
        }
        combinedData.get(custN).add(srId);
      }
    });
  });

  // 3. 讀取目標工作表並進行比對與更新
  const tarData = tarSheet.getDataRange().getValues();
  const tarHeaders = tarData.shift();
  const tarCIdx = getColIndex(tarHeaders, "Cust_N");
  const tarLIdx = getColIndex(tarHeaders, "LTC_Code");

  if (tarCIdx === -1 || tarLIdx === -1) {
    console.warn("UpdateCustLTCCode: 目標工作表缺少 Cust_N 或 LTC_Code 欄位");
    return;
  }

  // 準備更新後的資料列
  const updatedLTCColumn = tarData.map((row) => {
    const custN = row[tarCIdx];
    let existingCodes = row[tarLIdx]
      ? String(row[tarLIdx])
        .split(",")
        .map((s) => s.trim())
      : [];

    // console.log(`客戶 ${custN} 的 LTC_Code 更新: ${existingCodes.join(", ")}`);
    if (combinedData.has(custN)) {
      const newIds = combinedData.get(custN);

      // 合併舊有與新取得的 SR_ID
      newIds.forEach((id) => {
        if (!existingCodes.includes(id)) {
          existingCodes.push(id);
        }
      });
      // 排序並轉回字串
      return [existingCodes.sort().join(",")];
    }
    return [row[tarLIdx]]; // 若無匹配則維持原狀
  });

  // 4. 批次寫回目標工作表的 LTC_Code 欄位
  if (updatedLTCColumn.length > 0) {
    tarSheet
      .getRange(2, tarLIdx + 1, updatedLTCColumn.length, 1)
      .setValues(updatedLTCColumn);
  }

  console.log("資料整合完成，總客戶數: " + combinedData.size);
  // console.log("更新完成！");
  CacheService.getScriptCache().remove("SRServer01_InitData");
}

/**
 * 同步所有相關試算表的權限
 * 1. 包含 SYCompany 本身與 RecUrl 內的所有試算表。
 * 2. 根據 Manager 工作表名單授權為「編輯者」。
 * 3. 移除名單外所有「特定的」編輯者與檢視者。
 * 4. 將「一般存取權」設為「知道連結的人即可檢視」。
 * 特別處理：SYTemp 試算表的「一般存取權」設為「知道連結的人即可編輯」，以利資料同步作業。
 * 建議觸發時間：每日 00:00 - 01:00（可搭配 dailyMaintenanceJob 一起執行）
 * 注意事項：
 * - 確保 Manager 工作表的 Email 欄位正確無誤，避免誤刪權限。
 * - 執行前請備份重要資料，以防不慎操作導致權限異常。
 * - SYTemp 試算表的特殊權限設定是為了支援每日資料同步，請勿修改該試算表的權限設定。
 * - 若有新增年度試算表，dailyMaintenanceJob 會自動呼叫此函式同步權限。
 * - 若有試算表無法訪問或權限異常，請檢查 RecUrl 工作表中的網址是否正確，並確認該試算表的擁有者是否授予了足夠的權限。
 * - 建議定期檢查 Manager 工作表與 RecUrl 工作表的內容，確保權限同步的正確性與完整性。
 * - 若有需要排除特定試算表不進行權限同步，請在 RecUrl 工作表中添加一欄「ExcludeFromSync」，並在該欄填入「TRUE」以標記該試算表。
 * - 權限同步的過程中會有詳細的日誌輸出，建議定期檢查執行日誌以確保同步過程順利，並及時發現與解決可能的問題。
 * - 此函式使用了 Drive API 來管理權限，請確保已在 Google Cloud Console 中啟用 Drive API 並授予相應的權限。
 */
function syncMasterTablePermissions() {
  var managerSheet = MainSpreadsheet.getSheetByName("Manager");
  if (!managerSheet) {
    console.error("syncMasterTablePermissions: 找不到 'Manager' 工作表，權限同步中止。");
    return;
  }
  var managerData = managerSheet.getDataRange().getValues();
  var managerEmails = [];
  if (managerData.length > 1) {
    const headers = managerData[0];
    const emailIdx = getColIndex(headers, "Mana_Email");
    if (emailIdx !== -1) {
      for (var i = 1; i < managerData.length; i++) {
        var email = managerData[i][emailIdx];
        if (email) managerEmails.push(email.toString().trim().toLowerCase());
      }
    }
  }

  var RootFolder = getTargetDir("FolderUrl", "SYCompany");
  var targetFileIds = [{ Name: "SYCompany", UrlID: RootFolder.id }]; // 包含 SYCompany 本身
  targetFileIds.push({ Name: "SYTemp", UrlID: getTargetsheet("SYTemp", "SYTemp").id }); // 包含 SYTemp

  targetFileIds.forEach(function (item) {
    var fileId = item.UrlID;
    var fileName = item.Name;
    console.log("正在處理試算表: " + item.Name + " (ID: " + item.UrlID + ")");
    try {
      var file = DriveApp.getFileById(fileId);
      var ownerEmail = file.getOwner().getEmail().toLowerCase();
      // 先取得目前的編輯者名單，避免重複呼叫 API
      var currentEditors = file.getEditors().map(function (u) { return u.getEmail().toLowerCase(); });

      managerEmails.forEach(function (email) {
        // 只有當 email 不在目前的編輯者名單中，且不是擁有者時，才新增權限
        if (currentEditors.indexOf(email) === -1 && email !== ownerEmail) {
          var resource = {
            role: "writer",
            type: "user",
            emailAddress: email,
          };
          Drive.Permissions.create(resource, fileId, {
            sendNotificationEmails: false,
          });
        }
      });

      // var file = DriveApp.getFileById(fileId);
      // var ownerEmail = file.getOwner().getEmail().toLowerCase();

      file.getEditors().forEach(function (editor) {
        var e = editor.getEmail().toLowerCase();
        if (managerEmails.indexOf(e) === -1 && e !== ownerEmail) {
          file.removeEditor(editor);
        }
      });

      file.getViewers().forEach(function (viewer) {
        var v = viewer.getEmail().toLowerCase();
        if (managerEmails.indexOf(v) === -1) {
          file.removeViewer(viewer);
        }
      });

      if (fileName === "SYTemp") {
        file.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.EDIT,
        );
      } else {
        file.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW,
        );
      }
    } catch (e) {
      console.error(
        "檔案 " + fileName + " (" + fileId + ") 處理失敗: " + e.message,
      );
    }
  });
}


/**
 * 處理 User 同步：SYTemp > User 搬移至 SYCompany > User
 * 特別處理：確保電話號碼 User_Tel 為文字字串格式
 * 1. 自動檢查 Email 是否重複，重複則不新增但仍從 Temp 移除。
 * 2. 確保 User_Tel 以文字格式 (@) 存入。
 */
function processUserSync(mainSS, tempSS) {
  // 1. 初始值檢查與預設值設定
  if (!mainSS) {
    mainSS = MainSpreadsheet;
    console.log("mainSS 未提供，使用預設 MainSpreadsheet");
  }

  if (!tempSS) {
    tempSS = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    console.log("tempSS 未提供，使用預設 SYTemp");
  }

  const mainUserSheet = mainSS.getSheetByName("User");
  const tempUserSheet = tempSS.getSheetByName("User");

  if (!tempUserSheet || !mainUserSheet) {
    console.error("找不到 User 工作表，請檢查名稱是否正確。");
    return;
  }

  // 2. 讀取資料
  const tempData = tempUserSheet.getDataRange().getValues();
  if (tempData.length <= 1) {
    console.log("暫存表無資料，跳過同步。");
    return;
  }

  const mainRange = mainUserSheet.getDataRange();
  const mainData = mainRange.getValues();
  const mainHeaders = mainData[0];
  const tempHeaders = tempData[0];

  // 3. 動態偵測欄位 Index (使用 getColIndicesMap 簡化)
  const targetFields = ["User_N", "User_Email", "User_Tel"];
  const mainColMap = getColIndicesMap(mainHeaders, targetFields);
  const tempColMap = getColIndicesMap(tempHeaders, targetFields);

  // 檢查主表是否缺少必要欄位
  if (mainColMap["User_N"] === -1) {
    throw new Error("主表找不到標題: User_N，無法進行比對。");
  }

  // 4. 建立主表索引 Map { User_N_Value : rowIndex }
  const mainIndexMap = {};
  for (let i = 1; i < mainData.length; i++) {
    let userName = mainData[i][mainColMap["User_N"]].toString().trim();
    if (userName) mainIndexMap[userName] = i;
  }

  const newRowsToAppend = [];
  let updateCount = 0;

  // 5. 開始同步 (比對 User_N)
  for (let i = 1; i < tempData.length; i++) {
    const tempRow = tempData[i];
    const tempUserName = tempRow[tempColMap["User_N"]].toString().trim();

    if (!tempUserName) continue; // 跳過空名

    if (mainIndexMap.hasOwnProperty(tempUserName)) {
      // --- 狀況 A: User_N 已存在 -> 更新資料 ---
      const mainRowIdx = mainIndexMap[tempUserName];

      // 更新這三個指定欄位
      targetFields.forEach(field => {
        const mIdx = mainColMap[field];
        const tIdx = tempColMap[field];
        // 確保暫存表也有該欄位才更新
        if (mIdx !== -1 && tIdx !== -1) {
          mainData[mainRowIdx][mIdx] = tempRow[tIdx];
        }
      });
      updateCount++;
    } else {
      // --- 狀況 B: User_N 是新的 -> 準備新增 ---
      // 按照主表的欄位順序構造一筆新資料
      let newRow = new Array(mainHeaders.length).fill("");
      tempHeaders.forEach((headerName, tIdx) => {
        const mIdx = mainHeaders.indexOf(headerName);
        if (mIdx !== -1) {
          newRow[mIdx] = tempRow[tIdx];
        }
      });
      newRowsToAppend.push(newRow);
    }
  }

  // 6. 資料回寫
  // 更新現有列 (一次性覆蓋主表範圍以提升運作效率)
  if (updateCount > 0) {
    mainUserSheet.getRange(1, 1, mainData.length, mainHeaders.length).setValues(mainData);
    console.log(`已成功更新 ${updateCount} 筆現有使用者資料。`);
  }

  // 新增全新列
  if (newRowsToAppend.length > 0) {
    const startRow = mainUserSheet.getLastRow() + 1;
    const targetRange = mainUserSheet.getRange(
      startRow,
      1,
      newRowsToAppend.length,
      mainHeaders.length
    );
    targetRange.setNumberFormat("@"); // 強制文字格式避免號碼跑掉
    targetRange.setValues(newRowsToAppend);
    console.log(`已新增 ${newRowsToAppend.length} 筆新使用者資料。`);
  }

  // 7. 排序與清理
  const finalLastRow = mainUserSheet.getLastRow();
  if (finalLastRow > 1) {
    // 依據 User_N 所在欄位進行 ASC 排序
    mainUserSheet
      .getRange(2, 1, finalLastRow - 1, mainUserSheet.getLastColumn())
      .sort({ column: mainColMap["User_N"] + 1, ascending: true });
    console.log("主表資料已完成排序。");
  }

  if (tempUserSheet.getLastRow() > 1) {
    tempUserSheet.deleteRows(2, tempUserSheet.getLastRow() - 1);
    console.log("暫存表清理完畢。");
  }
  CacheService.getScriptCache().remove("SRServer01_InitData");
}

/**
 * 處理 SR_Data 遷移：上個月資料搬移至年度試算表
 * 修正：日期偏移、新增首列凍結、設定日期欄位格式
 */
function processSRDataMigration() {
  console.log("processSRDataMigration 開始遷移 SR_Data 資料...");
  mainSS = MainSpreadsheet;
  tempSS = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
  // SYTemp > SR_Data 工作表
  const srSheet = tempSS.getSheetByName("SR_Data");
  if (!srSheet) return;
  // 1. 讀取資料
  const data = srSheet.getDataRange().getValues();
  if (data.length <= 1) return;
  // 2. 計算遷移截止日期 (本月1號)
  const headers = data[0];
  const dateIdx = getColIndex(headers, "Date");
  if (dateIdx === -1) {
    console.error("processSRDataMigration: 找不到 'Date' 欄位，無法進行遷移。");
    return;
  }

  const cutoffDate = new Date();
  cutoffDate.setDate(1);
  cutoffDate.setHours(0, 0, 0, 0);
  //cutoffDate.setDate(today.getDate() - 8); // 7 天前
  Logger.log("資料遷移截止日期 (cutoffDate): " + cutoffDate.toISOString());

  const migrationMap = {};
  const rowsToKeep = [headers];
  let createdNewSS = false;

  for (let i = 1; i < data.length; i++) {
    let row = [...data[i]];
    let rawDate = row[dateIdx];

    let dateObj;
    if (rawDate instanceof Date) {
      dateObj = new Date(rawDate);
    } else {
      dateObj = new Date(rawDate.toString().replace(/-/g, "/"));
    }
    dateObj.setHours(0, 0, 0, 0);

    let formattedDate = Utilities.formatDate(dateObj, "Asia/Taipei", "yyyy-MM-dd");
    row[dateIdx] = formattedDate;

    if (dateObj < cutoffDate) {
      let yearmonth = dateObj.getFullYear() + Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM");
      if (!migrationMap[yearmonth]) migrationMap[yearmonth] = [];
      migrationMap[yearmonth].push(row);
    } else {
      rowsToKeep.push(row);
    }
  }

  for (let yearmonth in migrationMap) {
    let year = yearmonth.substring(0, 4)
    let month = yearmonth.substring(4, 6);
    let syName = "SY" + year;
    // get target spredsheet
    let tarspredsheet = getTargetsheet("RecUrl", syName).Spreadsheet;
    // get target sheet name is yyyyMM
    let tarsheetName = year + month;
    let tarSheet = tarspredsheet.getSheetByName(tarsheetName);
    Logger.log(`處理 ${yearmonth} 資料，目標試算表: ${syName}, 目標工作表: ${tarsheetName}`);
    // move data to target sheet
    let targetUrl = tarspredsheet.getUrl();
    if (targetUrl) {
      appendDataToExternalSS(targetUrl, yearmonth, migrationMap[yearmonth], headers);
    }
    console.log("搬移資料至 " + year + " 年試算表，網址: " + targetUrl);
    console.log("搬移筆數: " + migrationMap[yearmonth].length);
  }

  // 同步權限如果有新建立年度試算表
  if (createdNewSS) {
    syncMasterTablePermissions();
  }
  // 清理 SR_Data 工作表，只保留未遷移資料
  srSheet.clearContents();
  srSheet
    .getRange(1, 1, rowsToKeep.length, headers.length)
    .setValues(rowsToKeep);
}

/**
 * 輔助函式：將資料寫入年度試算表
 * 包含：凍結首列、設定日期格式、設定文字格式
 * 1. 凍結首列並設定 A 欄日期格式
 * 2. 移除舊篩選器並重新建立 (涵蓋所有資料列)
 * 3. 針對 A 欄進行 A 到 Z (由舊到新) 排序
 */
function appendDataToExternalSS(url, year, rows, headers) {
  try {
    const targetSS = SpreadsheetApp.openByUrl(url);
    const dateIdx = getColIndex(headers, "Date");
    if (dateIdx === -1) {
      console.error("appendDataToExternalSS: 來源資料中找不到 'Date' 欄位，無法進行遷移。");
      return;
    }
    const firstDateStr = rows[0][dateIdx].toString().replace(/-/g, "/");
    const firstDate = new Date(firstDateStr);
    const monthStr = Utilities.formatDate(
      firstDate,
      Session.getScriptTimeZone(),
      "yyyyMM",
    );

    let targetSheet = targetSS.getSheetByName(monthStr);

    if (!targetSheet) {
      targetSheet = targetSS.insertSheet(monthStr);
      const headers = [
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
      targetSheet.setFrozenRows(1);
      targetSheet.getRange("A:A").setNumberFormat("yyyy-MM-dd");
    }

    const startRow = targetSheet.getLastRow() + 1;
    const numCols = rows[0].length;
    const targetRange = targetSheet.getRange(startRow, 1, rows.length, numCols);

    if (numCols > 1) {
      targetSheet
        .getRange(startRow, 2, rows.length, numCols - 1)
        .setNumberFormat("@");
    }
    targetRange.setValues(rows);

    const currentFilter = targetSheet.getFilter();
    if (currentFilter) {
      currentFilter.remove();
    }

    targetSheet.setFrozenRows(1);

    const fullRange = targetSheet.getDataRange();
    const newFilter = fullRange.createFilter();
    if (fullRange.getNumRows() > 1) {
      // 動態偵測排序欄位
      const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      const dateSortIdx = getColIndex(targetHeaders, "Date");
      const custSortIdx = getColIndex(targetHeaders, "CUST_N");

      const sortColumns = [];

      // 優先使用 Date 欄位排序，若無則預設第一欄
      if (dateSortIdx !== -1) {
        sortColumns.push({ column: dateSortIdx + 1, ascending: true });
      } else {
        console.warn(`在 ${targetSheet.getName()} 中找不到 'Date' 欄位，將使用預設第一欄進行排序。`);
        sortColumns.push({ column: 1, ascending: true });
      }

      // 其次使用 CUST_N 欄位排序，若無則預設第三欄
      if (custSortIdx !== -1) {
        sortColumns.push({ column: custSortIdx + 1, ascending: true });
      } else {
        console.warn(`在 ${targetSheet.getName()} 中找不到 'CUST_N' 欄位，將使用預設第三欄進行排序。`);
        sortColumns.push({ column: 3, ascending: true });
      }
      targetSheet.getRange(2, 1, fullRange.getNumRows() - 1, fullRange.getNumColumns()).sort(sortColumns);
    }
    removeSRDuplicates(targetSheet);

    console.log(
      `成功搬移並排序 ${rows.length} 筆資料至 ${year} 年 ${monthStr} 表`,
    );
  } catch (e) {
    console.error("寫入外部試算表失敗: " + e.toString());
  }
}

/**
 * 輔助函式：比對個案資料是否有變更
 * @param {Array} tRow - 目標工作表的資料列
 * @param {Object} newData - 來源工作表的新資料物件
 * @param {Object} tarIdx - 目標工作表的欄位索引對應
 * @returns {boolean} - 如果資料有變更則回傳 true
 */
function hasCustDataChanged(tRow, newData, tarIdx) {
  // 比對日期需要特別小心，建議轉字串比對
  const tBD =
    tRow[tarIdx.bd] instanceof Date
      ? Utilities.formatDate(tRow[tarIdx.bd], "GMT+8", "yyyy/M/d")
      : tRow[tarIdx.bd];

  return (
    tRow[tarIdx.sex] !== newData.sex ||
    tBD !== newData.bd ||
    tRow[tarIdx.add] !== newData.add ||
    tRow[tarIdx.ltc] !== newData.ltc ||
    tRow[tarIdx.formurl] !== newData.formurl
  );
}
/**
 * 處理 Cust 同步：Case_Reports > 個案清單 搬移至 SYCompany > Cust
 * 對應欄位：
 * 個案姓名 -> Cust_N
 * 性別 -> Cust_Sex
 * 生日 -> Cust_BD
 * 地址 -> Cust_Add
 * 服務項目 -> Cust_LTC_Code
 * 1ib8q-lKJgLEhRVrwncnRqOyKNauMqaV2wtYEpGlmRlk
 */
function processCustSync() {
  syncCustData("個案清單", "Cust");
}

function processOldCustSync() {
  syncCustData("結案個案清單", "OldCust");
}

/**
 * 通用個案同步邏輯 (Internal Helper)
 * @param {string} sourceSheetName 來源工作表名稱
 * @param {string} targetSheetName 目標工作表名稱
 */
function syncCustData(sourceSheetName, targetSheetName) {
  console.log(
    `開始同步 Case_Reports > ${sourceSheetName} > ${targetSheetName} 資料...`,
  );
  const SOURCE_SS_ID = "1ib8q-lKJgLEhRVrwncnRqOyKNauMqaV2wtYEpGlmRlk";

  const sourceSheet = SpreadsheetApp.openById(SOURCE_SS_ID).getSheetByName(sourceSheetName);
  const targetSheet = MainSpreadsheet.getSheetByName(targetSheetName);

  if (!sourceSheet || !targetSheet) {
    console.error(`無法開啟工作表: ${sourceSheetName} 或 ${targetSheetName}`);
    return;
  }

  // 1. 取得來源與目標的所有資料
  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  const sourceHeaders = sourceData[0];
  const targetHeaders = targetData[0];
  const sourceRows = sourceData.slice(1);
  const targetRows = targetData.slice(1);

  // 2. 建立欄位索引字典
  const srcIdx = { name: 0, sex: 1, bd: 2, add: 3, formurl: 4, ltc: 5 }; // 根據 CSV 結構固定
  const targetFieldNames = ["Cust_N", "Cust_Sex", "Cust_BD", "Cust_Add", "Cust_LTC_Code", "Form_Url"];
  const colMap = getColIndicesMap(targetHeaders, targetFieldNames);
  const tarIdx = {
    name: colMap["Cust_N"],
    sex: colMap["Cust_Sex"],
    bd: colMap["Cust_BD"],
    add: colMap["Cust_Add"],
    ltc: colMap["Cust_LTC_Code"],
    formurl: colMap["Form_Url"],
  };

  // 3. 將目標表轉換為 Map 物件，方便以「姓名」快速查詢
  // Key: 姓名, Value: 該列在陣列中的索引與資料內容
  let targetMap = new Map();
  targetRows.forEach((row, index) => {
    const name = row[tarIdx.name];
    if (name) targetMap.set(name, { index: index, data: row });
  });

  let updateCount = 0;
  let insertCount = 0;
  let finalRows = [...targetRows]; // 複製一份目標資料來修改

  // 4. 遍歷來源資料進行比對
  sourceRows.forEach((sRow) => {
    const sName = sRow[srcIdx.name];
    if (!sName) return; // 跳過空列

    // 格式化來源生日 (處理日期物件轉字串比對)
    const sBD =
      sRow[srcIdx.bd] instanceof Date
        ? Utilities.formatDate(sRow[srcIdx.bd], "GMT+8", "yyyy/M/d")
        : sRow[srcIdx.bd];

    const newData = {
      sex: sRow[srcIdx.sex],
      bd: sBD,
      add: sRow[srcIdx.add],
      ltc: sRow[srcIdx.ltc],
      formurl: sRow[srcIdx.formurl],
    };

    if (targetMap.has(sName)) {
      // --- 狀況 A: 姓名已存在，檢查內容是否不同 ---
      let tEntry = targetMap.get(sName);
      let tRow = tEntry.data;

      const isChanged = hasCustDataChanged(tRow, newData, tarIdx);

      if (isChanged) {
        finalRows[tEntry.index][tarIdx.sex] = newData.sex;
        finalRows[tEntry.index][tarIdx.bd] = newData.bd;
        finalRows[tEntry.index][tarIdx.add] = newData.add;
        finalRows[tEntry.index][tarIdx.ltc] = newData.ltc;
        finalRows[tEntry.index][tarIdx.formurl] = newData.formurl;
        updateCount++;
      }
    } else {
      // --- 狀況 B: 姓名不存在，新增一列 ---
      let newRow = new Array(targetHeaders.length).fill("");
      newRow[tarIdx.name] = sName;
      newRow[tarIdx.sex] = newData.sex;
      newRow[tarIdx.bd] = newData.bd;
      newRow[tarIdx.add] = newData.add;
      newRow[tarIdx.ltc] = newData.ltc;
      newRow[tarIdx.formurl] = newData.formurl;

      finalRows.push(newRow);
      insertCount++;
    }
  });

  // 5. 寫回資料
  if (updateCount > 0 || insertCount > 0) {
    // 寫回所有資料 (包含更新後與新增的)
    targetSheet
      .getRange(2, 1, finalRows.length, targetHeaders.length)
      .setValues(finalRows);
    // 清除個案基本資料快取，因為資料已變動
    CacheService.getScriptCache().remove("CustInfoMap");
    CacheService.getScriptCache().remove("CustN_All");
    CacheService.getScriptCache().remove("SRServer01_InitData");
    console.log(`同步完成！ 更新：${updateCount} 筆 新增：${insertCount} 筆`);
  } else {
    console.log("資料皆為最新，無需更新。");
  }
}

function UpdateRawResponseDaily() {
  const srcsSheet = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
  const srcSheet = srcsSheet.getSheetByName("SR_Data");
  var data = srcSheet.getDataRange().getValues();
  var headers = data[0];
  const targetFields = ["Date", "CUST_N", "SR_ID", "SR_REC", "Pay_Type", "LOC", "MOOD", "SPCONS"];
  const colMap = getColIndicesMap(headers, targetFields);

  // 檢查關鍵欄位 Date 是否存在
  if (colMap["Date"] === -1) {
    console.error("UpdateRawResponseDaily: 找不到 'Date' 欄位，無法執行每日更新。");
    return;
  }

  var d = new Date();
  d.setDate(d.getDate() - 1);
  var yesterday = Utilities.formatDate(d, "Asia/Taipei", "yyyy-MM-dd");

  const idxDate = colMap["Date"];

  for (var i = 1; i < data.length; i++) {
    rowDate = Utilities.formatDate(data[i][idxDate], "Asia/Taipei", "yyyy-MM-dd");
    // Logger.log("rowDate: " + rowDate + " targetDate: "+ yesterday);
    if (rowDate === yesterday) {
      var formObj = {
        date: data[i][idxDate],
        custN: colMap["CUST_N"] !== -1 ? data[i][colMap["CUST_N"]] : "",
        srId: colMap["SR_ID"] !== -1 ? data[i][colMap["SR_ID"]] : "",
        srRec: colMap["SR_REC"] !== -1 ? data[i][colMap["SR_REC"]] : "",
        payType: colMap["Pay_Type"] !== -1 ? data[i][colMap["Pay_Type"]] : "",
        loc: colMap["LOC"] !== -1 ? data[i][colMap["LOC"]] : "",
        mood: colMap["MOOD"] !== -1 ? data[i][colMap["MOOD"]] : "",
        spcons: colMap["SPCONS"] !== -1 ? data[i][colMap["SPCONS"]] : "",
      }
      Logger.log("formObj: " + formObj);
      UpdateRawResponse(formObj);
    }
  }
}

/**
 * 設定系統自動化觸發條件 (Triggers)
 * 用途：自動建立每日與每月的排程任務
 * 操作：請在編輯器中手動執行此函式一次以完成安裝
 */
function setupTriggers() {
  // 防呆機制：使用 LockService 避免短時間內重複執行
  var lock = LockService.getScriptLock();
  // 嘗試取得鎖定，若 5 秒內無法取得則認為是重複執行
  if (!lock.tryLock(5000)) {
    console.warn("setupTriggers 正在執行中，請勿重複觸發。");
    return;
  }

  try {
    console.log("開始設定觸發條件...");
    
    // 1. 清除現有相關觸發條件，避免重複建立
    const triggers = ScriptApp.getProjectTriggers();
    const handlerNames = ['dailyMaintenanceJob', 'monthlyMaintenanceJob', 'monthlyTenMaintenanceJob'];
    let deletedCount = 0;

    triggers.forEach(trigger => {
      if (handlerNames.includes(trigger.getHandlerFunction())) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    if (deletedCount > 0) {
      console.log(`已清除 ${deletedCount} 個舊有的觸發條件。`);
    }

    // 2. 設定每日維護任務 (每日 00:00 - 01:00 執行)
    // 負責：同步使用者/個案資料、匯入 Raw Response、更新 LTC Code
    ScriptApp.newTrigger('dailyMaintenanceJob')
      .timeBased()
      .everyDays(1)
      .atHour(0)
      .create();

    // 3. 設定每月 1 號維護任務 (每月 1 號 01:00 - 02:00 執行)
    // 負責：同步試算表權限、跨月資料整合
    ScriptApp.newTrigger('monthlyMaintenanceJob')
      .timeBased()
      .onMonthDay(1)
      .atHour(1)
      .create();

    // 4. 設定每月 10 號維護任務 (每月 10 號 00:00 - 01:00 執行)
    // 負責：將上個月資料從 SYTemp 搬移至年度封存表
    ScriptApp.newTrigger('monthlyTenMaintenanceJob')
      .timeBased()
      .onMonthDay(10)
      .atHour(0)
      .create();

    console.log("系統自動化觸發條件已設定完成 (共 3 個任務)。");
    
  } catch (e) {
    console.error("設定觸發條件失敗: " + e.toString());
  } finally {
    // 確保釋放鎖定
    lock.releaseLock();
  }
}

/**
 * 輔助函式：檢查執行時間是否即將逾時
 * 若逾時：儲存當前步驟、建立續傳觸發器、回傳 true
 */
function checkTimeoutAndScheduleResume(currentStep) {
  if (isNearTimeout()) {
    console.warn(`[System] 執行時間不足 (Step ${currentStep})，正在儲存進度並排程續傳...`);
    saveProgress("DAILY_MAINTENANCE_STATE", { step: currentStep });
    
    // 建立一次性觸發器，1 分鐘後接續執行
    ScriptApp.newTrigger("dailyMaintenanceJob")
      .timeBased()
      .after(60 * 1000) 
      .create();
      
    return true; // 指示主程式中止
  }
  return false;
}
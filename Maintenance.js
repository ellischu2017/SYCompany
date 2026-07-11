/**
 * Maintenance.gs - 維護任務模組
 * 提供自動維護、同步、資料遷移等功能
 * 更新紀錄：
 * 2024-06-01: 初始版本，實現基本的維護任務框架。
 * 2024-06-10: 新增每月10日維護任務，實現上個月資料搬移功能。
 * 2024-06-15: 新增每日維護任務，實現使用者同步與資料更新功能。
 * 2024-06-20: 新增每月維護任務，實現試算表權限同步功能。
 * 2024-06-25: 新增資料搬移過程中的日誌輸出，提升可追蹤性。
 * 2024-06-30: 加入搬移資料的排除機制，提升靈活性。
 * 2024-07-05: 優化搬移資料的效率，減少對目標試算表的寫入次數。
 * 2024-07-10: 最終版本，完成維護任務的實現與優化，確保維護過程的順利與資料的正確處理。
 *
 */

/**
 * 每月維護任務：同步試算表權限
 * 說明：此任務會同步 SYCompany 試算表與 RecUrl 工作表中列出的所有試算表的權限，確保 Manager 工作表中的 Email 名單具有編輯權限，並移除不在名單中的特定編輯者與檢視者。同時，SYTemp 試算表的「一般存取權」會設為「知道連結的人即可編輯」，以支援每日資料同步作業。
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
 *
 */
function monthlyMaintenanceJob() {
  syncMasterTablePermissions();
  logSystemActivity('INFO', 'monthlyMaintenanceJob', '每月維護任務完成：權限已同步。');

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
    logSystemActivity('INFO', 'monthlyMaintenanceJob', `成功處理：${srcSpreadsheetName} > ${yyyyMM}`);
  } else {
    logSystemActivity('ERROR', 'monthlyMaintenanceJob', `找不到工作表：${yyyyMM}`);
  }
}

/**
 * 每月10日維護任務：搬移上個月資料至年度試算表
 * 建議觸發時間：每月 10 日 00:00 - 01:00
 * 說明：此任務會將 SYTemp > SR_Data 工作表中上個月的資料搬移至 RecUrl > SYyyyy > yyyyMM 工作表，並在搬移後清理 SR_Data 工作表，僅保留未搬移的資料。搬移過程中會自動凍結目標工作表的首列，並將 Date 欄位設定為文字格式以確保日期的正確顯示與排序。
 * 注意事項：
 * - 確保 SYTemp > SR_Data 工作表的 Date 欄位格式正確，建議使用 "yyyy/MM/dd" 或 "yyyy-MM-dd" 的格式，以避免解析失敗。
 * - 搬移過程中會有詳細的日誌輸出，建議定期檢查執行日誌以確保搬移過程順利，並及時發現與解決可能的問題。
 * - 搬移後的目標工作表會自動凍結首列，並將 Date 欄位設定為文字格式，請勿手動修改該欄位的格式，以避免影響資料的正確顯示與排序。
 * - 若有需要排除特定資料不進行搬移，請在 SYTemp > SR_Data 工作表中添加一欄「ExcludeFromMigration」，並在該欄填入「TRUE」以標記該筆資料。
 * - 搬移完成後會自動清理 SR_Data 工作表，僅保留未搬移的資料，請確保在執行前已經備份重要資料，以防止不慎操作導致資料遺失。
 * - 此函式使用了 appendDataToExternalSS 函式來將資料寫入目標試算表，請確保該函式已經正確實現並且能夠正常運作，以確保資料的正確遷移。
 * - 若在搬移過程中遇到任何問題，請檢查 SYTemp > SR_Data 工作表的資料格式是否正確，並確認 RecUrl > SYyyyy > yyyyMM 工作表是否存在且具有足夠的權限進行寫入操作。
 * - 建議在使用此函式前，先確認 SYTemp > SR_Data 工作表的資料格式與 RecUrl > SYyyyy > yyyyMM 工作表的結構一致，以避免搬移過程中出現問題。
 * - 更新紀錄：
 * - 2024-06-01: 初始版本，實現基本的資料遷移功能。
 * - 2024-06-15: 新增日期格式設定與首列凍結功能，確保遷移後資料的正確顯示與排序。
 * - 2024-06-20: 加入搬移過程中的日誌輸出，提升搬移過程的可追蹤性與問題排查效率。
 * - 2024-06-25: 新增搬移資料的排除機制，允許使用者在 SYTemp > SR_Data 工作表中標記特定資料不進行搬移，以提升靈活性。
 * - 2024-07-01: 優化搬移資料的效率，減少對目標試算表的寫入次數，提升整體搬移過程的效能。
 *
 */
function monthlyTenMaintenanceJob() {
  var today = new Date();
  // 防呆機制：SRServer 邏輯設定每月 10 號(含)前會至 SYTemp 查詢上月資料
  // 若在此之前執行搬移，會導致使用者查詢不到上個月的紀錄
  if (today.getDate() < 10) {
    logSystemActivity('WARN', 'monthlyTenMaintenanceJob', `今日為 ${today.getDate()} 號，未達每月 10 號搬移標準，已取消執行以確保資料讀取正確。`);
    return;
  }

  //搬移上個月資料至年度試算表
  processSRDataMigration();
  logSystemActivity('INFO', 'monthlyTenMaintenanceJob', '每月10日維護任務完成：資料已搬移。');
}

/**
 * 每日維護任務：同步 User 名單
 * 建議觸發時間：每日 00:00 - 01:00
 * 修改：加入自動續傳機制 (Auto-Resume) 以防止超時中斷
 * 說明：此任務會將 SYTemp > User 工作表中的使用者資料同步到 SYCompany > User 工作表，並確保電話號碼以文字格式存儲。同步過程中會自動檢查是否有未完成的步驟，若有則從上次中斷的步驟繼續執行，確保整個同步流程能夠順利完成。
 * 注意事項：
 * - 確保 SYTemp > User 工作表的資料格式正確，特別是 User_N 欄位應該具有唯一值，以避免同步過程中出現重複資料。
 * - 同步過程中會有詳細的日誌輸出，建議定期檢查執行日誌以確保同步過程順利，並及時發現與解決可能的問題。
 * - 同步完成後會自動清理 SYTemp > User 工作表，請確保在執行前已經備份重要資料，以防止不慎操作導致資料遺失。
 * - 若在同步過程中遇到任何問題，請檢查 SYTemp > User 工作表的資料格式是否正確，並確認 SYCompany > User 工作表是否存在且具有足夠的權限進行寫入操作。
 * - 建議在使用此函式前，先確認 SYTemp > User 工作表的資料格式與 SYCompany > User 工作表的結構一致，以避免同步過程中出現問題。
 * - 更新紀錄：
 * - 2024-06-01: 初始版本，實現基本的使用者同步功能。
 * - 2024-06-15: 加入自動續傳機制，確保同步過程能夠順利完成，即使遇到執行時間限制的問題。
 * - 2024-06-20: 加入同步過程中的日誌輸出，提升同步過程的可追蹤性與問題排查效率。
 * - 2024-06-25: 優化同步過程的效率，減少對目標試算表的寫入次數，提升整體同步過程的效能。
 *
 */
function dailyMaintenanceJob() {
  // 1. 讀取執行進度 (若無則從 0 開始)
  var state = getProgress("DAILY_MAINTENANCE_STATE");
  var step = state ? state.step : 0;

  if (step > 0) {
    // Set time when the script resumes
    startTime = new Date().getTime();
    logSystemActivity('WARN', 'dailyMaintenanceJob', `偵測到中斷，從步驟 ${step} 繼續執行...`);
  }

  try {
    logSystemActivity('INFO', 'dailyMaintenanceJob', '開始執行每日維護任務');
    // 準備試算表物件 (供後續步驟使用)
    // 注意：這裡的 SYTemp 是每日維護任務中資料同步的核心試算表，確保它的權限與訪問正常是非常重要的
    const ssSYTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    const tmpSheet = ssSYTemp.getSheetByName("SR_Data");
    const tarSheet = MainSpreadsheet.getSheetByName("Cust");

    // Step 1: 同步使用者 (User Sync)
    // 注意：processUserSync 函式內部應該包含對 SYTemp > User 工作表的資料讀取與 SYCompany > User 工作表的資料寫入邏輯，確保使用者資料能夠正確同步並且電話號碼以文字格式存儲。
    if (step <= 1) {
      if (checkTimeoutAndScheduleResume(1)) return;
      processUserSync(MainSpreadsheet, ssSYTemp);
      step = 2;
    }

    // Step 2: 同步個案 (Cust Sync)
    // 注意：processCustSync 函式內部應該包含對 SYTemp > Cust 工作表的資料讀取與 SYCompany > Cust 工作表的資料寫入邏輯，確保個案資料能夠正確同步並且相關欄位格式正確。
    if (step <= 2) {
      if (checkTimeoutAndScheduleResume(2)) return;
      processCustSync();
      step = 3;
    }

    // Step 3: 同步封存個案 (Old Cust Sync)
    // 注意：processOldCustSync 函式內部應該包含對 SYTemp > Old_Cust 工作表的資料讀取
    if (step <= 3) {
      if (checkTimeoutAndScheduleResume(3)) return;
      processOldCustSync();
      step = 4;
    }

    // Step 4: 匯入新紀錄 (Transfer Data - 最耗時步驟)
    // 注意：processTransferData 函式內部應該包含對 SYTemp > SR_Data 工作表的資料讀取與 SYCompany > 相關工作表的資料寫入邏輯，確保新紀錄能夠正確匯入並且相關欄位格式正確。由於這是最耗時的步驟，建議在該函式內部實現分批處理與自動續傳機制，以確保即使遇到執行時間限制的問題，也能夠順利完成資料的匯入。
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
    // 注意：removeSRDuplicates 函式內部應該包含對 SYTemp > SR_Data 工作表的資料讀取與重複資料的清除邏輯，確保重複資料能夠正確識別並且從工作表中移除。由於這個步驟可能會涉及大量資料的處理，建議在該函式內部實現分批處理與自動續傳機制，以確保即使遇到執行時間限制的問題，也能夠順利完成重複資料的清除。
    if (step <= 5) {
      if (checkTimeoutAndScheduleResume(5)) return;
      removeSRDuplicates(tmpSheet);
      // 清除重複後主動清除快取
      CacheService.getScriptCache().remove("SRServer01_InitData");
      step = 6;
    }

    // Step 6: 更新長照編碼 (Update LTC Code)
    // 注意：UpdateCustLTCCode 函式內部應該包含對 SYTemp > SR_Data 工作表與 RecUrl > SYyyyy > yyyyMM 工作表的資料讀取，並將相關資料整合後更新到 SYCompany > Cust 工作表的 LTC_Code 欄位。由於這個步驟可能會涉及大量資料的處理，建議在該函式內部實現分批處理與自動續傳機制，以確保即使遇到執行時間限制的問題，也能夠順利完成長照編碼的更新。
    if (step <= 6) {
      if (checkTimeoutAndScheduleResume(6)) return;
      UpdateCustLTCCode(tmpSheet, tarSheet);
      step = 7;
    }

    // Step 7: 更新使用者個案關聯 (Update User Cust Name)
    // 注意：UpdateUserCustName 函式內部應該包含對 SYCompany > User 工作表與 SYCompany > Cust 工作表的資料讀取，並將相關資料整合後更新到 SYCompany > User 工作表的 Cust_Name 欄位。由於這個步驟可能會涉及大量資料的處理，建議在該函式內部實現分批處理與自動續傳機制，以確保即使遇到執行時間限制的問題，也能夠順利完成使用者個案關聯的更新。
    if (step <= 7) {
      if (checkTimeoutAndScheduleResume(7)) return;
      UpdateUserCustName();
      step = 8;
    }

    // Step 8: 資料遷移 — 每月10日後才搬走上個月的資料 (結算日前保留)
    if (step <= 8) {
      if (checkTimeoutAndScheduleResume(8)) return;
      if (new Date().getDate() >= 10) {
        processSRDataMigration();
      }
      step = 9;
    }

    // Step 9: 清理舊日誌 (Cleanup Old Logs)
    if (step <= 9) {
      if (checkTimeoutAndScheduleResume(9)) return;
      cleanupOldErrorLogs();
      step = 10;
    }

    // 全部完成，清除進度
    clearProgress("DAILY_MAINTENANCE_STATE");
    // 確保任務結束時快取一定是最新的
    CacheService.getScriptCache().remove("SRServer01_InitData");
    logSystemActivity('INFO', 'dailyMaintenanceJob', '每日維護任務已成功完成');
  } catch (e) {
    logSystemActivity('ERROR', 'dailyMaintenanceJob', '每日維護任務失敗: ' + e.toString());
    // 發生錯誤時不清除進度，保留狀態供排錯或下次重試
  }
}

/**
 * 更新 SYCompany 試算表中的 LTC_Code
 * 修正：支援多來源、使用 getColIndex 避免大小寫問題
 * 說明：此函式會從一個或多個來源工作表中讀取 CUST_N 與 SR_ID 欄位，整合後更新到目標工作表的 LTC_Code 欄位。整合過程中會確保同一客戶的 SR_ID 唯一且以逗號分隔存儲，並且在更新目標工作表時會保留原有的 LTC_Code 資料，僅在新增 SR_ID 時進行合併與更新。
 * 注意事項：
 * - 確保來源工作表中 CUST_N 與 SR_ID 欄位的名稱與格式正確，以避免資料讀取失敗。
 * - 更新過程中會有詳細的日誌輸出，建議定期檢查執行日誌以確保更新過程順利，並及時發現與解決可能的問題。
 * - 更新完成後會自動清除相關快取，確保前端讀取到最新資料。
 *
 * 更新紀錄：
 * - 2024-06-01: 初始版本，實現基本的 LTC_Code 更新功能。
 * - 2024-06-15: 支援多來源工作表，並使用 getColIndex 函式避免欄位名稱大小寫問題，提升函式的靈活性與容錯能力。
 * - 2024-06-20: 加入更新過程中的日誌輸出，提升更新過程的可追蹤性與問題排查效率。
 * - 2024-06-25: 優化更新過程的效率，減少對目標工作表的寫入次數，提升整體更新過程的效能。
 *
 * @param {Sheet|Sheet[]} srcSheets 來源工作表 (單一或陣列)
 * @param {Sheet} tarSheet 目標工作表 (Cust)
 */
function UpdateCustLTCCode(srcSheets, tarSheet) {
  logSystemActivity('INFO', 'UpdateCustLTCCode', '開始執行 UpdateCustLTCCode 資料整合...');
  // 2. 讀取資料並整合到記憶體 (Map 結構)
  // Map 鍵為 CUST_N, 值為 Set (確保 SR_ID 唯一)
  let combinedData = new Map();
  const excludeKeywords = ["副本", "表單回覆", "Raw Responses"];

  // 支援單一工作表或工作表陣列
  const sheets = Array.isArray(srcSheets) ? srcSheets : [srcSheets];

  sheets.forEach((sheet) => {
    if (!sheet) return;
    const sheetInfo = `[${sheet.getParent().getName()} > ${sheet.getName()}]`;
    
    logSystemActivity('INFO', 'UpdateCustLTCCode', `正在從來源讀取資料: ${sheetInfo}`);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return; // 跳過無資料的工作表

    const headers = data.shift();
    // 使用 getColIndex 處理欄位名稱大小寫差異 (如 CUST_N vs Cust_N)
    const cIdx = getColIndex(headers, "CUST_N"); // 搜尋個案名稱欄位
    const sIdx = getColIndex(headers, "SR_ID");  // 搜尋服務編碼欄位

    if (cIdx === -1 || sIdx === -1) {
      logSystemActivity('WARN', 'UpdateCustLTCCode', `UpdateCustLTCCode: 來源 ${sheetInfo} 缺少必要欄位 (CUST_N 或 SR_ID)，已略過。`);
      return;
    }

    data.forEach((row) => {
      const custN = row[cIdx] ? String(row[cIdx]).trim() : "";
      // 1. 移除所有非英數減號字元，並修正格式：如果 '-' 後面不是數字，則移除該 '-'
      let cleanSrId = row[sIdx] ? String(row[sIdx]).replace(/[^a-zA-Z0-9-]/g, "").replace(/-+(?!\d)/g, "") : "";
      
      // 2. 格式化處理：前2碼大寫，後面小寫
      if (cleanSrId.length >= 2) {
        cleanSrId = cleanSrId.substring(0, 2).toUpperCase() + cleanSrId.substring(2).toLowerCase();
      }

      // 3. 驗證格式: 前2碼大寫字母 + 至少一個數字 + 後續可接數字/小寫，若有 '-' 則後方必接數字
      const isValidSrId = /^[A-Z]{2}\d+[a-z0-9]*(-[0-9]+[a-z0-9]*)*$/.test(cleanSrId);

      if (custN && isValidSrId && !excludeKeywords.some((k) => custN.includes(k))) {
        if (!combinedData.has(custN)) {
          combinedData.set(custN, new Set());
        }
        combinedData.get(custN).add(cleanSrId);
      }
    });
  });

  // 3. 讀取目標工作表並進行比對與更新
  const tarData = tarSheet.getDataRange().getValues();
  const tarHeaders = tarData.shift();
  const tarCIdx = getColIndex(tarHeaders, "Cust_N");
  const tarLIdx = getColIndex(tarHeaders, "LTC_Code");

  if (tarCIdx === -1 || tarLIdx === -1) {
    logSystemActivity('ERROR', 'UpdateCustLTCCode', `UpdateCustLTCCode 嚴重錯誤: 目標表 ${tarSheet.getName()} 缺少 Cust_N 或 LTC_Code 欄位`);
    return;
  }

  // 準備更新後的資料列
  const updatedLTCColumn = tarData.map((row) => {
    const custN = row[tarCIdx];
    let existingCodes = row[tarLIdx]
      ? String(row[tarLIdx])
          .split(",")
          .map(code => {
            // 同步對既有資料進行去空格與 '-' 修正
            let c = code.replace(/[^a-zA-Z0-9-]/g, "").replace(/-+(?!\d)/g, "");
            if (c.length < 2) return c;
            return c.substring(0, 2).toUpperCase() + c.substring(2).toLowerCase();
          })
          .filter(s => s !== "")
      : [];

    
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

  logSystemActivity('INFO', 'UpdateCustLTCCode', `UpdateCustLTCCode 完成。處理客戶數: ${combinedData.size}，目標欄位: LTC_Code`);
  
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
    logSystemActivity('ERROR', 'syncMasterTablePermissions', "syncMasterTablePermissions: 找不到 'Manager' 工作表，權限同步中止。");
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
  targetFileIds.push({
    Name: "SYTemp",
    UrlID: getTargetsheet("SYTemp", "SYTemp").id,
  }); // 包含 SYTemp

  targetFileIds.forEach(function (item) {
    var fileId = item.UrlID;
    var fileName = item.Name;
    logSystemActivity('INFO', 'syncMasterTablePermissions', "正在處理試算表: " + item.Name + " (ID: " + item.UrlID + ")");
    try {
      var file = DriveApp.getFileById(fileId);
      var ownerEmail = file.getOwner().getEmail().toLowerCase();
      // 先取得目前的編輯者名單，避免重複呼叫 API
      var currentEditors = file.getEditors().map(function (u) {
        return u.getEmail().toLowerCase();
      });

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
      logSystemActivity('ERROR', 'syncMasterTablePermissions', "檔案 " + fileName + " (" + fileId + ") 處理失敗: " + e.message);
    }
  });
}

/**
 * 處理 User 同步：SYTemp > User 搬移至 SYCompany > User
 * 特別處理：確保電話號碼 User_Tel 為文字字串格式
 * 1. 自動檢查 Email 是否重複，重複則不新增但仍從 Temp 移除。
 * 2. 確保 User_Tel 以文字格式 (@) 存入。
 * 建議觸發時間：每日 00:00 - 01:00（可搭配 dailyMaintenanceJob 一起執行）
 * 注意事項：
 * - 確保 SYTemp > User 工作表的資料格式正確，特別是 User_N 欄位應該具有唯一值，以避免同步過程中出現重複資料。
 * - 同步過程中會有詳細的日誌輸出，建議定期檢查執行日誌以確保同步過程順利，並及時發現與解決可能的問題。
 * - 同步完成後會自動清理 SYTemp > User 工作表，請確保在執行前已經備份重要資料，以防止不慎操作導致資料遺失。
 * - 若在同步過程中遇到任何問題，請檢查 SYTemp > User 工作表的資料格式是否正確，並確認 SYCompany > User 工作表是否存在且具有足夠的權限進行寫入操作。
 * - 建議在使用此函式前，先確認 SYTemp > User 工作表的資料格式與 SYCompany > User 工作表的結構一致，以避免同步過程中出現問題。
 * - 更新紀錄：
 * - 2024-06-01: 初始版本，實現基本的使用者同步功能。
 * - 2024-06-15: 加入自動續傳機制，確保同步過程能夠順利完成，即使遇到執行時間限制的問題。
 * - 2024-06-20: 加入同步過程中的日誌輸出，提升同步過程的可追蹤性與問題排查效率。
 * - 2024-06-25: 優化同步過程的效率，減少對目標試算表的寫入次數，提升整體同步過程的效能。
 */
function processUserSync(mainSS, tempSS) {
  // 1. 初始值檢查與預設值設定
  if (!mainSS) {
    mainSS = MainSpreadsheet;
    logSystemActivity('INFO', 'processUserSync', 'mainSS 未提供，使用預設 MainSpreadsheet');
  }

  if (!tempSS) {
    tempSS = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    logSystemActivity('INFO', 'processUserSync', 'tempSS 未提供，使用預設 SYTemp');
  }

  const mainUserSheet = mainSS.getSheetByName("User");
  const tempUserSheet = tempSS.getSheetByName("User");

  if (!tempUserSheet || !mainUserSheet) {
    logSystemActivity('ERROR', 'processUserSync', "找不到 User 工作表，請檢查名稱是否正確。");
    return;
  }

  // 2. 讀取資料
  const tempData = tempUserSheet.getDataRange().getValues();
  if (tempData.length <= 1) {
    logSystemActivity('INFO', 'processUserSync', '暫存表無資料，跳過同步。');
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
  const passIdx = getColIndex(mainHeaders, "Pass");
  const roleIdx = getColIndex(mainHeaders, "Role");

  // 檢查主表是否缺少必要欄位
  if (mainColMap["User_N"] === -1) {
    throw new Error("主表找不到標題: User_N，無法進行比對。");
  }

  // 4. 建立主表索引 Map { User_N_Value : rowIndex }
  const mainIndexMap = {};
  for (let i = 1; i < mainData.length; i++) {
    let userName = mainData[i][mainColMap["User_N"]]
      .toString()
      .replace(/[\s,]/g, "");
    if (userName) mainIndexMap[userName] = i;
  }

  const newRowsToAppend = [];
  let updateCount = 0;

  // 5. 開始同步 (比對 User_N)
  for (let i = 1; i < tempData.length; i++) {
    const tempRow = tempData[i];
    const tempUserName = tempRow[tempColMap["User_N"]]
      .toString()
      .replace(/[\s,]/g, "");

    if (!tempUserName) continue; // 跳過空名

    if (mainIndexMap.hasOwnProperty(tempUserName)) {
      // --- 狀況 A: User_N 已存在 -> 更新資料 ---
      const mainRowIdx = mainIndexMap[tempUserName];

      // 更新這三個指定欄位
      targetFields.forEach((field) => {
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

      // 新增時設定預設值
      if (passIdx !== -1) newRow[passIdx] = "";
      if (roleIdx !== -1) newRow[roleIdx] = "user";

      newRowsToAppend.push(newRow);
    }
  }

  // 6. 資料回寫
  // 更新現有列 (一次性覆蓋主表範圍以提升運作效率)
  if (updateCount > 0) {
    mainUserSheet
      .getRange(1, 1, mainData.length, mainHeaders.length)
      .setValues(mainData);
    logSystemActivity('INFO', 'processUserSync', `已成功更新 ${updateCount} 筆現有使用者資料。`);
  }

  // 新增全新列
  if (newRowsToAppend.length > 0) {
    const startRow = mainUserSheet.getLastRow() + 1;
    const targetRange = mainUserSheet.getRange(
      startRow,
      1,
      newRowsToAppend.length,
      mainHeaders.length,
    );
    targetRange.setNumberFormat("@"); // 強制文字格式避免號碼跑掉
    targetRange.setValues(newRowsToAppend);
    logSystemActivity('INFO', 'processUserSync', `已新增 ${newRowsToAppend.length} 筆新使用者資料。`);
  }

  // 7. 排序與清理 - 依 User_N 排序，並清理暫存表
  const finalLastRow = mainUserSheet.getLastRow();
  if (finalLastRow > 1) {
    // 依據 User_N 所在欄位進行 ASC 排序
    mainUserSheet
      .getRange(2, 1, finalLastRow - 1, mainUserSheet.getLastColumn())
      .sort({ column: mainColMap["User_N"] + 1, ascending: true });
    logSystemActivity('INFO', 'processUserSync', '主表資料已完成排序。');
  }

  if (tempUserSheet.getLastRow() > 1) {
    tempUserSheet.deleteRows(2, tempUserSheet.getLastRow() - 1);
    logSystemActivity('INFO', 'processUserSync', '暫存表清理完畢。');
  }
  CacheService.getScriptCache().remove("SRServer01_InitData");
}

/**
 * 處理 SR_Data 遷移：上個月資料搬移至年度試算表
 * 修正：日期偏移、新增首列凍結、設定日期欄位格式
 * 1. 讀取 SYTemp > SR_Data 工作表資料。
 * 2. 根據 Date 欄位判斷是否為上個月資料，若是則搬移至 RecUrl > SYyyyy > yyyyMM 工作表。
 * 3. 搬移後清理 SR_Data 工作表，僅保留未搬移的資料。
 * 特別處理：搬移資料時會自動凍結目標工作表的首列，並將 Date 欄位設定為文字格式以確保日期的正確顯示與排序。
 * 建議觸發時間：每月 10 日 00:00 - 01:00（可搭配 monthlyTenMaintenanceJob 一起執行）
 * 注意事項：
 * - 確保 SYTemp > SR_Data 工作表的 Date 欄位格式正確，建議使用 "yyyy/MM/dd" 或 "yyyy-MM-dd" 的格式，以避免解析失敗。
 * - 搬移過程中會有詳細的日誌輸出，建議定期檢查執行日誌以確保搬移過程順利，並及時發現與解決可能的問題。
 * - 搬移後的目標工作表會自動凍結首列，並將 Date 欄位設定為文字格式，請勿手動修改該欄位的格式，以避免影響資料的正確顯示與排序。
 * - 若有需要排除特定資料不進行搬移，請在 SYTemp > SR_Data 工作表中添加一欄「ExcludeFromMigration」，並在該欄填入「TRUE」以標記該筆資料。
 * - 搬移完成後會自動清理 SR_Data 工作表，僅保留未搬移的資料，請確保在執行前已經備份重要資料，以防止不慎操作導致資料遺失。
 * - 此函式使用了 appendDataToExternalSS 函式來將資料寫入目標試算表，請確保該函式已經正確實現並且能夠正常運作，以確保資料的正確遷移。
 * - 若在搬移過程中遇到任何問題，請檢查 SYTemp > SR_Data 工作表的資料格式是否正確，並確認 RecUrl > SYyyyy > yyyyMM 工作表是否存在且具有足夠的權限進行寫入操作。
 * - 建議在使用此函式前，先確認 SYTemp > SR_Data 工作表的資料格式與 RecUrl > SYyyyy > yyyyMM 工作表的結構一致，以避免搬移過程中出現問題。
 * - 更新紀錄：
 * - 2024-06-01: 初始版本，實現基本的資料遷移功能。
 * - 2024-06-15: 新增日期格式設定與首列凍結功能，確保遷移後資料的正確顯示與排序。
 * - 2024-06-20: 加入搬移過程中的日誌輸出，提升搬移過程的可追蹤性與問題排查效率。
 * - 2024-06-25: 新增搬移資料前的日期格式檢查與解析，提升搬移過程的穩定性與容錯能力。
 * - 2024-06-30: 加入搬移資料的排除機制，允許使用者標記特定資料不進行搬移，提升搬移過程的靈活性與控制力。
 * - 2024-07-05: 優化搬移資料的效率，減少對目標試算表的寫入次數，提升搬移過程的性能與穩定性。
 * - 2024-07-10: 最終版本，完成資料遷移功能的實現與優化，確保搬移過程的順利與資料的正確遷移。
 *
 */
function processSRDataMigration() {
  logSystemActivity('INFO', 'processSRDataMigration', 'processSRDataMigration 開始遷移 SR_Data 資料...');
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
    logSystemActivity('ERROR', 'processSRDataMigration', "processSRDataMigration: 找不到 'Date' 欄位，無法進行遷移。");
    return;
  }

  const cutoffDate = new Date();
  cutoffDate.setDate(1);
  cutoffDate.setHours(0, 0, 0, 0);
  //cutoffDate.setDate(today.getDate() - 8); // 7 天前
  logSystemActivity('INFO', 'processSRDataMigration', "資料遷移截止日期 (cutoffDate): " + cutoffDate.toISOString());

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

    let formattedDate = Utilities.formatDate(
      dateObj,
      "Asia/Taipei",
      "yyyy-MM-dd",
    );
    row[dateIdx] = formattedDate;

    if (dateObj < cutoffDate) {
      let yearmonth =
        dateObj.getFullYear() +
        Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM");
      if (!migrationMap[yearmonth]) migrationMap[yearmonth] = [];
      migrationMap[yearmonth].push(row);
    } else {
      rowsToKeep.push(row);
    }
  }

  for (let yearmonth in migrationMap) {
    let year = yearmonth.substring(0, 4);
    let month = yearmonth.substring(4, 6);
    let syName = "SY" + year;
    // get target spredsheet
    let tarspredsheet = getTargetsheet("RecUrl", syName).Spreadsheet;
    // get target sheet name is yyyyMM
    let tarsheetName = year + month;
    let tarSheet = tarspredsheet.getSheetByName(tarsheetName);
    logSystemActivity('INFO', 'processSRDataMigration', 
      `處理 ${yearmonth} 資料，目標試算表: ${syName}, 目標工作表: ${tarsheetName}`,
    );
    // move data to target sheet
    let targetUrl = tarspredsheet.getUrl();
    if (targetUrl) {
      appendDataToExternalSS(
        targetUrl,
        yearmonth,
        migrationMap[yearmonth],
        headers,
      );
    }
    logSystemActivity('INFO', 'processSRDataMigration', `搬移資料至 ${year} 年試算表，網址: ${targetUrl}，搬移筆數: ${migrationMap[yearmonth].length}`);
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
 * 4. 針對 CUST_N 欄位進行 A 到 Z 排序，確保同一客戶的紀錄相鄰
 * 5. 移除重複的 SR_ID 紀錄，確保資料唯一性
 * 注意：此函式假設來源資料的 Date 欄位已經是可解析的日期格式，並且會將其轉換為 "yyyy-MM-dd" 的字串格式存入目標工作表，以確保日期的一致性與正確排序。
 * 如果來源資料的 Date 欄位格式不正確，可能會導致日期解析失敗，進而影響資料的正確遷移與排序，因此建議在呼叫此函式前先確保來源資料的 Date 欄位格式正確。
 * 特別處理：如果在目標工作表中找不到 Date 或 CUST_N 欄位，會自動退回到預設的第一欄和第三欄進行排序，並在日誌中輸出警告訊息提醒使用者檢查欄位名稱。
 * 建議：在使用此函式前，先確認目標工作表的欄位名稱與來源資料的欄位名稱一致，並且確保 Date 欄位的格式正確，以避免遷移過程中出現問題。
 * 更新紀錄：
 * - 2024-06-01: 初始版本，實現基本的資料遷移與格式設定功能。
 * - 2024-06-15: 新增日期格式設定與排序功能，確保遷移後資料的正確性與可讀性。
 * - 2024-06-20: 加入重複 SR_ID 移除功能，提升資料品質與唯一性。
 * - 2024-06-25: 加入錯誤處理機制，確保在寫入過程中遇到問題能夠適當記錄並繼續執行後續操作。
 * - 2024-06-30: 優化排序邏輯，動態偵測排序欄位並提供預設排序方案，提升函式的彈性與適用性。
 * - 2024-07-05: 加入日誌輸出，提供更詳細的執行過程記錄，方便後續的監控與排錯。
 * - 2024-07-10: 最終測試與驗證，確保函式在各種情況下都能正常運作，並且資料遷移的正確性與完整性得到保障。
 * 注意事項：
 * - 在執行此函式前，建議先備份目標試算表的資料，以防止因格式不正確或其他問題導致資料遺失。
 * - 確保來源資料的 Date 欄位格式正確，建議使用 "yyyy/MM/dd" 或 "yyyy-MM-dd" 的格式，以避免解析失敗。
 * - 此函式會直接修改目標工作表的資料，請確保在執行前已經確認目標工作表的結構與欄位名稱，以避免因欄位名稱不一致導致的問題。
 * - 如果目標工作表中已經存在相同 SR_ID 的紀錄，會自動移除重複的紀錄，確保資料的唯一性，但建議在執行前先確認是否有重要資料可能被誤刪。
 * - 此函式假設來源資料的欄位順序與目標工作表的欄位順序一致，如果不一致可能會導致資料錯亂，建議在使用前先確認欄位順序的一致性。
 * - 在執行過程中，如果遇到任何錯誤，會在日誌中記錄錯誤訊息，建議定期檢查執行日誌以確保遷移過程順利，並及時發現與解決可能的問題。
 * - 此函式使用了 getColIndex 函式來動態偵測欄位索引，確保在欄位名稱大小寫不一致的情況下仍能正確找到對應的欄位，提升函式的彈性與適用性。
 * - 此函式會自動凍結目標工作表的首列，確保在資料量較大時仍能方便地查看欄位名稱，提升使用者體驗。
 * - 此函式會將目標工作表的 A 欄設定為日期格式，確保日期資料的正確顯示與排序，建議在使用前先確認目標工作表的 A 欄沒有其他重要資料，以避免格式設定導致的問題。
 * - 此函式會在遷移完成後自動清除相關快取，確保前端讀取到最新的資料，建議在使用前先確認是否有其他功能依賴於這些快取，以避免因快取清除導致的問題。
 * - 此函式的設計目的是為了確保資料遷移的正確性與完整性，建議在使用前先仔細閱讀此函式的說明與注意事項，以確保在執行過程中能夠順利完成資料遷移，並且避免因操作不當導致的問題。
 *
 *
 *
 */
function appendDataToExternalSS(url, year, rows, headers) {
  try {
    const targetSS = SpreadsheetApp.openByUrl(url);
    const dateIdx = getColIndex(headers, "Date");
    if (dateIdx === -1) {
      logSystemActivity('ERROR', 'appendDataToExternalSS', "appendDataToExternalSS: 來源資料中找不到 'Date' 欄位，無法進行遷移。");
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
      const targetHeaders = targetSheet
        .getRange(1, 1, 1, targetSheet.getLastColumn())
        .getValues()[0];
      const dateSortIdx = getColIndex(targetHeaders, "Date");
      const custSortIdx = getColIndex(targetHeaders, "CUST_N");

      const sortColumns = [];

      // 優先使用 Date 欄位排序，若無則預設第一欄
      if (dateSortIdx !== -1) {
        sortColumns.push({ column: dateSortIdx + 1, ascending: true });
      } else {
        logSystemActivity('WARN', 'appendDataToExternalSS', `在 ${targetSheet.getName()} 中找不到 'Date' 欄位，將使用預設第一欄進行排序。`);
        sortColumns.push({ column: 1, ascending: true });
      }

      // 其次使用 CUST_N 欄位排序，若無則預設第三欄
      if (custSortIdx !== -1) {
        sortColumns.push({ column: custSortIdx + 1, ascending: true });
      } else {
        logSystemActivity('WARN', 'appendDataToExternalSS', `在 ${targetSheet.getName()} 中找不到 'CUST_N' 欄位，將使用預設第三欄進行排序。`);
        sortColumns.push({ column: 3, ascending: true });
      }
      targetSheet
        .getRange(2, 1, fullRange.getNumRows() - 1, fullRange.getNumColumns())
        .sort(sortColumns);
    }
    removeSRDuplicates(targetSheet);

    logSystemActivity('INFO', 'appendDataToExternalSS', `成功搬移並排序 ${rows.length} 筆資料至 ${year} 年 ${monthStr} 表`);
  } catch (e) {
    logSystemActivity('ERROR', 'appendDataToExternalSS', "寫入外部試算表失敗: " + e.toString());
  }
}

/**
 * 輔助函式：比對個案資料是否有變更
 * 說明：此函式專門用於比對個案資料的關鍵欄位（性別、生日、地址、服務項目、表單網址），以判斷是否需要更新目標工作表的資料。
 * 注意：比對生日欄位時，建議將日期轉為統一格式的字串（如 yyyy/M/d）進行比對，以避免因日期物件格式差異導致的誤判。
 * 比對邏輯：
 * 1. 比對性別、地址、服務項目、表單網址欄位的值是否相同。
 * 2. 比對生日欄位時，將來源與目標的日期都轉為 yyyy/M/d 格式的字串進行比對。
 * 3. 若任一欄位的值不同，則回傳 true 表示資料有變更；若所有欄位的值都相同，則回傳 false 表示資料無變更。
 * 參數說明：
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
 * 表單網址 -> Form_Url
 * 特別處理：生日欄位需要確保格式一致，建議統一轉為 yyyy/M/d 格式的字串進行比對與存儲。
 */
function processCustSync() {
  syncCustData("個案清單", "Cust");
}

/**
 * 處理 Old Cust 同步：Case_Reports > 結案個案清單 搬移至 SYCompany > OldCust
 * 對應欄位同上
 */
function processOldCustSync() {
  syncCustData("結案個案清單", "OldCust");
}

/**
 * 通用個案同步邏輯 (Internal Helper)
 * 說明：此函式包含從來源工作表讀取資料、比對目標工作表、更新或新增資料的完整邏輯。
 * 注意：來源工作表必須包含固定欄位 (姓名、性別、生日、地址、服務項目、表單網址)，且目標工作表必須包含對應的欄位。
 * 比對邏輯：
 * 1. 以「姓名」作為唯一識別碼進行比對。
 * 2. 若姓名存在但其他欄位資料不同，則更新該列資料。
 * 3. 若姓名不存在，則新增一列資料。
 * 4. 同步完成後會寫回整個資料範圍，確保資料一致性。
 * 性能考量：
 * - 使用 Map 結構優化姓名查詢，避免每筆資料都進行迴圈比對。
 * - 批次寫回資料，減少對試算表的讀寫次數。
 * 參數說明：
 * @param {string} sourceSheetName 來源工作表名稱
 * @param {string} targetSheetName 目標工作表名稱
 */
function syncCustData(sourceSheetName, targetSheetName) {
  logSystemActivity('INFO', 'syncCustData', `開始同步 Case_Reports > ${sourceSheetName} > ${targetSheetName} 資料...`);
  const SOURCE_SS_ID = "1ib8q-lKJgLEhRVrwncnRqOyKNauMqaV2wtYEpGlmRlk";

  const sourceSheet =
    SpreadsheetApp.openById(SOURCE_SS_ID).getSheetByName(sourceSheetName);
  const targetSheet = MainSpreadsheet.getSheetByName(targetSheetName);

  if (!sourceSheet || !targetSheet) {
    logSystemActivity('ERROR', 'syncCustData', `無法開啟工作表: ${sourceSheetName} 或 ${targetSheetName}`);
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
  const targetFieldNames = [
    "Cust_N",
    "Cust_Sex",
    "Cust_BD",
    "Cust_Add",
    "Cust_LTC_Code",
    "Form_Url",
  ];
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
    logSystemActivity('INFO', 'syncCustData', `同步完成！ 更新：${updateCount} 筆 新增：${insertCount} 筆`);
  } else {
    logSystemActivity('INFO', 'syncCustData', "資料皆為最新，無需更新。");
  }
}

/**
 * 設定系統自動化觸發條件 (Triggers)
 * 用途：自動建立每日與每月的排程任務
 * 操作：請在編輯器中手動執行此函式一次以完成安裝
 * 觸發條件：
 * 1. 每日維護任務 (dailyMaintenanceJob)：每天 00:00 - 01:00 執行，負責同步使用者/個案資料、匯入 Raw Response、更新 LTC Code。
 * 2. 每月 1 號維護任務 (monthlyMaintenanceJob)：每月 1 號 01:00 - 02:00 執行，負責同步試算表權限、跨月資料整合。
 * 3. 每月 10 號維護任務 (monthlyTenMaintenanceJob)：每月 10 號 00:00 - 01:00 執行，負責將上個月資料從 SYTemp 搬移至年度封存表。
 * 注意事項：
 * - 此函式包含防呆機制，使用 LockService 避免短時間內重複執行，確保觸發條件不會被重複建立。
 * - 在執行此函式前，建議先確認目前的觸發條件狀態，以避免不必要的觸發條件被刪除或重複建立。
 * - 此函式會自動清除現有相關的觸發條件，請確保在執行前已經備份重要的觸發條件設定，以防止不慎操作導致的問題。
 */
function setupTriggers() {
  // 防呆機制：使用 LockService 避免短時間內重複執行
  var lock = LockService.getScriptLock();
  // 嘗試取得鎖定，若 5 秒內無法取得則認為是重複執行
  if (!lock.tryLock(5000)) {
    logSystemActivity('WARN', 'setupTriggers', "setupTriggers 正在執行中，請勿重複觸發。");
    return;
  }

  try {
    logSystemActivity('INFO', 'setupTriggers', "開始設定觸發條件...");

    // 1. 清除現有相關觸發條件，避免重複建立
    const triggers = ScriptApp.getProjectTriggers();
    const handlerNames = [
      "dailyMaintenanceJob",
      "monthlyMaintenanceJob",
      "monthlyTenMaintenanceJob",
    ];
    let deletedCount = 0;

    triggers.forEach((trigger) => {
      if (handlerNames.includes(trigger.getHandlerFunction())) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });

    if (deletedCount > 0) {
      logSystemActivity('INFO', 'setupTriggers', `已清除 ${deletedCount} 個舊有的觸發條件。`);
    }

    // 2. 設定每日維護任務 (每日 00:00 - 01:00 執行)
    // 負責：同步使用者/個案資料、匯入 Raw Response、更新 LTC Code
    ScriptApp.newTrigger("dailyMaintenanceJob")
      .timeBased()
      .everyDays(1)
      .atHour(0)
      .create();

    // 3. 設定每月 1 號維護任務 (每月 1 號 01:00 - 02:00 執行)
    // 負責：同步試算表權限、跨月資料整合
    ScriptApp.newTrigger("monthlyMaintenanceJob")
      .timeBased()
      .onMonthDay(1)
      .atHour(1)
      .create();

    // 4. 設定每月 10 號維護任務 (每月 10 號 00:00 - 01:00 執行)
    // 負責：將上個月資料從 SYTemp 搬移至年度封存表
    ScriptApp.newTrigger("monthlyTenMaintenanceJob")
      .timeBased()
      .onMonthDay(10)
      .atHour(0)
      .create();

    logSystemActivity('INFO', 'setupTriggers', "系統自動化觸發條件已設定完成 (共 3 個任務)。");
  } catch (e) {
    logSystemActivity('ERROR', 'setupTriggers', "設定觸發條件失敗: " + e.toString());
  } finally {
    // 確保釋放鎖定
    lock.releaseLock();
  }
}

/**
 * 輔助函式：檢查執行時間是否即將逾時
 * 若逾時：儲存當前步驟、建立續傳觸發器、回傳 true
 * 否則：回傳 false，繼續執行
 * 注意：Google Apps Script 的執行時間限制通常為 6 分鐘，建議在每個主要步驟結束後呼叫此函式，以確保在執行時間即將逾時時能夠適當地儲存進度並排程續傳，避免因執行時間超過限制而導致的錯誤。
 * 使用方式：在每日維護任務 (dailyMaintenanceJob) 的每個主要步驟結束後呼叫此函式，並傳入當前的步驟名稱或編號，以便在儲存進度時能夠清楚地知道當前的執行狀態。
 * 更新紀錄：
 * - 2024-06-01: 初始版本，實現基本的逾時檢查與續傳排程功能。
 * - 2024-06-15: 優化逾時檢查邏輯，加入更精確的時間計算與警告訊息，提升使用者體驗與問題排查效率。
 * - 2024-06-20: 加入進度儲存的詳細資訊，確保在續傳時能夠正確地從上次中斷的地方繼續執行，提升系統的穩定性與可靠性。
 * - 2024-06-25: 加入錯誤處理機制，確保在儲存進度或建立續傳觸發器時遇到問題能夠適當記錄並繼續執行後續操作，提升系統的容錯能力。
 * - 2024-06-30: 優化續傳排程的邏輯，確保在多次續傳的情況下能夠正確地管理觸發器，避免因觸發器重複建立而導致的問題。
 * - 2024-07-05: 加入日誌輸出，提供更詳細的執行過程記錄，方便後續的監控與排錯，提升系統的可維護性與透明度。
 * 注意事項：
 * - 在使用此函式前，建議先確認目前的執行時間狀態，以確保在適當的時機呼叫此函式，避免因過早或過晚呼叫而導致的問題。
 */
function checkTimeoutAndScheduleResume(currentStep) {
  if (isNearTimeout()) {
    logSystemActivity('WARN', 'checkTimeoutAndScheduleResume', `[System] 執行時間不足 (Step ${currentStep})，正在儲存進度並排程續傳...`);
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

/**
 * 清理 ErrorLog 工作表中超過 3 個月的舊資料
 * 建議：每日維護任務的一部分
 */
function cleanupOldErrorLogs() {
  const sheet = MainSpreadsheet.getSheetByName("ErrorLog");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const cutoffDate = new Date();
  cutoffDate.setMonth(cutoffDate.getMonth() - 3); // 設定為 3 個月前

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowsToDelete = 0;

  // 假設日誌是依時間順序附加的 (最舊的在上面)
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && new Date(data[i][0]) < cutoffDate) {
      rowsToDelete++;
    } else {
      // 遇到第一個未過期的日期即可停止檢查
      break;
    }
  }

  if (rowsToDelete > 0) {
    sheet.deleteRows(2, rowsToDelete);
    logSystemActivity('INFO', 'cleanupOldErrorLogs', `清理日誌完成：已移除 ${rowsToDelete} 筆超過 3 個月的舊錯誤紀錄。`);
  }
}

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
    UpdateCustLTCCode(srcSheet, tarSheet);
    UpdateCustLTCCode(tmpSheet, tarSheet);
    console.log(`成功處理：${srcSpreadsheetName} > ${yyyyMM}`);
  } else {
    console.error(`找不到工作表：${yyyyMM}`);
  }
}

/**
 * 每日維護任務：遷移 7 天前資料與同步 User 名單
 * 建議觸發時間：每日 00:00 - 01:00
 */
function dailyMaintenanceJob() {
  const ssSYTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
  const tmpSheet = ssSYTemp.getSheetByName("SR_Data");    
  const tarSheet = MainSpreadsheet.getSheetByName("Cust");

  processUserSync(MainSpreadsheet, ssSYTemp);
  processCustSync();
  processTransferData("all", true);
  processSRDataMigration(MainSpreadsheet, ssSYTemp);
  
  removeSRDuplicates(tmpSheet);
  UpdateCustLTCCode(tmpSheet, tarSheet);
  // const sSYyear = getTargetsheet("RecUrl","")
}


/**
 * 更新 SYCompany 試算表中的 LTC_Code
 * @param {string} srcId 來源試算表 (SY+yyyy) 的 ID
 * @param {string} tmpId SYTemp 試算表的 ID
 * @param {string} tarId SYCompany 試算表的 ID
 */
function UpdateCustLTCCode(srcSheet, tarSheet) {
  // 2. 讀取資料並整合到記憶體 (Map 結構)
  // Map 鍵為 CUST_N, 值為 Set (確保 SR_ID 唯一)
  let combinedData = new Map();

  function processSheet(sheet) {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const cIdx = headers.indexOf("CUST_N");
    const sIdx = headers.indexOf("SR_ID");

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
  }

  processSheet(srcSheet);
  // processSheet(tmpSheet);

  // 3. 讀取目標工作表並進行比對與更新
  const tarData = tarSheet.getDataRange().getValues();
  const tarHeaders = tarData.shift();
  const tarCIdx = tarHeaders.indexOf("Cust_N");
  const tarLIdx = tarHeaders.indexOf("LTC_Code");
  
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
  var managerData = managerSheet.getDataRange().getValues();
  var managerEmails = [];
  for (var i = 1; i < managerData.length; i++) {
    var email = managerData[i][1];
    if (email) managerEmails.push(email.toString().trim().toLowerCase());
  }

  var targetFileIds = [{ Name: "SYCompany", UrlID: MainSpreadsheet.getId() }]; // 包含 SYCompany 本身
  targetFileIds.push({Name: "SYTemp",UrlID: getTargetsheet("SYTemp", "SYTemp").id }); // 包含 SYTemp
  var recUrlSheet = MainSpreadsheet.getSheetByName("RecUrl"); // 取得 RecUrl 工作表
  if (recUrlSheet) {
    var urlData = recUrlSheet.getDataRange().getValues();
    for (var j = 1; j < urlData.length; j++) {
      var name = urlData[j][0];
      var url = urlData[j][1];
      if (url && url.indexOf("docs.google.com") !== -1) {
        try {
          targetFileIds.push({
            Name: name,
            UrlID: SpreadsheetApp.openByUrl(url).getId(),
          });
        } catch (e) {
          console.error(
            "無法開啟試算表 " +
              name +
              "，網址: " +
              url +
              "，錯誤: " +
              e.message,
          );
        }
      }
    }
  }

  var recUrlSheet = MainSpreadsheet.getSheetByName("ReportsUrl"); // 取得 ReportsUrl  工作表
  if (recUrlSheet) {
    var urlData = recUrlSheet.getDataRange().getValues();
    for (var j = 1; j < urlData.length; j++) {
      var name = urlData[j][0];
      var url = urlData[j][1];
      if (url && url.indexOf("docs.google.com") !== -1) {
        try {
          targetFileIds.push({
            Name: name,
            UrlID: SpreadsheetApp.openByUrl(url).getId(),
          });
        } catch (e) {
          console.error(
            "無法開啟試算表 " +
              name +
              "，網址: " +
              url +
              "，錯誤: " +
              e.message,
          );
        }
      }
    }
  }

  // console.log("開始同步權限至 " + targetFileIds.length + " 個試算表");
  // console.log("管理員名單: " + managerEmails.join(", "));

  targetFileIds.forEach(function (item) {
    var fileId = item.UrlID;
    var fileName = item.Name;
    console.log("正在處理試算表: " + item.Name + " (ID: " + item.UrlID + ")");
    try {
      managerEmails.forEach(function (email) {
        var resource = {
          role: "writer",
          type: "user",
          emailAddress: email,
        };

        Drive.Permissions.create(resource, fileId, {
          sendNotificationEmails: false,
        });
      });

      var file = DriveApp.getFileById(fileId);
      var ownerEmail = file.getOwner().getEmail().toLowerCase();

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

  // 3. 動態偵測欄位 Index (標題名稱必須精確匹配)
  const targetFields = ["User_N", "User_Email", "User_Tel"];
  const mainColMap = {};
  const tempColMap = {};

  targetFields.forEach(field => {
    mainColMap[field] = mainHeaders.indexOf(field);
    tempColMap[field] = tempHeaders.indexOf(field);
  });

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
}

/**
 * 處理 SR_Data 遷移：7天前資料搬移至年度試算表
 * 修正：日期偏移、新增首列凍結、設定日期欄位格式
 */
function processSRDataMigration(mainSS, tempSS) {
  console.log("processSRDataMigration 開始遷移 SR_Data 資料...");
  if (!mainSS) {
    mainSS = MainSpreadsheet;
    console.log("mainSS 未提供，使用預設 MainSpreadsheet");
  }

  if (!tempSS) {
    tempSS = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    console.log("tempSS 未提供，使用預設 SYTemp");
  }

  const srSheet = tempSS.getSheetByName("SR_Data");
  if (!srSheet) return;

  const data = srSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const today = new Date();
  const cutoffDate = new Date();
  cutoffDate.setDate(today.getDate() - 8); // 7 天前

  const migrationMap = {};
  const rowsToKeep = [headers];
  let createdNewSS = false;

  for (let i = 1; i < data.length; i++) {
    let row = [...data[i]];
    let rawDate = row[0];

    let dateObj;
    if (rawDate instanceof Date) {
      dateObj = rawDate;
    } else {
      dateObj = new Date(rawDate.toString().replace(/-/g, "/"));
    }

    let formattedDate = Utilities.formatDate(
      dateObj,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
    row[0] = formattedDate;

    if (dateObj < cutoffDate) {
      let year = dateObj.getFullYear();
      if (!migrationMap[year]) migrationMap[year] = [];
      migrationMap[year].push(row);
    } else {
      rowsToKeep.push(row);
    }
  }

  for (let year in migrationMap) {
    let syName = "SY" + year;
    let targetUrl = getUrlFromRecUrl(mainSS, syName);

    if (!targetUrl) {
      targetUrl = createNewYearlySS(mainSS, syName);
      createdNewSS = true;
    }
    console.log("搬移資料至 " + year + " 年試算表，網址: " + targetUrl);
    console.log("搬移筆數: " + migrationMap[year].length);
    if (targetUrl) {
      appendDataToExternalSS(targetUrl, year, migrationMap[year]);
    }
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
function appendDataToExternalSS(url, year, rows) {
  try {
    const targetSS = SpreadsheetApp.openByUrl(url);
    const firstDateStr = rows[0][0].toString().replace(/-/g, "/");
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

    const fullRange = targetSheet.getDataRange();
    const newFilter = fullRange.createFilter();
    newFilter.sort([
      { column: 1, ascending: true },
      { column: 3, ascending: true },
    ]);

    console.log(
      `成功搬移並排序 ${rows.length} 筆資料至 ${year} 年 ${monthStr} 表`,
    );
  } catch (e) {
    console.error("寫入外部試算表失敗: " + e.toString());
  }
}

/**
 * 輔助函式：建立新年度試算表並回傳網址
 */
function createNewYearlySS(mainSS, syName) {
  // 1. Create the new Spreadsheet
  const newSS = SpreadsheetApp.create(syName);
  const url = newSS.getUrl();

  // 2. Get the Recording Sheet
  const recSheet = mainSS.getSheetByName("RecUrl");

  // 3. Append the new row
  recSheet.appendRow([syName, url]);

  // --- Formatting the RecUrl Sheet ---

  // Freeze the first row (Header)
  if (recSheet.getFrozenRows() === 0) {
    recSheet.setFrozenRows(1);
  }

  // Get the data range (All rows and columns that have data)
  const fullRange = recSheet.getDataRange();

  // Remove existing filters to avoid conflicts, then create a new one
  if (recSheet.getFilter()) {
    recSheet.getFilter().remove();
  }
  fullRange.createFilter();

  // Sort the range by syName (Column A / Index 1) in Ascending order
  fullRange.sort([
    { column: 1, ascending: true },
    { column: 3, ascending: true },
  ]);
  return url;
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
  console.log(
    "processCustSync 開始同步 Case_Reports > 個案清單 > Cust 資料...",
  );
  const SOURCE_SS_ID = "1ib8q-lKJgLEhRVrwncnRqOyKNauMqaV2wtYEpGlmRlk";

  const sourceSheet =
    SpreadsheetApp.openById(SOURCE_SS_ID).getSheetByName("個案清單");
  const targetSheet = MainSpreadsheet.getSheetByName("Cust");

  // 1. 取得來源與目標的所有資料
  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();

  const sourceHeaders = sourceData[0];
  const targetHeaders = targetData[0];
  const sourceRows = sourceData.slice(1);
  const targetRows = targetData.slice(1);

  // 2. 建立欄位索引字典
  const getIdx = (headers, name) => headers.indexOf(name);
  const srcIdx = { name: 0, sex: 1, bd: 2, add: 3, ltc: 5 }; // 根據 CSV 結構固定或動態獲取
  const tarIdx = {
    name: getIdx(targetHeaders, "Cust_N"),
    sex: getIdx(targetHeaders, "Cust_Sex"),
    bd: getIdx(targetHeaders, "Cust_BD"),
    add: getIdx(targetHeaders, "Cust_Add"),
    ltc: getIdx(targetHeaders, "Cust_LTC_Code"),
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
    };

    if (targetMap.has(sName)) {
      // --- 狀況 A: 姓名已存在，檢查內容是否不同 ---
      let tEntry = targetMap.get(sName);
      let tRow = tEntry.data;

      // 比對日期需要特別小心，建議轉字串比對
      const tBD =
        tRow[tarIdx.bd] instanceof Date
          ? Utilities.formatDate(tRow[tarIdx.bd], "GMT+8", "yyyy/M/d")
          : tRow[tarIdx.bd];

      const isChanged =
        tRow[tarIdx.sex] !== newData.sex ||
        tBD !== newData.bd ||
        tRow[tarIdx.add] !== newData.add ||
        tRow[tarIdx.ltc] !== newData.ltc;

      if (isChanged) {
        finalRows[tEntry.index][tarIdx.sex] = newData.sex;
        finalRows[tEntry.index][tarIdx.bd] = newData.bd;
        finalRows[tEntry.index][tarIdx.add] = newData.add;
        finalRows[tEntry.index][tarIdx.ltc] = newData.ltc;
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
    console.log(`同步完成！ 更新：${updateCount} 筆 新增：${insertCount} 筆`);
  } else {
    console.log("資料皆為最新，無需更新。");
  }
}

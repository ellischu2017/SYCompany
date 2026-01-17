/**
 * Maintenance.gs - 維護任務模組
 * 提供自動維護、同步、資料遷移等功能
 */

/**
 * 同步所有相關試算表的權限
 * 1. 包含 SYCompany 本身與 RecUrl 內的所有試算表。
 * 2. 根據 Manager 工作表名單授權為「編輯者」。
 * 3. 移除名單外所有「特定的」編輯者與檢視者。
 * 4. 將「一般存取權」設為「知道連結的人即可檢視」。
 */
function syncMasterTablePermissions() {
  var managerSheet = MainSpreadsheet.getSheetByName("Manager");
  var managerData = managerSheet.getDataRange().getValues();
  var managerEmails = [];
  for (var i = 1; i < managerData.length; i++) {
    var email = managerData[i][1];
    if (email) managerEmails.push(email.toString().trim().toLowerCase());
  }

  var targetFileIds = [MainSpreadsheet.getId()];
  var recUrlSheet = MainSpreadsheet.getSheetByName("RecUrl");
  if (recUrlSheet) {
    var urlData = recUrlSheet.getDataRange().getValues();
    for (var j = 1; j < urlData.length; j++) {
      var url = urlData[j][1];
      if (url && url.indexOf("docs.google.com") !== -1) {
        try {
          targetFileIds.push(SpreadsheetApp.openByUrl(url).getId());
        } catch (e) {}
      }
    }
  }

  console.log("開始同步權限至 " + targetFileIds.length + " 個試算表");  
  console.log("管理員名單: " + managerEmails.join(", "));
  console.log("目標試算表 ID 列表: " + targetFileIds.join(", ")); 

  targetFileIds.forEach(function (fileId) {
    try {
      managerEmails.forEach(function (email) {
        var resource = {
          role: "writer",
          type: "user",
          value: email,
        };

        Drive.Permissions.insert(resource, fileId, {
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

      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
    } catch (e) {
      console.error("檔案 ID " + fileId + " 處理失敗: " + e.message);
    }
  });
}

/**
 * 每日維護任務：遷移 7 天前資料與同步 User 名單
 * 建議觸發時間：每日 00:00 - 01:00
 */
function dailyMaintenanceJob() {
  const tempSS = getTargetsheet("SYTemp", "SYTemp");

  processUserSync(MainSpreadsheet, tempSS);
  processSRDataMigration(MainSpreadsheet, tempSS);
}

/**
 * 處理 User 同步：SYTemp > User 搬移至 SYCompany > User
 * 特別處理：確保電話號碼 User_Tel 為文字字串格式
 * 1. 自動檢查 Email 是否重複，重複則不新增但仍從 Temp 移除。
 * 2. 確保 User_Tel 以文字格式 (@) 存入。
 */
function processUserSync(mainSS, tempSS) {
  const tempUserSheet = tempSS.getSheetByName("User");
  const mainUserSheet = mainSS.getSheetByName("User");

  if (!tempUserSheet || !mainUserSheet) return;

  const tempData = tempUserSheet.getDataRange().getValues();
  if (tempData.length <= 1) return;

  const mainData = mainUserSheet.getDataRange().getValues();
  const existingEmails = mainData
    .slice(1)
    .map((row) => row[1].toString().trim().toLowerCase());

  const newRowsToAppend = [];
  const headers = tempData[0];

  for (let i = 1; i < tempData.length; i++) {
    let row = tempData[i];
    let tempEmail = row[1].toString().trim().toLowerCase();

    if (existingEmails.indexOf(tempEmail) === -1) {
      newRowsToAppend.push(row);
    } else {
      console.log("Email 已存在，跳過新增: " + tempEmail);
    }
  }

  if (newRowsToAppend.length > 0) {
    const startRow = mainUserSheet.getLastRow() + 1;
    const targetRange = mainUserSheet.getRange(
      startRow,
      1,
      newRowsToAppend.length,
      headers.length
    );

    targetRange.setNumberFormat("@");
    targetRange.setValues(newRowsToAppend);
    console.log("已新增 " + newRowsToAppend.length + " 筆新居服員資料。");
  }

  if (tempUserSheet.getLastRow() > 1) {
    tempUserSheet.deleteRows(2, tempUserSheet.getLastRow() - 1);
    console.log("SYTemp > User 已清理完畢。");
  }
}

/**
 * 處理 SR_Data 遷移：7天前資料搬移至年度試算表
 * 修正：日期偏移、新增首列凍結、設定日期欄位格式
 */
function processSRDataMigration(mainSS, tempSS) {
  const srSheet = tempSS.getSheetByName("SR_Data");
  if (!srSheet) return;

  const data = srSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const today = new Date();
  const cutoffDate = new Date();
  cutoffDate.setDate(today.getDate() - 7);

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
      "yyyy-MM-dd"
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

    if (targetUrl) {
      appendDataToExternalSS(targetUrl, year, migrationMap[year]);
    }
  }

  if (createdNewSS) {
    syncMasterTablePermissions();
  }

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
      "yyyyMM"
    );

    let targetSheet = targetSS.getSheetByName(monthStr);

    if (!targetSheet) {
      targetSheet = targetSS.insertSheet(monthStr);
      const headers = [
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
    newFilter.sort(1, true);

    console.log(
      `成功搬移並排序 ${rows.length} 筆資料至 ${year} 年 ${monthStr} 表`
    );
  } catch (e) {
    console.error("寫入外部試算表失敗: " + e.toString());
  }
}

/**
 * 輔助函式：建立新年度試算表並回傳網址
 */
function createNewYearlySS(mainSS, syName) {
  const newSS = SpreadsheetApp.create(syName);
  const url = newSS.getUrl();
  const recSheet = mainSS.getSheetByName("RecUrl");
  recSheet.appendRow([syName, url]);
  return url;
}

/**
 * Utilities.gs - 工具函式模組
 * 提供通用的工具函式
 */

// 全域試算表參考 SYCompany
const MainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

/**
 * 測試用的 include 函數 (若未來需要拆分 CSS/JS 檔案時使用)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 從試算表動態取得意見反應連結並生成 Footer HTML
 */
function includeFooter() {
  let suggestUrl = "";

  try {
    const sheet = MainSpreadsheet.getSheetByName("SYTemp");

    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === "SYSuggest") {
          suggestUrl = data[i][1];
          break;
        }
      }
    }
  } catch (e) {
    console.log("取得意見反應網址失敗: " + e.message);
  }

  const template = HtmlService.createTemplateFromFile("Footer");
  template.suggestUrl = suggestUrl;
  return template.evaluate().getContent();
}

/**
 * 提供給 HTML 範本呼叫，用來載入導航列組件
 */
function includeNav() {
  var template = HtmlService.createTemplateFromFile("Nav");
  return template.evaluate().getContent();
}

/**
 * 取得當前 Web App 的 URL
 * 用於前端按鈕跳轉
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * 輔助函式：從 SYCompany (本腳本綁定之試算表) 的 sheetName 工作表取得外部試算表物件
 */

function getTargetsheet(sheetName, targetName) {
  var sheet = MainSpreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error("找不到 SYCompany 中的" + sheetName + "工作表");

  var data = sheet.getDataRange().getValues();
  var url = "";

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === targetName) {
      url = data[i][1];
      break;
    }
  }

  if (!url) {
    if (sheetName === "ReportsUrl") {
      var year = targetName.substring(2, 6);
      var month = targetName.substring(6, 8);

      // 如果沒有，拷貝 Template
      // 尋找 Template URL
      var tempUrl = "";
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === "RPTemplate") {
          tempUrl = data[i][1];
          break;
        }
      }
      if (!tempUrl) throw new Error("找不到 RPTemplate");

      var templateFile = DriveApp.getFileById(getIdFromUrl(tempUrl));
      var folderName = "RP" + year;
      var folders = DriveApp.getfoldersByName(folderName);
      var destinationFolder;
      if (folders.hasNext()) {
        destinationFolder = folders.next();
      } else {
        destinationFolder = DriveApp.createFolder(folderName);
      }
      var newFile = templateFile.makeCopy(targetName, destinationFolder);
      url = newFile.getUrl();

      // 把名稱及網址存到 ss4
      sheet.appendRow([targetName, url]);
    } else {
      console.log(
        "無法在" + sheetName + "工作表中找到名稱為" + targetName + "的對應網址",
      );
      return url;
    }
  }
  return {
    url: url,
    id: getIdFromUrl(url),
    Spreadsheet: SpreadsheetApp.openByUrl(url),
  };
}

function removeSR() {
  const ssSYTemp = getTargetsheet("RecUrl", "SY2026").Spreadsheet;
  const sSRData = ssSYTemp.getSheetByName("202601");
  removeSRDuplicates(sSRData);
}

function removeSRDuplicates(sheet) {
  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  // const sheet = ss.getSheetByName("SR_Data");
  console.log("開始處理 SR_Data 重複資料...");
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return; // 如果沒資料就結束

  const headers = data[0];
  const rows = data.slice(1);

  // 定義要檢查的目標欄位名稱
  const targetFields = ["Date", "SRTimes", "CUST_N", "USER_N", "SR_ID"];

  // 自動尋找標題對應的索引位置
  const indices = targetFields.map((field) => headers.indexOf(field));

  // 檢查是否所有欄位都存在
  if (indices.includes(-1)) {
    const missing = targetFields.filter((_, i) => indices[i] === -1);
    SpreadsheetApp.getUi().alert("找不到欄位: " + missing.join(", "));
    return;
  }

  const seen = new Set();
  const uniqueRows = [];

  rows.forEach((row) => {
    // 根據索引組合唯一鍵值
    const key = indices.map((idx) => row[idx]).join("|");

    if (!seen.has(key)) {
      uniqueRows.push(row);
      seen.add(key);
    }
  });

  // 清除舊資料並貼上
  sheet.clearContents();
  const output = [headers].concat(uniqueRows);
  sheet.getRange(1, 1, output.length, headers.length).setValues(output);

  console.log("處理完成！剩餘筆數：" + uniqueRows.length);
}

// 在 Utilities.js 加入
var startTime = new Date().getTime();

/** 檢查是否快要超時 (設定為 5 分鐘以確保安全) */
function isNearTimeout() {
  return new Date().getTime() - startTime > 20 * 1000;
}

/** 儲存/讀取進度 (PropertiesService 會存在雲端專案屬性中) */
function saveProgress(data) {
  PropertiesService.getScriptProperties().setProperty(
    "REPORT_JOB",
    JSON.stringify(data),
  );
}

function getProgress() {
  var p = PropertiesService.getScriptProperties().getProperty("REPORT_JOB");
  return p ? JSON.parse(p) : null;
}

function clearProgress() {
  PropertiesService.getScriptProperties().deleteProperty("REPORT_JOB");
}
/**
 * 將工作表依照名稱反向排序 (選用)
 */
function sortSheetsDesc(ss) {
  var sheets = ss.getSheets();
  sheets.sort(function (a, b) {
    return b.getName().localeCompare(a.getName());
  });
  for (var i = 0; i < sheets.length; i++) {
    ss.setActiveSheet(sheets[i]);
    ss.moveActiveSheet(i + 1);
  }
}

/**
 * 將試算表分頁依名稱升冪排序 (A -> Z, 0 -> 9)
 */
function sortSheetsAsc(ss) {
  var sheets = ss.getSheets();

  // 關鍵修改：使用 a 對比 b 達成升冪
  sheets.sort(function (a, b) {
    return a.getName().localeCompare(b.getName());
  });

  // 重新排列分頁位置
  for (var i = 0; i < sheets.length; i++) {
    ss.setActiveSheet(sheets[i]);
    ss.moveActiveSheet(i + 1);
  }
}

/**
 * 尋找欄位索引（不分大小寫與底線）
 */
function getColIndex(headers, name) {
  var idx = headers.indexOf(name);
  if (idx !== -1) return idx;
  idx = headers.indexOf(name.toUpperCase());
  if (idx !== -1) return idx;
  idx = headers.indexOf(name.toLowerCase());
  if (idx !== -1) return idx;
  // 特別處理 Cust_N 與 CUST_N 等可能的情形
  for (var i = 0; i < headers.length; i++) {
    if (
      headers[i].toString().replace("_", "").toUpperCase() ===
      name.replace("_", "").toUpperCase()
    ) {
      return i;
    }
  }
  return -1;
}

/**
 * 從網址中提取試算表 ID
 * @param {string} url 試算表的網址
 * @return {string} 試算表 ID
 */
function getIdFromUrl(url) {
  var match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 *轉換 LTC_Code 欄位的資料格式，從 Cust_LTC_Code 讀取原始資料，處理後存回 LTC_Code
  - 原始資料可能包含多個 LTC Code，以逗號分隔，且可能有空白或不規則格式
  - 處理步驟：
    1. 分割字串成陣列
    2. 去除每個項目的空白
    3. 提取符合規則的部分（英文字母、數字、連字號）
    4. 排序
    5. 去重（選用）
    6. 結合成字串並存回 LTC_Code 欄位
 * @returns 
 */

function processLTCCodes() {
  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sCust = MainSpreadsheet.getSheetByName("Cust");
  const data = sCust.getDataRange().getValues();
  const headers = data[0];

  // 找出目標欄位的索引 (Index)
  const srcIdx = headers.indexOf("Cust_LTC_Code");
  const tarIdx = headers.indexOf("LTC_Code");

  if (srcIdx === -1 || tarIdx === -1) {
    SpreadsheetApp.getUi().alert("找不到指定的欄位名稱！");
    return;
  }

  // 從第二列開始處理資料
  // const updates = [];
  for (let i = 1; i < data.length; i++) {
    let rawStr = data[i][srcIdx] ? data[i][srcIdx].toString() : "";
    if (rawStr) {
      let items = rawStr.split(",");

      let processedArr = items
        .map((item) => {
          // 1. 去除空白
          let trimmed = item.trim();
          // 2. 提取開頭符合 [英、數、-] 的部分
          // ^[a-zA-Z0-9-]+ 代表從頭開始匹配多個符合的字元
          let match = trimmed.match(/^[a-zA-Z0-9-]+/);
          return match ? match[0] : null;
        })
        .filter((val) => val !== null); // 移除不符合規則的項目

      // 3. 排序 (Array.sort() 預設為字母/數字排序)
      processedArr.sort();

      // 4. 去重 (選用，避免重複項目)
      processedArr = [...new Set(processedArr)];
      // 5. 結合並存入 LTC_Code 欄位
      sCust.getRange(i + 1, tarIdx + 1).setValue(processedArr.join(","));
    }
  }
}

function UpdateUserName() {
  const now = new Date();
  now.setMonth(now.getMonth() - 1);
  const timeZone = Session.getScriptTimeZone();
  const yyyy = Utilities.formatDate(now, timeZone, "yyyy");
  const yyyyMM = Utilities.formatDate(now, timeZone, "yyyyMM");

  const srcSpreadsheet = getTargetsheet("RecUrl", "SY" + yyyy).Spreadsheet;
  const syyyyMM = srcSpreadsheet.getSheetByName(yyyyMM);
  const SYTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
  const tempSheet = SYTemp.getSheetByName("SR_Data");

  if (!syyyyMM || !tempSheet) return;

  // 1. 聚合來源資料：Map<User_N, Set<Cust_N>>
  const userMap = new Map();
  const processData = (sheet) => {
    const values = sheet.getDataRange().getValues();
    const headers = values.shift();
    const uIdx = getColIndex(headers, "USER_N");
    const cIdx = getColIndex(headers, "CUST_N");
    values.forEach(row => {
      const u = String(row[uIdx] || "").trim();
      const c = String(row[cIdx] || "").trim();
      if (u && c) {
        if (!userMap.has(u)) userMap.set(u, new Set());
        userMap.get(u).add(c);
      }
    });
  };
  processData(syyyyMM);
  processData(tempSheet);

  // 2. 處理目標工作表 (SYCompany > User)
  const tarSheet = MainSpreadsheet.getSheetByName("User");
  const tarData = tarSheet.getDataRange().getValues();
  const tarHeaders = tarData[0];

  // 動態獲取目標表的欄位索引
  const tarUserIdx = getColIndex(tarHeaders, "User_N");
  const tarCustIdx = getColIndex(tarHeaders, "Cust_N");

  if (tarUserIdx === -1 || tarCustIdx === -1) {
    throw new Error("找不到目標欄位 User_N 或 Cust_N，請檢查標題名稱是否完全一致");
  }

  // 3. 準備更新後的資料陣列 (保留原始結構，僅修改 Cust_N 欄位)
  // 我們只處理從第 2 列開始的資料
  const rowsToUpdate = tarData.slice(1);
  const processedUsers = new Set();

  const updatedRows = rowsToUpdate.map(row => {
    const userName = String(row[tarUserIdx] || "").trim();
    if (userMap.has(userName)) {
      // 找到匹配的 User，更新其 Cust_N
      const custSet = userMap.get(userName);
      row[tarCustIdx] = Array.from(custSet).sort().join(",");
      processedUsers.add(userName); // 記錄已更新的 User
    }
    return row;
  });

  // 4. (選填) 處理「目標表原本不存在」的新 User
  userMap.forEach((custSet, userName) => {
    if (!processedUsers.has(userName)) {
      const newRow = new Array(tarHeaders.length).fill("");
      newRow[tarUserIdx] = userName;
      newRow[tarCustIdx] = Array.from(custSet).sort().join(",");
      updatedRows.push(newRow);
    }
  });

  // 5. 將所有資料排序並寫回
  updatedRows.sort((a, b) => String(a[tarUserIdx]).localeCompare(String(b[tarUserIdx])));

  // 寫回目標區域 (從 A2 開始，涵蓋整張表的寬度)
  tarSheet.getRange(2, 1, updatedRows.length, tarHeaders.length).setValues(updatedRows);

  console.log("資料更新完成！");
}

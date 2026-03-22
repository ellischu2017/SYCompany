/**
 * Utilities.gs - 工具函式模組
 * 提供通用的工具函式
 */

// 注意：這個檔案主要放置與試算表操作、URL 解析、資料處理等相關的工具函式，與業務邏輯無關的部分都可以放在這裡，以保持程式碼的模組化和可維護性。
// 例如：getTarget、getIdFromUrl、removeSRDuplicates、processLTCCodes 等函式都適合放在這裡。
// 這樣的結構也方便未來如果需要拆分成多個檔案（如 Utilities.js、SpreadsheetUtils.js、DriveUtils.js 等）時，可以更清晰地管理不同類型的工具函式。
// 注意：這裡的函式應該盡量保持純粹的工具性質，不應該直接操作 UI 或特定業務邏輯，這樣才能在不同的情境下重複使用。
// 注意：如果有需要與前端 HTML 模板互動的函式（如 includeFooter、includeNav），也可以放在這裡，但要確保它們的職責僅限於生成 HTML 內容，不應該包含過多的業務邏輯。

// 目錄結構：
// SYCompany : 主目錄
// ├── LTCRecord : 主要電子表單
// |   └── SYyyyy : 每年相關的電子記錄試算表
// ├── RPyyyy : 每年相關的報表試算表
// |   ├── Pdf : 每月客戶報表Pdf檔。
// |   |   └── PDyyyymm_custn：每月客戶報表Pdf檔。
// |   ├── PDyyyymm : 每月報表總表Pdf檔。
// |   └── RPyyyyMM : 每月報表相關的試算表
// ├── Template : 存放報表模板的資料夾，包含 RPSample、SYSample 等模板試算表
// ├── SYCompany.gxlsx : 這個腳本綁定的試算表，包含 SYTemp、User、Cust 等工作表
// ├── SYTemp : 存放臨時資料的工作表，包含 SYSuggest 等設定
//  * 此檔案匯入所有模組化的 .gs 檔案
/* 
 * 模組結構：
 * - Utilities.gs: 共用工具函式
 * - Auth.gs: 認證和權限檢查
 * - Cust.gs: 個案管理
 * - User.gs: 居服員管理
 * - Manager.gs: 管理員管理
 * - LtcCode.gs: 服務編碼管理
 * - RecUrl.gs: 網址管理
 * - SRServer.gs: 服務紀錄管理
 * - Maintenance.gs: 維護和同步任務
*/

// 全域試算表參考 SYCompany
const MainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

// 全域執行時間限制 (分鐘)，比 Apps Script 的 6 分鐘限制少一點以策安全
const EXECUTION_TIMEOUT_MINUTES = 5;


/**
 * 從試算表動態取得意見反應連結並生成 Footer HTML
 */
function includeFooter() {
  let suggestUrl = "";
  const cache = CacheService.getScriptCache();
  const cacheKey = "SYSuggestUrl";

  // 1. 嘗試從快取讀取，如果有就直接使用
  suggestUrl = cache.get(cacheKey);

  if (!suggestUrl) {
    try {
      const sheet = MainSpreadsheet.getSheetByName("SYTemp");
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length > 1) {
          const headers = data[0];
          // 嘗試偵測欄位，若無則使用預設索引 0, 1
          let idxKey = getColIndex(headers, "Setting_Name");
          if (idxKey === -1) idxKey = 0;
          let idxVal = getColIndex(headers, "Setting_Value");
          if (idxVal === -1) idxVal = 1;

          for (let i = 1; i < data.length; i++) {
            if (data[i][idxKey] === "SYSuggest") {
              suggestUrl = data[i][idxVal];
              // 2. 寫入快取，設定 6 小時 (21600秒) 後過期
              cache.put(cacheKey, suggestUrl, 21600);
              break;
            }
          }
        }
      }
    } catch (e) {
      console.log("取得意見反應網址失敗: " + e.message);
    }
  }

  const template = HtmlService.createTemplateFromFile("Footer");
  template.suggestUrl = suggestUrl;
  return template.evaluate().getContent();
}

/**
 * 提供給 HTML 範本呼叫，用來載入導航列組件
 */
function includeNav() {
  // 加入快取機制，避免每次都重新讀取檔案
  var cache = CacheService.getScriptCache();
  var cachedNav = cache.get("NavHTML_v2");
  if (cachedNav) return cachedNav;

  var template = HtmlService.createTemplateFromFile("Nav");
  var content = template.evaluate().getContent();
  cache.put("NavHTML_v2", content, 21600); // 快取 6 小時
  return content;
}

/**
 * 取得當前 Web App 的 URL
 * 用於前端按鈕跳轉
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * 輔助函式：自動建立報表資料夾結構
 * 規則：
 * 1. 若 targetName 為 RPyyyy (長度6)，父資料夾為 SYCompany
 * 2. 若 targetName 為 RPyyyyMM (長度>6)，父資料夾為 RPyyyy
 * @param {string} targetName 目標資料夾名稱 (e.g., "RP2024" or "RP202401")
 * @returns {string} 新建立資料夾的 URL
 */
function createReportFolder(targetName) {
  const sheet = MainSpreadsheet.getSheetByName("FolderUrl");
  const year = targetName.substring(2, 6);
  const upfolderName = "RP" + year;
  let upfolder;

  if (targetName.length === 6) {
    // 建立年份資料夾 (e.g., RP2024)，父資料夾為 SYCompany
    upfolder = getTargetDir("FolderUrl", "SYCompany").folder;
  } else {
    // 建立月份資料夾 (e.g., RP202401)，父資料夾為 RP2024
    // 注意：這裡遞迴呼叫 getTargetDir，若 RP2024 不存在會自動建立
    upfolder = getTargetDir("FolderUrl", upfolderName).folder;
  }

  const folder = upfolder.createFolder(targetName);
  const res = folder.getUrl();
  sheet.appendRow([targetName, res]);

  return res;
}

/**
* 輔助函式：從 SYCompany (本腳本綁定之試算表) 的 sheetName 工作表取得外部試算表物
 * @param {*} sheetName SYCompany 中的工作表名稱 
 * @param {*} targetName 試算表名稱（如 RecUrl 中的 SY202401）
 * @returns {Object} 包含 url、id、Spreadsheet 物件的物件
 * @throws {Error} 如果找不到對應的工作表或試算表，會丟出錯誤
 */
function getTargetDir(sheetName, targetName) {
  // console.log("嘗試取得目標資料夾，sheetName: " + sheetName + ", targetName: " + targetName);
  var res = getTarget(sheetName, targetName);

  if (!res && targetName.substring(0, 2) === "RP") {
    res = createReportFolder(targetName);
  }

  var folderId = getIdFromUrl(res);

  // 2. 檢查解析出來的 ID 是否有效
  if (!folderId || typeof folderId !== 'string') {
    throw new Error("無法從 URL 解析出有效的 ID: " + res);
  }

  try {
    // 3. 建議直接用 getFolderById，除非你確定 Res 是檔案 ID
    var folder = DriveApp.getFolderById(folderId);

    return {
      url: res,
      id: folderId,
      folder: folder,
      folderName: folder.getName() // 順便測試是否真的抓到了
    };
  } catch (e) {
    throw new Error("DriveApp 找不到該 ID 的資料夾，請檢查權限或 ID 是否正確。錯誤訊息: " + e.message);
  }
}

/**
 * 輔助函式：從模板複製並建立新的試算表，並在註冊表中登記
 * @param {string} registrySheetName 註冊表名稱 (e.g., "ReportsUrl")
 * @param {string} newFileName 新檔案的名稱 (e.g., "RP202401")
 * @param {string} templateName 模板檔案在註冊表中的名稱 (e.g., "RPSample")
 * @param {Folder} destinationFolder 目標資料夾物件
 * @returns {string} 新建立檔案的 URL
 */
function createSheetFromTemplate(registrySheetName, newFileName, templateName, destinationFolder) {
  // 1. 取得模板檔案
  const templateFile = DriveApp.getFileById(getTargetsheet(registrySheetName, templateName).id);

  // 2. 複製模板
  const newFile = templateFile.makeCopy(newFileName, destinationFolder);
  const newUrl = newFile.getUrl();

  // 3. 在註冊表中登記新檔案
  const registrySheet = MainSpreadsheet.getSheetByName(registrySheetName);
  registrySheet.appendRow([newFileName, newUrl]);

  return newUrl;
}

function getTargetsheet(sheetName, targetName) {
  var res = getTarget(sheetName, targetName);

  if (!res) {
    if (sheetName === "ReportsUrl") {
      const destinationFolder = getTargetDir("FolderUrl", targetName).folder;
      res = createSheetFromTemplate(sheetName, targetName, "RPSample", destinationFolder);
    } else if (sheetName === "RecUrl") {
      const destinationFolder = getTargetDir("FolderUrl", "LTCRecord").folder;
      res = createSheetFromTemplate(sheetName, targetName, "SYSample", destinationFolder);
    }
  }

  if (!res) {
    console.log(
      "無法在" + sheetName + "工作表中找到名稱為" + targetName + "的對應網址",
    );
    return;
  }

  var fileId = getIdFromUrl(res);
  if (!fileId || typeof fileId !== 'string') {
    throw new Error("無法從 URL 解析出有效的 ID: " + res);
  }
  var spreadsheet = SpreadsheetApp.openById(fileId);
  return {
    url: res,
    id: fileId,
    Spreadsheet: spreadsheet
  };
}


/**
 * 輔助函式：從 SYCompany (本腳本綁定之試算表) 的 sheetName 工作表取得外部試算表物件
 * @param {string} sheetName SYCompany 中的工作表名稱
 * @param {string} targetName 試算表名稱（如 RecUrl 中的 SY202401）
 * @returns {string} 試算表的 URL
 * @throws {Error} 如果找不到對應的工作表或試算表，會丟出錯誤
 */
function getTarget(sheetName, targetName) {
  var sheet = MainSpreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error("找不到 SYCompany 中的" + sheetName + "工作表");

  var data = sheet.getDataRange().getValues();
  var url = "";
  if (data.length < 2) return url;

  var headers = data[0];
  // 嘗試偵測欄位 (支援 RecUrl 的 SY_N/SY_Url 或通用 Name/Url)
  var idxKey = getColIndex(headers, "SY_N");
  if (idxKey === -1) idxKey = getColIndex(headers, "Name");
  if (idxKey === -1) idxKey = 0; // 防呆預設

  var idxUrl = getColIndex(headers, "SY_Url");
  if (idxUrl === -1) idxUrl = getColIndex(headers, "Url");
  if (idxUrl === -1) idxUrl = 1; // 防呆預設

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxKey] === targetName) {
      url = data[i][idxUrl];
      break;
    }
  }
  // console.log("取得試算表網址: " + url);
  return url;
}

function setTargetUrl(sheetName, targetName, url) {
  var sheet = MainSpreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error("找不到 SYCompany 中的" + sheetName + "工作表");
  var data = sheet.getDataRange().getValues();// 讀取整個資料範圍，包含標題列
  var headers = data[0];

  // 動態偵測欄位
  var idxKey = getColIndex(headers, "SY_N");
  if (idxKey === -1) idxKey = getColIndex(headers, "Name");
  if (idxKey === -1) idxKey = 0;

  var idxUrl = getColIndex(headers, "SY_Url");
  if (idxUrl === -1) idxUrl = getColIndex(headers, "Url");
  if (idxUrl === -1) idxUrl = 1;

  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][idxKey] === targetName) {
      data[i][idxUrl] = url;
      found = true;
      break;
    }
  }
  // 如果沒找到，就新增一列
  if (!found) {
    // 確保新增的列長度與標題一致，避免 jagged array 問題
    var newRow = new Array(headers.length).fill("");
    newRow[idxKey] = targetName;
    newRow[idxUrl] = url;
    data.push(newRow);
  }
  // 寫回整個資料範圍，包含標題列
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  // 以 Key 欄排序
  sheet.getRange(2, 1, data.length - 1, data[0].length).sort({ column: idxKey + 1, ascending: true });
}


/**
 * 清除 SR_Data 工作表中的重複資料，根據 Date、SRTimes、CUST_N、USER_N、SR_ID 這幾個欄位的組合來判斷是否重複
 * @param {*} sheet 
 * @returns 
 */
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
  const indices = targetFields.map((field) => getColIndex(headers, field));

  // 檢查是否所有欄位都存在
  if (indices.includes(-1)) {
    const missing = targetFields.filter((_, i) => indices[i] === -1);
    console.error("removeSRDuplicates 錯誤: 找不到欄位 " + missing.join(", "));
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
 
/** 
 * 檢查執行時間是否接近超時限制
 */
function isNearTimeout() {
  return new Date().getTime() - startTime > EXECUTION_TIMEOUT_MINUTES * 60 * 1000;
}

/** 儲存/讀取進度 (PropertiesService 會存在雲端專案屬性中)
 * @param {Object} data 進度物件，會被序列化成 JSON 字串存儲
 */
function saveProgress(propname, data) {
  PropertiesService.getScriptProperties().setProperty(
    propname,
    JSON.stringify(data),
  );
}

/**
 * 讀取進度
 * @ returns {Object|null} 進度物件，如果沒有則回傳 null * 
 */
function getProgress(propname) {
  var p = PropertiesService.getScriptProperties().getProperty(propname);
  return p ? JSON.parse(p) : null;
}

/** 移除進度 */
function clearProgress(propname) {
  PropertiesService.getScriptProperties().deleteProperty(propname);
}

/**
 * 將工作表依照名稱反向排序 (選用)
 * @param {*} ss spreadsheet 物件
 */
function sortSheetsDesc(ss) {
  var sheets = ss.getSheets();
  sheets.sort(function (a, b) {
    return b.getName().localeCompare(a.getName(), undefined, { numeric: true });
  });
  for (var i = 0; i < sheets.length; i++) {
    ss.setActiveSheet(sheets[i]);
    ss.moveActiveSheet(i + 1);
  }
}

/**
 * 將試算表分頁依名稱升冪排序 (A -> Z, 0 -> 9)
 * @param {*} ss spreadsheet 物件 
 */
function sortSheetsAsc(ss) {
  var sheets = ss.getSheets();

  // 關鍵修改：使用 a 對比 b 達成升冪
  sheets.sort(function (a, b) {
    return a.getName().localeCompare(a.getName(), undefined, { numeric: true });
  });

  // 重新排列分頁位置
  for (var i = 0; i < sheets.length; i++) {
    ss.setActiveSheet(sheets[i]);
    ss.moveActiveSheet(i + 1);
  }
}

/**
 * 尋找欄位索引（不分大小寫與底線）
 * @param {Array} headers 欄位名稱陣列
 * @param {string} name 欄位名稱
 * @return {number} 欄位索引，找不到則回傳 -1 
 */
function getColIndex(headers, name) {
  // 1. 優先進行最快、最精確的比對 (大小寫完全相符)
  const exactIndex = headers.indexOf(name);
  if (exactIndex !== -1) {
    return exactIndex;
  }

  // 2. 若找不到，則進行標準化比對 (忽略大小寫、空格、底線)
  const cleanName = name.toString().replace(/[_ ]/g, "").toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    // 確保 header[i] 是字串，避免因空儲存格導致錯誤
    const cleanHeader = (headers[i] || "").toString().replace(/[_ ]/g, "").toLowerCase();
    if (cleanHeader === cleanName) {
      return i;
    }
  }

  return -1;
}

/**
 * 批次取得欄位索引映射表
 * @param {Array} headers 標題列陣列
 * @param {Array} targetFields 目標欄位名稱陣列
 * @returns {Object} { "Date": 0, "Name": 1, ... }
 */
function getColIndicesMap(headers, targetFields) {
  var map = {};
  targetFields.forEach(function(field) {
    map[field] = getColIndex(headers, field);
  });
  return map;
}

/**
 * 根據欄位索引映射表將資料列標準化 (擷取指定欄位)
 * @param {Array} row 原始資料列
 * @param {Object} colIndicesMap getColIndicesMap 回傳的索引物件
 * @param {Array} targetFields 目標欄位順序
 * @returns {Array} 標準化後的資料陣列
 */
function normalizeRow(row, colIndicesMap, targetFields) {
  return targetFields.map(function(field) {
    var idx = colIndicesMap[field];
    return (idx !== -1 && row[idx] !== undefined) ? row[idx] : "";
  });
}

/**
 * 從網址中提取試算表 ID
 * @param {string} url 試算表的網址
 * @return {string} 試算表 ID
 */
function getIdFromUrl(url) {
  if (!url || typeof url !== 'string') return "";

  // 優先匹配 /d/ID, /folders/ID, 或 id=ID 的格式
  const match = url.match(/(?:d|folders)\/([-\w]{25,})|id=([-\w]{25,})/);

  if (match) {
    // 回傳第一個或第二個捕獲組的內容 (match[1] 或 match[2])
    return match[1] || match[2] || "";
  }
  return "";
}

/**
 * 格式化 LTC Code 字串
 * 處理步驟：分割、清理、排序、去重、合併
 * @param {string} rawStr - 原始的 LTC Code 字串 (e.g., " BA01, BA02a-1, 無效碼, BA01 ")
 * @returns {string} - 格式化後的字串 (e.g., "BA01,BA02a-1")
 */
function formatLtcCodeString(rawStr) {
  if (!rawStr || typeof rawStr !== 'string') return "";

  const items = rawStr.split(",");

  const processedArr = items
    .map((item) => {
      // 1. 去除空白並提取有效部分
      const match = item.trim().match(/^[a-zA-Z0-9-]+/);
      return match ? match[0] : null;
    })
    .filter((val) => val !== null); // 2. 移除不符合規則的項目

  // 3. 去重並排序
  const uniqueSortedArr = [...new Set(processedArr)].sort();

  // 4. 結合
  return uniqueSortedArr.join(",");
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

function processCustLTCCodes() {
  // const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sCust = MainSpreadsheet.getSheetByName("Cust");
  const data = sCust.getDataRange().getValues();
  const headers = data[0];

  // 找出目標欄位的索引 (Index)
  const srcIdx = getColIndex(headers, "Cust_LTC_Code");
  const tarIdx = getColIndex(headers, "LTC_Code");

  if (srcIdx === -1 || tarIdx === -1) {
    SpreadsheetApp.getUi().alert("找不到指定的欄位名稱！");
    return;
  }

  // 從第二列開始處理資料
  // const updates = [];
  for (let i = 1; i < data.length; i++) {
    let rawStr = data[i][srcIdx] ? data[i][srcIdx].toString() : "";
    // 只有在原始字串包含有效內容時才處理
    if (rawStr.trim()) {
      const formattedCodes = formatLtcCodeString(rawStr);
      sCust.getRange(i + 1, tarIdx + 1).setValue(formattedCodes);
    }
  }
  CacheService.getScriptCache().remove("SRServer01_InitData");
}

/**
 * 
 * @returns 
 */

function UpdateUserCustName() {
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
    const colMap = getColIndicesMap(headers, ["USER_N", "CUST_N"]);
    const uIdx = colMap["USER_N"];
    const cIdx = colMap["CUST_N"];
    values.forEach(row => {
      const u = (uIdx !== -1) ? String(row[uIdx] || "").trim() : "";
      const c = (cIdx !== -1) ? String(row[cIdx] || "").trim() : "";
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
  const tarColMap = getColIndicesMap(tarHeaders, ["User_N", "Cust_N"]);
  const tarUserIdx = tarColMap["User_N"];
  const tarCustIdx = tarColMap["Cust_N"];

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
  CacheService.getScriptCache().remove("SRServer01_InitData");
}

/**
 * 格式化日期物件或字串
 * @param {Date|string} date - 日期物件或可被 new Date() 解析的字串
 * @param {string} [format='yyyy-MM-dd'] - (選填) 日期格式，預設為 'yyyy-MM-dd'
 * @param {string} [timezone='Asia/Taipei'] - (選填) 時區，預設為 'Asia/Taipei'
 * @returns {string} 格式化後的日期字串，若輸入無效則回傳原始字串
 */
function formatDate(date, format = 'yyyy-MM-dd', timezone = 'Asia/Taipei') {
  if (!date) return "";
  try {
    const d = new Date(date);
    // 檢查日期是否有效
    if (isNaN(d.getTime())) {
      return String(date); // 如果是無效日期，直接回傳原始字串
    }
    return Utilities.formatDate(d, timezone, format);
  } catch (e) {
    // 發生錯誤時，回傳原始字串以利除錯
    return String(date);
  }
}

/**
 * 手動清除所有系統快取
 * 包含導航列、下拉選單快取、以及各模組的資料快取
 */
function clearAllCaches() {
  var userEmail = Session.getActiveUser().getEmail();
  console.log(`[System] 開始手動清除快取... 觸發者: ${userEmail || 'Unknown'}，時間: ${new Date().toISOString()}`);

  const cache = CacheService.getScriptCache();
  const keysToRemove = [
    "NavHTML_v2",
    "SYSuggestUrl",
    "ManagerData",
    "ManaList",
    "UserData",
    "CustInfoMap",
    "CustN_All",
    "SRServer01_InitData"
  ];
  cache.removeAll(keysToRemove);
  console.log(`[System] 快取清除完成。已移除以下鍵值: ${keysToRemove.join(', ')}`);
}

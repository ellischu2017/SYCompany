/**
 * 報表欄位標題與資料物件屬性的對應表
 * key: 報表模板中的欄位標題 (e.g., "日期")
 * value: groupDataInLocalMemory 產生的 dataObjs 物件中的屬性名稱 (e.g., "date")
 */
const REPORT_COLUMN_MAPPING = {
  "日期": "date",
  "服務次數": "srTimes",
  "居服員": "userN",
  "意識狀況": "loc",
  "身心狀況": "mood",
  "特殊狀況": "spcons",
  "服務內容": "srData",
};

/**
 * PDF 匯出參數設定
 * ref: https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
 */
const PDF_EXPORT_OPTIONS = {
  format: "pdf",
  size: "A4",               // 紙張大小
  portrait: "false",      // false = 橫向, true = 直向
  fitw: "true",           // true = 寬度符合頁面
  gridlines: "false",     // false = 不顯示格線
  fzr: "true",            // true = 重複凍結列 (作為頁首)
  top_margin: "0.5",      // 邊界 (英吋)
  bottom_margin: "0.5",
  left_margin: "0.5",
  right_margin: "0.5",
  pagenumbers: "false",   // false = 不顯示頁碼
  printtitle: "false",    // false = 不顯示標題列
  pagenum: "CENTER",      // 頁碼位置
  sheetnames: "false",    // false = 不顯示工作表名稱
};

/**
 * 獲取個案清單供前端下拉選單使用
 */
function getCustN(yearmonth) {
  // 加入快取機制，避免重複讀取試算表
  var cache = CacheService.getScriptCache();
  var cacheKey = "CustN_" + (yearmonth || "All");
  var cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  var custNames = [];

  if (yearmonth) {
    var allDataMap = getAllDataMapFromSources(yearmonth);
    custNames = Object.keys(allDataMap);
    custNames.sort(function (a, b) {
      return a.localeCompare(b);
    });
  } else {
    // 直接複用 getCustInfoMap 的邏輯，它會從 Cust 和 OldCust 讀取並快取
    var custInfo = getCustInfoMap();
    custNames = custInfo.names.sort((a, b) => a.localeCompare(b));
  }
  
  // 寫入快取，設定 20 分鐘 (1200秒)
  try {
    cache.put(cacheKey, JSON.stringify(custNames), 1200);
  } catch (e) {
    Logger.log("快取失敗 (資料過大): " + e.toString());
  }
  return custNames;
}

/**
 * 主進入點：處理單一或全部個案
 */
function genreport(yearmonth, custn, regen) {
  var year = yearmonth.substring(0, 4);
  var month = yearmonth.substring(4, 6);

  // 0. 定義固定的欄位索引對映表 (colMap)
  // 對應 getAllDataMapFromSources 中標準化後的資料順序
  // ["Date", "SRTimes", "User_N", "LOC", "MOOD", "SPCONS", "SR_ID", "SR_REC", "Cust_N"]
  var colMap = {
    Date: 0,
    SRTimes: 1,
    User_N: 2,
    LOC: 3,
    MOOD: 4,
    SPCONS: 5,
    SR_ID: 6,
    SR_REC: 7
    // Cust_N: 8 (報表分組已用，此處僅需紀錄相關欄位)
  };

  // --- 如果是單一個案處理 ---
  if (custn !== "all") {
    // 1. 準備單一案主所需的資料 (跟批次邏輯一樣，先抓取必要的 Map)
    var allDataMap = getAllDataMapFromSources(yearmonth);
    var custBase = getCustInfoMap();
    var ssReportFile = getTargetsheet("ReportsUrl", "RP" + yearmonth).Spreadsheet;
    var templateSheet =
      ssReportFile.getSheetByName("Template") || ssReportFile.getSheets()[0];

    // 2. 檢查該案主是否有資料
    if (!allDataMap[custn])
      return {
        status: "complete",
        message: "該個案此月份無紀錄。",
        btntext: "返回",
        currentIndex: 1,
        total: 1,
      };

    // 3. 呼叫正確的函式名稱：processSingleReport
    var url = processSingleReport(
      ssReportFile,
      templateSheet,
      custn,
      allDataMap[custn],
      custBase.info[custn],
      year,
      month,
      regen,
      colMap // 傳入欄位索引
    );

    return {
      status: "complete",
      currentIndex: 1,
      total: 1,
      message: "個案 " + custn + " 報表處理完成！",
      btntext: "查看報表",
      url: url,
    };
  }

  // --- 批次處理邏輯 (支援續傳) ---
  var progress = getProgress("REPORT_JOB");
  var currentIndex = 0;
  var allCusts = [];

  // 如果是第一次執行或月份變了，就初始化
  if (!progress || progress.yearmonth !== yearmonth) {
    allCusts = getCustN(yearmonth);
    currentIndex = 0;
  } else {
    allCusts = progress.allCusts;
    currentIndex = progress.currentIndex;
  }

  // 預讀資料 (優化效能)
  // var year = yearmonth.substring(0, 4);
  // var month = yearmonth.substring(4, 6);
  var allDataMap = getAllDataMapFromSources(yearmonth);
  var custBase = getCustInfoMap();
  var ssReportFile = getTargetsheet("ReportsUrl", "RP" + yearmonth).Spreadsheet;

  var templateSheet =
    ssReportFile.getSheetByName("Template") || ssReportFile.getSheets()[0];

  for (var i = currentIndex; i < allCusts.length; i++) {
    // 檢查是否快要超時
    if (isNearTimeout()) {
      saveProgress("REPORT_JOB", {
        yearmonth: yearmonth,
        allCusts: allCusts,
        currentIndex: i,
      });
      return {
        status: "continue",
        message:
          "已處理 " + i + " / " + allCusts.length + " 個個案，正在續傳...",
        currentIndex: i,
        total: allCusts.length,
      };
    }

    var name = allCusts[i];
    if (allDataMap[name]) {
      processSingleReport(
        ssReportFile,
        templateSheet,
        name,
        allDataMap[name],
        custBase.info[name],
        year,
        month,
        regen,
        colMap // 傳入欄位索引
      );
    }
  }

  // 全部完成
  clearProgress("REPORT_JOB");
  sortSheetsDesc(ssReportFile); // 選用：完成後將工作表依名稱反向排序
  return {
    status: "complete",
    message: "全部個案處理完成！",
    btntext: "查看報表",
    url: ssReportFile.getUrl(),
  };
}

/**
 * 核心邏輯：生成單一案主報表
 */
function processSingleReport(
  ssFile,
  template,
  custn,
  rawRows,
  info,
  year,
  month,
  regen,
  colMap // 新增參數
) {
  var sheet = ssFile.getSheetByName(custn);

  // 檢查是否跳過
  if (sheet && sheet.getLastRow() >= 7 && !regen) {
    return ssFile.getUrl() + "#gid=" + sheet.getSheetId();
  }

  // 初始化或清空工作表
  if (!sheet) {
    sheet = template.copyTo(ssFile).setName(custn);
  } else {
    var lastR = sheet.getLastRow();
    if (lastR >= 7) {
      // 清除舊資料，從第 7 行開始，範圍涵蓋所有欄
      sheet
        .getRange(7, 1, lastR - 6, sheet.getLastColumn())
        .clearContent()
        .setBorder(false, false, false, false, false, false);
    }
  }

  // 寫入 Header (一次寫入一個 Range 效率較高)
  sheet
    .getRange("A2")
    .setValue(
      parseInt(year) -
      1911 +
      "年" +
      parseInt(month) +
      "月　　居家服務照顧內容紀錄單",
    );
  sheet
    .getRange("A3:C3")
    .merge()
    .setValue("個案姓名：" + custn);
  sheet
    .getRange("D3:E3")
    .merge()
    .setValue("性別：" + (info ? info.sex : ""));
  sheet
    .getRange("F3:G3")
    .merge()
    .setValue("出生年月日：" + (info ? info.bd : ""));
  sheet
    .getRange("A4:G4")
    .merge()
    .setValue("住址：" + (info ? info.add : ""));

  // --- 在記憶體中整理資料結構 ---
  var finalValues = [];
  var dataObjs = [];
  groupDataInLocalMemory(rawRows, dataObjs, colMap);

  if (dataObjs.length > 0) {
    // 1. 假設報表資料標題在第 6 行，並讀取所有有內容的欄位
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return ssFile.getUrl() + "#gid=" + sheet.getSheetId(); // 防呆：若無任何欄位則直接返回
    const reportHeaders = sheet.getRange(6, 1, 1, lastCol).getValues()[0];

    finalValues = dataObjs.map(function (obj) {
      // 根據實際標題順序建立資料列
      return reportHeaders.map(header => {
        const cleanHeader = String(header).trim();
        const propName = REPORT_COLUMN_MAPPING[cleanHeader];
        // 如果找不到對應屬性，或物件中沒有該值，則回傳空字串
        return (propName && obj.hasOwnProperty(propName)) ? obj[propName] : "";
      });
    });

    // 批次寫入所有數據 (關鍵優化！)
    // 使用 reportHeaders.length 來確保寫入的欄數與標題欄數一致
    var writeRange = sheet.getRange(7, 1, finalValues.length, reportHeaders.length);
    writeRange
      .setValues(finalValues)
      .setFontSize(10)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("top")
      .setWrap(true)
      .setBorder(true, true, true, true, true, true);

    // 設置行高 (此動作較慢，但已大幅減少調用次數)
    for (var i = 0; i < finalValues.length; i++) {
      if (sheet.getRowHeight(7 + i) < 50) sheet.setRowHeight(7 + i, 50);
    }
  }
  sortSheetsAsc(ssFile); // 選用：每次生成後將工作表依名稱反向排序
  return ssFile.getUrl() + "#gid=" + sheet.getSheetId();
}

async function genPdfFile(yearmonth, custn, regen) {
  // var year = yearmonth.substring(0, 4);
  // var month = yearmonth.substring(4, 6);
  const ssReportFile = getTargetsheet("ReportsUrl", "RP" + yearmonth).Spreadsheet;
  // get all sheets except Template (優化效能，避免在迴圈裡重複呼叫 getSheetByName)
  var allSheets = ssReportFile.getSheets();
  var allSheetNames = [];
  for (var i = 0; i < allSheets.length; i++) {
    allSheetNames.push(allSheets[i].getName());
  }
  allSheetNames.splice(allSheetNames.indexOf("Template"), 1);
  //Logger.log("所有個案工作表名稱: " + allSheetNames.join(", ") + "，總數: " + allSheetNames.length + "個");
  if (custn !== "all") {
    // 取得該個案的Pdf Url 
    var pdfUrl = getTarget("PdfUrl", "PD" + yearmonth + "_" + custn);
    if (pdfUrl && !regen) {
      return {
        status: "complete",
        currentIndex: 1,
        total: 1,
        message: "個案 " + custn + " 報表處理完成！",
        btntext: "查看報表",
        url: pdfUrl,
      };
    } else {
      // 呼叫 genreport 生成報表，這裡可以直接呼叫 processSingleReport 以避免重複讀取資料，但為了保持邏輯清晰，我們先呼叫 genreport 
      if (!allSheetNames.includes(custn)) {
        // 代表該個案無資料，無法生成報表，因此也無法生成 PDF
        return {
          status: "complete",
          message: "該個案此月份無紀錄。",
          btntext: "返回",
          currentIndex: 1,
          total: 1,
        };
      } else {
        //2. 生成 PDF
        var pdfUrl = await processSinglePdf(yearmonth, custn, regen);
        return {
          status: "complete",
          currentIndex: 1,
          total: 1,
          message: "個案 " + custn + " 報表處理完成！",
          btntext: "查看報表",
          url: pdfUrl,
        };
      }
    }
  }

  // --- 批次處理邏輯 (支援續傳) ---
  var progress = getProgress("PDF_JOB");
  var currentIndex = 0;
  var allCusts = [];

  // 如果是第一次執行或月份變了，就初始化
  if (!progress || progress.yearmonth !== yearmonth) {

    allCusts = allSheetNames;
    currentIndex = 0;
    clearProgress("PDF_JOB");
  } else {
    allCusts = progress.allCusts;
    currentIndex = progress.currentIndex;
  }

  // 預讀資料 (優化效能)
  // var year = yearmonth.substring(0, 4);
  // var month = yearmonth.substring(4, 6);
  // var allDataMap = getAllDataMapFromSources(yearmonth, year, month);
  for (var i = currentIndex; i < allCusts.length; i++) {
    // 檢查是否快要超時
    if (isNearTimeout()) {
      saveProgress("PDF_JOB", {
        yearmonth: yearmonth,
        allCusts: allCusts,
        currentIndex: i,
      });
      return {
        status: "continue",
        message:
          "已處理 " + i + " / " + allCusts.length + " 個個案，正在續傳...",
        currentIndex: i,
        total: allCusts.length,
      };
    }

    var name = allCusts[i];
    var pdfUrl = getTarget("PdfUrl", "PD" + yearmonth + "_" + name);
    if (pdfUrl && !regen) {
      continue;
    } else {
      // 生成 PDF
      await processSinglePdf(yearmonth, name, regen);
      // 加入延遲以避免觸發 Google API 速率限制 (Too Many Requests)
      // 增加至 3 秒以減少 429 錯誤發生的機率
      Utilities.sleep(3000);
    }

  }
  //把所有個案的 PDF 都生成完後，再把 PDF 資料夾裡的檔案合併成一個 PDF (選用，視需求而定)
  var folder = getTargetDir("FolderUrl", "RP" + yearmonth).folder;
  var mergedPdfUrl = null;
  var pdfFiles = folder.getFoldersByName("PDF").hasNext() ? folder.getFoldersByName("PDF").next().getFiles() : null;
  if (pdfFiles) {
    var fileList = [];
    while (pdfFiles.hasNext()) {
      fileList.push(pdfFiles.next());
    }
    // 依檔名排序，確保合併順序正確
    fileList.sort(function (a, b) { return a.getName().localeCompare(b.getName()); });

    var pdfIds = fileList.map(function (f) { return f.getId(); });

    // 這裡可以呼叫一個合併 PDF 的函式，將 pdfUrls 中的 PDF 合併成一個檔案，並儲存到 Drive 中，最後回傳合併後 PDF 的網址
    mergedPdfUrl = await mergePdfs(pdfIds, yearmonth);
    // return { status: "complete", message: "全部個案處理完成！", btntext: "查看合併報表", url: mergedPdfUrl };
  }
  // 全部完成
  clearProgress("PDF_JOB");
  return {
    status: "complete",
    message: "全部個案處理完成！",
    btntext: "查看報表",
    url: mergedPdfUrl,
  };

}

/**
 * 使用 pdf-lib 庫將多個 PDF 檔案合併成一個 PDF，並儲存到 Drive 中
 * @param {Array} pdfIds - 要合併的 PDF 檔案 ID 陣列
 * @param {String} yearmonth - 年月字串，用於命名合併後的 PDF 檔案
 * @returns {String} 合併後 PDF 的網址
 */
async function mergePdfs(pdfIds, yearmonth) {
  // --- 1. 載入 pdf-lib 庫 (只載入一次) ---
  var setTimeout = function (f, t) { Utilities.sleep(t || 0); return f(); };
  var clearTimeout = function () { };
  eval(fetchPdfLib());

  // --- 2. 建立一個新的 PDF 文件作為合併容器 ---
  const mergedPdfDoc = await PDFLib.PDFDocument.create();

  // --- 3. 遍歷所有 PDF ID 進行合併 ---
  for (const id of pdfIds) {
    const pdfBytes = DriveApp.getFileById(id).getBlob().getBytes();
    const uint8Array = new Uint8Array(pdfBytes);

    // 載入來源 PDF
    const pdfDoc = await PDFLib.PDFDocument.load(uint8Array);

    // 複製所有頁面
    const copiedPages = await mergedPdfDoc.copyPages(pdfDoc, pdfDoc.getPageIndices());

    // 將複製的頁面加入合併文件
    copiedPages.forEach((page) => mergedPdfDoc.addPage(page));
  }

  // --- 4. 儲存合併後的 PDF ---
  const mergedPdfBytes = await mergedPdfDoc.save();

  // 上傳合併後的 PDF 到 Drive
  var fileName = "PD" + yearmonth + ".pdf";
  var blob = Utilities.newBlob(mergedPdfBytes, "application/pdf", fileName);
  var folder = getTargetDir("FolderUrl", "RP" + yearmonth).folder;

  // 檢查是否已存在同名檔案，若有則刪除舊檔 (選用)
  var existingFiles = folder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  var file = folder.createFile(blob);

  // 儲存 PDF 的網址到 Properties 以供前端使用
  setTargetUrl("PdfUrl", "PD" + yearmonth, file.getUrl());

  // 回傳檔案網址供後續使用
  return file.getUrl();
}

/**
 * 核心邏輯：生成單一案主報表
 */
async function processSinglePdf(yearmonth, custn, regen) {
  const ssReportFile = getTargetsheet("ReportsUrl", "RP" + yearmonth).Spreadsheet;
  const sheet = ssReportFile.getSheetByName(custn);
  //把工作表 匯出成暫存試算表
  const ssId = ssReportFile.getId();
  const sheetId = sheet.getSheetId();
  const token = ScriptApp.getOAuthToken();

  // --- 1. 動態建立匯出 URL ---
  const exportParams = Object.entries(PDF_EXPORT_OPTIONS)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join('&');

  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?${exportParams}&gid=${sheetId}`;


  // --- 2. 抓取 PDF 原始內容 ---
  var response;
  var attempt = 0;
  var maxAttempts = 5;

  while (attempt < maxAttempts) {
    response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 429) {
      Logger.log("PDF Export encountered 429 Too Many Requests. Retrying... Attempt " + (attempt + 1));
      Utilities.sleep(5000 * Math.pow(2, attempt)); // 指數退避: 5s, 10s, 20s, 40s, 80s
      attempt++;
    } else if (response.getResponseCode() !== 200) {
      throw new Error("PDF Export Failed (" + response.getResponseCode() + "): " + response.getContentText());
    } else {
      break; // Success
    }
  }

  if (response.getResponseCode() !== 200) {
     throw new Error("PDF Export Failed (" + response.getResponseCode() + ") after " + maxAttempts + " attempts: " + response.getContentText());
  }

  const pdfBytes = response.getContent();
  // 轉換成 Uint8Array 以供 pdf-lib 使用
  const uint8Array = new Uint8Array(pdfBytes);

  // --- 3. 載入 pdf-lib 庫 ---
  var setTimeout = function (f, t) { Utilities.sleep(t || 0); return f(); };
  var clearTimeout = function () { };
  eval(fetchPdfLib());

  // --- 4. 解析 PDF 並檢查頁數 ---
  let pdfDoc = await PDFLib.PDFDocument.load(uint8Array);
  // --- 5. 確保為偶數頁，以利雙面列印 ---
  pdfDoc = await ensureEvenPages(pdfDoc);

  // --- 6. 儲存最終 PDF ---
  const finalPdfBytes = await pdfDoc.save();
  const fileName = "PD" + yearmonth + "_" + custn + ".pdf";
  const blob = Utilities.newBlob(finalPdfBytes, "application/pdf", fileName);

  // 移動到指定資料夾 (選用)
  var folder = getTargetDir("FolderUrl", "RP" + yearmonth).folder;
  // 檢查folder 是否已經有 PDF 子資料夾
  var pdfFolder = null;
  var folders = folder.getFoldersByName("PDF");
  if (folders.hasNext()) {
    pdfFolder = folders.next();
  } else {
    pdfFolder = folder.createFolder("PDF");
  }

  // 如果 regen == true，先檢查並刪除舊檔
  if (regen) {
    var existingFiles = pdfFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }
  }

  const file = DriveApp.createFile(blob);
  pdfFolder.addFile(file);
  folder.removeFile(file); // 從根目錄移除，避免混亂
  // 儲存 PDF 的網址到 Properties 以供前端使用
  setTargetUrl("PdfUrl", "PD" + yearmonth + "_" + custn, file.getUrl());

  // 回傳檔案網址供後續使用
  return file.getUrl();
}

/**
 * 輔助函式：確保 PDF 文件為偶數頁
 * 若為奇數頁，則在末尾補上一張與最後一頁同尺寸的空白頁。
 * @param {PDFDocument} pdfDoc - 從 pdf-lib 載入的 PDF 文件物件
 * @returns {Promise<PDFDocument>} - 處理過後的 PDF 文件物件
 */
async function ensureEvenPages(pdfDoc) {
  const pageCount = pdfDoc.getPageCount();
  Logger.log("原始頁數: " + pageCount);

  if (pageCount % 2 !== 0) {
    Logger.log("偵測到奇數頁，正在加入空白頁以符合雙面列印需求...");
    const pages = pdfDoc.getPages();
    const lastPage = pages[pages.length - 1];
    const { width, height } = lastPage.getSize();
    pdfDoc.addPage([width, height]);
  } else {
    Logger.log("頁數已是偶數，無需補頁。");
  }
  return pdfDoc;
}


/**
 * 預先讀取所有資料並分類 (Map: Name -> Rows[])
 */
function getAllDataMapFromSources(yearmonth) {
  // 1. 加入快取機制，避免重複讀取 (設定 30 分鐘)
  var cache = CacheService.getScriptCache();
  var cacheKey = "DataMap_" + yearmonth;
  var cached = cache.get(cacheKey);
  if (cached) {
    // Logger.log("從快取讀取 " + cacheKey);
    return JSON.parse(cached);
  }

  var year = yearmonth.substring(0, 4);
  var month = yearmonth.substring(4, 6);
  var ssAnnual = getTargetsheet("RecUrl", "SY" + year).Spreadsheet.getSheetByName(
    yearmonth,
  );
  var ssTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet.getSheetByName("SR_Data");
  var dataMap = {};
  var excludeIDs = ["GA05", "GA09", "SY01"];
  // 定義標準化欄位順序 (與 genreport 中的 colMap 對應)
  var targetFields = ["Date", "SRTimes", "User_N", "LOC", "MOOD", "SPCONS", "SR_ID", "SR_REC", "Cust_N"];

  [ssAnnual, ssTemp].forEach(function (sheet, idx) {
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    var h = values[0];
    // 取得所有目標欄位的索引
    var colIndices = getColIndicesMap(h, targetFields);

    // 檢查關鍵欄位是否存在
    if (colIndices.Date === -1 || colIndices.Cust_N === -1) return;

    var isTemp = idx === 1;
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      
      // 建立標準化後的資料列 (Normalize)
      var normalizedRow = normalizeRow(r, colIndices, targetFields);

      var rDate = normalizedRow[0]; // Date is index 0
      var name = normalizedRow[8];  // Cust_N is index 8
      var srId = String(normalizedRow[6]); // SR_ID is index 6

      // 過濾條件
      if (excludeIDs.indexOf(srId) !== -1) continue;
      if (isTemp) {
        var d = new Date(rDate);
        if (d.getFullYear() != year || d.getMonth() + 1 != month) continue;
      }

      if (!dataMap[name]) dataMap[name] = [];
      dataMap[name].push(normalizedRow);
    }
  });

  // 2. 寫入快取 (如果資料過大則跳過)
  try {
    cache.put(cacheKey, JSON.stringify(dataMap), 1800); // 1800 秒 = 30 分鐘
  } catch (e) {
    Logger.log("快取失敗 (資料過大): " + e.toString());
    // 資料過大時直接返回，不快取
  }
  return dataMap;
}

/**
 * 預先讀取所有個案基本資料
 */
function getCustInfoMap() {
  // 1. 加入快取機制
  var cache = CacheService.getScriptCache();
  var cacheKey = "CustInfoMap";
  var cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  var infoMap = {};
  var custNames = [];
  const sheetNames = ["Cust", "OldCust"];
  const targetFields = ["Cust_N", "Cust_Sex", "Cust_BD", "Cust_Add"];

  sheetNames.forEach(function(sheetName) {
    const sheet = MainSpreadsheet.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return; // 沒有資料列

    const headers = data[0];
    const colMap = getColIndicesMap(headers, targetFields);

    // 確保關鍵欄位存在
    if (colMap["Cust_N"] === -1) {
      console.warn(`[getCustInfoMap] 工作表 ${sheetName} 缺少 'Cust_N' 欄位，略過處理。`);
      return;
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[colMap["Cust_N"]];
      // 如果 name 存在且尚未加入 map，才進行處理
      if (name && !infoMap[name]) {
        custNames.push(name);
        infoMap[name] = {
          sex: colMap["Cust_Sex"] !== -1 ? row[colMap["Cust_Sex"]] : "",
          bd: colMap["Cust_BD"] !== -1 ? formatDate(row[colMap["Cust_BD"]], "yyyy/MM/dd") : "",
          add: colMap["Cust_Add"] !== -1 ? row[colMap["Cust_Add"]] : "",
        };
      }
    }
  });

  var result = { names: custNames, info: infoMap };

  // 2. 寫入快取 (設定 6 小時, 如果資料過大則跳過)
  try {
    cache.put(cacheKey, JSON.stringify(result), 21600);
  } catch (e) {
    Logger.log("快取失敗 (資料過大): " + e.toString());
  }
  return result;
}

/**
 * 日期/次數/人員分組邏輯 (維持原邏輯)
 */
function groupDataInLocalMemory(rows, allDataObjs, col) {
  // col 由外部動態傳入，避免硬編碼索引
  // 確保 Date, SRTimes, User_N, LOC, MOOD, SPCONS, SR_ID, SR_REC 存在

  // 取得唯一日期並排序
  var uniqueDates = [
    ...new Set(rows.map((r) => formatDate(r[col.Date], "yyyy/MM/dd"))),
  ].sort();

  uniqueDates.forEach(function (dStr) {
    var dateRows = rows.filter((r) => formatDate(r[col.Date], "yyyy/MM/dd") === dStr);
    var uniqueTimes = [...new Set(dateRows.map((r) => r[col.SRTimes]))].sort(
      (a, b) => a - b,
    );

    uniqueTimes.forEach(function (tVal) {
      var tRows = dateRows.filter((r) => r[col.SRTimes] === tVal);
      var uniqueUsers = [...new Set(tRows.map((r) => r[col.User_N]))];

      uniqueUsers.forEach(function (uName) {
        var uRows = tRows.filter((r) => r[col.User_N] === uName);
        var first = uRows[0];

        // 整理 SR_REC 資料
        var recItems = uRows.map(
          (r) =>
            String(r[col.SR_ID]) +
            String(r[col.SR_REC]).replace(/\n/g, "").replace(/,/g, "，"),
        );
        var srStr = "";
        if (recItems.length <= 3) {
          srStr = recItems.join("\n");
        } else {
          var lines = [];
          for (var i = 0; i < 3; i++) {
            lines.push(
              recItems[i] + (recItems[i + 3] ? "  " + recItems[i + 3] : ""),
            );
          }
          srStr = lines.join("\n");
        }

        var srStr = formatServiceRecords(uRows, col);
        allDataObjs.push({
          date: dStr,
          srTimes: tVal,
          userN: uName,
          loc: first[col.LOC],
          mood: first[col.MOOD],
          spcons: first[col.SPCONS],
          srData: srStr,
        });
      });
    });
  });
}

/**
 * 格式化服務內容字串
 * 將多筆服務紀錄 (SR_ID + SR_REC) 組合，若超過 3 筆則分欄顯示 (左右兩欄，最多6筆)
 */
function formatServiceRecords(rows, col) {
  var recItems = rows.map(
    (r) =>
      String(r[col.SR_ID]) +
      String(r[col.SR_REC]).replace(/\n/g, "").replace(/,/g, "，"),
  );

  if (recItems.length <= 3) {
    return recItems.join("\n");
  } else {
    var lines = [];
    for (var i = 0; i < 3; i++) {
      lines.push(
        recItems[i] + (recItems[i + 3] ? "  " + recItems[i + 3] : ""),
      );
    }
    return lines.join("\n");
  }
}

// --- 內部輔助函式庫 ---

var _pdfLibCache = null;
function fetchPdfLib() {
  if (!_pdfLibCache) {
    _pdfLibCache = UrlFetchApp.fetch("https://unpkg.com/pdf-lib/dist/pdf-lib.min.js").getContentText();
  }
  return _pdfLibCache;
}

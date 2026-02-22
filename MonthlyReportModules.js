/**
 * 獲取個案清單供前端下拉選單使用
 */
function getCustN() {
  // getTargetsheet 定義在 Utilities.js 中
  // var ss = getTargetsheet("SYTemp", "SYTemp");
  var sCust = MainSpreadsheet.getSheetByName("Cust");
  var data = sCust.getDataRange().getValues();
  var custNames = [];

  // 假設 Cust 工作表第一欄是 Cust_N
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) custNames.push(data[i][0]);
  }
  return custNames;
}

/**
 * 主進入點：處理單一或全部個案
 */
function genreport(yearmonth, custn, regen) {
  var year = yearmonth.substring(0, 4);
  var month = yearmonth.substring(4, 6);

  // --- 如果是單一個案處理 ---
  if (custn !== "all") {
    // 1. 準備單一案主所需的資料 (跟批次邏輯一樣，先抓取必要的 Map)
    var allDataMap = getAllDataMapFromSources(yearmonth, year, month);
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
    allCusts = getCustN();
    currentIndex = 0;
  } else {
    allCusts = progress.allCusts;
    currentIndex = progress.currentIndex;
  }

  // 預讀資料 (優化效能)
  // var year = yearmonth.substring(0, 4);
  // var month = yearmonth.substring(4, 6);
  var allDataMap = getAllDataMapFromSources(yearmonth, year, month);
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
    if (lastR >= 6) {
      sheet
        .getRange(6, 1, lastR - 5, 7)
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
  groupDataInLocalMemory(rawRows, dataObjs);

  if (dataObjs.length > 0) {
    finalValues = dataObjs.map(function (obj) {
      return [
        obj.date,
        obj.userN,
        obj.loc,
        obj.mood,
        obj.spcons,
        obj.srData,
        "",
      ];
    });

    // 批次寫入所有數據 (關鍵優化！)
    var writeRange = sheet.getRange(7, 1, finalValues.length, 7);
    writeRange
      .setValues(finalValues)
      .setFontSize(10)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("top")
      .setWrap(true)
      .setBorder(true, true, true, true, true, true);

    // 設置行高 (此動作較慢，但已大幅減少調用次數)
    for (var i = 0; i < finalValues.length; i++) {
      if (sheet.getRowHeight(6 + i) < 50) sheet.setRowHeight(6 + i, 50);
    }
  }
  sortSheetsAsc(ssFile); // 選用：每次生成後將工作表依名稱反向排序
  return ssFile.getUrl() + "#gid=" + sheet.getSheetId();
}

async function genPdfFile(yearmonth, custn, regen) {
  var year = yearmonth.substring(0, 4);
  var month = yearmonth.substring(4, 6);
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
      return pdfUrl;
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
  var allDataMap = getAllDataMapFromSources(yearmonth, year, month);
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

async function processSinglePdf(yearmonth, custn, regen) {
  const ssReportFile = getTargetsheet("ReportsUrl", "RP" + yearmonth).Spreadsheet;
  const sheet = ssReportFile.getSheetByName(custn);
  //把工作表 匯出成暫存試算表
  const ssId = ssReportFile.getId();
  const sheetId = sheet.getSheetId();
  const token = ScriptApp.getOAuthToken();
  //匯出成PDF
  // --- 1. 定義匯出 URL (含凍結列、頁碼、頁尾設定) ---
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?" +
    "format=pdf&" +
    "size=A4&" +
    "portrait=false&" +     // 橫向
    "fitw=true&" +          // 寬度符合頁面
    "gridlines=false&" +    // 不顯示格線
    "fzr=true&" +           // 重要：重複凍結列 (作為每一頁的頁首)
    "top_margin=0.5&" + //邊界：窄  (預設是 0.75 英吋，A4 紙寬約 8.27 英吋，這裡設定為 0.5 英吋，實際內容寬度約 7.27 英吋)
    "bottom_margin=0.5&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "pagenumbers=false&" +  // 不顯示頁碼
    "printtitle=false&" +   // 不顯示標題列
    "pagenum=CENTER&" +     // 頁碼置中 (雙面列印最保險的位置)
    "sheetnames=false&" +    // 不顯示工作表名稱
    "gid=" + sheetId;


  // --- 2. 抓取 PDF 原始內容 ---
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("PDF Export Failed (" + response.getResponseCode() + "): " + response.getContentText());
  }

  const pdfBytes = response.getContent();
  // 轉換成 Uint8Array 以供 pdf-lib 使用
  const uint8Array = new Uint8Array(pdfBytes);

  // --- 3. 載入 pdf-lib 庫 ---
  var setTimeout = function (f, t) { Utilities.sleep(t || 0); return f(); };
  var clearTimeout = function () { };
  eval(fetchPdfLib());

  // --- 4. 解析 PDF 並檢查頁數 ---
  const pdfDoc = await PDFLib.PDFDocument.load(uint8Array);
  const pageCount = pdfDoc.getPageCount();
  Logger.log("原始頁數: " + pageCount);

  // --- 5. 若為奇數頁，則補上一張空白頁 ---
  if (pageCount % 2 !== 0) {
    Logger.log("偵測到奇數頁，正在加入空白頁以符合雙面列印需求...");

    // 取得最後一頁的大小，建立相同大小的空白頁
    const pages = pdfDoc.getPages();
    const lastPage = pages[pages.length - 1];
    const { width, height } = lastPage.getSize();

    pdfDoc.addPage([width, height]);

    // (選填) 你可以在空白頁上加一行字，防止別人以為是漏印
    // const font = await pdfDoc.embedFont(PDFLib.StandardFonts.Helvetica);
    // pdfDoc.getPages()[pageCount].drawText('Blank Page for Duplex Printing', { x: 50, y: 50, size: 10, font: font });
  } else {
    Logger.log("頁數已是偶數，無需補頁。");
  }

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
  const file = DriveApp.createFile(blob);
  pdfFolder.addFile(file);
  folder.removeFile(file); // 從根目錄移除，避免混亂
  // 儲存 PDF 的網址到 Properties 以供前端使用
  setTargetUrl("PdfUrl", "PD" + yearmonth + "_" + custn, file.getUrl());

  // 回傳檔案網址供後續使用
  return file.getUrl();
}


/**
 * 預先讀取所有資料並分類 (Map: Name -> Rows[])
 */
function getAllDataMapFromSources(yearmonth, year, month) {
  var ssAnnual = getTargetsheet("RecUrl", "SY" + year).Spreadsheet.getSheetByName(
    yearmonth,
  );
  var ssTemp = getTargetsheet("SYTemp", "SYTemp").Spreadsheet.getSheetByName("SR_Data");
  var dataMap = {};
  var excludeIDs = ["GA05", "GA09"];

  [ssAnnual, ssTemp].forEach(function (sheet, idx) {
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    var h = values[0];
    var col = {
      Date: getColIndex(h, "Date"),
      Cust_N: getColIndex(h, "Cust_N"),
      SR_ID: getColIndex(h, "SR_ID"),
    };

    var isTemp = idx === 1;
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      var name = r[col.Cust_N];
      var srId = String(r[col.SR_ID]);

      // 過濾條件
      if (excludeIDs.indexOf(srId) !== -1) continue;
      if (isTemp) {
        var d = new Date(r[col.Date]);
        if (d.getFullYear() != year || d.getMonth() + 1 != month) continue;
      }

      if (!dataMap[name]) dataMap[name] = [];
      dataMap[name].push(r);
    }
  });
  return dataMap;
}

/**
 * 預先讀取所有個案基本資料
 */
function getCustInfoMap() {
  var sCust = MainSpreadsheet.getSheetByName("Cust");
  var data = sCust.getDataRange().getValues();
  var infoMap = {};
  var custNames = [];

  for (var i = 1; i < data.length; i++) {
    var name = data[i][0];
    if (name) {
      custNames.push(name);
      infoMap[name] = {
        sex: data[i][1],
        bd: formatDate(data[i][2]),
        add: data[i][3],
      };
    }
  }
  return { names: custNames, info: infoMap };
}

/**
 * 日期/次數/人員分組邏輯 (維持原邏輯)
 */
function groupDataInLocalMemory(rows, allDataObjs) {
  // 預設索引 (這部分建議在讀取時標準化)
  var col = {
    Date: 0,
    SRTimes: 1,
    User_N: 3,
    LOC: 7,
    MOOD: 8,
    SPCONS: 9,
    SR_ID: 5,
    SR_REC: 6,
  };

  // 取得唯一日期並排序
  var uniqueDates = [
    ...new Set(rows.map((r) => formatDate(r[col.Date]))),
  ].sort();

  uniqueDates.forEach(function (dStr) {
    var dateRows = rows.filter((r) => formatDate(r[col.Date]) === dStr);
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

        allDataObjs.push({
          date: dStr,
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

// --- 內部輔助函式庫 ---

/**
 * 合併並過濾來自年度與 Temp 的資料 (排除 GA02/GA09)
 */
function getCombinedFilteredData(ss1, ss2, custn, year, month) {
  var allRows = [];
  var excludeIDs = ["GA02", "GA09"];
  var columnInfo = null;

  [ss1, ss2].forEach(function (sheet, index) {
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    var h = data[0];
    var col = {
      Date: getColIndex(h, "Date"),
      Cust_N: getColIndex(h, "Cust_N"),
      SR_ID: getColIndex(h, "SR_ID"),
      SRTimes: getColIndex(h, "SRTimes"),
      User_N: getColIndex(h, "User_N"),
      LOC: getColIndex(h, "LOC"),
      MOOD: getColIndex(h, "MOOD"),
      SPCONS: getColIndex(h, "SPCONS"),
      SR_REC: getColIndex(h, "SR_REC"),
    };

    if (!columnInfo) columnInfo = col; // 存一份索引供後續處理使用

    var isTemp = index === 1;
    var filtered = data.slice(1).filter(function (r) {
      // 1. 案主姓名比對
      if (r[col.Cust_N] !== custn) return false;
      // 2. 排除 ID (GA02, GA09)
      if (excludeIDs.indexOf(String(r[col.SR_ID])) !== -1) return false;
      // 3. 日期篩選 (針對 SYTemp 需限定當月份)
      if (isTemp) {
        var d = new Date(r[col.Date]);
        if (d.getFullYear() != year || d.getMonth() + 1 != month) return false;
      }
      return true;
    });
    allRows = allRows.concat(filtered);
  });

  return { rows: allRows, col: columnInfo };
}

/**
 * 將原始資料列依照 日期 -> 次數 -> 人員 進行分組物件化
 */
function processRowsIntoDataObjs(rows, col, allDataObjs) {
  // 取得唯一日期清單並排序 (yyyy/MM/dd)
  var uniqueDates = [
    ...new Set(rows.map((r) => formatDate(r[col.Date]))),
  ].sort();

  uniqueDates.forEach(function (dStr) {
    var dateRows = rows.filter((r) => formatDate(r[col.Date]) === dStr);
    var uniqueTimes = [...new Set(dateRows.map((r) => r[col.SRTimes]))].sort(
      (a, b) => a - b,
    );

    uniqueTimes.forEach(function (timeVal) {
      var timeRows = dateRows.filter((r) => r[col.SRTimes] === timeVal);
      var uniqueUsers = [...new Set(timeRows.map((r) => r[col.User_N]))];

      uniqueUsers.forEach(function (uName) {
        var userRows = timeRows.filter((r) => r[col.User_N] === uName);
        if (userRows.length === 0) return;
        var first = userRows[0];

        var obj = {
          date: dStr,
          userN: uName,
          loc: first[col.LOC],
          mood: first[col.MOOD],
          spcons: first[col.SPCONS],
          srData: "",
        };

        // 組合 SR_ID + SR_REC，並處理換行與逗號
        var recItems = userRows.map(
          (r) =>
            String(r[col.SR_ID]) +
            String(r[col.SR_REC]).replace(/\n/g, "").replace(/,/g, "，"),
        );

        // 分欄邏輯 (若資料多則併行顯示)
        if (recItems.length <= 3) {
          obj.srData = recItems.join("\n");
        } else {
          var lines = [];
          for (var i = 0; i < 3; i++) {
            var s =
              recItems[i] + (recItems[i + 3] ? " " + recItems[i + 3] : "");
            lines.push(s);
          }
          obj.srData = lines.join("\n");
        }
        allDataObjs.push(obj);
      });
    });
  });
}

/**
 * 寫入報表標題與個人資料
 */
function writeReportHeader(sheet, custn, year, month) {
  // var ssMain = getTargetsheet("SYTemp", "SYTemp");
  var custSheet = MainSpreadsheet.getSheetByName("Cust");
  var custData = custSheet.getDataRange().getValues();
  var info = { sex: "", bd: "", add: "" };

  for (var i = 1; i < custData.length; i++) {
    if (custData[i][0] === custn) {
      info.sex = custData[i][1];
      info.bd = formatDate(custData[i][2]);
      info.add = custData[i][3];
      break;
    }
  }

  // Row 2: 大標題
  sheet
    .getRange("A2")
    .setValue(
      parseInt(year) -
      1911 +
      "年" +
      parseInt(month) +
      "月　　居家服務照顧內容紀錄單",
    );
  // Row 3: 個資
  sheet
    .getRange("A3:C3")
    .merge()
    .setValue("個案姓名：" + custn);
  sheet
    .getRange("D3:E3")
    .merge()
    .setValue("性別：" + info.sex);
  sheet
    .getRange("F3:G3")
    .merge()
    .setValue("出生年月日：" + info.bd);
  // Row 4: 住址
  sheet
    .getRange("A4:G4")
    .merge()
    .setValue("住址：" + info.add);
}

/**
 * 格式化日期為 yyyy/MM/dd
 */
function formatDate(dateObj) {
  if (!dateObj) return "";
  var d = new Date(dateObj);
  if (isNaN(d.getTime())) return String(dateObj);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

var _pdfLibCache = null;
function fetchPdfLib() {
  if (!_pdfLibCache) {
    _pdfLibCache = UrlFetchApp.fetch("https://unpkg.com/pdf-lib/dist/pdf-lib.min.js").getContentText();
  }
  return _pdfLibCache;
}

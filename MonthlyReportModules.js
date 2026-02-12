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
  var progress = getProgress();
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
      saveProgress({
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
        false,
      );
    }
  }

  // 全部完成
  clearProgress();
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
  if (sheet && sheet.getLastRow() >= 6 && !regen) {
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
    var writeRange = sheet.getRange(6, 1, finalValues.length, 7);
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

function genPdfFile(yearmonth, custn, regen) {
  var year = yearmonth.substring(0, 4);
  var month = yearmonth.substring(4, 6);
  if (custn !== "all") {
    //
  }

}

function processSingleReport() {


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






/**
 * 取得所有試算表資料，供前端快取與篩選
 * 回傳結構包含 main (主程式+暫存), rec (紀錄), reports (報表), pdfs (PDF)
 */
function getOpenSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    main: [],
    rec: [],
    reports: [],
    pdfs: []
  };

  // 1. 加入 Main Spreadsheet (SYCompany)
  result.main.push({ name: "主程式 (SYCompany)", url: ss.getUrl() });

  // 2. 加入 SYTemp
  const tempSheet = ss.getSheetByName("SYTemp");
  if (tempSheet) {
    const tempData = tempSheet.getDataRange().getValues();
    for (let i = 1; i < tempData.length; i++) {
      if (tempData[i][0] === 'SYTemp') {
        result.main.push({ name: "暫存檔 (SYTemp)", url: tempData[i][1] });
        break;
      }
    }
  }

  // 3. 取得 RecUrl
  const recSheet = ss.getSheetByName("RecUrl");
  if (recSheet) {
    const recData = recSheet.getDataRange().getValues();
    for (let i = 1; i < recData.length; i++) {
      if (recData[i][0] && recData[i][1] && recData[i][0] !== "SYSample") {
        result.rec.push({ name: recData[i][0], url: recData[i][1] });
      }
    }
  }

  // 4. 取得 ReportsUrl
  const reportsheet = ss.getSheetByName("ReportsUrl");
  if (reportsheet) {
    const reportData = reportsheet.getDataRange().getValues();
    for (let i = 1; i < reportData.length; i++) {
      if (reportData[i][0] && reportData[i][1] && reportData[i][0] !== "RPSample") {
        result.reports.push({ name: reportData[i][0], url: reportData[i][1] });
      }
    }
  }

  // 5. 取得 PdfUrl
  const pdfsheet = ss.getSheetByName("PdfUrl");
  if (pdfsheet) {
    const pdfData = pdfsheet.getDataRange().getValues();
    for (let i = 1; i < pdfData.length; i++) {
      if (pdfData[i][0] && pdfData[i][1]) {
        result.pdfs.push({ name: pdfData[i][0], url: pdfData[i][1] });
      }
    }
  }

  return result;
}
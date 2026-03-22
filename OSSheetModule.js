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
    if (recData.length > 1) {
      const headers = recData[0];
      const nameIdx = getColIndex(headers, "SY_N");
      const urlIdx = getColIndex(headers, "SY_Url");
      for (let i = 1; i < recData.length; i++) {
        const name = recData[i][nameIdx !== -1 ? nameIdx : 0];
        const url = recData[i][urlIdx !== -1 ? urlIdx : 1];
        if (name && url && name !== "SYSample") {
          result.rec.push({ name, url });
        }
      }
    }
  }

  // 4. 取得 ReportsUrl
  const reportsheet = ss.getSheetByName("ReportsUrl");
  if (reportsheet) {
    const reportData = reportsheet.getDataRange().getValues();
    if (reportData.length > 1) {
      const headers = reportData[0];
      const nameIdx = getColIndex(headers, "SY_N");
      const urlIdx = getColIndex(headers, "SY_Url");
      for (let i = 1; i < reportData.length; i++) {
        const name = reportData[i][nameIdx !== -1 ? nameIdx : 0];
        const url = reportData[i][urlIdx !== -1 ? urlIdx : 1];
        if (name && url && name !== "RPSample") {
          result.reports.push({ name, url });
        }
      }
    }
  }

  // 5. 取得 PdfUrl
  const pdfsheet = ss.getSheetByName("PdfUrl");
  if (pdfsheet) {
    const pdfData = pdfsheet.getDataRange().getValues();
    if (pdfData.length > 1) {
      const headers = pdfData[0];
      const nameIdx = getColIndex(headers, "SY_N");
      const urlIdx = getColIndex(headers, "SY_Url");
      for (let i = 1; i < pdfData.length; i++) {
        const name = pdfData[i][nameIdx !== -1 ? nameIdx : 0];
        const url = pdfData[i][urlIdx !== -1 ? urlIdx : 1];
        if (name && url) {
          result.pdfs.push({ name, url });
        }
      }
    }
  }

  return result;
}
/**
 * 取得試算表清單供前端下拉選單使用
 */
function getOpenSheetOptions() {
  const options = [];
  
  // 1. 加入 Main Spreadsheet (SYCompany)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  options.push({ name: "主程式 (SYCompany)", url: ss.getUrl() });

  // 2. 加入 SYTemp
  // 假設您有名為 'SYTemp' 的工作表，且網址在某個欄位，或直接根據規則定義
  // 這裡需根據您的實際工作表結構調整
  const tempSheet = ss.getSheetByName("SYTemp");
  if(tempSheet) {
    const tempData = tempSheet.getDataRange().getValues();
    // 假設第二列開始，第一欄是 'SYTemp'，第二欄是網址
    tempData.forEach(row => {
      if(row[0] === 'SYTemp') {
        options.push({ name: "暫存檔 (SYTemp)", url: row[1] }); // 這裡 row[1] 應為 SYT_Url
      }
    });
  }

  // 3. 遍歷 SY+yyyy (歷年紀錄)
  const recSheet = ss.getSheetByName("RecUrl");
  if(recSheet) {
    const recData = recSheet.getDataRange().getValues();
    // 跳過標題列
    for(let i = 1; i < recData.length; i++) {
      const name = recData[i][0]; // SY_N
      const url = recData[i][1];  // SY_Url
      if(name && name.toString().startsWith("SY")) {
        options.push({ name: "歷年紀錄 - " + name, url: url });
      }
    }
  }

  return options;
}
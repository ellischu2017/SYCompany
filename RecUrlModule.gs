/**
 * RecUrl.gs - 服務紀錄單網址管理模組
 * 提供網址資料的 CRUD 操作
 */

/**
 * 取得「RecUrl」工作表的個案姓名列表
 */
function getRecUrlList() {
  const sheet = MainSpreadsheet.getSheetByName("RecUrl");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢特定個案的網址資料
 */
function queryRecUrlData(syName) {
  const sheet = MainSpreadsheet.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == syName) {
      return {
        found: true,
        rowResult: {
          SY_N: data[i][0],
          SY_Url: data[i][1],
        },
      };
    }
  }
  return { found: false };
}

/**
 * 儲存網址資料 (新增或更新)
 */
function saveRecUrlData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.syName) {
      sheet.getRange(i + 1, 2).setValue(formObj.recUrl);
      return { success: true, message: "網址更新成功！" };
    }
  }

  sheet.appendRow([formObj.syName, formObj.recUrl]);
  return { success: true, message: "網址新增成功！" };
}

/**
 * 刪除網址資料
 */
function deleteRecUrlData(syName) {
  const sheet = MainSpreadsheet.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == syName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "資料已刪除！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

/**
 * 輔助函式：根據 SY_N 取得 RecUrl 內的網址
 */
function getUrlFromRecUrl(mainSS, syName) {
  const recSheet = mainSS.getSheetByName("RecUrl");
  const data = recSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === syName) return data[i][1];
  }
  return null;
}

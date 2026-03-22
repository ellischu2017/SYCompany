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
 * 取得所有 RecUrl 資料 (用於前端快取)
 */
function getAllRecUrlData() {
  const sheet = MainSpreadsheet.getSheetByName("RecUrl");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["SY_N", "SY_Url"];
  const colMap = getColIndicesMap(headers, targetFields);
  
  if (colMap["SY_N"] === -1) return [];

  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][colMap["SY_N"]]) {
      result.push({
        SY_N: data[i][colMap["SY_N"]],
        SY_Url: colMap["SY_Url"] !== -1 ? data[i][colMap["SY_Url"]] : ""
      });
    }
  }
  return result;
}

/**
 * 查詢特定個案的網址資料
 */
function queryRecUrlData(syName) {
  const sheet = MainSpreadsheet.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["SY_N", "SY_Url"];
  const colMap = getColIndicesMap(headers, targetFields);
  
  if (colMap["SY_N"] === -1) return { found: false };

  for (let i = 1; i < data.length; i++) {
    if (data[i][colMap["SY_N"]] == syName) {
      return {
        found: true,
        rowResult: {
          SY_N: data[i][colMap["SY_N"]],
          SY_Url: colMap["SY_Url"] !== -1 ? data[i][colMap["SY_Url"]] : "",
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
  const headers = data[0];
  const idxName = getColIndex(headers, "SY_N");
  const idxUrl = getColIndex(headers, "SY_Url");

  if (idxName === -1) return { success: false, message: "錯誤：找不到 SY_N 欄位" };

  for (let i = 1; i < data.length; i++) {
    if (data[i][idxName] == formObj.syName) {
      if (idxUrl !== -1) sheet.getRange(i + 1, idxUrl + 1).setValue(formObj.recUrl);
      return { success: true, message: "網址更新成功！" };
    }
  }

  // 新增
  const newRow = new Array(headers.length).fill("");
  newRow[idxName] = formObj.syName;
  if (idxUrl !== -1) newRow[idxUrl] = formObj.recUrl;

  sheet.appendRow(newRow);

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow > 1) {
    sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .sort({ column: idxName + 1, ascending: true });
  }
  return { success: true, message: "網址新增成功！" };
}

/**
 * 刪除網址資料
 */
function deleteRecUrlData(syName) {
  const sheet = MainSpreadsheet.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxName = getColIndex(headers, "SY_N");

  for (let i = 1; i < data.length; i++) {
    if (data[i][idxName] == syName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "資料已刪除！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

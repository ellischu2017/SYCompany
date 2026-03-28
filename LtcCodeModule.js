/**
 * LtcCode.gs - 長照編碼管理模組
 * 提供服務編碼資料的 CRUD 操作
 */

/**
 * 取得所有服務編碼資料 (用於前端初始化 Session)
 */
function getAllLtcCodeData() {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const targetFields = ["SR_ID", "SR_Name", "SR_Detail", "SR_Cont"];
  const colMap = getColIndicesMap(headers, targetFields);

  const idxId = colMap["SR_ID"];
  // 支援 SR_Name 或 SR_Cont (Service Content) 作為名稱欄位
  const idxName =
    colMap["SR_Name"] !== -1 ? colMap["SR_Name"] : colMap["SR_Cont"];
  const idxDetail = colMap["SR_Detail"];

  if (idxId === -1) return [];

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idxId]) {
      // SR_ID 不為空
      result.push({
        SR_ID: row[idxId],
        SR_Name: idxName !== -1 ? row[idxName] : "",
        SR_Detail: idxDetail !== -1 ? row[idxDetail] : "",
      });
    }
  }
  return result;
}

/**
 * 取得「LTC_Code」工作表的服務編碼列表 (SR_ID)
 */
function getLtcCodeList() {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 先讀取標題以確認 SR_ID 欄位位置
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxId = getColIndex(headers, "SR_ID");

  if (idxId === -1) return [];

  const data = sheet.getRange(2, idxId + 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢服務編碼詳細資料
 */
function queryLtcCodeData(srId) {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["SR_ID", "SR_Name", "SR_Detail", "SR_Cont"];
  const colMap = getColIndicesMap(headers, targetFields);

  const idxId = colMap["SR_ID"];
  const idxName =
    colMap["SR_Name"] !== -1 ? colMap["SR_Name"] : colMap["SR_Cont"];
  const idxDetail = colMap["SR_Detail"];

  if (idxId === -1) return { found: false };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) == srId) {
      return {
        found: true,
        rowResult: {
          SR_ID: data[i][idxId],
          SR_Name: idxName !== -1 ? data[i][idxName] : "",
          SR_Detail: idxDetail !== -1 ? data[i][idxDetail] : "",
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增服務編碼
 */
function addLtcCodeData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  const list = getLtcCodeList();

  if (list.includes(formObj.srId)) {
    return { success: false, message: "該服務編碼已存在！" };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["SR_ID", "SR_Name", "SR_Detail", "SR_Cont"];
  const colMap = getColIndicesMap(headers, targetFields);

  const idxId = colMap["SR_ID"];
  const idxName =
    colMap["SR_Name"] !== -1 ? colMap["SR_Name"] : colMap["SR_Cont"];
  const idxDetail = colMap["SR_Detail"];

  if (idxId === -1)
    return { success: false, message: "錯誤：找不到 SR_ID 欄位" };

  const newRow = new Array(headers.length).fill("");
  newRow[idxId] = formObj.srId;
  if (idxName !== -1) newRow[idxName] = formObj.srName;
  if (idxDetail !== -1) newRow[idxDetail] = formObj.srDetail;

  sheet.appendRow(newRow);
  // 2. 執行排序 (ORDER BY Cust_N A->Z)
  // getRange(row, column, numRows, numColumns)
  // 從第 2 列開始排（避開標題列），針對所有已使用的範圍進行排序
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow > 1) {
    sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .sort({ column: idxId + 1, ascending: true });
  }
  CacheService.getScriptCache().remove("SRServer01_InitData");
  return { success: true, message: "編碼新增成功！" };
}

/**
 * 更新服務編碼資料
 */
function updateLtcCodeData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["SR_ID", "SR_Name", "SR_Detail", "SR_Cont"];
  const colMap = getColIndicesMap(headers, targetFields);

  const idxId = colMap["SR_ID"];
  const idxName =
    colMap["SR_Name"] !== -1 ? colMap["SR_Name"] : colMap["SR_Cont"];
  const idxDetail = colMap["SR_Detail"];

  if (idxId === -1)
    return { success: false, message: "錯誤：找不到 SR_ID 欄位" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) == formObj.srId) {
      const rowNum = i + 1;
      if (idxName !== -1)
        sheet.getRange(rowNum, idxName + 1).setValue(formObj.srName);
      if (idxDetail !== -1)
        sheet.getRange(rowNum, idxDetail + 1).setValue(formObj.srDetail);
      CacheService.getScriptCache().remove("SRServer01_InitData");
      return { success: true, message: "編碼資料更新成功！" };
    }
  }
  return { success: false, message: "找不到資料，無法更新。" };
}

/**
 * 刪除服務編碼
 */
function deleteLtcCodeData(srId) {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxId = getColIndex(headers, "SR_ID");

  if (idxId === -1)
    return { success: false, message: "錯誤：找不到 SR_ID 欄位" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]) == srId) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove("SRServer01_InitData");
      return { success: true, message: "編碼刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

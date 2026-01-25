/**
 * LtcCode.gs - 長照編碼管理模組
 * 提供服務編碼資料的 CRUD 操作
 */

/**
 * 取得「LTC_Code」工作表的服務編碼列表 (SR_ID)
 */
function getLtcCodeList() {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢服務編碼詳細資料
 */
function queryLtcCodeData(srId) {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == srId) {
      return {
        found: true,
        rowResult: {
          SR_ID: data[i][0],
          SR_Name: data[i][1],
          SR_Detail: data[i][2],
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

  const newRow = [formObj.srId, formObj.srName, formObj.srDetail];

  sheet.appendRow(newRow);
  // 2. 執行排序 (ORDER BY Cust_N A->Z)
  // getRange(row, column, numRows, numColumns)
  // 從第 2 列開始排（避開標題列），針對所有已使用的範圍進行排序
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow > 1) {
    sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .sort({ column: 1, ascending: true });
  }
  return { success: true, message: "編碼新增成功！" };
}

/**
 * 更新服務編碼資料
 */
function updateLtcCodeData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.srId) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 2).setValue(formObj.srName);
      sheet.getRange(rowNum, 3).setValue(formObj.srDetail);
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

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == srId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "編碼刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

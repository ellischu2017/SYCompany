/**
 * Manager.gs - 管理員管理模組
 * 提供管理員資料的 CRUD 操作
 */

/**
 * 取得「Manager」工作表的管理員姓名列表
 */
function getManaList() {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢管理員詳細資料
 */
function queryManaData(manaName) {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == manaName) {
      return {
        found: true,
        rowResult: {
          Mana_N: data[i][0],
          Mana_Email: data[i][1],
          Mana_Tel: data[i][2].toString(),
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增管理員資料
 */
function addManaData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  const list = getManaList();

  if (list.includes(formObj.manaName)) {
    return { success: false, message: "該管理員姓名已存在！" };
  }

  const newRow = [
    formObj.manaName,
    formObj.manaEmail,
    "'" + formObj.manaTel,
  ];

  sheet.appendRow(newRow);
  return { success: true, message: "新增成功！" };
}

/**
 * 更新管理員資料
 */
function updateManaData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.manaName) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 2).setValue(formObj.manaEmail);
      sheet.getRange(rowNum, 3).setValue("'" + formObj.manaTel);
      return { success: true, message: "資料更新成功！" };
    }
  }
  return { success: false, message: "找不到資料，無法更新。" };
}

/**
 * 刪除管理員資料
 */
function deleteManaData(manaName) {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == manaName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

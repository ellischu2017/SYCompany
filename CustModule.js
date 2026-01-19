/**
 * Cust.gs - 個案管理模組
 * 提供個案資料的 CRUD 操作
 */

/**
 * 取得「Cust」工作表的個案姓名列表，用於下拉選單
 */
function getCustList() {
  const sheet = MainSpreadsheet.getSheetByName("Cust");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 根據個案姓名查詢詳細資料
 */
function queryCustData(custName) {
  const sheet = MainSpreadsheet.getSheetByName("Cust");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == custName) {
      const row = data[i];

      let birthDate = row[1];
      if (birthDate instanceof Date) {
        birthDate = Utilities.formatDate(
          birthDate,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
      }

      return {
        found: true,
        rowResult: {
          Cust_N: row[0],
          Cust_BD: birthDate,
          Cust_Add: row[2],
          Sup_N: row[3],
          DSup_N: row[4],
          Sup_Tel: row[5].toString(),
          Office_Tel: row[6].toString(),
          Cust_EC: row[7],
          Cust_EC_Tel: row[8].toString(),
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增個案資料
 */
function addCustData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("Cust");

  const list = getCustList();
  if (list.includes(formObj.custName)) {
    return { success: false, message: "該個案姓名已存在！" };
  }

  const newRow = [
    formObj.custName,
    "'" + formObj.custBirth,
    formObj.custAddr,
    formObj.supName,
    formObj.dsupName,
    "'" + formObj.supTel,
    "'" + formObj.officeTel,
    formObj.ecName,
    "'" + formObj.ecTel,
  ];

  sheet.appendRow(newRow);
  return { success: true, message: "新增成功！" };
}

/**
 * 更新個案資料
 */
function updateCustData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("Cust");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.custName) {
      const rowNum = i + 1;

      sheet.getRange(rowNum, 2).setValue(formObj.custBirth);
      sheet.getRange(rowNum, 3).setValue(formObj.custAddr);
      sheet.getRange(rowNum, 4).setValue(formObj.supName);
      sheet.getRange(rowNum, 5).setValue(formObj.dsupName);
      sheet.getRange(rowNum, 6).setValue("'" + formObj.supTel);
      sheet.getRange(rowNum, 7).setValue("'" + formObj.officeTel);
      sheet.getRange(rowNum, 8).setValue(formObj.ecName);
      sheet.getRange(rowNum, 9).setValue("'" + formObj.ecTel);

      return { success: true, message: "資料更新成功！" };
    }
  }
  return { success: false, message: "找不到該個案資料，無法更新。" };
}

/**
 * 刪除個案資料
 */
function deleteCustData(custName) {
  const sheet = MainSpreadsheet.getSheetByName("Cust");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == custName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "找不到資料，刪除失敗。" };
}

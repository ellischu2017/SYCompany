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

      let birthDate = row[2];
      if (birthDate instanceof Date) {
        birthDate = Utilities.formatDate(
          birthDate,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd",
        );
      }

      return {
        found: true,
        rowResult: {
          Cust_N: row[0].toString().trim(),
          Cust_Sex: row[1] || "",
          Cust_BD: birthDate,
          Cust_Add: row[3] || "",
          Sup_N: row[4].trim() || "",
          DSup_N: row[5].trim() || "",
          Sup_Tel: row[6].toString(),
          Office_Tel: row[7].toString(),
          Cust_EC: row[8].toString().trim() || "",
          Cust_EC_Tel: row[9].toString(),
          Cust_LTC_Code: row[10] || "",
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
    formObj.custSex,
    "'" + formObj.custBirth,
    formObj.custAddr,
    formObj.supName,
    formObj.dsupName,
    "'" + formObj.supTel,
    "'" + formObj.officeTel,
    formObj.ecName,
    "'" + formObj.ecTel,
    formObj.ltcCode || "",
  ];

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
  return { success: true, message: "新增成功！新增成功並已完成姓名排序！" };
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
      sheet.getRange(rowNum, 3).setValue(formObj.custSex);
      sheet.getRange(rowNum, 4).setValue(formObj.custAddr);
      sheet.getRange(rowNum, 5).setValue(formObj.supName);
      sheet.getRange(rowNum, 6).setValue(formObj.dsupName);
      sheet.getRange(rowNum, 7).setValue("'" + formObj.supTel);
      sheet.getRange(rowNum, 8).setValue("'" + formObj.officeTel);
      sheet.getRange(rowNum, 9).setValue(formObj.ecName);
      sheet.getRange(rowNum, 10).setValue("'" + formObj.ecTel);
      sheet.getRange(rowNum, 11).setValue(formObj.ltcCode || "");

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

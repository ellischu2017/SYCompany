/**
 * User.gs - 居服員管理模組
 * 提供居服員資料的 CRUD 操作
 */

/**
 * 取得「User」工作表的居服員姓名列表
 */
function getUserList() {
  const sheet = MainSpreadsheet.getSheetByName("User");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢居服員詳細資料
 */
function queryUserData(userName) {
  const sheet = MainSpreadsheet.getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userName) {
      return {
        found: true,
        rowResult: {
          User_N: data[i][0],
          User_Email: data[i][1],
          User_Tel: data[i][2].toString(),
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增居服員資料
 */
function addUserData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("User");
  const list = getUserList();

  if (list.includes(formObj.userName)) {
    return { success: false, message: "該居服員姓名已存在！" };
  }

  const newRow = [formObj.userName, formObj.userEmail, "'" + formObj.userTel];

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
  return { success: true, message: "新增成功！" };
}

/**
 * 更新居服員資料
 */
function updateUserData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.userName) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 2).setValue(formObj.userEmail);
      sheet.getRange(rowNum, 3).setValue("'" + formObj.userTel);
      return { success: true, message: "資料更新成功！" };
    }
  }
  return { success: false, message: "找不到資料，無法更新。" };
}

/**
 * 刪除居服員資料
 */
function deleteUserData(userName) {
  const sheet = MainSpreadsheet.getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

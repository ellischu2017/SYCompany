/**
 * User.gs - 居服員管理模組
 * 提供居服員資料的 CRUD 操作
 */

/**
 * 取得所有居服員資料 (用於前端快取)
 */
function getAllUserData() {
  // 1. 嘗試從快取讀取
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("UserData");
  if (cachedData) return JSON.parse(cachedData);

  const sheet = MainSpreadsheet.getSheetByName("User");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["User_N", "User_Email", "User_Tel"];
  const colMap = getColIndicesMap(headers, targetFields);

  if (colMap["User_N"] === -1) return [];

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[colMap["User_N"]];
    if (name) {
      result.push({
        User_N: name.toString(),
        User_Email: colMap["User_Email"] !== -1 ? String(row[colMap["User_Email"]]) : "",
        User_Tel: colMap["User_Tel"] !== -1 ? String(row[colMap["User_Tel"]]) : ""
      });
    }
  }

  // 2. 寫入快取 (設定 30 分鐘)
  try {
    cache.put("UserData", JSON.stringify(result), 1800);
  } catch (e) {
    console.log("快取 UserData 失敗: " + e.toString());
  }
  return result;
}

/**
 * 查詢居服員詳細資料
 */
function queryUserData(userName) {
  const sheet = MainSpreadsheet.getSheetByName("User");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["User_N", "User_Email", "User_Tel"];
  const colMap = getColIndicesMap(headers, targetFields);
  
  if (colMap["User_N"] === -1) return { found: false };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[colMap["User_N"]] == userName) {
      return {
        found: true,
        rowResult: {
          User_N: row[colMap["User_N"]],
          User_Email: colMap["User_Email"] !== -1 ? row[colMap["User_Email"]] : "",
          User_Tel: colMap["User_Tel"] !== -1 ? String(row[colMap["User_Tel"]]) : "",
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
  // 改用動態讀取標題與資料，避免依賴 getUserList 的硬編碼順序
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idxName = getColIndex(headers, "User_N");
  const idxEmail = getColIndex(headers, "User_Email");
  const idxTel = getColIndex(headers, "User_Tel");

  if (idxName === -1) return { success: false, message: "錯誤：找不到 User_N 欄位" };

  // 檢查是否重複
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxName]) === formObj.userName) {
      return { success: false, message: "該居服員姓名已存在！" };
    }
  }

  // 動態建立新資料列
  const newRow = new Array(headers.length).fill("");
  newRow[idxName] = formObj.userName;
  if (idxEmail !== -1) newRow[idxEmail] = formObj.userEmail;
  if (idxTel !== -1) newRow[idxTel] = "'" + formObj.userTel;

  sheet.appendRow(newRow);
  
  // 執行排序
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow > 1) {
    sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .sort({ column: idxName + 1, ascending: true });
  }
  CacheService.getScriptCache().remove("SRServer01_InitData");
  CacheService.getScriptCache().remove("UserData");
  return { success: true, message: "新增成功！" };
}

/**
 * 更新居服員資料
 */
function updateUserData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("User");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxName = getColIndex(headers, "User_N");
  const idxEmail = getColIndex(headers, "User_Email");
  const idxTel = getColIndex(headers, "User_Tel");

  if (idxName === -1) return { success: false, message: "錯誤：找不到 User_N 欄位" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxName]) == formObj.userName) {
      const rowNum = i + 1;
      if (idxEmail !== -1) sheet.getRange(rowNum, idxEmail + 1).setValue(formObj.userEmail);
      if (idxTel !== -1) sheet.getRange(rowNum, idxTel + 1).setValue("'" + formObj.userTel);
      CacheService.getScriptCache().remove("SRServer01_InitData");
      CacheService.getScriptCache().remove("UserData");
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
  const headers = data[0];
  const idxName = getColIndex(headers, "User_N");

  if (idxName === -1) return { success: false, message: "錯誤：找不到 User_N 欄位" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxName]) == userName) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove("SRServer01_InitData");
      CacheService.getScriptCache().remove("UserData");
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "找不到資料，無法刪除。" };
}

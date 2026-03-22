/**
 * Manager.gs - 管理員管理模組
 * 提供管理員資料的 CRUD 操作
 */

/**
 * 取得所有管理員資料 (用於前端快取)
 */
function getAllManagerData() {
  // 1. 嘗試從快取讀取
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("ManagerData");
  if (cachedData) return JSON.parse(cachedData);

  const sheet = MainSpreadsheet.getSheetByName("Manager");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["Mana_N", "Mana_Email", "Mana_Tel"];
  const colMap = getColIndicesMap(headers, targetFields);

  if (colMap["Mana_N"] === -1) return [];

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[colMap["Mana_N"]]) {
      result.push({
        Mana_N: row[colMap["Mana_N"]].toString(),
        Mana_Email: colMap["Mana_Email"] !== -1 ? String(row[colMap["Mana_Email"]]) : "",
        Mana_Tel: colMap["Mana_Tel"] !== -1 ? String(row[colMap["Mana_Tel"]]) : ""
      });
    }
  }

  // 2. 寫入快取 (設定 30 分鐘)
  try {
    cache.put("ManagerData", JSON.stringify(result), 1800);
  } catch (e) {
    console.log("快取 ManagerData 失敗: " + e.toString());
  }
  return result;
}

/**
 * 取得「Manager」工作表的管理員姓名列表
 */
function getManaList() {
  // 1. 嘗試從快取讀取
  const cache = CacheService.getScriptCache();
  const cachedList = cache.get("ManaList");
  if (cachedList) return JSON.parse(cachedList);

  const sheet = MainSpreadsheet.getSheetByName("Manager");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const list = data.map((r) => r[0]).filter((n) => n !== "");

  // 2. 寫入快取 (設定 30 分鐘過期)
  cache.put("ManaList", JSON.stringify(list), 1800);
  return list;
}

/**
 * 查詢管理員詳細資料
 */
function queryManaData(manaName) {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const targetFields = ["Mana_N", "Mana_Email", "Mana_Tel"];
  const colMap = getColIndicesMap(headers, targetFields);

  if (colMap["Mana_N"] === -1) return { found: false };

  for (let i = 1; i < data.length; i++) {
    if (data[i][colMap["Mana_N"]] == manaName) {
      return {
        found: true,
        rowResult: {
          Mana_N: data[i][colMap["Mana_N"]],
          Mana_Email: colMap["Mana_Email"] !== -1 ? data[i][colMap["Mana_Email"]] : "",
          Mana_Tel: colMap["Mana_Tel"] !== -1 ? data[i][colMap["Mana_Tel"]].toString() : "",
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
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxName = getColIndex(headers, "Mana_N");
  const idxEmail = getColIndex(headers, "Mana_Email");
  const idxTel = getColIndex(headers, "Mana_Tel");

  if (idxName === -1) return { success: false, message: "錯誤：找不到 Mana_N 欄位" };

  // 檢查重複
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxName]) === formObj.manaName) {
      return { success: false, message: "該管理員姓名已存在！" };
    }
  }

  // 動態建立新資料列
  const newRow = new Array(headers.length).fill("");
  newRow[idxName] = formObj.manaName;
  if (idxEmail !== -1) newRow[idxEmail] = formObj.manaEmail;
  if (idxTel !== -1) newRow[idxTel] = "'" + formObj.manaTel;

  sheet.appendRow(newRow);

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow > 1) {
    sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .sort({ column: idxName + 1, ascending: true });
  }
  // 清除快取，確保下次讀取到最新名單
  CacheService.getScriptCache().remove("ManaList");
  CacheService.getScriptCache().remove("ManagerData");
  syncMasterTablePermissions();
  return { success: true, message: "新增成功！" };
}

/**
 * 更新管理員資料
 */
function updateManaData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxName = getColIndex(headers, "Mana_N");
  const idxEmail = getColIndex(headers, "Mana_Email");
  const idxTel = getColIndex(headers, "Mana_Tel");

  if (idxName === -1) return { success: false, message: "錯誤：找不到 Mana_N 欄位" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxName]) == formObj.manaName) {
      const rowNum = i + 1;
      if (idxEmail !== -1) sheet.getRange(rowNum, idxEmail + 1).setValue(formObj.manaEmail);
      if (idxTel !== -1) sheet.getRange(rowNum, idxTel + 1).setValue("'" + formObj.manaTel);
      // 清除快取
      CacheService.getScriptCache().remove("ManaList");
      CacheService.getScriptCache().remove("ManagerData");
      syncMasterTablePermissions();
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
  const headers = data[0];
  const idxName = getColIndex(headers, "Mana_N");

  if (idxName === -1) return { success: false, message: "錯誤：找不到 Mana_N 欄位" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxName]) == manaName) {
      sheet.deleteRow(i + 1);
      // 清除快取
      CacheService.getScriptCache().remove("ManaList");
      CacheService.getScriptCache().remove("ManagerData");
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

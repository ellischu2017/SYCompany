/**
 * Manager.gs - 管理員管理模組
 * 提供管理員資料的 CRUD 操作
 */

/**
 * 取得所有管理員資料 (用於前端快取)
 */
function getAllManagerData() {
  const sheet = MainSpreadsheet.getSheetByName("Manager");
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const result = [];

  // 從第 2 列開始 (索引 1)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      result.push({
        Mana_N: data[i][0].toString(),
        Mana_Email: data[i][1] ? data[i][1].toString() : "",
        Mana_Tel: data[i][2] ? data[i][2].toString() : ""
      });
    }
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

  const newRow = [formObj.manaName, formObj.manaEmail, "'" + formObj.manaTel];

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
  // 清除快取，確保下次讀取到最新名單
  CacheService.getScriptCache().remove("ManaList");
  syncMasterTablePermissions();
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
      // 清除快取
      CacheService.getScriptCache().remove("ManaList");
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

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == manaName) {
      sheet.deleteRow(i + 1);
      // 清除快取
      CacheService.getScriptCache().remove("ManaList");
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

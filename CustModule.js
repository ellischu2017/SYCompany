/**
 * Cust.gs - 個案管理模組
 * 提供個案資料的 CRUD 操作
 */

/**
 * 從指定工作表讀取所有資料並格式化為物件陣列
 * @param {string} sheetName 工作表名稱
 * @returns {Array<Object>} 物件陣列
 */
function getFormattedDataFromSheet(sheetName) {
  const sheet = MainSpreadsheet.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // 取得並移除標題列

  // 建立標題與索引的對應
  const headerMap = headers.reduce((acc, header, i) => {
    acc[header] = i;
    return acc;
  }, {});

  const results = data.map(row => {
    let birthDate = row[headerMap["Cust_BD"]];
    if (birthDate instanceof Date) {
      birthDate = Utilities.formatDate(birthDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    return {
      Cust_N: row[headerMap["Cust_N"]]?.toString().trim() || "",
      Cust_Sex: row[headerMap["Cust_Sex"]] || "",
      Cust_BD: birthDate || "",
      Cust_Add: row[headerMap["Cust_Add"]] || "",
      Sup_N: row[headerMap["Sup_N"]]?.toString().trim() || "",
      DSup_N: row[headerMap["DSup_N"]]?.toString().trim() || "",
      Sup_Tel: row[headerMap["Sup_Tel"]]?.toString() || "",
      Office_Tel: row[headerMap["Office_Tel"]]?.toString() || "",
      Cust_EC: row[headerMap["Cust_EC"]]?.toString().trim() || "",
      Cust_EC_Tel: row[headerMap["Cust_EC_Tel"]]?.toString() || "",
      Cust_LTC_Code: row[headerMap["Cust_LTC_Code"]] || "",
      LTC_Code: row[headerMap["LTC_Code"]] || "",
      Form_Url: row[headerMap["Form_Url"]] || "",
    };
  });
  return results;
}

/**
 * 取得「Cust」與「OldCust」工作表的完整個案資料
 * @returns {Object} { active: Array<Object>, archived: Array<Object> }
 */
function getAllCustData() {
  const activeData = getFormattedDataFromSheet("Cust");
  const archivedData = getFormattedDataFromSheet("OldCust");
  return { active: activeData, archived: archivedData };
}

/**
 * 新增個案資料
 */
function addCustData(formObj) {
  const sheet = MainSpreadsheet.getSheetByName("Cust");

  const allData = getAllCustData();
  const allNames = [...allData.active.map(c => c.Cust_N), ...allData.archived.map(c => c.Cust_N)];
  if (allNames.includes(formObj.custName)) {
    return { success: false, message: "該個案姓名已存在 (包含封存名單)！" };
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
    formObj.custltcCode || "",
    formObj.ltcCode || "",
    formObj.formUrl || "",
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
  // 清除個案基本資料快取
  CacheService.getScriptCache().remove("CustInfoMap");
  CacheService.getScriptCache().remove("CustN_All");
  CacheService.getScriptCache().remove("SRServer01_InitData");
  return { success: true, message: "新增成功！新增成功並已完成姓名排序！" };
}

/**
 * 更新個案資料
 */
function updateCustData(formObj, sheetName = "Cust") {
  const sheet = MainSpreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, message: `找不到工作表: ${sheetName}` };
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameColIdx = headers.indexOf("Cust_N");

  if (nameColIdx === -1) {
    return { success: false, message: `工作表 ${sheetName} 中找不到 'Cust_N' 欄位。` };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameColIdx] == formObj.custName) {
      const rowNum = i + 1;

      // 根據標題順序建立要更新的資料陣列
      const newRowData = headers.map(header => {
        switch (header) {
          case 'Cust_N': return formObj.custName;
          case 'Cust_Sex': return formObj.custSex;
          case 'Cust_BD': return formObj.custBirth;
          case 'Cust_Add': return formObj.custAddr;
          case 'Sup_N': return formObj.supName;
          case 'DSup_N': return formObj.dsupName;
          case 'Sup_Tel': return "'" + formObj.supTel;
          case 'Office_Tel': return "'" + formObj.officeTel;
          case 'Cust_EC': return formObj.ecName;
          case 'Cust_EC_Tel': return "'" + formObj.ecTel;
          case 'Cust_LTC_Code': return formObj.custltcCode || "";
          case 'LTC_Code': return formObj.ltcCode || "";
          case 'Form_Url': return formObj.formUrl || "";
          default:
            const colIdx = headers.indexOf(header);
            return data[i][colIdx]; // 保留其他欄位的原始值
        }
      });

      sheet.getRange(rowNum, 1, 1, newRowData.length).setValues([newRowData]);

      // 清除個案基本資料快取
      CacheService.getScriptCache().remove("CustInfoMap");
      CacheService.getScriptCache().remove("SRServer01_InitData");
      return { success: true, message: "資料更新成功！" };
    }
  }
  return { success: false, message: `在 ${sheetName} 找不到該個案資料，無法更新。` };
}

/**
 * 刪除個案資料
 */
function deleteCustData(custName, sheetName = "OldCust") {
  const sheet = MainSpreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return { success: false, message: `找不到工作表: ${sheetName}` };
  }
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == custName) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove("CustInfoMap");
      CacheService.getScriptCache().remove("CustN_All");
      CacheService.getScriptCache().remove("SRServer01_InitData");
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: `在 ${sheetName} 找不到資料，刪除失敗。` };
}

/**
 * 移動個案資料 (封存/解封)
 */
function moveCustomer(custName, fromSheetName, toSheetName) {
  const fromSheet = MainSpreadsheet.getSheetByName(fromSheetName);
  const toSheet = MainSpreadsheet.getSheetByName(toSheetName);

  if (!fromSheet || !toSheet) return { success: false, message: "來源或目標工作表不存在。" };

  const fromData = fromSheet.getDataRange().getValues();
  let rowToMove = null;
  let rowIndex = -1;

  for (let i = 1; i < fromData.length; i++) {
    if (fromData[i][0] == custName) {
      rowToMove = fromData[i];
      rowIndex = i + 1;
      break;
    }
  }

  if (!rowToMove) return { success: false, message: `在 ${fromSheetName} 中找不到 ${custName}。` };

  // 檢查目標工作表是否已存在相同個案
  const toData = toSheet.getDataRange().getValues();
  let duplicateRowIndex = -1;

  for (let i = 1; i < toData.length; i++) {
    if (toData[i][0] == custName) {
      duplicateRowIndex = i + 1;
      break;
    }
  }

  if (duplicateRowIndex !== -1) {
    // 如果存在，則更新該列資料
    toSheet.getRange(duplicateRowIndex, 1, 1, rowToMove.length).setValues([rowToMove]);
  } else {
    // 如果不存在，則新增
    toSheet.appendRow(rowToMove);
  }

  fromSheet.deleteRow(rowIndex);

  // 排序目標工作表
  const toLastRow = toSheet.getLastRow();
  if (toLastRow > 1) {
    toSheet.getRange(2, 1, toLastRow - 1, toSheet.getLastColumn()).sort({ column: 1, ascending: true });
  }

  CacheService.getScriptCache().remove("CustInfoMap");
  CacheService.getScriptCache().remove("CustN_All");
  CacheService.getScriptCache().remove("SRServer01_InitData");

  return { success: true, message: `已將 ${custName} 從 ${fromSheetName} 移至 ${toSheetName}。` };
}

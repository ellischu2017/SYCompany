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
  // 載入前先執行全量同步與資料清理
  syncAllCustDataFromSource();
  const activeData = getFormattedDataFromSheet("Cust");
  const archivedData = getFormattedDataFromSheet("OldCust");
  return { active: activeData, archived: archivedData };
}

/**
 * 從來源 (Case Reports) 同步個案資料，並執行去重、衝突處理與編碼合併
 * 流程符合 1-8 步驟要求
 */
function syncAllCustDataFromSource() {
  logSystemActivity('INFO', 'CustModule', '開始執行全量個案資料同步與清理...');
  let sourceSS;
  try {
    const sourceObj = getTargetsheet("SYTemp", "Case_Reports");
    if (!sourceObj || !sourceObj.Spreadsheet) throw new Error("在 SYTemp 中找不到 Case_Reports 的設定");
    sourceSS = sourceObj.Spreadsheet;
  } catch (e) {
    logSystemActivity('ERROR', 'CustModule', '同步中止，無法取得來源試算表: ' + e.toString());
    return;
  }
  
  const activeSource = sourceSS.getSheetByName("個案清單");
  const archiveSource = sourceSS.getSheetByName("結案個案清單");
  const custSheet = MainSpreadsheet.getSheetByName("Cust");
  const oldCustSheet = MainSpreadsheet.getSheetByName("OldCust");

  const performSync = (srcSheet, tarSheet) => {
    if (!srcSheet || !tarSheet) return null;
    const srcData = srcSheet.getDataRange().getValues();
    const tarData = tarSheet.getDataRange().getValues();
    const srcHeaders = srcData[0];
    const tarHeaders = tarData[0];
    const srcRows = srcData.slice(1);
    const tarRows = tarData.slice(1);

    const srcIdx = { name: 0, sex: 1, bd: 2, add: 3, formurl: 4, ltc: 5 };
    const tarColMap = getColIndicesMap(tarHeaders, ["Cust_N", "Cust_Sex", "Cust_BD", "Cust_Add", "Cust_LTC_Code", "Form_Url", "LTC_Code"]);
    
    if (tarColMap.Cust_N === -1) return null;

    const tarNameMap = new Map();
    tarRows.forEach((row, i) => {
      const name = String(row[tarColMap.Cust_N] || "").trim();
      if (name) tarNameMap.set(name, i);
    });

    srcRows.forEach(sRow => {
      const sName = String(sRow[srcIdx.name] || "").trim();
      if (!sName) return;
      
      const sBD = sRow[srcIdx.bd] instanceof Date 
        ? Utilities.formatDate(sRow[srcIdx.bd], "GMT+8", "yyyy/M/d") 
        : sRow[srcIdx.bd];

      const newData = { sex: sRow[srcIdx.sex], bd: sBD, add: sRow[srcIdx.add], ltc: sRow[srcIdx.ltc], formurl: sRow[srcIdx.formurl] };

      if (tarNameMap.has(sName)) {
        const idx = tarNameMap.get(sName);
        tarRows[idx][tarColMap.Cust_Sex] = newData.sex;
        tarRows[idx][tarColMap.Cust_BD] = newData.bd;
        tarRows[idx][tarColMap.Cust_Add] = newData.add;
        tarRows[idx][tarColMap.Cust_LTC_Code] = newData.ltc;
        tarRows[idx][tarColMap.Form_Url] = newData.formurl;
      } else {
        const newRow = new Array(tarHeaders.length).fill("");
        newRow[tarColMap.Cust_N] = sName;
        newRow[tarColMap.Cust_Sex] = newData.sex;
        newRow[tarColMap.Cust_BD] = newData.bd;
        newRow[tarColMap.Cust_Add] = newData.add;
        newRow[tarColMap.Cust_LTC_Code] = newData.ltc;
        newRow[tarColMap.Form_Url] = newData.formurl;
        tarRows.push(newRow);
      }
    });
    return { headers: tarHeaders, rows: tarRows, colMap: tarColMap };
  };

  let activeRes = performSync(activeSource, custSheet);
  let archiveRes = performSync(archiveSource, oldCustSheet);
  if (!activeRes || !archiveRes) return;

  // 5. & 6. 清理重複：確保 Cust_N 唯一
  const dedupe = (res) => {
    const seen = new Set();
    res.rows = res.rows.filter(row => {
      const n = String(row[res.colMap.Cust_N] || "").trim();
      if (!n || seen.has(n)) return false;
      seen.add(n);
      return true;
    });
  };
  dedupe(activeRes); dedupe(archiveRes);

  // 8. 處理 LTC_Code 合併：將 Cust_LTC_Code 的編碼增量存入 LTC_Code (先行處理以便後續拷貝)
  const processCodes = (res) => {
    res.rows.forEach(row => {
      row[res.colMap.LTC_Code] = mergeLtcCodeStrings(
        row[res.colMap.Cust_LTC_Code],
        row[res.colMap.LTC_Code]
      );
    });
  };
  processCodes(activeRes); processCodes(archiveRes);

  // 7. 交叉檢查與遷移：若發現 OldCust 中有相同個案，則從 Cust 拷貝資料到 OldCust，再刪除 Cust 中的資料
  const activeMap = new Map();
  activeRes.rows.forEach(r => {
    const n = String(r[activeRes.colMap.Cust_N] || "").trim().toUpperCase();
    if (n) activeMap.set(n, r);
  });

  const finalArchiveNames = new Set();
  archiveRes.rows.forEach(arcRow => {
    const n = String(arcRow[archiveRes.colMap.Cust_N] || "").trim().toUpperCase();
    if (!n) return;
    finalArchiveNames.add(n);

    if (activeMap.has(n)) {
      const actRow = activeMap.get(n);
      // 執行「拷貝」動作：將執行中表的最新資料覆蓋至結案表
      arcRow[archiveRes.colMap.Cust_Sex] = actRow[activeRes.colMap.Cust_Sex];
      arcRow[archiveRes.colMap.Cust_BD] = actRow[activeRes.colMap.Cust_BD];
      arcRow[archiveRes.colMap.Cust_Add] = actRow[activeRes.colMap.Cust_Add];
      arcRow[archiveRes.colMap.Form_Url] = actRow[activeRes.colMap.Form_Url];
      arcRow[archiveRes.colMap.LTC_Code] = mergeLtcCodeStrings(actRow[activeRes.colMap.LTC_Code], arcRow[archiveRes.colMap.LTC_Code]);
    }
  });

  // 執行「刪除」動作：從 activeRes 中過濾掉所有已存在於 OldCust 的名單
  activeRes.rows = activeRes.rows.filter(r => {
    const n = String(r[activeRes.colMap.Cust_N] || "").trim().toUpperCase();
    return !finalArchiveNames.has(n);
  });

  const finalize = (sheet, res) => {
    sheet.clearContents();
    const out = [res.headers, ...res.rows];
    sheet.getRange(1, 1, out.length, res.headers.length).setValues(out);
    sheet.getRange(2, 1, res.rows.length, res.headers.length).sort({ column: 1, ascending: true });
  };
  finalize(custSheet, activeRes); finalize(oldCustSheet, archiveRes);

  CacheService.getScriptCache().remove("CustInfoMap");
  CacheService.getScriptCache().remove("CustN_All");
  CacheService.getScriptCache().remove("SRServer01_InitData");
  logSystemActivity('INFO', 'CustModule', '全量個案同步與清理完成。');
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
  const nameColIdx = getColIndex(headers, "Cust_N");

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
  const headers = data[0];
  const nameColIdx = getColIndex(headers, "Cust_N");

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameColIdx] == custName) {
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

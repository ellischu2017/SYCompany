/**
 * SRServer.gs - 服務紀錄單管理模組
 * 提供服務紀錄的查詢、新增、修改、刪除操作
 */

/**
 * 初始化管理員頁面資料 (包含需求 1: 同步檢查)
 */
function getAdminInitialData() {
  const userSheet = MainSpreadsheet.getSheetByName("User");
  const users = userSheet
    ? userSheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map((r) => ({ name: r[0], email: r[1] }))
    : [];

  const custSheet = MainSpreadsheet.getSheetByName("Cust");
  const customers = custSheet
    ? custSheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map((r) => r[0])
    : [];

  const ltcSheet = MainSpreadsheet.getSheetByName("LTC_Code");
  const ltcCodes = ltcSheet
    ? ltcSheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map((r) => r[0])
    : [];

  const tempsheet = getTargetsheet("SYTemp", "SYTemp");
  const tempUserSheet = tempsheet.getSheetByName("User");
  const hasTempData = tempsheet && tempUserSheet.getLastRow() > 1;

  return {
    users: users,
    customers: customers,
    ltcCodes: ltcCodes,
    triggerSync: hasTempData,
  };
}

/**
 * 跨表查詢邏輯 (需求 5, 6, 7)
 * 核心查詢邏輯：使用字串比對防止日期偏移
 */
function queryServiceRecord(params) {
  const { date, custN, userN, payType, srId } = params;
  const year = date.split("-")[0];
  const syYearKey = "SY" + year;

  const recUrlData = queryRecUrlData(syYearKey);
  if (recUrlData.found) {
    try {
      const targetSs = SpreadsheetApp.openByUrl(recUrlData.rowResult.SY_Url);
      const result = searchAcrossSheets(targetSs, params);
      if (result)
        return { ...result, source: "SYCompany", ssId: targetSs.getId() };
    } catch (e) {
      console.log("年度表開啟失敗");
    }
  }

  const ssTemp = getTargetsheet("SYTemp", "SYTemp");
  const tempSheet = ssTemp.getSheetByName("SR_Data");
  const tempResult = searchSheet(tempSheet, params);

  if (tempResult) {
    return { ...tempResult, source: "SYTemp", ssId: ssTemp.getId() };
  }

  return { found: false };
}

/**
 * 搜尋所有工作表
 */
function searchAcrossSheets(ss, p) {
  const sheets = MainSpreadsheet.getSheets();
  for (let sheet of sheets) {
    const res = searchSheet(sheet, p);
    if (res) return { ...res, sheetName: sheet.getName() };
  }
  return null;
}

/**
 * 搜尋指定工作表
 */
function searchSheet(sheet, p) {
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    let rowDateStr =
      data[i][0] instanceof Date
        ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy-MM-dd")
        : String(data[i][0]);
    console.log(p);
    if (
      rowDateStr === p.date &&
      data[i][2] === p.custN &&
      data[i][3] === p.userN &&
      data[i][4] === p.payType &&
      data[i][5] === p.srId
    ) {
      return { found: true, data: data[i], rowIndex: i + 1 };
    }
  }
  return { found: false, data: null };
}

/**
 * 管理員 CRUD 操作
 */
function manageSRData(action, form, sourceInfo) {
  const row = [
    form.date,
    form.email,
    form.custN,
    form.userN,
    form.payType,
    form.srId,
    form.srRec,
    form.loc,
    form.mood,
    form.spcons,
  ];

  if (action === "add") {
    const tempSheet = MainSpreadsheet.getSheetByName("SR_Data");
    tempSheet.appendRow(row);
    tempSheet.getRange(tempSheet.getLastRow(), 1).setNumberFormat("yyyy-MM-dd");
    return "已成功新增至 SYTemp。";
  }

  const targetSs = SpreadsheetApp.openById(sourceInfo.ssId);
  const targetSheet =
    sourceInfo.source === "SYTemp"
      ? targetSs.getSheetByName("SR_Data")
      : targetSs.getSheetByName(sourceInfo.sheetName);

  if (action === "update") {
    targetSheet
      .getRange(sourceInfo.rowIndex, 1, 1, row.length)
      .setValues([row]);
    targetSheet.getRange(sourceInfo.rowIndex, 1).setNumberFormat("yyyy-MM-dd");
    return `資料已於 ${sourceInfo.source} 更新成功。`;
  } else if (action === "delete") {
    targetSheet.deleteRow(sourceInfo.rowIndex);
    return `資料已從 ${sourceInfo.source} 刪除。`;
  }
}

/**
 * 核心處理：處理服務紀錄 (查詢/新增/修改/刪除)
 * 對應需求 2 (儲存新 User) & 需求 3 (寫入 SR_Data)
 */
function processSRData(formObj, actionType) {
  try {
    var targetSs = getTargetsheet("SYTemp", "SYTemp");

    if (formObj.userTel && formObj.userName && formObj.email) {
      var userSheet = targetSs.getSheetByName("User");
      if (!userSheet) userSheet = targetSs.insertSheet("User");

      var uData = userSheet.getDataRange().getValues();
      var uExists = false;
      for (var k = 1; k < uData.length; k++) {
        if (uData[k][1] === formObj.email) {
          uExists = true;
          break;
        }
      }

      if (!uExists) {
        userSheet.appendRow([
          formObj.userName,
          formObj.email,
          "'" + formObj.userTel,
        ]);
      }
    }

    var targetSheet = targetSs.getSheetByName("SR_Data");
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet("SR_Data");
      targetSheet.appendRow([
        "Date",
        "E-mail",
        "CUST_N",
        "USER_N",
        "Pay_Type",
        "SR_ID",
        "SR_REC",
        "LOC",
        "MOOD",
        "SPCONS",
      ]);
    }

    const rowData = [
      formObj.date,
      formObj.email,
      formObj.custName,
      formObj.userName,
      formObj.payType || "補助",
      formObj.srId,
      formObj.srRec || "",
      formObj.loc || "清醒",
      formObj.mood || "穩定",
      formObj.spcons || "無",
    ];

    if (actionType === "add") {
      targetSheet.appendRow(rowData);
      return { success: true, message: "新增紀錄成功" };
    }

    var data = targetSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let sheetDate =
        data[i][0] instanceof Date
          ? Utilities.formatDate(
              data[i][0],
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : data[i][0].toString();

      if (
        sheetDate === formObj.date &&
        data[i][2].toString().trim() === formObj.custName.trim() &&
        data[i][3].toString().trim() === formObj.userName.trim() &&
        data[i][4].toString().trim() === formObj.payType.trim() &&
        data[i][5].toString().trim() === formObj.srId.trim()
      ) {
        if (actionType === "query") {
          return {
            found: true,
            data: {
              date: sheetDate,
              email: data[i][1],
              custName: data[i][2],
              userName: data[i][3],
              payType: data[i][4],
              srId: data[i][5],
              srRec: data[i][6],
              loc: data[i][7],
              mood: data[i][8],
              spcons: data[i][9],
            },
          };
        } else if (actionType === "update") {
          targetSheet.getRange(i + 1, 1, 1, 10).setValues([rowData]);
          return { success: true, message: "資料已更新" };
        } else if (actionType === "delete") {
          targetSheet.deleteRow(i + 1);
          return { success: true, message: "資料已刪除" };
        }
      }
    }

    if (actionType === "query") {
      return { found: false, message: "在該日期下找不到對應的紀錄" };
    } else {
      return { success: false, message: "操作失敗，找不到該筆資料" };
    }
  } catch (e) {
    return { success: false, message: "錯誤：" + e.toString() };
  }
}

/**
 * 初始化 SR_server01 頁面所需的所有下拉選單資料
 */
function getSRServer01InitData() {
  var userEmail = Session.getActiveUser().getEmail();
  var currentUserName = "";
  var found = false;

  var localUserSheet = MainSpreadsheet.getSheetByName("User");
  if (localUserSheet) {
    var localData = localUserSheet.getDataRange().getValues();
    for (var i = 1; i < localData.length; i++) {
      if (localData[i][1] === userEmail) {
        currentUserName = localData[i][0];
        found = true;
        break;
      }
    }
  }

  if (!found) {
    try {
      var remoteSS = getTargetsheet("SYTemp", "SYTemp");
      var remoteUserSheet = remoteSS.getSheetByName("User");
      if (remoteUserSheet) {
        var remoteData = remoteUserSheet.getDataRange().getValues();
        for (var j = 1; j < remoteData.length; j++) {
          if (remoteData[j][1] === userEmail) {
            currentUserName = remoteData[j][0];
            found = true;
            break;
          }
        }
      }
    } catch (e) {
      console.log("外部 User 表查詢失敗: " + e.toString());
    }
  }

  return {
    custNames: getCustList(),
    srIds: getLtcCodeList(),
    currentUserName: currentUserName,
    userEmail: userEmail || "",
  };
}

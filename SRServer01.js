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

/**
 * 核心處理：處理服務紀錄 (查詢/新增/修改/刪除)
 * 對應需求 1 (儲存新 User) & 需求 3 (寫入 SR_Data)
 */
function processSR01Data(formObj, actionType) {
  try {
    var targetSs = getTargetsheet("SYTemp", "SYTemp");

    if (formObj.userTel && formObj.userName && formObj.email) {
      var userSheet = targetSs.getSheetByName("User");
      if (!userSheet) userSheet = targetSs.insertSheet("User");

      var uData = userSheet.getDataRange().getValues();
      var uExists = false;
      for (var k = 0; k < uData.length; k++) {
        if (uData[k][0] === formObj.email) {
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
        "SRTimes",
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
      formObj.SRTimes || "1",
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
          ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy-MM-dd")
          : data[i][0].toString();

      if (
        sheetDate === formObj.date &&
        data[i][1].toString().trim() === formObj.SRTimes.trim() &&
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
              SRTimes: data[i][1],
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
 * 根據居服員姓名取得「常用個案」與「其他個案」
 */
function getCustomClientLists(userName) {
  try {
    // 1. 從 SYTemp > SR_Data 找出該人員服務過的個案 (常用名單)
    var targetSs = getTargetsheet("SYTemp", "SYTemp");
    var srDataSheet = targetSs.getSheetByName("SR_Data");
    var favClients = [];

    if (srDataSheet) {
      var data = srDataSheet.getDataRange().getValues();
      var nameSet = new Set();
      for (var i = 1; i < data.length; i++) {
        if (data[i][3] === userName) {
          // USER_N 在第 4 欄
          nameSet.add(data[i][2]); // CUST_N 在第 3 欄
        }
      }
      favClients = Array.from(nameSet).sort();
    }

    // 2. 從 SYCompany > Cust 取得所有個案
    var localCustSheet = MainSpreadsheet.getSheetByName("Cust");
    var allClients = [];
    if (localCustSheet) {
      var custData = localCustSheet.getDataRange().getValues();
      for (var j = 1; j < custData.length; j++) {
        if (custData[j][0]) allClients.push(custData[j][0].toString()); // 假設 Cust_N 在第 1 欄
      }
    }

    // 3. 過濾出「不在常用名單」中的其他個案
    var otherClients = allClients
      .filter(function (name) {
        return !favClients.includes(name);
      })
      .sort();

    return {
      favClients: favClients,
      otherClients: otherClients,
    };
  } catch (e) {
    console.log("取得客製化名單失敗: " + e.toString());
    return { favClients: [], otherClients: [] };
  }
}

/**
 * 根據個案姓名取得客製化的服務編碼清單 (優先顯示常用編碼)
 */
function getCustomSrIdList(custName) {
  try {
    // 1. 從 SYTemp > SR_Data 找出該個案曾使用過的編碼 (常用編碼)
    var targetSs = getTargetsheet("SYTemp", "SYTemp"); // 使用 Utilities.js 中的工具
    var srDataSheet = targetSs.getSheetByName("SR_Data");
    var favSrIds = [];

    if (srDataSheet) {
      var srData = srDataSheet.getDataRange().getValues();
      var idSet = new Set();
      for (var i = 1; i < srData.length; i++) {
        // CUST_N 在第 3 欄 (index 2)，SR_ID 在第 6 欄 (index 5)
        if (srData[i][2] === custName && srData[i][5]) {
          idSet.add(srData[i][5].toString());
        }
      }
      favSrIds = Array.from(idSet).sort();
    }

    // 2. 從 SYCompany (本機試算表) > LTC_Code 取得所有編碼
    var ltcSheet = MainSpreadsheet.getSheetByName("LTC_Code");
    var allSrIds = [];
    if (ltcSheet) {
      var ltcData = ltcSheet.getDataRange().getValues();
      for (var j = 1; j < ltcData.length; j++) {
        if (ltcData[j][0]) allSrIds.push(ltcData[j][0].toString()); // 假設 SR_ID 在 A 欄
      }
    }

    // 3. 過濾出尚未出現在常用名單中的其餘編碼
    var otherSrIds = allSrIds
      .filter(function (id) {
        return !favSrIds.includes(id);
      })
      .sort();

    // 4. 合併兩者：常用在前，其餘在後
    return favSrIds.concat(otherSrIds);
  } catch (e) {
    console.log("取得客製化編碼清單失敗: " + e.toString());
    // 發生錯誤時回傳預設的完整編碼清單
    return typeof getLtcCodeList === "function" ? getLtcCodeList() : [];
  }
}

// 修改原有的 getSRServer01InitData，讓它在初始化時就觸發名單載入
// 在 HTML 的 successHandler 中，如果 currentUserName 有值，就呼叫 fetchClientLists(data.currentUserName)

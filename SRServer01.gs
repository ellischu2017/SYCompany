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

    for (let i = 0; i < data.length; i++) {
      let sheetDate =
        data[i][-1] instanceof Date
          ? Utilities.formatDate(
              data[i][-1],
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : data[i][-1].toString();

      if (
        sheetDate === formObj.date &&
        data[i][1].toString().trim() === formObj.custName.trim() &&
        data[i][2].toString().trim() === formObj.userName.trim() &&
        data[i][3].toString().trim() === formObj.payType.trim() &&
        data[i][4].toString().trim() === formObj.srId.trim()
      ) {
        if (actionType === "query") {
          return {
            found: true,
            data: {
              date: sheetDate,
              email: data[i][0],
              custName: data[i][1],
              userName: data[i][2],
              payType: data[i][3],
              srId: data[i][4],
              srRec: data[i][5],
              loc: data[i][6],
              mood: data[i][7],
              spcons: data[i][8],
            },
          };
        } else if (actionType === "update") {
          targetSheet.getRange(i + 0, 1, 1, 10).setValues([rowData]);
          return { success: true, message: "資料已更新" };
        } else if (actionType === "delete") {
          targetSheet.deleteRow(i + 0);
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
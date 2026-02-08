/**
 * 初始化 SR_server01 頁面所需的所有資料
 * 修改點：同時讀取 SYCompany (主表) 與 SYTemp (更新表) 的 User 資料
 */
function getSRServer01InitData() {
  var userEmail = Session.getActiveUser().getEmail();
  
  // 1. 讀取 SYCompany > User (主名單)
  var userSheet = MainSpreadsheet.getSheetByName("User");
  var userMap = new Map(); // 使用 Map 以便用 Name 進行合併

  if (userSheet) {
    var rawUsers = userSheet.getDataRange().getValues();
    var headers = rawUsers[0];
    
    var idxName = getColIndex(headers, "User_N");
    var idxEmail = getColIndex(headers, "User_Email");
    if(idxEmail === -1) idxEmail = getColIndex(headers, "Email");
    var idxCust = getColIndex(headers, "Cust_N");
    
    if (idxName !== -1) {
      for (var i = 1; i < rawUsers.length; i++) {
        var row = rawUsers[i];
        var name = String(row[idxName]).trim();
        if(name) {
          userMap.set(name, {
            name: name,
            email: idxEmail !== -1 ? String(row[idxEmail]).trim() : "",
            favCustStr: idxCust !== -1 ? String(row[idxCust]) : ""
          });
        }
      }
    }
  }

  // 2. 讀取 SYTemp > User (更新名單) 並覆蓋/合併
  try {
    var tempSS = getTargetsheet("SYTemp", "SYTemp");
    var tempUserSheet = tempSS.getSheetByName("User");
    if (tempUserSheet) {
      var tempUsers = tempUserSheet.getDataRange().getValues();
      var tHeaders = tempUsers[0];
      var tNameIdx = getColIndex(tHeaders, "User_N");
      var tEmailIdx = getColIndex(tHeaders, "User_Email");
      if(tEmailIdx === -1) tEmailIdx = getColIndex(tHeaders, "Email");

      if (tNameIdx !== -1 && tEmailIdx !== -1) {
        for (var j = 1; j < tempUsers.length; j++) {
          var tRow = tempUsers[j];
          var tName = String(tRow[tNameIdx]).trim();
          var tEmail = String(tRow[tEmailIdx]).trim();
          
          if (tName) {
            if (userMap.has(tName)) {
              // 若主名單已有，則更新 Email (以 Temp 為準)
              var uObj = userMap.get(tName);
              if (tEmail) uObj.email = tEmail;
            } else {
              // 若主名單沒有，則新增 (視需求，通常只會更新既有員工的綁定)
              userMap.set(tName, {
                name: tName,
                email: tEmail,
                favCustStr: "" // Temp 表可能沒有個案關聯資料
              });
            }
          }
        }
      }
    }
  } catch (e) {
    console.log("讀取 SYTemp > User 失敗: " + e.toString());
  }

  // 轉回陣列
  var userData = Array.from(userMap.values());

  // 3. 取得 Cust 資料
  var custSheet = MainSpreadsheet.getSheetByName("Cust");
  var custData = [];
  if (custSheet) {
    var rawCusts = custSheet.getDataRange().getValues();
    var cHeaders = rawCusts[0];
    var idxCName = getColIndex(cHeaders, "Cust_N");
    var idxCLTC = getColIndex(cHeaders, "LTC_Code");
    
    if (idxCName !== -1 && idxCLTC !== -1) {
      for (var k = 1; k < rawCusts.length; k++) {
        custData.push({
          name: rawCusts[k][idxCName],
          ltcCodeStr: String(rawCusts[k][idxCLTC])
        });
      }
    }
  }

  // 4. 取得 LTC_Code 資料
  var ltcSheet = MainSpreadsheet.getSheetByName("LTC_Code");
  var ltcIds = [];
  if (ltcSheet) {
    var rawLtc = ltcSheet.getDataRange().getValues();
    var lHeaders = rawLtc[0];
    var idxSRID = getColIndex(lHeaders, "SR_ID");
    var targetIdx = idxSRID !== -1 ? idxSRID : 0;
    
    for (var m = 1; m < rawLtc.length; m++) {
       var code = String(rawLtc[m][targetIdx]).trim();
       if(code && !ltcIds.includes(code)) {
         ltcIds.push(code);
       }
    }
  }

  return {
    email: userEmail,
    users: userData,
    custs: custData,
    ltcCodes: ltcIds
  };
}

/**
 * 核心處理：處理服務紀錄 (查詢/新增/修改/刪除)
 * 修改點：若為新綁定 (isNewBinding)，寫入 SYTemp > User
 */
function processSR01Data(formObj, actionType) {
  try {
    var targetSs = getTargetsheet("SYTemp", "SYTemp");

    // --- 處理需求：若前端標記為新綁定，則更新 SYTemp > User 表 ---
    if ((actionType === "add" || actionType === "update") && formObj.isNewBinding === true) {
      var tempUserSheet = targetSs.getSheetByName("User");
      if (!tempUserSheet) {
        tempUserSheet = targetSs.insertSheet("User");
        tempUserSheet.appendRow(["User_N", "User_Email", "User_Tel"]); // 建立標題
      }

      var tData = tempUserSheet.getDataRange().getValues();
      var tHeaders = tData[0];
      var tNameIdx = getColIndex(tHeaders, "User_N");
      var tEmailIdx = getColIndex(tHeaders, "User_Email");
      if(tEmailIdx === -1) tEmailIdx = getColIndex(tHeaders, "Email");
      var tTelIdx = getColIndex(tHeaders, "User_Tel");

      // 如果找不到對應欄位，嘗試使用預設欄位索引
      if (tNameIdx === -1) tNameIdx = 0;
      if (tEmailIdx === -1) tEmailIdx = 1;
      if (tTelIdx === -1) tTelIdx = 2;

      var foundInTemp = false;

      // 檢查是否已存在於 SYTemp，若有則更新
      for (var k = 1; k < tData.length; k++) {
        if (String(tData[k][tNameIdx]) === formObj.userName) {
          tempUserSheet.getRange(k + 1, tEmailIdx + 1).setValue(formObj.email);
          tempUserSheet.getRange(k + 1, tTelIdx + 1).setValue(formObj.userTel);
          foundInTemp = true;
          break; 
        }
      }

      // 若 SYTemp 中沒有此人，則新增
      if (!foundInTemp) {
        var newRow = [];
        newRow[tNameIdx] = formObj.userName;
        newRow[tEmailIdx] = formObj.email;
        newRow[tTelIdx] = formObj.userTel;
        tempUserSheet.appendRow(newRow);
      }
    }

    // --- 以下為標準 SR_Data 處理邏輯 ---
    var targetSheet = targetSs.getSheetByName("SR_Data");
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet("SR_Data");
      targetSheet.appendRow([
        "Date", "SRTimes", "CUST_N", "USER_N", "Pay_Type", "SR_ID", "SR_REC", "LOC", "MOOD", "SPCONS"
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

function getSPCONS(formObj) {
  try {
    var targetSs = getTargetsheet("SYTemp", "SYTemp");
    var spconsSheet = targetSs.getSheetByName("SR_Data");
    if (!spconsSheet) return {data:{ loc: "清醒", mood: "穩定", spcons: "無" } };

    var data = spconsSheet.getDataRange().getValues();
    const targetDateStr = formObj.date;

    for (let i = data.length - 1; i >= 1; i--) {
      let row = data[i];
      if (
        row[3].toString().trim() === formObj.userName.trim() &&
        row[2].toString().trim() === formObj.custName.trim() &&
        row[1].toString().trim() === formObj.SRTimes.trim()
      ) {
        let sheetDate =
          row[0] instanceof Date
            ? Utilities.formatDate(row[0], "GMT+8", "yyyy-MM-dd")
            : row[0].toString();

        if (sheetDate === targetDateStr) {          
          return {data:{ loc: data[i][7], mood: data[i][8], spcons: data[i][9] }};          
        }
      }
    }
  } catch (e) {
    console.log("取得特殊狀況失敗: " + e.toString());
  }
  return {data:{ loc: "清醒", mood: "穩定", spcons: "無" } };
}
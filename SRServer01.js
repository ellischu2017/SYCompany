/**
 * 初始化 SR_server01 頁面所需的所有資料
 * 修改點：同時讀取 SYCompany (主表) 與 SYTemp (更新表) 的 User 資料
 */
function getSRServer01InitData() {
  var userMap = new Map(); // 使用 Map 以便用 Name 進行合併
  let data;

    // --- 以下為從試算表讀取資料的原始邏輯 ---
    // 1. 讀取 SYCompany > User (主名單)
    var userSheet = MainSpreadsheet.getSheetByName("User");

    if (userSheet) {
      var rawUsers = userSheet.getDataRange().getValues();
      if (rawUsers.length > 0) {
        var headers = rawUsers[0];
        const userFields = ["User_N", "User_Email", "Email", "Cust_N"];
        const colMap = getColIndicesMap(headers, userFields);

        var idxName = colMap["User_N"];
        var idxEmail = colMap["User_Email"] !== -1 ? colMap["User_Email"] : colMap["Email"];
        var idxCust = colMap["Cust_N"];

        if (idxName !== -1) {
          for (var i = 1; i < rawUsers.length; i++) {
            var row = rawUsers[i];
            var name = String(row[idxName]).trim();
            if (name) {
              userMap.set(name, {
                name: name,
                email: idxEmail !== -1 ? String(row[idxEmail]).trim() : "",
                favCustStr: idxCust !== -1 ? String(row[idxCust]) : ""
              });
            }
          }
        }
      }
    }

    // 2. 讀取 SYTemp > User (更新名單) 並覆蓋/合併
    // 優化：只呼叫一次 getTargetsheet
    let tempSS;
    try {
      tempSS = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
      var tempUserSheet = tempSS.getSheetByName("User");
      if (tempUserSheet) {
        var tempUsers = tempUserSheet.getDataRange().getValues();
        if (tempUsers.length > 0) {
          var tHeaders = tempUsers[0];
          const tFields = ["User_N", "User_Email", "Email"];
          const tColMap = getColIndicesMap(tHeaders, tFields);

          var tNameIdx = tColMap["User_N"];
          var tEmailIdx = tColMap["User_Email"] !== -1 ? tColMap["User_Email"] : tColMap["Email"];

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
      if (rawCusts.length > 0) {
        var cHeaders = rawCusts[0];
        const custFields = ["Cust_N", "LTC_Code"];
        const cColMap = getColIndicesMap(cHeaders, custFields);
        var idxCName = cColMap["Cust_N"];
        var idxCLTC = cColMap["LTC_Code"];

        if (idxCName !== -1 && idxCLTC !== -1) {
          for (var k = 1; k < rawCusts.length; k++) {
            custData.push({
              name: rawCusts[k][idxCName],
              ltcCodeStr: String(rawCusts[k][idxCLTC])
            });
          }
        }
      }
    }

    // 4. 取得 LTC_Code 資料
    var ltcSheet = MainSpreadsheet.getSheetByName("LTC_Code");
    var ltcIds = [];
    if (ltcSheet) {
      var rawLtc = ltcSheet.getDataRange().getValues();
      if (rawLtc.length > 0) {
        var lHeaders = rawLtc[0];
        const ltcFields = ["SR_ID", "SR_Cont"];
        const lColMap = getColIndicesMap(lHeaders, ltcFields);
        var idxSRID = lColMap["SR_ID"];
        var idxCont = lColMap["SR_Cont"];
        var targetIdx = idxSRID !== -1 ? idxSRID : 0; // 若找不到標題，預設第一欄
        var seen = {};

        for (var k = 1; k < rawLtc.length; k++) {
          var code = rawLtc[k][targetIdx].toString().trim();
          var cont = idxCont !== -1 ? rawLtc[k][idxCont].toString().trim() : "";
          if (code && !seen[code]) {
            seen[code] = true;
            ltcIds.push({ id: code, cont: cont });
          }
        }
      }
    }

    // 5. 取得 SYTemp > SR_Data
    var srDataSheet = tempSS ? tempSS.getSheetByName("SR_Data") : null;
    var srData = [];
    if (srDataSheet) {
      var rawSrData = srDataSheet.getDataRange().getValues();
      if (rawSrData.length > 0) {
        var srHeaders = rawSrData[0];
        var dateColIdx = getColIndex(srHeaders, "Date");

        // Convert date objects to string to be able to JSON.stringify them for cache.
        // Also, we need headers on the client side.
        srData = rawSrData.map(function (row, index) {
          if (index === 0) return row; // keep headers
          if (dateColIdx > -1 && row[dateColIdx] instanceof Date) {
            row[dateColIdx] = Utilities.formatDate(row[dateColIdx], Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          return row;
        });
      }
    }
    data = {
      users: userData,
      custs: custData,
      ltcCodes: ltcIds,
      srData: srData,
    };

  return data;
}

/**
 * 核心處理：處理服務紀錄 (查詢/新增/修改/刪除)
 * 修改點：若為新綁定 (isNewBinding)，寫入 SYTemp > User
 */
function processSR01Data(formObj, actionType) {
  // 根據日期預先產生快取 key，以便後續清除
  let yearmonth = "";
  if (formObj && formObj.date) {
    const d = new Date(formObj.date);
    yearmonth = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyyMM");
  }

  const maxRetries = 3;
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      var targetSs = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;

      // --- 處理需求：若前端標記為新綁定，則更新 SYTemp > User 表 ---
      if ((actionType === "add" || actionType === "update") && formObj.isNewBinding === true) {
        var tempUserSheet = targetSs.getSheetByName("User");
        if (!tempUserSheet) {
          tempUserSheet = targetSs.insertSheet("User");
          tempUserSheet.appendRow(["User_N", "User_Email", "User_Tel"]); // 建立標題
        }

        var tData = tempUserSheet.getDataRange().getValues();
        var tHeaders = tData[0];
        const tFields = ["User_N", "User_Email", "Email", "User_Tel"];
        const tColMap = getColIndicesMap(tHeaders, tFields);

        var tNameIdx = tColMap["User_N"];
        var tEmailIdx = tColMap["User_Email"] !== -1 ? tColMap["User_Email"] : tColMap["Email"];
        var tTelIdx = tColMap["User_Tel"];

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
        UpdateRawResponse(formObj);
        // 清除對應月份的案主列表快取
        if (yearmonth) {
          CacheService.getScriptCache().remove("CustN_" + yearmonth);
          // 同步清除報表資料來源快取，確保月報表能抓到最新資料
          CacheService.getScriptCache().remove("DataMap_" + yearmonth);
        }
        return { success: true, message: "新增紀錄成功" };
      }

      var data = targetSheet.getDataRange().getValues();
      var headers = data[0];
      const targetFields = ["Date", "SRTimes", "CUST_N", "USER_N", "Pay_Type", "SR_ID", "SR_REC", "LOC", "MOOD", "SPCONS"];
      const colMap = getColIndicesMap(headers, targetFields);

      // 檢查關鍵欄位是否存在，若不存在則無法進行準確比對
      if (colMap["Date"] === -1 || colMap["CUST_N"] === -1) return { success: false, message: "錯誤：資料表欄位缺失" };

      for (let i = 1; i < data.length; i++) {
        let row = data[i];
        let sheetDate =
          row[colMap["Date"]] instanceof Date
            ? Utilities.formatDate(row[colMap["Date"]], "GMT+8", "yyyy-MM-dd")
            : String(row[colMap["Date"]]);

        if (
          sheetDate === formObj.date &&
          String(row[colMap["SRTimes"]]).trim() === formObj.SRTimes.trim() &&
          String(row[colMap["CUST_N"]]).trim() === formObj.custName.trim() &&
          String(row[colMap["USER_N"]]).trim() === formObj.userName.trim() &&
          String(row[colMap["Pay_Type"]]).trim() === formObj.payType.trim() &&
          String(row[colMap["SR_ID"]]).trim() === formObj.srId.trim()
        ) {
          if (actionType === "query") {
            return {
              found: true,
              data: {
                date: sheetDate,
                SRTimes: row[colMap["SRTimes"]],
                custName: row[colMap["CUST_N"]],
                userName: row[colMap["USER_N"]],
                payType: row[colMap["Pay_Type"]],
                srId: row[colMap["SR_ID"]],
                srRec: row[colMap["SR_REC"]],
                loc: row[colMap["LOC"]],
                mood: row[colMap["MOOD"]],
                spcons: row[colMap["SPCONS"]],
              },
            };
          } else if (actionType === "update") {
            // 更穩健的寫法：根據 colMap 寫入
            targetFields.forEach((field, idx) => {
              let colIdx = colMap[field];
              if (colIdx !== -1) {
                targetSheet.getRange(i + 1, colIdx + 1).setValue(rowData[idx]);
              }
            });

            UpdateRawResponse(formObj);
            // 清除對應月份的案主列表快取
            if (yearmonth) {
              CacheService.getScriptCache().remove("CustN_" + yearmonth);
              // 同步清除報表資料來源快取
              CacheService.getScriptCache().remove("DataMap_" + yearmonth);
            }
            return { success: true, message: "資料已更新" };
          } else if (actionType === "delete") {
            targetSheet.deleteRow(i + 1);
            // 清除對應月份的案主列表快取
            if (yearmonth) {
              CacheService.getScriptCache().remove("CustN_" + yearmonth);
              // 同步清除報表資料來源快取
              CacheService.getScriptCache().remove("DataMap_" + yearmonth);
            }
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
      console.warn(`processSR01Data 執行失敗 (第 ${attempt} 次重試): ${e.toString()}`);
      if (attempt === maxRetries) {
        return { success: false, message: "錯誤 (已重試 " + maxRetries + " 次)：" + e.toString() };
      }
      Utilities.sleep(1500 * attempt); // 延遲重試：1.5s, 3.0s, 4.5s
    }
  }
}

function getSPCONS(formObj) {
  try {
    var targetSs = getTargetsheet("SYTemp", "SYTemp").Spreadsheet;
    var spconsSheet = targetSs.getSheetByName("SR_Data");
    if (!spconsSheet) return { data: { loc: "清醒", mood: "穩定", spcons: "無" } };

    var data = spconsSheet.getDataRange().getValues();
    var headers = data[0];
    const targetFields = ["Date", "SRTimes", "CUST_N", "USER_N", "LOC", "MOOD", "SPCONS"];
    const colMap = getColIndicesMap(headers, targetFields);
    const targetDateStr = formObj.date;

    for (let i = data.length - 1; i >= 1; i--) {
      let row = data[i];
      if (
        String(row[colMap["USER_N"]]).trim() === formObj.userName.trim() &&
        String(row[colMap["CUST_N"]]).trim() === formObj.custName.trim() &&
        String(row[colMap["SRTimes"]]).trim() === formObj.SRTimes.trim()
      ) {
        let sheetDate =
          row[colMap["Date"]] instanceof Date
            ? Utilities.formatDate(row[colMap["Date"]], "GMT+8", "yyyy-MM-dd")
            : String(row[colMap["Date"]]);

        if (sheetDate === targetDateStr) {
          return { data: { loc: row[colMap["LOC"]], mood: row[colMap["MOOD"]], spcons: row[colMap["SPCONS"]] } };
        }
      }
    }
  } catch (e) {
    console.log("取得特殊狀況失敗: " + e.toString());
  }
  return { data: { loc: "清醒", mood: "穩定", spcons: "無" } };
}

/**
 * 根據當前登入者 Email 於 User 分頁查找姓名
 */
function getUserNameByEmail() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("User");
  
  if (!userSheet) return "查無使用者資料表";
  
  const userData = userSheet.getDataRange().getValues();
  // 假設第一欄是 User_N (姓名), 第二欄是 User_Email
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][1] === email) {
      return { userName: userData[i][0] , userEmail: email};
    }
  }

  const managerSheet = ss.getSheetByName("Manager");
  
  if (!managerSheet) return "查無使用者資料表";
  
  const managerData = managerSheet.getDataRange().getValues();
  // 假設第一欄是 User_N (姓名), 第二欄是 User_Email
  for (let i = 1; i < managerData.length; i++) {
    if (managerData[i][1] === email) {
      return { userName: managerData[i][0] , userEmail: email};
    }
  }
  return { userName: "訪客", userEmail: email };
}

/**
 * 將建議存入 SYTemp 檔案中的 Suggest 工作表
 */
function addSuggestion(formData) {
  try {
    // 驗證 formData
    if (!formData || !formData.date) {
      return { success: false, message: "表單資料為空或不完整" };
    }

    let ss;
    try {
      ss = getTargetsheet("SYTemp", "SYTemp");
    } catch (error) {
      return { success: false, message: "無法取得 SYTemp 試算表: " + error.message };
    }
    
    if (!ss) {
      return { success: false, message: "找不到 SYTemp 試算表" };
    }

    // 檢查 Suggest 工作表是否存在，不存在則建立
    let tempSheet = ss.getSheetByName("Suggest");
    if (!tempSheet) {
      tempSheet = ss.insertSheet("Suggest");
      // 設定第一列為標題
      tempSheet.appendRow(["Date","Su_Email","Su_N","Su_Rec","Deal"]);
    }

    // 準備資料
    const rowData = [
      formData.date || "",
      formData.suEmail || "",
      formData.suName || "",
      formData.suRec || ""
    ];

    Logger.log("寫入資料: " + JSON.stringify(rowData));
    
    // 新增資料列
    tempSheet.appendRow(rowData);

    // 按日期由新到舊排序 (Z to A)
    const dataRange = tempSheet.getDataRange();
    tempSheet.sort(1, false); // 第 1 欄 (日期)，false = 降序 (Z to A)

    return { success: true, message: "資料已寫入", data: rowData };
  } catch (e) {
    Logger.log("addSuggestion 錯誤: " + e.toString());
    return { success: false, message: "錯誤: " + e.toString() };
  }
}
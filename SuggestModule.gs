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
    // 1. 取得 SYTemp 試算表 (請替換為您實際的 SYTemp ID 或透過搜尋取得)
    // 假設您在 RecUrl 或某處存有 SYTemp 的 URL/ID
    const ss = getTargetsheet("SYTemp", "SYTemp");
    const tempSheet = ss.getSheetByName("Suggest"); // 直接存於主檔或連動到 SYTemp
    
    if (!tempSheet) {
      return { success: false, message: "找不到 Suggest 工作表" };
    }

    // 2. 寫入資料
    tempSheet.appendRow([
      formData.date,
      formData.suEmail,
      formData.suName,
      formData.suRec
    ]);    

    return { success: true, message: "資料已寫入" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
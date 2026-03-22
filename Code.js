
function doGet(e) {
  var page = e.parameter.page;

  // 優先檢查：如果網址有指定 page 參數，則直接載入該頁面
  // 這允許沒有 Google 帳號的使用者透過 Auth.html 的 redirect 訪問 Index 或 SR_server01
  if (page) {
    try {
      var template = HtmlService.createTemplateFromFile(page);
      template.webAppUrl = getScriptUrl();
      template.authUser = "0";

      return template.evaluate()
        .setTitle("舒漾長照系統")
        .addMetaTag("viewport", "width=device-width, initial-scale=1")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (err) {
      // 若頁面不存在，回退到 Auth
      var errorTemplate = HtmlService.createTemplateFromFile("Auth");
      errorTemplate.webAppUrl = getScriptUrl();
      return errorTemplate.evaluate()
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }

  // 2. 只有在完全抓不到 Email 的情況下才顯示 AuthPortal
  var template = HtmlService.createTemplateFromFile("Auth");
  template.webAppUrl = getScriptUrl();
  // 使用強制帳號選擇器
  template.authUrl = "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(getScriptUrl());

  return template.evaluate()
    .setTitle("帳號驗證 - 舒漾長照")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}




/**
 * 取得當前使用者的名稱與 Email
 * 用於 Suggest.html 預填資料
 */
function getUserNameByEmail() {
  // 取消自動抓取 Google Email
  var email = "";
  var userName = "";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 定義搜尋目標：優先搜尋 User，若無則搜尋 Manager
  const searchTargets = [
    { sheet: "User", nameCol: "User_N", emailCols: ["User_Email", "Email"] },
    { sheet: "Manager", nameCol: "Mana_N", emailCols: ["Mana_Email"] }
  ];

  for (const target of searchTargets) {
    const sheet = ss.getSheetByName(target.sheet);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idxName = getColIndex(headers, target.nameCol);
    
    // 尋找 Email 欄位 (支援多個可能名稱)
    let idxEmail = -1;
    for (const col of target.emailCols) {
      idxEmail = getColIndex(headers, col);
      if (idxEmail !== -1) break;
    }

    if (idxName !== -1 && idxEmail !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idxEmail]).trim().toLowerCase() === email.toLowerCase()) {
          userName = data[i][idxName];
          break;
        }
      }
    }
    
    if (userName) break;
  }


  return {
    userName: userName,
    userEmail: email
  };
}
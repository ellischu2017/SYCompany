/**
 * Auth.gs - 認證模組
 * 處理使用者認證和權限檢查
 */

/**
 * 處理 HTTP GET 請求
 * 根據參數 'page' 決定顯示哪個 HTML 檔案
 */
function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail();
  var authUser = e.parameter.authuser;

  // 1. 如果已經抓到 Email，就直接進入身分分流，不再卡在 AuthPortal
  if (userEmail && userEmail !== "") {
    var isManager = checkManagerPrivilege(userEmail);
    var page = e.parameter.page || "Index";

    // 身分分流
    var pageToLoad = isManager || page == "Suggest" ? page : "SR_server01";

    var template = HtmlService.createTemplateFromFile(pageToLoad);
    template.userEmail = userEmail;
    template.webAppUrl = getScriptUrl();
    template.authUser = authUser || "0";

    return template.evaluate()
      .setTitle(isManager ? "舒漾長照管理系統" : "舒漾電子服務紀錄管理")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 2. 只有在完全抓不到 Email 的情況下才顯示 AuthPortal
  var template = HtmlService.createTemplateFromFile("AuthPortal");
  // 使用強制帳號選擇器
  template.authUrl = "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(getScriptUrl());
  
  return template.evaluate()
    .setTitle("帳號驗證 - 舒漾長照")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/**
 * 檢查 Email 是否存在於 Manager 工作表的 Mana_Email 欄位
 */
function checkManagerPrivilege(email) {
  try {
    var sheet = MainSpreadsheet.getSheetByName("Manager");
    if (!sheet) return false;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var emailColIndex = headers.indexOf("Mana_Email");
    if (emailColIndex === -1) return false;

    for (var i = 1; i < data.length; i++) {
      if (
        data[i][emailColIndex].toString().toLowerCase() === email.toLowerCase()
      ) {
        // 先執行同步
        processUserSync01();
        return true;
      }
    }
  } catch (f) {
    console.log("驗證過程出錯: " + f.toString());
  }
  return false;
}

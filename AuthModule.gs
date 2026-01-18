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
  if (!userEmail || userEmail === "") {
    var template = HtmlService.createTemplateFromFile("Auth");
    template.scriptUrl = getScriptUrl();
    return template
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var page = e.parameter.page || "Index";
  var isManager = checkManagerPrivilege(userEmail);
  var pageToLoad = isManager || page === "Suggest" ? page : "SR_server01";

  var template = HtmlService.createTemplateFromFile(pageToLoad);
  template.userEmail = userEmail;
  template.webAppUrl = getScriptUrl();

  return template
    .evaluate()
    .setTitle(isManager ? "舒漾長照管理系統" : "舒漾電子服務紀錄管理")
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

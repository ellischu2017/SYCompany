/**
 * Utilities.gs - 工具函式模組
 * 提供通用的工具函式
 */

// 全域試算表參考
const MainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

/**
 * 測試用的 include 函數 (若未來需要拆分 CSS/JS 檔案時使用)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 從試算表動態取得意見反應連結並生成 Footer HTML
 */
function includeFooter() {
  let suggestUrl = "";

  try {
    const sheet = MainSpreadsheet.getSheetByName("SYTemp");

    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === "SYSuggest") {
          suggestUrl = data[i][1];
          break;
        }
      }
    }
  } catch (e) {
    console.log("取得意見反應網址失敗: " + e.message);
  }

  const template = HtmlService.createTemplateFromFile("Footer");
  template.suggestUrl = suggestUrl;
  return template.evaluate().getContent();
}

/**
 * 提供給 HTML 範本呼叫，用來載入導航列組件
 */
function includeNav() {
  var template = HtmlService.createTemplateFromFile("Nav");
  return template.evaluate().getContent();
}

/**
 * 取得當前 Web App 的 URL
 * 用於前端按鈕跳轉
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * 輔助函式：從 SYCompany (本腳本綁定之試算表) 的 sheetName 工作表取得外部試算表物件
 */
function getTargetsheet(sheetName, targetName) {
  var sheet = MainSpreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error("找不到 SYCompany 中的" + sheetName + "工作表");

  var data = sheet.getDataRange().getValues();
  var url = "";

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === targetName) {
      url = data[i][1];
      break;
    }
  }

  if (!url)
    throw new Error(
      "無法在" + sheetName + "工作表中找到名稱為" + targetName + "的對應網址"
    );

  return SpreadsheetApp.openByUrl(url);
}

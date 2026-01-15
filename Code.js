/**
 * 處理 HTTP GET 請求
 * 根據參數 'page' 決定顯示哪個 HTML 檔案
 */
function doGet(e) {
  // 1. 取得當前登入使用者的 Email
  var userEmail = Session.getActiveUser().getEmail();
  // 如果因為權限設定抓不到 Email，會引導使用者登入
  if (!userEmail || userEmail === "") {
    var output = HtmlService.createHtmlOutput(
      "<div style='font-family: sans-serif; text-align: center; padding-top: 50px;'>" +
        "<h3>需要授權以存取系統</h3>" +
        "<p>請確保您已登入 Google 帳號。如果仍看到此訊息，請點擊下方按鈕進行授權。</p>" +
        "<a href='" +
        ScriptApp.getService().getUrl() +
        "' target='_top' " +
        "style='padding: 10px 20px; background: #4285f4; color: white; text-decoration: none; border-radius: 5px;'>重新驗證身份</a>" +
        "</div>"
    );
    return output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); //[cite: 6]
  }

  // 2. 檢查是否在 Manager 工作表的允許名單內
  var page = e.parameter.page || "Index";
  var isManager = checkManagerPrivilege(userEmail);

  // 3. 根據檢查結果決定顯示的頁面
  var pageToLoad = isManager ? page : "SR_server01";

  // 建立 HTML 模板
  var template = HtmlService.createTemplateFromFile(pageToLoad);

  // 傳遞參數給前端 (選用)
  template.userEmail = userEmail;
  template.webAppUrl = ScriptApp.getService().getUrl();

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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Manager");
    if (!sheet) return false;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // 尋找 Mana_Email 欄位的索引
    var emailColIndex = headers.indexOf("Mana_Email");
    if (emailColIndex === -1) return false;

    // 比對每一列的 Email (忽略大小寫)
    for (var i = 1; i < data.length; i++) {
      if (
        data[i][emailColIndex].toString().toLowerCase() === email.toLowerCase()
      ) {
        return true;
      }
    }
  } catch (f) {
    console.log("驗證過程出錯: " + f.toString());
  }
  return false;
}

/**
 * 取得當前 Web App 的 URL
 * 用於前端按鈕跳轉
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SYTemp");

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

  // --- 關鍵改寫部分 ---

  // 1. 從檔案建立模板
  const template = HtmlService.createTemplateFromFile("Footer");

  // 2. 將變數賦值給模板對象 (名稱必須與 HTML 內的 <?= ... ?> 一致)
  template.suggestUrl = suggestUrl;

  // 3. 執行模板並回傳 HTML 內容字串
  // 使用 evaluate().getContent() 取得最終的 HTML 字串
  return template.evaluate().getContent();
}

/**
 * 提供給 HTML 範本呼叫，用來載入導航列組件
 */
function includeNav() {
  // 使用 createTemplateFromFile 而非 createHtmlOutputFromFile
  var template = HtmlService.createTemplateFromFile("Nav");

  // 執行並回傳解析後的內容
  return template.evaluate().getContent();
}

// --- 個案管理系統 (Cust) 相關功能 ----------------------------------------------

/**
 * 取得「Cust」工作表的個案姓名列表，用於下拉選單
 * 來源:
 */
function getCustList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cust"); // [cite: 3]
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // 沒資料

  // 取得第一欄 (Cust_N) 的所有資料
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // 過濾空值並轉為一維陣列
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 根據個案姓名查詢詳細資料
 * 來源:
 */
function queryCustData(custName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cust");
  const data = sheet.getDataRange().getValues();

  // 遍歷資料尋找匹配的姓名 (假設第一列是標題，從第二列開始)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == custName) {
      // Cust_N 在第 1 欄 (index 0)
      const row = data[i];

      // 處理日期格式 yyyy-MM-dd [cite: 10]
      let birthDate = row[1];
      if (birthDate instanceof Date) {
        birthDate = Utilities.formatDate(
          birthDate,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd"
        );
      }

      return {
        found: true,
        rowResult: {
          Cust_N: row[0],
          Cust_BD: birthDate,
          Cust_Add: row[2],
          Sup_N: row[3],
          DSup_N: row[4],
          Sup_Tel: row[5].toString(), // 強制轉文字 [cite: 23]
          Office_Tel: row[6].toString(),
          Cust_EC: row[7],
          Cust_EC_Tel: row[8].toString(), // 強制轉文字 [cite: 33]
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增個案資料
 * 來源:
 */
function addCustData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cust");

  // 檢查是否已存在
  const list = getCustList();
  if (list.includes(formObj.custName)) {
    return { success: false, message: "該個案姓名已存在！" };
  }

  // 準備寫入的陣列，順序需對應工作表欄位 [cite: 4-33]
  // 格式: [Cust_N, Cust_BD, Cust_Add, Sup_N, DSup_N, Sup_Tel, Office_Tel, Cust_EC, Cust_EC_Tel]
  const newRow = [
    formObj.custName,
    "'" + formObj.custBirth, // 加單引號強制視為文字或保持格式，視需求而定，或是直接存 Date 物件
    formObj.custAddr,
    formObj.supName,
    formObj.dsupName,
    "'" + formObj.supTel, // 電話強制文字格式
    "'" + formObj.officeTel,
    formObj.ecName,
    "'" + formObj.ecTel,
  ];

  sheet.appendRow(newRow);
  return { success: true, message: "新增成功！" };
}

/**
 * 更新個案資料
 * 來源:
 */
function updateCustData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cust");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.custName) {
      // 根據姓名鎖定列
      // 欄位索引: 0=Name, 1=BD, 2=Add, 3=SupN, 4=DSupN, 5=SupTel, 6=OffTel, 7=EC, 8=ECTel
      const rowNum = i + 1; // 實際列號

      // 更新各欄位
      sheet.getRange(rowNum, 2).setValue(formObj.custBirth);
      sheet.getRange(rowNum, 3).setValue(formObj.custAddr);
      sheet.getRange(rowNum, 4).setValue(formObj.supName);
      sheet.getRange(rowNum, 5).setValue(formObj.dsupName);
      sheet.getRange(rowNum, 6).setValue("'" + formObj.supTel);
      sheet.getRange(rowNum, 7).setValue("'" + formObj.officeTel);
      sheet.getRange(rowNum, 8).setValue(formObj.ecName);
      sheet.getRange(rowNum, 9).setValue("'" + formObj.ecTel);

      return { success: true, message: "資料更新成功！" };
    }
  }
  return { success: false, message: "找不到該個案資料，無法更新。" };
}

/**
 * 刪除個案資料
 * 來源:
 */
function deleteCustData(custName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Cust");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == custName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "找不到資料，刪除失敗。" };
}

// --- 居服員管理系統 (User) 相關功能 -----------------------

/**
 * 取得「User」工作表的居服員姓名列表
 */
function getUserList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢居服員詳細資料
 */
function queryUserData(userName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userName) {
      // User_N 在第 1 欄
      return {
        found: true,
        rowResult: {
          User_N: data[i][0],
          User_Email: data[i][1],
          User_Tel: data[i][2].toString(),
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增居服員資料
 */
function addUserData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User");
  const list = getUserList();

  if (list.includes(formObj.userName)) {
    return { success: false, message: "該居服員姓名已存在！" };
  }

  const newRow = [
    formObj.userName,
    formObj.userEmail,
    "'" + formObj.userTel, // 強制文字格式
  ];

  sheet.appendRow(newRow);
  return { success: true, message: "新增成功！" };
}

/**
 * 更新居服員資料
 */
function updateUserData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.userName) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 2).setValue(formObj.userEmail);
      sheet.getRange(rowNum, 3).setValue("'" + formObj.userTel);
      return { success: true, message: "資料更新成功！" };
    }
  }
  return { success: false, message: "找不到資料，無法更新。" };
}

/**
 * 刪除居服員資料
 */
function deleteUserData(userName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("User");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == userName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

// --- 管理員管理系統 (Manager) 相關功能 -----------------------

/**
 * 取得「Manager」工作表的管理員姓名列表
 */
function getManaList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manager");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢管理員詳細資料
 */
function queryManaData(manaName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == manaName) {
      // User_N 在第 1 欄
      return {
        found: true,
        rowResult: {
          Mana_N: data[i][0],
          Mana_Email: data[i][1],
          Mana_Tel: data[i][2].toString(),
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增管理員資料
 */
function addManaData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manager");
  const list = getManaList();

  if (list.includes(formObj.manaName)) {
    return { success: false, message: "該管理員姓名已存在！" };
  }

  const newRow = [
    formObj.manaName,
    formObj.manaEmail,
    "'" + formObj.manaTel, // 強制文字格式
  ];

  sheet.appendRow(newRow);
  return { success: true, message: "新增成功！" };
}

/**
 * 更新管理員資料
 */
function updateManaData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.manaName) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 2).setValue(formObj.manaEmail);
      sheet.getRange(rowNum, 3).setValue("'" + formObj.manaTel);
      return { success: true, message: "資料更新成功！" };
    }
  }
  return { success: false, message: "找不到資料，無法更新。" };
}

/**
 * 刪除管理員資料
 */
function deleteManaData(manaName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Manager");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == manaName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

// --- 長照編碼管理 (LTC_Code) 相關功能 --------------------------

/**
 * 取得「LTC_Code」工作表的服務編碼列表 (SR_ID)
 */
function getLtcCodeList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LTC_Code");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // 取得第一欄 (SR_ID) 的所有資料
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢服務編碼詳細資料
 */
function queryLtcCodeData(srId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == srId) {
      return {
        found: true,
        rowResult: {
          SR_ID: data[i][0],
          SR_Name: data[i][1],
          SR_Detail: data[i][2], // 新增第 3 欄：服務內容
        },
      };
    }
  }
  return { found: false };
}

/**
 * 新增服務編碼
 */
function addLtcCodeData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LTC_Code");
  const list = getLtcCodeList();

  if (list.includes(formObj.srId)) {
    return { success: false, message: "該服務編碼已存在！" };
  }

  const newRow = [
    formObj.srId,
    formObj.srName,
    formObj.srDetail, // 新增 SR_Detail 資料
  ];

  sheet.appendRow(newRow);
  return { success: true, message: "編碼新增成功！" };
}

/**
 * 更新服務編碼資料
 */
function updateLtcCodeData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.srId) {
      const rowNum = i + 1;
      sheet.getRange(rowNum, 2).setValue(formObj.srName); // 更新 SR_Name
      sheet.getRange(rowNum, 3).setValue(formObj.srDetail); // 更新第 3 欄
      return { success: true, message: "編碼資料更新成功！" };
    }
  }
  return { success: false, message: "找不到資料，無法更新。" };
}

/**
 * 刪除服務編碼
 */
function deleteLtcCodeData(srId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LTC_Code");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == srId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "編碼刪除成功！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

// --- 服務紀錄單網址管理 (RecUrl) 相關功能 ----------------------------------

/**
 * 取得「RecUrl」工作表的個案姓名列表
 */
function getRecUrlList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RecUrl");
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return data.map((r) => r[0]).filter((n) => n !== "");
}

/**
 * 查詢特定個案的網址資料
 */
function queryRecUrlData(syName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == syName) {
      // Cust_N 在第 1 欄
      return {
        found: true,
        rowResult: {
          SY_N: data[i][0],
          SY_Url: data[i][1],
        },
      };
    }
  }
  return { found: false };
}

/**
 * 儲存網址資料 (新增或更新)
 */
function saveRecUrlData(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();

  // 檢查是否已存在，存在則更新，不存在則新增
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formObj.syName) {
      sheet.getRange(i + 1, 2).setValue(formObj.recUrl);
      return { success: true, message: "網址更新成功！" };
    }
  }

  sheet.appendRow([formObj.syName, formObj.recUrl]);
  return { success: true, message: "網址新增成功！" };
}

/**
 * 刪除網址資料
 */
function deleteRecUrlData(syName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("RecUrl");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == syName) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "資料已刪除！" };
    }
  }
  return { success: false, message: "刪除失敗。" };
}

// --- 服務紀錄單管理 (SR_server) 相關功能 ----------------------------------

/**
 * 初始化管理員頁面資料 (包含需求 1: 同步檢查)
 * 管理員頁面初始化資料
 */
function getAdminInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const userSheet = ss.getSheetByName("User");
  const users = userSheet
    ? userSheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map((r) => ({ name: r[0], email: r[1] }))
    : [];

  const custSheet = ss.getSheetByName("Cust");
  const customers = custSheet
    ? custSheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map((r) => r[0])
    : [];

  const ltcSheet = ss.getSheetByName("LTC_Code");
  const ltcCodes = ltcSheet
    ? ltcSheet
        .getDataRange()
        .getValues()
        .slice(1)
        .map((r) => r[0])
    : [];

  // 檢查 SYTemp 是否有待同步資料
  const tempsheet = getSYTempSpreadsheet();
  const tempUserSheet = tempsheet.getSheetByName("User");
  const hasTempData = tempsheet && tempUserSheet.getLastRow() > 1;

  return {
    users: users,
    customers: customers,
    ltcCodes: ltcCodes,
    triggerSync: hasTempData,
  };
}

/**
 * 跨表查詢邏輯 (需求 5, 6, 7)
 * 核心查詢邏輯：使用字串比對防止日期偏移
 */
function queryServiceRecord(params) {
  const { date, custN, userN, payType, srId } = params; // date 為 "yyyy-MM-dd" 字串
  const year = date.split("-")[0];
  const syYearKey = "SY" + year;

  // 1. 查詢年度表 (SYCompany 體系)
  const recUrlData = queryRecUrlData(syYearKey); // 假設您已有此函式獲取 URL
  if (recUrlData.found) {
    try {
      const targetSs = SpreadsheetApp.openByUrl(recUrlData.rowResult.SY_Url);
      const result = searchAcrossSheets(targetSs, params);
      if (result)
        return { ...result, source: "SYCompany", ssId: targetSs.getId() };
    } catch (e) {
      console.log("年度表開啟失敗");
    }
  }

  // 2. 查詢臨時表 (SYTemp)
  const ssTemp = getSYTempSpreadsheet();
  const tempSheet = ssTemp.getSheetByName("SR_Data");
  const tempResult = searchSheet(tempSheet, params);

  if (tempResult) {
    return { ...tempResult, source: "SYTemp", ssId: ssTemp.getId() };
  }

  return { found: false };
}

function searchAcrossSheets(ss, p) {
  const sheets = ss.getSheets();
  for (let sheet of sheets) {
    const res = searchSheet(sheet, p);
    if (res) return { ...res, sheetName: sheet.getName() };
  }
  return null;
}

function searchSheet(sheet, p) {
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    // 關鍵修正：將試算表日期轉為字串進行比對，防止少一天
    let rowDateStr =
      data[i][0] instanceof Date
        ? Utilities.formatDate(data[i][0], "GMT+8", "yyyy-MM-dd")
        : String(data[i][0]);
    console.log(p);
    if (
      rowDateStr === p.date &&
      data[i][2] === p.custN &&
      data[i][3] === p.userN &&
      data[i][4] === p.payType &&
      data[i][5] === p.srId
    ) {
      return { found: true, data: data[i], rowIndex: i + 1 };
    }
  }
  return { found: false, data: null };
}

/**
 * 管理員 CRUD 操作
 */
function manageSRData(action, form, sourceInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const row = [
    form.date,
    form.email,
    form.custN,
    form.userN,
    form.payType,
    form.srId,
    form.srRec,
    form.loc,
    form.mood,
    form.spcons,
  ];

  if (action === "add") {
    const tempSheet = ss.getSheetByName("SR_Data");
    tempSheet.appendRow(row);
    tempSheet.getRange(tempSheet.getLastRow(), 1).setNumberFormat("yyyy-MM-dd");
    return "已成功新增至 SYTemp。";
  }

  const targetSs = SpreadsheetApp.openById(sourceInfo.ssId);
  const targetSheet =
    sourceInfo.source === "SYTemp"
      ? targetSs.getSheetByName("SR_Data")
      : targetSs.getSheetByName(sourceInfo.sheetName);

  if (action === "update") {
    targetSheet
      .getRange(sourceInfo.rowIndex, 1, 1, row.length)
      .setValues([row]);
    targetSheet.getRange(sourceInfo.rowIndex, 1).setNumberFormat("yyyy-MM-dd");
    return `資料已於 ${sourceInfo.source} 更新成功。`;
  } else if (action === "delete") {
    targetSheet.deleteRow(sourceInfo.rowIndex);
    return `資料已從 ${sourceInfo.source} 刪除。`;
  }
}

// --- 服務紀錄單管理 (SR_server01) 相關功能 ----------------------------------
/**
 * 核心處理：處理服務紀錄 (查詢/新增/修改/刪除)
 * 對應需求 2 (儲存新 User) & 需求 3 (寫入 SR_Data)
 */

function processSRData(formObj, actionType) {
  try {
    var targetSs = getSYTempSpreadsheet(); // 取得 SYTemp 試算表

    // --- 處理新使用者邏輯 (需求 2) ---
    // 如果前端傳來了電話號碼，代表這是新使用者，需要寫入 SYTemp > User
    if (formObj.userTel && formObj.userName && formObj.email) {
      var userSheet = targetSs.getSheetByName("User");
      if (!userSheet) userSheet = targetSs.insertSheet("User");

      // 雙重檢查避免重複 (依 Email)
      var uData = userSheet.getDataRange().getValues();
      var uExists = false;
      for (var k = 1; k < uData.length; k++) {
        if (uData[k][1] === formObj.email) {
          uExists = true;
          break;
        }
      }

      if (!uExists) {
        // 寫入格式：User_N, User_Email, User_Tel
        userSheet.appendRow([
          formObj.userName,
          formObj.email,
          "'" + formObj.userTel,
        ]);
      }
    }

    // --- 處理服務紀錄邏輯 (需求 3) ---
    // 改為讀寫 SYTemp > SR_Data 工作表
    var targetSheet = targetSs.getSheetByName("SR_Data");
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet("SR_Data");
      // 確保標頭存在
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

    // 準備寫入的資料列
    // 注意：這裡假設前端傳來的 userTel 只有在新用戶時才有值，
    // 寫入 SR_Data 時不需要寫入電話，只需寫入 User_N
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

    // 新增模式
    if (actionType === "add") {
      targetSheet.appendRow(rowData);
      return { success: true, message: "新增紀錄成功" };
    }

    // 查詢、更新、刪除模式
    var data = targetSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let sheetDate =
        data[i][0] instanceof Date
          ? Utilities.formatDate(
              data[i][0],
              Session.getScriptTimeZone(),
              "yyyy-MM-dd"
            )
          : data[i][0].toString();

      // 比對條件：日期 + 個案名 + 居服員 + 服務編碼
      if (
        sheetDate === formObj.date &&
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
              email: data[i][1],
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
 * 初始化 SR_server01 頁面所需的所有下拉選單資料
 * 這是前端呼叫 getSRServer01InitData() 時對應的後端邏輯
 */

function getSRServer01InitData() {
  var userEmail = Session.getActiveUser().getEmail();
  var currentUserName = "";
  var found = false;

  // 1. 先檢查 SYCompany (本地) 的 User 表
  var localSS = SpreadsheetApp.getActiveSpreadsheet();
  var localUserSheet = localSS.getSheetByName("User");
  if (localUserSheet) {
    var localData = localUserSheet.getDataRange().getValues();
    for (var i = 1; i < localData.length; i++) {
      if (localData[i][1] === userEmail) {
        // User_Email column
        currentUserName = localData[i][0]; // User_N
        found = true;
        break;
      }
    }
  }

  // 2. 如果本地沒找到，檢查 SYTemp (外部) 的 User 表
  if (!found) {
    try {
      var remoteSS = getSYTempSpreadsheet();
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

  // 回傳結果
  return {
    custNames: getCustList(),
    srIds: getLtcCodeList(),
    currentUserName: currentUserName, // 若為空字串，前端將啟用輸入框
    userEmail: userEmail || "",
  };
}

/**
 * 輔助函式：從 SYCompany (本腳本綁定之試算表) 的 SYTemp 工作表取得外部試算表物件
 * 對應需求 4
 */
function getSYTempSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // SYCompany (ReadOnly context)
  var sheet = ss.getSheetByName("SYTemp"); // 設定檔所在的工作表
  if (!sheet) throw new Error("找不到 SYCompany 中的 SYTemp 設定工作表");

  var data = sheet.getDataRange().getValues();
  var url = "";

  // 尋找 SYT_N 為 "SYTemp" 的網址
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "SYTemp") {
      url = data[i][1];
      break;
    }
  }

  if (!url)
    throw new Error("無法在 SYTemp 工作表中找到名稱為 'SYTemp' 的對應網址");

  return SpreadsheetApp.openByUrl(url); // 回傳外部試算表物件
}

//------自動更新部份------------------------------------

/**
 * 同步所有相關試算表的權限
 * 1. 包含 SYCompany 本身與 RecUrl 內的所有試算表。
 * 2. 根據 Manager 工作表名單授權為「編輯者」。
 * 3. 移除名單外所有「特定的」編輯者與檢視者。
 * 4. 將「一般存取權」設為「知道連結的人即可檢視」。
 */
function syncMasterTablePermissions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 取得管理員 Email 名單
  var managerSheet = ss.getSheetByName("Manager");
  var managerData = managerSheet.getDataRange().getValues();
  var managerEmails = [];
  for (var i = 1; i < managerData.length; i++) {
    var email = managerData[i][1];
    if (email) managerEmails.push(email.toString().trim().toLowerCase());
  }

  // 2. 取得所有目標檔案 ID (包含 SYCompany 與 RecUrl 中的網址)
  var targetFileIds = [ss.getId()];
  var recUrlSheet = ss.getSheetByName("RecUrl");
  if (recUrlSheet) {
    var urlData = recUrlSheet.getDataRange().getValues();
    for (var j = 1; j < urlData.length; j++) {
      var url = urlData[j][1];
      if (url && url.indexOf("docs.google.com") !== -1) {
        try {
          targetFileIds.push(SpreadsheetApp.openByUrl(url).getId());
        } catch (e) {}
      }
    }
  }

  // 3. 執行權限操作
  targetFileIds.forEach(function (fileId) {
    try {
      // --- 核心：使用 Drive API 授權但不發信 ---
      managerEmails.forEach(function (email) {
        var resource = {
          role: "writer", // 設為編輯者
          type: "user",
          value: email,
        };

        // 關鍵參數：sendNotificationEmails: false
        Drive.Permissions.insert(resource, fileId, {
          sendNotificationEmails: false,
        });
      });

      // --- 移除不在名單內的人員 (這部分仍可使用 DriveApp) ---
      var file = DriveApp.getFileById(fileId);
      var ownerEmail = file.getOwner().getEmail().toLowerCase();

      // 清理編輯者
      file.getEditors().forEach(function (editor) {
        var e = editor.getEmail().toLowerCase();
        if (managerEmails.indexOf(e) === -1 && e !== ownerEmail) {
          file.removeEditor(editor);
        }
      });

      // 清理檢視者
      file.getViewers().forEach(function (viewer) {
        var v = viewer.getEmail().toLowerCase();
        if (managerEmails.indexOf(v) === -1) {
          file.removeViewer(viewer);
        }
      });

      // 設定一般存取權：知道連結的人可檢視
      file.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
    } catch (e) {
      console.error("檔案 ID " + fileId + " 處理失敗: " + e.message);
    }
  });
}

/**
 * 每日維護任務：遷移 7 天前資料與同步 User 名單
 * 建議觸發時間：每日 00:00 - 01:00
 */
function dailyMaintenanceJob() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 目前應為 SYCompany
  const tempSS = getSYTempSpreadsheet(); // 呼叫您現有的函式取得 SYTemp 試算表

  // --- 任務 A: 處理 User 名單同步 (需求 4) ---
  processUserSync(ss, tempSS);

  // --- 任務 B: 處理過期 SR_Data 遷移 (需求 2 & 3) ---
  processSRDataMigration(ss, tempSS);
}

/**
 * 處理 User 同步：SYTemp > User 搬移至 SYCompany > User
 * 特別處理：確保電話號碼 User_Tel 為文字字串格式
 * 1. 自動檢查 Email 是否重複，重複則不新增但仍從 Temp 移除。
 * 2. 確保 User_Tel 以文字格式 (@) 存入。
 */
function processUserSync(mainSS, tempSS) {
  const tempUserSheet = tempSS.getSheetByName("User");
  const mainUserSheet = mainSS.getSheetByName("User");

  if (!tempUserSheet || !mainUserSheet) return;

  const tempData = tempUserSheet.getDataRange().getValues();
  if (tempData.length <= 1) return; // 只有標題列

  // 1. 取得主表 (SYCompany) 現有的 Email 名單
  const mainData = mainUserSheet.getDataRange().getValues();
  const existingEmails = mainData
    .slice(1)
    .map((row) => row[1].toString().trim().toLowerCase());

  const newRowsToAppend = [];
  const headers = tempData[0];

  // 2. 遍歷 Temp 資料進行比對
  for (let i = 1; i < tempData.length; i++) {
    let row = tempData[i];
    let tempEmail = row[1].toString().trim().toLowerCase();

    // 如果 Email 不在現有名單中，才加入「準備寫入」的清單
    if (existingEmails.indexOf(tempEmail) === -1) {
      newRowsToAppend.push(row);
    } else {
      console.log("Email 已存在，跳過新增: " + tempEmail);
    }
  }

  // 3. 執行寫入主表 (如果有新資料)
  if (newRowsToAppend.length > 0) {
    const startRow = mainUserSheet.getLastRow() + 1;
    const targetRange = mainUserSheet.getRange(
      startRow,
      1,
      newRowsToAppend.length,
      headers.length
    );

    // 強制設定為文字格式防止電話 0 消失
    targetRange.setNumberFormat("@");
    targetRange.setValues(newRowsToAppend);
    console.log("已新增 " + newRowsToAppend.length + " 筆新居服員資料。");
  }

  // 4. 無論是否有新增到主表，最後都清空 SYTemp > User 的資料 (保留標題)
  if (tempUserSheet.getLastRow() > 1) {
    tempUserSheet.deleteRows(2, tempUserSheet.getLastRow() - 1);
    console.log("SYTemp > User 已清理完畢。");
  }
}

/**
 * 處理 SR_Data 遷移：7天前資料搬移至年度試算表
 * 修正：日期偏移、新增首列凍結、設定日期欄位格式
 */
function processSRDataMigration(mainSS, tempSS) {
  const srSheet = tempSS.getSheetByName("SR_Data");
  if (!srSheet) return;

  const data = srSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const today = new Date();
  const cutoffDate = new Date();
  cutoffDate.setDate(today.getDate() - 7);

  const migrationMap = {};
  const rowsToKeep = [headers];
  let createdNewSS = false;

  for (let i = 1; i < data.length; i++) {
    let row = [...data[i]];
    let rawDate = row[0];

    // 修正日期少一天的關鍵：處理字串並指定時區
    let dateObj;
    if (rawDate instanceof Date) {
      dateObj = rawDate;
    } else {
      // 避免字串解析時區偏移，將 "-" 取代為 "/" 有助於部分瀏覽器引擎正確解析
      dateObj = new Date(rawDate.toString().replace(/-/g, "/"));
    }

    // 強制格式化為 yyyy-MM-dd 字串，確保搬運過程不失真
    let formattedDate = Utilities.formatDate(
      dateObj,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
    row[0] = formattedDate;

    if (dateObj < cutoffDate) {
      let year = dateObj.getFullYear();
      if (!migrationMap[year]) migrationMap[year] = [];
      migrationMap[year].push(row);
    } else {
      rowsToKeep.push(row);
    }
  }

  for (let year in migrationMap) {
    let syName = "SY" + year;
    let targetUrl = getUrlFromRecUrl(mainSS, syName);

    if (!targetUrl) {
      targetUrl = createNewYearlySS(mainSS, syName);
      createdNewSS = true;
    }

    if (targetUrl) {
      appendDataToExternalSS(targetUrl, year, migrationMap[year]);
    }
  }

  if (createdNewSS) {
    syncMasterTablePermissions();
  }

  srSheet.clearContents();
  srSheet
    .getRange(1, 1, rowsToKeep.length, headers.length)
    .setValues(rowsToKeep);
}

/**
 * 輔助函式：根據 SY_N 取得 RecUrl 內的網址
 */
function getUrlFromRecUrl(mainSS, syName) {
  const recSheet = mainSS.getSheetByName("RecUrl");
  const data = recSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === syName) return data[i][1];
  }
  return null;
}

/**
 * 輔助函式：將資料寫入年度試算表
 * 包含：凍結首列、設定日期格式、設定文字格式
 * 1. 凍結首列並設定 A 欄日期格式
 * 2. 移除舊篩選器並重新建立 (涵蓋所有資料列)
 * 3. 針對 A 欄進行 A 到 Z (由舊到新) 排序
 */
function appendDataToExternalSS(url, year, rows) {
  try {
    const targetSS = SpreadsheetApp.openByUrl(url);
    const firstDateStr = rows[0][0].toString().replace(/-/g, "/");
    const firstDate = new Date(firstDateStr);
    const monthStr = Utilities.formatDate(
      firstDate,
      Session.getScriptTimeZone(),
      "yyyyMM"
    );

    let targetSheet = targetSS.getSheetByName(monthStr);

    // 如果是新工作表
    if (!targetSheet) {
      targetSheet = targetSS.insertSheet(monthStr);
      const headers = [
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
      ];
      targetSheet.appendRow(headers);
      targetSheet.setFrozenRows(1);
      targetSheet.getRange("A:A").setNumberFormat("yyyy-MM-dd");
    }

    // 1. 寫入新資料
    const startRow = targetSheet.getLastRow() + 1;
    const numCols = rows[0].length;
    const targetRange = targetSheet.getRange(startRow, 1, rows.length, numCols);

    // 寫入前先針對非日期欄位設定純文字格式
    if (numCols > 1) {
      targetSheet
        .getRange(startRow, 2, rows.length, numCols - 1)
        .setNumberFormat("@");
    }
    targetRange.setValues(rows);

    // 2. 處理篩選器與排序
    // 先移除現有的篩選器 (若有的話)，以確保篩選範圍包含新加入的列
    const currentFilter = targetSheet.getFilter();
    if (currentFilter) {
      currentFilter.remove();
    }

    // 取得目前所有資料範圍 (從 A1 到最後一列最後一欄)
    const fullRange = targetSheet.getDataRange();

    // 建立新的篩選器
    const newFilter = fullRange.createFilter();

    // 3. 執行排序：針對第 1 欄 (Date) 進行 A 到 Z 排序
    // 參數 1 代表第一欄，true 代表由小到大 (A-Z)
    newFilter.sort(1, true);

    console.log(
      `成功搬移並排序 ${rows.length} 筆資料至 ${year} 年 ${monthStr} 表`
    );
  } catch (e) {
    console.error("寫入外部試算表失敗: " + e.toString());
  }
}

/**
 * 輔助函式：建立新年度試算表並回傳網址 (依照您的設計文件)
 */
function createNewYearlySS(mainSS, syName) {
  const newSS = SpreadsheetApp.create(syName);
  const url = newSS.getUrl();
  const recSheet = mainSS.getSheetByName("RecUrl");
  recSheet.appendRow([syName, url]);
  return url;
}

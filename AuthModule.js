/**
 * Auth.gs - 認證模組
 * 處理使用者認證和權限檢查
 */

/**
 * 處理 HTTP GET 請求
 * 根據參數 'page' 決定顯示哪個 HTML 檔案
 * @e 
 */



/**
 * 檢查 Email 是否存在於 Manager 工作表的 Mana_Email 欄位
 * @email : 
 */
function checkManagerPrivilege(email) {
  try {
    var sheet = MainSpreadsheet.getSheetByName("Manager");
    if (!sheet) return false;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    var emailColIndex = getColIndex(headers, "Mana_Email");
    if (emailColIndex === -1) return false;

    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][emailColIndex]).trim().toLowerCase() === email.toLowerCase()
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

/**
 * 讀取 User 與 Manager 工作表，建立授權使用者清單
 * 用於 Auth.html 前端登入驗證
 * @return {Array} 包含 {User, Pass, Role} 物件的陣列
 */
function getAuthUsers() {
  const authList = [];

  // --- 1. 處理 User 工作表 (居服員) ---
  const sheetUser = MainSpreadsheet.getSheetByName("User");
  if (sheetUser) {
    const data = sheetUser.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      // 使用 Utilities.js 中的 getColIndex 來確保欄位名稱比對正確 (忽略大小寫)
      const idxName = getColIndex(headers, "User_N");
      const idxPass = getColIndex(headers, "Pass"); // 密碼欄位
      const idxRole = getColIndex(headers, "Role"); // 角色欄位

      if (idxName !== -1) {
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const name = String(row[idxName] || "").trim();
          if (!name) continue;

          // 若 Pass 欄位不存在，則為空字串
          const pass = (idxPass !== -1) ? String(row[idxPass]).trim() : "";
          // 若 Role 欄位不存在或為空，預設為 'User' (統一首字大寫)
          let role = (idxRole !== -1) ? String(row[idxRole]).trim() : "";
          if (!role) role = "User";

          authList.push({
            User: name,
            Pass: pass,
            Role: role
          });
        }
      }
    }
  }

  // --- 2. 處理 Manager 工作表 (管理員) ---
  const sheetMana = MainSpreadsheet.getSheetByName("Manager");
  if (sheetMana) {
    const data = sheetMana.getDataRange().getValues();
    if (data.length > 1) {
      const headers = data[0];
      const idxName = getColIndex(headers, "Mana_N");
      const idxPass = getColIndex(headers, "Pass");
      const idxRole = getColIndex(headers, "Role");

      if (idxName !== -1) {
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const name = String(row[idxName] || "").trim();
          if (!name) continue;

          const pass = (idxPass !== -1) ? String(row[idxPass]).trim() : "";
          // 若 Role 欄位不存在或為空，預設為 'Admin'
          let role = (idxRole !== -1) ? String(row[idxRole]).trim() : "";
          if (!role) role = "Admin";

          authList.push({
            User: name,
            Pass: pass,
            Role: role
          });
        }
      }
    }
  }

  return authList;
}

function changeUserPassword(username, newPass) {
  var ss = MainSpreadsheet;
  var found = false;

  // 定義要搜尋的工作表與對應的帳號欄位
  var targets = [
    { sheet: "User", idCol: "User_N" },
    { sheet: "Manager", idCol: "Mana_N" }
  ];

  for (var k = 0; k < targets.length; k++) {
    var sheet = ss.getSheetByName(targets[k].sheet);
    if (!sheet) continue;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) continue; // 只有標題或空的

    var headers = data[0];
    var idxId = getColIndex(headers, targets[k].idCol);
    var idxPass = getColIndex(headers, "Pass");

    if (idxId === -1) continue; // 找不到帳號欄位

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idxId]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
        // 驗證：檢查觸發修改的使用者是否為本人或管理員
        const userEmail = Session.getActiveUser().getEmail().toLowerCase();
        
        // 安全獲取 Email 欄位值，避免索引為 -1 導致讀取 undefined
        const idxManaEmail = getColIndex(headers, 'Mana_Email');
        const idxUserEmail = getColIndex(headers, 'User_Email');
        const rowManaEmail = idxManaEmail !== -1 ? String(data[i][idxManaEmail]).trim().toLowerCase() : "";
        const rowUserEmail = idxUserEmail !== -1 ? String(data[i][idxUserEmail]).trim().toLowerCase() : "";

        const isSelf = (rowManaEmail === userEmail || rowUserEmail === userEmail);
        const isAdmin = checkManagerPrivilege(userEmail);

        if (!isSelf && !isAdmin) {
          console.warn(`[Security] ${userEmail} 嘗試修改 ${username} 的密碼但權限不足！`);
          throw new Error("您沒有權限修改其他使用者的密碼。");
        }

        // 記錄：在修改密碼前加入日誌
        console.log(`[Security] 使用者 ${userEmail} 正在修改 ${username} 的密碼。`);

        // 密碼安全：檢查新密碼強度 (長度至少 8 碼)
        if (!newPass || newPass.length < 8) {
          throw new Error("新密碼長度不得少於 8 碼。");
        }

        // 如果找不到 Pass 欄位，自動新增在最後一欄
        if (idxPass === -1) {
          idxPass = headers.length;
          sheet.getRange(1, idxPass + 1).setValue("Pass");
        }

        // 更新密碼 (設定為文字格式，避免數字密碼變形)
        sheet.getRange(i + 1, idxPass + 1).setNumberFormat("@").setValue(newPass);
        found = true;
        break;
      }
    }
    if (found) break;
  }

  if (found) {
    return true;
  } else {
    throw new Error("使用者不存在或無法更新密碼");
  }
}

/**
 * 取得 Web App 的部署網址
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * SPA 模式核心：取得特定頁面的 HTML 片段
 * @param {string} pageName 頁面檔案名稱 (不含 .html)
 * @return {string} 經過評估後的 HTML 內容字串
 */
function getPagePart(pageName) {
  const template = HtmlService.createTemplateFromFile(pageName);
  // 注入變數給模板評估時使用，確保子頁面中的 <?!= getScriptUrl() ?> 或變數引用不會報錯
  template.webAppUrl = getScriptUrl();
  return template.evaluate().getContent();
}

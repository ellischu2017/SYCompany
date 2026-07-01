# AGENTS.md — SYCompany GAS Project

## What this is

A **Google Apps Script Web App** (SPA) for long-term care management (Taiwan).  
Entrypoint: `Code.js` → `doGet()` renders HTML templates via `HtmlService.createTemplateFromFile()`.

## Developer commands

```bash
npm run push:gas    # clasp push
npm run pull:gas    # clasp pull
npm run deploy:gas  # clasp deploy → 自動加入版本號說明 (v{version})
```

**注意**：公司網路可能阻擋 `oauth2.googleapis.com`，所有 clasp 指令已透過 `clasp-fix.js` 繞過 SSL 驗證。若需手動執行 clasp，請先設：
```bash
$env:NODE_OPTIONS='-r E:\MyProg\GAS\SYCompany\Code\clasp-fix.js'
```

There are **no** test, lint, build, or typecheck scripts.

## Architecture

- **Frontend**: HTML files with embedded `<?!= ... ?>` GAS template tags.  
  Communicates with backend via `google.script.run` (async RPC).
- **Backend**: Plain JS modules – each manages one domain (Auth, Cust, User, Manager, SR, LTC codes, etc.).
- **Data**: Google Sheets (the script is **bound** to the `SYCompany` sheet).  
  Yearly archives (`SYyyyy`), monthly reports (`RPyyyyMM`), and temp data (`SYTemp`) are referenced by URL in registry sheets (`RecUrl`, `ReportsUrl`).
- **AI**: `Utilities.js` `generateAIServiceRecord()` calls Gemini API with a fallback chain.
- **PDF**: `PDFApp.js` is a bundled external library from [tanaikech/PDFApp](https://github.com/tanaikech/PDFApp) (do not edit upstream).
- **Scheduled jobs**: `Maintenance.js` (`dailyMaintenanceJob`, `monthlyMaintenanceJob`, `monthlyTenMaintenanceJob`) — uses `PropertiesService` for auto-resume state.

## GAS specifics

- **Runtime**: V8 (`appsscript.json`)
- **Timezone**: `Asia/Taipei`
- **Auth**: OAuth scopes for Spreadsheets, Drive, Script container UI, external requests, scriptapp
- **`webapp` access**: `ANYONE`, executes as `USER_DEPLOYING`
- `.clasp.json` is tracked in git (private repo); contains the GAS script ID
- All UI text is Traditional Chinese

## Key files

| File | Purpose |
|------|---------|
| `Code.js` | Webapp entrypoint (`doGet`) |
| `appsscript.json` | GAS manifest (scopes, runtime, webapp config) |
| `.clasp.json` | clasp project config (script ID) |
| `jsconfig.json` | VS Code JS config (no type checking: `checkJs: false`) |

## What is NOT here

- No README, no CI/CD, no test files, no linter/formatter config
- `node_modules/` and `package-lock.json` are gitignored
- No build step — GAS runs the raw JS files directly

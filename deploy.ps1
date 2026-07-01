# clasp deploy wrapper — 自動在部署說明中加入版本號 (v{version})
$env:NODE_OPTIONS = '-r E:\MyProg\GAS\SYCompany\Code\clasp-fix.js'

# `clasp deploy` 每次會建立一個新版本，所以預測版本號 = 目前最新 + 1
$versions = clasp versions 2>&1 | Out-String
$verMatch = [regex]::Matches($versions, '^(\d+)', [System.Text.RegularExpressions.RegexOptions]::Multiline)
$latestVer = if ($verMatch.Count -gt 0) { [int]$verMatch[$verMatch.Count - 1].Groups[1].Value } else { 0 }
$nextVer = $latestVer + 1
$desc = "v$nextVer"

Write-Host "部署版本 @$nextVer，說明：$desc"
$result = clasp deploy --description $desc 2>&1 | Out-String
Write-Host $result

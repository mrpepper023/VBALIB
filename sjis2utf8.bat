@(echo '> NUL
echo off)
setlocal enableextensions
set "THIS_PATH=%~f0"
set "PARAM_1=%~1"
PowerShell.exe -Command "iex -Command ((gc \"%THIS_PATH:`=``%\") -join \"`n\")"
exit /b %errorlevel%
-- 縺薙・1縺､荳翫・陦後∪縺ｧ繝舌ャ繝√ヵ繧｡繧､繝ｫ
') | sv -Name TempVar

# 縺薙％縺九ｉPowerShell繧ｹ繧ｯ繝ｪ繝励ヨ
Convert-Path .
pause

$files = Get-ChildItem -Path . -Include "*.bas","*.cls","*.frm","*.frx","*.bat" -Recurse
foreach ($file in $files) {
    $file.FullName
    Get-Content -Path $file.FullName -Encoding String `
    | Out-String `
    | % { [Text.Encoding]::UTF8.GetBytes($_) } `
    | Set-Content -Path $file.FullName -Encoding Byte
}
pause

@(echo '> NUL
echo off)
setlocal enableextensions
set "THIS_PATH=%~f0"
set "PARAM_1=%~1"
PowerShell.exe -Command "iex -Command ((gc \"%THIS_PATH:`=``%\") -join \"`n\")"
exit /b %errorlevel%
-- この1つ上の行までバッチファイル
') | sv -Name TempVar

# ここからPowerShellスクリプト
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

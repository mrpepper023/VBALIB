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
$currentTime = [System.DateTime]::Now

# テンポラリフォルダ作成
$tmp = $env:TEMP | Join-Path -ChildPath $([System.Guid]::NewGuid().Guid)
New-Item -ItemType Directory -Path $tmp | Push-Location

# テンポラリフォルダ名をテンポラリフォルダのmoduleimporter.txtに保存
$nm = $env:TEMP | Join-Path -ChildPath "moduleimporter.txt"
Set-Content -Path $nm -Value $tmp

echo $nm

# ダウンロードフォルダの最新アーカイブを特定
$shellapp = New-Object -ComObject Shell.Application
$dlfolder = $shellapp.Namespace("shell:Downloads").Self.Path
$targetgl = Join-Path $dlfolder VBALIB-main*.zip
$items = Get-ChildItem $targetgl -File

$newestnm = ""
$newestlwt = 0
foreach ($item in $items) {
	if($newestnm -eq ""){
		$newestnm = $item.Name
		$newestlwt = $item.LastWriteTime
	}
	if($newestlwt -lt $item.LastWriteTime){
		$newestnm = $item.Name
		$newestlwt = $item.LastWriteTime
		echo "update!"
	}
	echo $item.Name
	echo $item.LastWriteTime
}

# 最新アーカイブのZIP解凍
# 解答したファイルの文字コード変換 to SJIS

Pop-Location

pause

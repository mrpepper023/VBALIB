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
$currentTime = [System.DateTime]::Now

# 繝・Φ繝昴Λ繝ｪ繝輔か繝ｫ繝菴懈・
$tmp = $env:TEMP | Join-Path -ChildPath $([System.Guid]::NewGuid().Guid)
New-Item -ItemType Directory -Path $tmp | Push-Location

# 繝・Φ繝昴Λ繝ｪ繝輔か繝ｫ繝蜷阪ｒ繝・Φ繝昴Λ繝ｪ繝輔か繝ｫ繝縺ｮmoduleimporter.txt縺ｫ菫晏ｭ・
$nm = $env:TEMP | Join-Path -ChildPath "moduleimporter.txt"
Set-Content -Path $nm -Value $tmp -Force

echo $nm

# 繝繧ｦ繝ｳ繝ｭ繝ｼ繝峨ヵ繧ｩ繝ｫ繝縺ｮ譛譁ｰ繧｢繝ｼ繧ｫ繧､繝悶ｒ迚ｹ螳・
$shellapp = New-Object -ComObject Shell.Application
$dlfolder = $shellapp.Namespace("shell:Downloads").Self.Path
$targetgl = Join-Path $dlfolder VBALIB-main*.zip
$items = Get-ChildItem $targetgl -File

$newestnm = ""
$newestlwt = 0
foreach ($item in $items) {
	if($newestnm -eq ""){
		$newestnm = $item.FullName
		$newestlwt = $item.LastWriteTime
		echo "init"
	}
	if($newestlwt -lt $item.LastWriteTime){
		$newestnm = $item.FullName
		$newestlwt = $item.LastWriteTime
		echo "update!"
	}
	echo $item.FullName
	echo $item.LastWriteTime
}
echo "newest---"
echo $newestnm 
echo $newestlwt 

# 譛譁ｰ繧｢繝ｼ繧ｫ繧､繝悶・ZIP隗｣蜃・
Expand-Archive -Path $newestnm -DestinationPath $tmp

# 隗｣蜃阪＠縺溘ヵ繧｡繧､繝ｫ縺ｮ譁・ｭ励さ繝ｼ繝牙､画鋤 to SJIS
$files = Get-ChildItem -Path $nm -Include "*.bas","*.cls","*.frm","*.frx","*.bat" -Recurse
foreach ($file in $files) {
    $file.FullName
    Get-Content -Path $file.FullName -Encoding UTF8 `
    | Set-Content -Path $file.FullName -Value $content -Encoding String
}

Pop-Location

pause

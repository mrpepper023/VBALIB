@powershell -NoProfile -ExecutionPolicy Unrestricted "$s=[scriptblock]::create((gc 
¥"%~f0¥"|?{$_.readcount -gt 1})-join¥"`n¥");&$s" %*&goto:eof
# ここからPowerShellスクリプト
$currentTime = [System.DateTime]::Now

# テンポラリフォルダ作成
$tmp = $env:TEMP | Join-Path -ChildPath $([System.Guid]::NewGuid().Guid)
New-Item -ItemType Directory -Path $tmp | Push-Location

# テンポラリフォルダ名をテンポラリフォルダのmoduleimporter.txtに保存
$nm = $env:TEMP | Join-Path -ChildPath "moduleimporter.txt"
Set-Content -Path $nm -Value $tmp -Force

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

# 最新アーカイブをZIP展開
Expand-Archive -Path $newestnm -DestinationPath $tmp

# 展開したファイルのエンコーディング変換 to SJIS
$files = Get-ChildItem -Path $nm -Include "*.bas","*.cls","*.frm","*.frx","*.bat" -Recurse
foreach ($file in $files) {
    $file.FullName
    $content = Get-Content -Path $file.FullName -Encoding UTF8
    Set-Content -Path $file.FullName -Value $content -Encoding String
}

Pop-Location

pause

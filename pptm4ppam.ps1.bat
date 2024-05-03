@set PSROOT=%CD%&&set ARGS=%*&&powershell -NoProfile -ExecutionPolicy Unrestricted "$s=[scriptblock]::create((gc \"%~f0\"|?{$_.readcount -gt 1})-join\"`n\");&$s" %*&goto:eof

#' --------------------------------------------------------------------
#' From here, write it normally in Powershell
#' --------------------------------------------------------------------
Param(
    [Alias("Source", "s")]$SourcePath
)

if ("$SourcePath" -eq "") {
	Write-Host "PowerPoint Add-In Converter"
	Write-Host "---------------------------"
	Write-Host "The macro you want to add-in is [HOGE].pptm. "
	Write-Host "[HOGE] is an arbitrary name that is descriptive. "
	Write-Host ""
	Write-Host "Create and place the menu configuration "
	Write-Host "you want to reflect on the ribbon "
	Write-Host "with the file name [HOGE].xml "
	Write-Host "in the same location as [HOGE].pptm."
	Write-Host ""
	Write-Host "And if you drag and drop [HOGE].pptm "
	Write-Host "into this script, it will be converted "
	Write-Host "to [HOGE]_ppam.pptm."
	Write-Host ""
	if ($true) {Read-Host "Press [ENTER]"; Exit}
}

Get-Location 
$folder = Split-Path "$SourcePath"
if ($folder -eq "") {
	$folder = "."
}
$name = [System.IO.Path]::GetFileNameWithoutExtension("$SourcePath")
$ext = [System.IO.Path]::GetExtension("$SourcePath")
echo $folder
echo $name
echo $ext
if ($ext -ne ".pptm") {Read-Host "The target file isn't *.pptm"; Exit}

New-Item ".\.pptm4addin" -ItemType Directory
New-Item ".\.pptm4addin\arc" -ItemType Directory
Copy-Item -Path "$SourcePath" -Destination ".\.pptm4addin"
Rename-Item -Path ".\.pptm4addin\${name}${ext}" -NewName "${name}.zip"
Expand-Archive -Path ".\.pptm4addin\${name}.zip" -DestinationPath ".\.pptm4addin\arc"
If (-!(Test-Path ".\.pptm4addin\arc\customui")) {
	New-Item ".\.pptm4addin\arc\customui" -ItemType Directory
}
If (Test-Path "${folder}\${name}.xml") {
	Write-Host "${name}.xml is found!" 
	Copy-Item -Path "${folder}\${name}.xml" -Destination ".\.pptm4addin\arc\customui"
	if (Test-Path ".\.pptm4addin\arc\customui\customui.xml") {
		Remove-Item -Path ".\.pptm4addin\arc\customui\customui.xml"
	}
	Rename-Item -Path ".\.pptm4addin\arc\customui\${name}.xml" -NewName "customui.xml"
} Else {
	Write-Host "${folder}\${name}.xml is not found!" 
}
if (Test-Path ".\.pptm4addin\arc\customui\customui.xml") {
	if (Test-Path ".\.pptm4addin\arc\_rels\.rels") {
		$replace_settings = ".\.pptm4addin\arc\_rels\.rels"
		$edited = Select-String -Path "${replace_settings}" -Pattern "myCustomUI"
		if ($edited -ne $null){
		    Write-Host ".rels has already edited."
		}else{
			$before = '</Relationships>'
			$after = '<Relationship Id="myCustomUI" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customui/customui.xml"/></Relationships>'
			$str = Get-Content "${replace_settings}" -Encoding Default | % { $_ -replace "$before","$after" }
			$str | Out-File "${replace_settings}" -Encoding Default 
		}
	}
}

Compress-Archive -Path ".\.pptm4addin\arc\*" -DestinationPath ".\${name}_ppam.zip" -Force
If (Test-Path ".\${name}_ppam.pptm") {
	Remove-Item ".\${name}_ppam.pptm"
}
Rename-Item -Path ".\${name}_ppam.zip" -NewName ".\${name}_ppam.pptm"

Remove-Item -Path ".\.pptm4addin" -Recurse

if ($true) {Read-Host "Press [ENTER]"; Exit}

<#
.SYNOPSIS
Get-SpecialFolderの出力をファイルに保存する
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$osVersion = (Get-CimInstance Win32_OperatingSystem).Version
$cpu = $env:PROCESSOR_ARCHITECTURE
$edition = Get-ItemPropertyValue 'HKLM:/SOFTWARE/Microsoft/Windows NT/CurrentVersion' 'EditionID'
$now = Get-Date -Format yyyyMMdd-HHmmss

Import-Module "$PSScriptRoot/../src/PSSpecialFolder.psd1" -Force

Get-SpecialFolder -Debug -InformationAction Continue 6>&1 |
	ForEach-Object {
		if ($_ -is [System.Management.Automation.InformationRecord]) { [pscustomobject]@{
			Information = $_.ToString().Replace("`n", '')
		} }
		elseif (!$_.FolderItem) { $_ }
		else { [pscustomobject]@{
			Name = $_.Name
			Dir = $_.Dir
			Path = $_.Path
			DisplayName = $_.FolderItem.Name
			Type = $_.FolderItem.Type
		} }
	} |
	ConvertTo-Html -As List -Head '<meta charset="UTF-8">' |
	# ps5.1では必要 (https://github.com/PowerShell/PowerShell/pull/2184)
	ForEach-Object { $_.ToString().Replace('<td>*:</td>', '<td>Information:</td>') } |
	Out-File "$PSScriptRoot/$osVersion $cpu $edition $now.html" -Encoding utf8

Push-Location $PSScriptRoot

$txtFiles = Get-ChildItem "$osVersion $cpu $edition *.html" | Sort-Object -Property LastWriteTime -Descending
if (@($txtFiles).Length -ge 2) { fc.exe /n /20 $txtFiles[1].Name $txtFiles[0].Name }

Pop-Location

<#
.SYNOPSIS
Get-SpecialFolderの出力をファイルに保存する
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$osVersion = (Get-CimInstance Win32_OperatingSystem).Version
$cpu = $env:PROCESSOR_ARCHITECTURE
$edition = Get-ItemPropertyValue "HKLM:/SOFTWARE/Microsoft/Windows NT/CurrentVersion" "EditionID"
$now = Get-Date -Format yyyyMMdd-HHmmss

Import-Module "$PSScriptRoot/../src/PSSpecialFolder.psd1" -Force

Get-SpecialFolder -Debug -InformationAction Continue 6>&1 |
	Where-Object { $_ } |
	ForEach-Object {
		if ($_ -is [System.Management.Automation.InformationRecord] -or !$_.FolderItem) { $_ }
		else { [pscustomobject]@{
			Title = $_.Title
			Dir = $_.Dir
			Path = $_.Path
			DisplayName = $_.FolderItem.Name
			Type = $_.FolderItem.Type
		} }
	} |
	Out-File "$PSScriptRoot\$osVersion $cpu $edition $now.txt" -Encoding utf8

using namespace System.IO
using namespace System.Management.Automation

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

# 開発用以外のバージョンをアンロードする
Get-Module PSSpecialFolder | Remove-Module

$module = Import-Module "$PSScriptRoot/../src/PSSpecialFolder.psd1" -PassThru
$encoding = if ($PSVersionTable['PSVersion'] -ge '6.0') { 'utf8BOM' } else { 'utf8' }

$shell = New-Object -ComObject Shell.Application

& {
	$InformationPreference = 'Continue'
	Write-Information "Module Version: $((Get-Module PSSpecialFolder).Version.ToString())"
	Get-SpecialFolder -Debug
} 6>&1 |
	ForEach-Object {
		if ($_ -is [InformationRecord]) {
			[pscustomobject]@{ Information = $_.ToString().Replace("`n", '') }
		} elseif ($_.FolderItem) {
			[pscustomobject]@{
				Name = if ($_.ClassName) { "$($_.Name) ($($_.ClassName))" } else { $_.Name }
				Dir = $_.Dir
				Path = $_.Path
				DisplayName = $_.FolderItem.Name
				Type = $_.FolderItem.Type
			}
		} else {
			$folder = try { $shell.NameSpace($_.Path) } catch { $null }
			if ($folder) { Write-Warning "$($folder.Self.Name): $($_.Path)" }
			[pscustomobject]@{
				Name = if ($_.ClassName) { "$($_.Name) ($($_.ClassName))" } else { $_.Name }
				Path = $_.Path
			}
		}
	} |
	ConvertTo-Html -As List -Head '<meta charset="UTF-8">' |
	# ps5.1では必要 (https://github.com/PowerShell/PowerShell/pull/2184)
	ForEach-Object { $_.ToString().Replace('<td>*:</td>', '<td>Information:</td>') } |
	Out-File "$PSScriptRoot/$osVersion $cpu $edition $now.html" -Encoding $encoding

Remove-Module $module

Push-Location $PSScriptRoot
$txtFiles = `
	[FileInfo[]]@(Get-ChildItem "$osVersion $cpu $edition *.html" | Sort-Object -Property Name -Descending)
if ($txtFiles.Length -ge 2) {
	# fc.exeはUTF-8が文字化けするのでdiff.exeがあるならこちらを使う
	$diff = "$($Env:ProgramFiles)/Git/usr/bin/diff.exe"

	if (Test-Path $diff) { & $diff -su1 $txtFiles[1].Name $txtFiles[0].Name }
	else { fc.exe /n /20 $txtFiles[1].Name $txtFiles[0].Name }
}
Pop-Location

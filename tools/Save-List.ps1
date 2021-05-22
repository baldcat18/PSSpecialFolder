using namespace System.IO

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

$txtPath = "$PSScriptRoot/$osVersion $cpu $edition $now.txt"
$outFileArgs = @{
	LiteralPath = $txtPath
	Encoding = if ($PSVersionTable['PSVersion'] -ge '6.0') { 'utf8BOM' } else { 'utf8' }
}
"Module Version: $((Get-Module PSSpecialFolder).Version.ToString())`n" | Out-File @outFileArgs

$shell = New-Object -ComObject Shell.Application
Get-SpecialFolder -Debug |
	ForEach-Object {
		$name = if ($_.ClassName) { "$($_.Name) ($($_.ClassName))" } else { $_.Name }
		if ($_.FolderItem) {
			[pscustomobject]@{
				Name = $name
				Dir = $_.Dir
				Path = $_.Path
				DisplayName = $_.FolderItem.Name
				Type = $_.FolderItem.Type
				Category = $_.Category
			}
		} else {
			$folder = try { $shell.NameSpace($_.Path) } catch { $null }
			if ($folder) { Write-Warning "$($folder.Self.Name): $($_.Path)" }
			[pscustomobject]@{ Name = $name; Path = $_.Path; Category = $_.Category }
		}
	} |
	Format-List -GroupBy Category |
	oss -Width ([int]::MaxValue) |
	Where-Object {
		# カテゴリ名は Format-List が表示するのでオブジェクトごとに入れる必要はない
		$_ -cnotmatch '^Cate'
	} |
	Out-File @outFileArgs -Append

Remove-Module $module

Push-Location $PSScriptRoot
$txtFiles = `
	[FileInfo[]]@(Get-ChildItem "$osVersion $cpu $edition *.txt" | Sort-Object -Property Name -Descending)
if ($txtFiles.Length -ge 2) {
	# fc.exeはUTF-8が文字化けするのでdiff.exeがあるならこちらを使う
	$diff = "$($Env:ProgramFiles)/Git/usr/bin/diff.exe"

	if (Test-Path $diff) { & $diff -su1 $txtFiles[1].Name $txtFiles[0].Name }
	else { fc.exe /n /20 $txtFiles[1].Name $txtFiles[0].Name }
}
Pop-Location

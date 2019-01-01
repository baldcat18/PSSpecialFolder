Set-StrictMode -Version Latest

class SpecialFolder {
	[string]$Title
	[string]$Path
	
	hidden [string]$Dir
	hidden [__ComObject]$FolderItem
}

class FolderOption {
	[string]$Title
	[string]$Path
}

function const ([string]$name, $value) {
	New-Variable -Name $name -Value $value -Option Constant -Scope 1
}

const shell (New-Object -ComObject Shell.Application)

function newSpecialFolder {
	[OutputType([SpecialFolder])]
	param ([string]$Dir, [FolderOption]$Option = (@{}))
	
	if (!$Dir) { return }
	if ($Dir.Substring(0, 2) -eq "\\") { $Dir = "file:" + $Dir }
	elseif ($Dir.Substring(0, 6) -ne "shell:" -and $Dir.Substring(0, 5) -ne "file:") { $Dir = "file:\\\" + $Dir }
	
	try { const folder $shell.NameSpace($Dir) }
	catch { return }
	
	if (!$folder) { return }
	const folderItem $shell.NameSpace($Dir).Self
	
	$title =
		if ($Dir -match "^shell:(?:(?:\w|\s)+)$") { $Dir.Substring(6) }
		elseif ($Dir -match "^shell:::.*\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\}$") {
			const clsid $Dir.Substring($Dir.Length - 38)
			(Get-Item "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$clsid").GetValue("")
		}
		else { $Dir -replace "^.+\\(.+?)$", "`$1" }
	
	$path = $folderItem.Path
	if ($path.Substring(0,2) -eq "::") { $path = "shell:" + $path }
	
	if ($Option.Title) { $title = $Option.Title }
	if ($Option.Path) { $path = $Option.Path }
	
	return [SpecialFolder]@{ Title = $title; Path = $path; Dir = $Dir; FolderItem = $folderItem }
}

function newShellCommand {
	[OutputType([SpecialFolder])]
	param ([string]$Dir)
	
	if (!$Dir) { return }
	
	const path "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$($Dir.Substring($Dir.Length - 38))"
	if (!(Test-Path $path)) { return }
	
	return [SpecialFolder]@{ Title = (Get-Item $path).GetValue(""); Path = $Dir; Dir = $Dir }
}

<#
.SYNOPSIS
Gets the special folders for Windows. This function supports the virtual folders, e.g. Control Panel and Recycle Bin.
.OUTPUTS
SpecialFolder[]
#>
function Get-SpecialFolder {
	[CmdletBinding()]
	[OutputType([SpecialFolder[]])]
	param ()
	
	const osVerion [Environment]::OSVersion.Version
	
	# Win8.1以降
	const win81 ($osVerion -gt [version]::new(6, 3))
	# Win10以降
	const win10 ($osVerion -gt [version]::new(10, 0))
	# Win10 1607以降
	const win10_1607 ($osVerion -gt [version]::new(10, 0, 14393))
	# Win10 1703以降
	const win10_1703 ($osVerion -gt [version]::new(10, 0, 15063))
	# Win10 1709以降
	const win10_1709 ($osVerion -gt [version]::new(10, 0, 16299))
	# Win10 1803以降
	const win10_1803 ($osVerion -gt [version]::new(10, 0, 17134))
	
	const is64bit [System.Environment]::Is64BitProcess
	
	const userShellFoldersKey (Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\\User Shell Folders")
	const currentVersionKey (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion")
}

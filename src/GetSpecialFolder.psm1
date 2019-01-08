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

function const {
	param ([string]$name, $value)
	
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
	
	const osVerion ([Environment]::OSVersion.Version)
	
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
	
	const is64bit ([System.Environment]::Is64BitProcess)
	
	const userShellFoldersKey (Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\\User Shell Folders")
	const currentVersionKey (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion")
	
	Write-Information "Module Version: $((Get-Module GetSpecialFolder).Version.ToString())`n"
	
	Write-Information "Category: User's Files"
	
	# shell:Profile
	# shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}
	# shell:ThisDeviceFolder / shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9} ([このデバイス]) (Win10 1703から)
	# %USERPROFILE%
	# %HOMEDRIVE%%HOMEPATH%
	Write-Output (newSpecialFolder "shell:UsersFilesFolder")
	# Win10からサポート
	# shell:UsersFilesFolder\3D Objects
	# shell:MyComputerFolder\::{0DB7E03F-FC29-4DC6-9020-FF41B59E513A} (Win10 1709から)
	# Win10 1507から1703では3D Builderを起動した時に自動生成される
	Write-Output (newSpecialFolder "shell:3D Objects")
	# shell:MyComputerFolder\::{B4BFCC3A-DB2C-424C-B029-7FE99A87C641} (Win8.1から)
	Write-Output $(if ($win81) { newSpecialFolder "shell:ThisPCDesktopFolder"} else { newSpecialFolder ([Environment]::GetFolderPath("DesktopDirectory")) @{ Title = "DesktopDirectory" } })
	# shell:Local Documents / shell:MyComputerFolder\::{D3162B92-9365-467A-956B-92703ACA08AF} (Win10から)
	# shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103} ([マイ ドキュメント] (Win8.1から))
	# shell:MyComputerFolder\::{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0} (Win8.1から)
	Write-Output (newSpecialFolder "shell:Personal" @{ Title = "My Documents" })
	# shell:Local Downloads / shell:MyComputerFolder\::{088E3905-0323-4B02-9826-5D99428E115F} (Win10から)
	# shell:MyComputerFolder\::{374DE290-123F-4565-9164-39C4925E467B} (Win8.1から)
	Write-Output (newSpecialFolder "shell:Downloads")
	
	# shell:Local Music / shell:MyComputerFolder\::{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE} (Win10から)
	# shell:MyComputerFolder\::{1CF1260C-4DD0-4EBB-811F-33C572699FDE} (Win8.1から)
	Write-Output (newSpecialFolder "shell:My Music")
	# shell:My Music\Playlists
	# WMPやGroove ミュージックで再生リストを作成する時に自動生成される
	Write-Output (newSpecialFolder "shell:Playlists")
	
	# shell:Local Pictures / shell:MyComputerFolder\::{24AD3AD4-A569-4530-98E1-AB02F9417AA8} (Win10から)
	# shell:MyComputerFolder\::{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA} (Win8.1から)
	Write-Output (newSpecialFolder "shell:My Pictures")
	# Win8.1からサポート
	# shell:My Pictures\Camera Roll
	# カメラアプリで写真や動画を撮影する時に自動生成される
	Write-Output (newSpecialFolder "shell:Camera Roll")
	# Win10からサポート
	# shell:My Pictures\Saved Pictures
	Write-Output (newSpecialFolder "shell:SavedPictures")
	# Win8からサポート
	# shell:My Pictures\Screenshots
	# Win＋PrtScrでスクリーンショットを保存する時に自動生成される
	Write-Output (newSpecialFolder "shell:Screenshots")
	# shell:My Pictures\Slide Shows
	# 手動でフォルダーを作成しても使用可
	Write-Output (newSpecialFolder "shell:PhotoAlbums")
	
	# shell:Local Videos / shell:MyComputerFolder\::{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A} (Win10から)
	# shell:MyComputerFolder\::{A0953C92-50DC-43BF-BE83-3742FED03C9C} (Win8.1から)
	Write-Output (newSpecialFolder "shell:My Video")
	# Win10からサポート
	# shell:My Video\Captures
	# ゲームバーで動画やスクリーンショットを保存する時に自動生成される
	Write-Output (newSpecialFolder "shell:Captures")
	
	# Win10 1703からサポート
	# shell:UsersFilesFolder\AppMods
	Write-Output (newSpecialFolder "shell:AppMods")
	# shell:UsersFilesFolder\{56784854-C6CB-462B-8169-88E350ACB882}
	Write-Output (newSpecialFolder "shell:Contacts")
	Write-Output (newSpecialFolder "shell:Favorites")
	# shell:::{323CA680-C24D-4099-B94D-446DD2D7249E} ([お気に入り])
	# shell:::{D34A6CA6-62C2-4C34-8A7C-14709C1AD938} ([Common Places FS Folder])
	# shell:UsersFilesFolder\{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}
	Write-Output (newSpecialFolder "shell:Links")
	# Win10からサポート
	# shell:UsersFilesFolder\Recorded Calls
	Write-Output (newSpecialFolder "shell:Recorded Calls")
	# shell:UsersFilesFolder\{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}
	Write-Output (newSpecialFolder "shell:SavedGames")
	# shell:UsersFilesFolder\{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}
	Write-Output (newSpecialFolder "shell:Searches")
	
	# OneDriveカテゴリのフォルダーはすべてWin8.1からサポート
	Write-Information "Category: OneDrive`n"
	
	# shell:UsersFilesFolder\OneDrive
	# Win8.1ではMicrosoftアカウントでサインインする時に自動生成される
	# shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}\::{8E74D236-7F35-4720-B138-1FED0B85EA75} (Win8.1のみ)
	# shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}\::{018D5C66-4533-4307-9B53-224DE2ED1FE6} (Win10から)
	# %OneDrive% (Win10 1607から)
	Write-Output (newSpecialFolder "shell:OneDrive")
	# shell:OneDrive\Documents
	Write-Output (newSpecialFolder $(if ($win10) { "shell:OneDriveDocuments" } else { "shell:SkyDriveDocuments" }))
	# shell:OneDrive\Music
	Write-Output (newSpecialFolder $(if ($win10) { "shell:OneDriveMusic" } else { "shell:SkyDriveMusic" }))
	# shell:OneDrive\Pictures
	Write-Output (newSpecialFolder $(if ($win10) { "shell:OneDrivePictures" } else { "shell:SkyDrivePictures" }))
	# shell:OneDrive\Pictures\Camera Roll
	Write-Output (newSpecialFolder $(if ($win10) { "shell:OneDriveCameraRoll" } else { "shell:SkyDriveCameraRoll" }))
}

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
	
	const title $(
		if ($null -ne $Option.Title) { $Option.Title }
		elseif ($Dir -match "^shell:(?:(?:\w|\s)+)$") { $Dir.Substring(6) }
		elseif ($Dir -match "^shell:.*::\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\}$") {
			const clsid $Dir.Substring($Dir.Length - 38)
			(Get-Item "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$clsid").GetValue("")
		}
		else { $Dir -replace "^.+\\(.+?)$", "`$1" }
	)
	
	if ($null -ne $Option.Path) { $path = $Option.Path }
	else {
		$path = $folderItem.Path
		if ($path.Substring(0,2) -eq "::") { $path = "shell:" + $path }
	}
	
	return [SpecialFolder]@{ Title = $title; Path = $path; Dir = $Dir; FolderItem = $folderItem }
}

function newShellCommand {
	[OutputType([SpecialFolder])]
	param ([string]$Path, [string]$Title = "")
	
	if (!$Path) { return }
	
	const clsidPath "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$($Path.Substring($Path.Length - 38))"
	if (!(Test-Path $clsidPath)) { return }
	
	return [SpecialFolder]@{ Title = if ($Title) { $Title } else { (Get-Item $clsidPath).GetValue("") }; Path = $Path }
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
	const win81 ($osVerion -gt [Version]::new(6, 3))
	# Win10以降
	const win10 ($osVerion -gt [Version]::new(10, 0))
	# Win10 1607以降
	const win10_1607 ($osVerion -gt [Version]::new(10, 0, 14393))
	# Win10 1703以降
	const win10_1703 ($osVerion -gt [Version]::new(10, 0, 15063))
	# Win10 1709以降
	const win10_1709 ($osVerion -gt [Version]::new(10, 0, 16299))
	# Win10 1803以降
	const win10_1803 ($osVerion -gt [Version]::new(10, 0, 17134))
	
	const is64bitOS ([Environment]::Is64BitOperatingSystem)
	const isWow64 ($is64bitOS -and ![Environment]::Is64BitProcess)
	
	const userShellFoldersKey (Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
	const currentVersionKey (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion")
	if ($win81) { const appxKey (Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx") }
	
	Write-Information "Module Version: $((Get-Module GetSpecialFolder).Version.ToString())`n"
	
	Write-Information "Category: User's Files"
	
	# shell:Profile
	# shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}
	# shell:ThisDeviceFolder / shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9} (Win10 1703から)
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
	# shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103} ([My Documents] (Win8.1から))
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
	# shell:::{323CA680-C24D-4099-B94D-446DD2D7249E} ([Favorites])
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
	Write-Information "`nCategory: OneDrive`n"
	
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
	
	Write-Information "`nCategory: AppData`n"
	
	# %APPDATA%
	Write-Output (newSpecialFolder "shell:AppData")
	Write-Output (newSpecialFolder "shell:CredentialManager")
	Write-Output (newSpecialFolder "shell:CryptoKeys")
	Write-Output (newSpecialFolder "shell:DpapiKeys")
	Write-Output (newSpecialFolder "shell:SystemCertificates")
		
	Write-Output (newSpecialFolder "shell:Quick Launch")
	# shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}
	Write-Output (newSpecialFolder "shell:User Pinned")
	Write-Output (newSpecialFolder "shell:ImplicitAppShortcuts")
	
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:AccountPictures")
	# Win8.1以降では[LocalAppData]カテゴリになるので非表示に
	if (!$win81) { Write-Output (newSpecialFolder "shell:Cookies") }
	Write-Output (newSpecialFolder "shell:NetHood")
	# shell:::{ED50FC29-B964-48A9-AFB3-15EBB9B97F36} ([printhood delegate folder])
	Write-Output (newSpecialFolder "shell:PrintHood")
	Write-Output (newSpecialFolder "shell:Recent")
	Write-Output (newSpecialFolder "shell:SendTo")
	Write-Output (newSpecialFolder "shell:Templates")
	
	Write-Information "`nCategory: Libraries`n"
	
	$librariesPath = $userShellFoldersKey.GetValue("{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}")
	if (!$librariesPath) { $librariesPath = "$([Environment]::GetFolderPath("ApplicationData"))\Microsoft\Windows\Libraries" }
	
	# shell:UsersLibrariesFolder
	# shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}
	Write-Output (newSpecialFolder "shell:Libraries" @{ Path = $librariesPath })
	# Win10からサポート
	# shell:Libraries\CameraRoll.library-ms
	# shell:Libraries\{2B20DF75-1EDA-4039-8097-38798227D5B7}
	$cameraRollLibraryPath = $userShellFoldersKey.GetValue("{2B20DF75-1EDA-4039-8097-38798227D5B7}")
	if (!$cameraRollLibraryPath) { $cameraRollLibraryPath = "$librariesPath\CameraRoll.library-ms" }
	Write-Output (newSpecialFolder "shell:CameraRollLibrary" @{ Path = $cameraRollLibraryPath })
	# shell:Libraries\{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}
	$documentsLibraryPath = $userShellFoldersKey.GetValue("{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}")
	if (!$documentsLibraryPath) { $documentsLibraryPath = "$librariesPath\Documents.library-ms" }
	Write-Output (newSpecialFolder "shell:DocumentsLibrary" @{ Path = $documentsLibraryPath })
	# shell:Libraries\{2112AB0A-C86A-4FFE-A368-0DE96E47012E}
	$musicLibraryPath = $userShellFoldersKey.GetValue("{2112AB0A-C86A-4FFE-A368-0DE96E47012E}")
	if (!$musicLibraryPath) { $musicLibraryPath = "$librariesPath\Music.library-ms" }
	Write-Output (newSpecialFolder "shell:MusicLibrary" @{ Path = $musicLibraryPath })
	# shell:Libraries\{A990AE9F-A03B-4E80-94BC-9912D7504104}
	$picturesLibraryPath = $userShellFoldersKey.GetValue("{A990AE9F-A03B-4E80-94BC-9912D7504104}")
	if (!$picturesLibraryPath) { $picturesLibraryPath = "$librariesPath\Pictures.library-ms" }
	Write-Output (newSpecialFolder "shell:PicturesLibrary" @{ Path = $picturesLibraryPath })
	# Win10からサポート
	# shell:Libraries\SavedPictures.library-ms
	# shell:Libraries\{E25B5812-BE88-4BD9-94B0-29233477B6C3}
	$savedPicturesLibraryPath = $userShellFoldersKey.GetValue("{E25B5812-BE88-4BD9-94B0-29233477B6C3}")
	if (!$savedPicturesLibraryPath) { $savedPicturesLibraryPath = "$librariesPath\SavedPictures.library-ms" }
	Write-Output (newSpecialFolder "shell:SavedPicturesLibrary" @{ Path = $savedPicturesLibraryPath })
	# shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}\{491E922F-5643-4AF4-A7EB-4E7A138D8174}
	$videosLibraryPath = $userShellFoldersKey.GetValue("{491E922F-5643-4AF4-A7EB-4E7A138D8174}")
	if (!$videosLibraryPath) { $videosLibraryPath = "$librariesPath\Videos.library-ms" }
	Write-Output (newSpecialFolder "shell:VideosLibrary" @{ Path = $videosLibraryPath })
	
	Write-Information "`nCategory: StartMenu`n"
	
	Write-Output (newSpecialFolder "shell:Start Menu")
	Write-Output (newSpecialFolder "shell:Programs")
	Write-Output (newSpecialFolder "shell:Administrative Tools")
	Write-Output (newSpecialFolder "shell:Startup")
	
	Write-Information "`nCategory: LocalAppData`n"
	
	# %LOCALAPPDATA%
	Write-Output (newSpecialFolder "shell:Local AppData")
	Write-Output (newSpecialFolder "shell:LocalAppDataLow")
		
	# Win10 1709からサポート
	# shell:Local AppData\Desktop
	Write-Output (newSpecialFolder "shell:AppDataDesktop")
	# Win10からサポート
	# shell:Local AppData\DevelopmentFiles
	Write-Output (newSpecialFolder "shell:Development Files")
	# Win10 1709からサポート
	# shell:Local AppData\Documents
	Write-Output (newSpecialFolder "shell:AppDataDocuments")
	# Win10 1709からサポート
	# shell:Local AppData\Favorites
	Write-Output (newSpecialFolder "shell:AppDataFavorites")
	# ストアアプリの設定
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:Local AppData\Packages" @{ Title = "Settings of the Windows Apps" })
	# Win10 1709からサポート
	# shell:Local AppData\ProgramData
	Write-Output (newSpecialFolder "shell:AppDataProgramData")
	# %TEMP%
	# %TMP%
	Write-Output (newSpecialFolder ([System.IO.Path]::GetTempPath()) @{ Title = "Temporary Folder" })
	Write-Output (newSpecialFolder "shell:Local AppData\VirtualStore")
		
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:Application Shortcuts")
	Write-Output (newSpecialFolder "shell:CD Burning")
	Write-Output (newSpecialFolder "shell:GameTasks")
	Write-Output (newSpecialFolder "shell:History")
	Write-Output (newSpecialFolder "shell:Cache")
	# Win8.1でこのカテゴリに移動
	if ($win81) { Write-Output (newSpecialFolder "shell:Cookies") }
	Write-Output (newSpecialFolder "shell:Ringtones")
	# Win8からサポート
	# shell:Local AppData\Microsoft\Windows\RoamedTileImages
	Write-Output (newSpecialFolder "shell:Roamed Tile Images")
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:Roaming Tiles")
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:Local AppData\Microsoft\Windows\WinX")
		
	# Win8.1からサポート
	# shell:Local AppData\Microsoft\Windows\ConnectedSearch\History
	Write-Output (newSpecialFolder "shell:SearchHistoryFolder")
	# Win8.1からサポート
	# shell:Local AppData\Microsoft\Windows\ConnectedSearch\Templates
	Write-Output (newSpecialFolder "shell:SearchTemplatesFolder")
		
	Write-Output (newSpecialFolder $(if ($win81) { "shell:Local AppData\Microsoft\Windows Sidebar\Gadgets" } else { "shell:Gadgets" }))
	# shell:Local AppData\Microsoft\Windows Photo Gallery\Original Images
	# フォトギャラリーでファイルを編集する時に自動生成される
	Write-Output (newSpecialFolder "shell:Original Images")
		
	# shell:Local AppData\Programs
	Write-Output (newSpecialFolder "shell:UserProgramFiles")
	# shell:Local AppData\Programs\Common
	Write-Output (newSpecialFolder "shell:UserProgramFilesCommon")
	
	Write-Information "`nCategory: Public`n"
	
	# shell:::{4336A54D-038B-4685-AB02-99BB52D3FB8B}
	# shell:ThisDeviceFolder (Win10 1507から1607まで)
	# shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90} (Win10から)
	# %PUBLIC%
	Write-Output (newSpecialFolder "shell:Public")
	# Win8からサポート
	# shell:Public\AccountPictures
	Write-Output (newSpecialFolder "shell:PublicAccountPictures")
	Write-Output (newSpecialFolder "shell:Common Desktop")
	Write-Output (newSpecialFolder "shell:Common Documents")
	Write-Output (newSpecialFolder "shell:CommonDownloads")
	Write-Output (newSpecialFolder "shell:PublicLibraries")
	Write-Output (newSpecialFolder "shell:CommonMusic")
	# shell:CommonMusic\Sample Music
	Write-Output (newSpecialFolder "shell:SampleMusic")
	# Win7までサポート
	# shell:CommonMusic\Sample Playlists
	Write-Output (newSpecialFolder "shell:SamplePlaylists")
	Write-Output (newSpecialFolder "shell:CommonPictures")
	# shell:CommonPictures\Sample Pictures
	Write-Output (newSpecialFolder "shell:SamplePictures")
	Write-Output (newSpecialFolder "shell:CommonVideo")
	# shell:CommonVideo\Sample Videos
	Write-Output (newSpecialFolder "shell:SampleVideos")
	
	Write-Information "`nCategory: ProgramData`n"
	
	# %ALLUSERSPROFILE%
	# %ProgramData%
	Write-Output (newSpecialFolder "shell:Common AppData")
	# %ALLUSERSPROFILE%\OEM Links
	Write-Output (newSpecialFolder "shell:OEM Links")
		
	# Win8からサポート
	if ($win81) { Write-Output (newSpecialFolder $appxKey.GetValue("PackageRepositoryRoot") @{ Title = "Repositories of the Windows Apps" }) }
	Write-Output (newSpecialFolder "shell:Device Metadata Store")
	Write-Output (newSpecialFolder "shell:PublicGameTasks")
	# Win10からサポート
	# shell:Common AppData\Microsoft\Windows\RetailDemo
	# 市販デモ モードで使用される
	Write-Output (newSpecialFolder "shell:Retail Demo")
	Write-Output (newSpecialFolder "shell:CommonRingtones")
	Write-Output (newSpecialFolder "shell:Common Templates")
	
	Write-Information "`nCategory: CommonStartMenu`n"
	
	Write-Output (newSpecialFolder "shell:Common Start Menu")
	Write-Output (newSpecialFolder "shell:Common Programs")
	# shell:ControlPanelFolder\::{D20EA4E1-3957-11D2-A40B-0C5020524153}
	Write-Output (newSpecialFolder "shell:Common Administrative Tools")
	Write-Output (newSpecialFolder "shell:Common Startup")
	# Win10からサポート
	Write-Output (newSpecialFolder "shell:Common Start Menu Places")
	
	Write-Information "`nCategory: Windows`n"
	
	# %SystemRoot%
	# %windir%
	Write-Output (newSpecialFolder "shell:Windows")
	# shell:::{1D2680C9-0E2A-469D-B787-065558BC7D43} ([Fusion Cache]) (.NET3.5まで)
	# CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
	Write-Output (newSpecialFolder "shell:Windows\assembly" @{ Title = ".NET Framework Assemblies" })
	# shell:ControlPanelFolder\::{BD84B380-8CA2-1069-AB1D-08000948F534}
	Write-Output (newSpecialFolder "shell:Fonts")
	
	Write-Output (newSpecialFolder "shell:ResourceDir")
	# shell:ResourceDir\xxxx (xxxxはロケールIDの16進数4桁 日本語では0411)
	Write-Output (newSpecialFolder "shell:LocalizedResourcesDir")
	
	Write-Output (newSpecialFolder $(if (!$isWow64) { "shell:System" } else { "shell:SystemX86" } ) )
	if ($is64bitOS) {
		Write-Output (newSpecialFolder $(if (!$isWow64) { "shell:SystemX86" } else { "shell:Windows\SysWOW64" } ) )
	}
	
	Write-Information "`nCategory: UserProfiles`n"
	
	Write-Output (newSpecialFolder "shell:UserProfiles")
	Write-Output (newSpecialFolder (Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" "Default") @{ Title = "DefaultUserProfile" })
	
	Write-Information "`nCategory: ProgramFiles`n"
	
	# shell:ProgramFilesX64 (64ビットアプリのみ)
	# %ProgramFiles%
	Write-Output (newSpecialFolder "shell:ProgramFiles")
	if ($is64bitOS) {
		if (!$isWow64) { Write-Output (newSpecialFolder "shell:ProgramFilesX86") }
		else { Write-Output (newSpecialFolder $currentVersionKey.GetValue("ProgramW6432Dir") @{Title = "ProgramFilesX64"}) }
	}
	# shell:ProgramFilesCommonX64 (64ビットアプリのみ)
	# %CommonProgramFiles%
	Write-Output (newSpecialFolder "shell:ProgramFilesCommon")
	if ($is64bitOS) {
		if (!$isWow64) { Write-Output (newSpecialFolder "shell:ProgramFilesCommonX86") }
		else { Write-Output (newSpecialFolder $currentVersionKey.GetValue("CommonW6432Dir") @{Title = "ProgramFilesCommonX64"}) }
	}
	# Win8からサポート
	if ($win81) { Write-Output (newSpecialFolder $appxKey.GetValue("PackageRoot") @{ Title = "Windows Apps" }) }
	if ($win81) { Write-Output (newSpecialFolder "shell:ProgramFiles\Windows Sidebar\Gadgets" @{ Title = "Default Gadgets" }) } else { Write-Output (newSpecialFolder "shell:Default Gadgets") }
	Write-Output (newSpecialFolder "shell:ProgramFiles\Windows Sidebar\Shared Gadgets")
	
	Write-Information "`nCategory: Desktop / $(if ($win81) { "ThisPC" } else { "Computer" })`n"
	
	Write-Output (newSpecialFolder "shell:Desktop")
	# shell:MyComputerFolderはWin10 1507/1511だとなぜかデスクトップになってしまう
	Write-Output (newSpecialFolder "shell:MyComputerFolder")
	# Recent Places Folder
	Write-Output (newSpecialFolder "shell:::{22877A6D-37A1-461A-91B0-DBDA5AAEBC99}")
	# Win10からサポート
	# shell:::{4564B25E-30CD-4787-82BA-39E73A750B14} ([Recent Items Instance Folder])
	Write-Output (newSpecialFolder "shell:::{3134EF9C-6B18-4996-AD04-ED5912E00EB5}" @{ Title = "Recent files" })
	# Portable Devices
	Write-Output (newSpecialFolder "shell:::{35786D3C-B075-49B9-88DD-029876E11C01}")
	# Frequent Places Folder
	# Win10からサポート
	Write-Output (newSpecialFolder "shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}")
	Write-Output (newSpecialFolder "shell:RecycleBinFolder")
	# Win10からサポート
	Write-Output (newSpecialFolder "shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}" @{ Title = "Quick access" })
	# Removable Storage Devices
	# Win8からサポート
	# Win8/8.1では[PC]と同じなので非表示に
	if ($win10) { Write-Output (newSpecialFolder "shell:::{A6482830-08EB-41E2-84C1-73920C2BADB9}") }
	Write-Output (newSpecialFolder "shell:HomeGroupFolder")
	Write-Output (newSpecialFolder "shell:NetworkPlacesFolder")
	# Removable Drives
	# Win10からサポート
	Write-Output (newSpecialFolder "shell:::{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}")
	
	Write-Information "`nCategory: ControlPanel`n"
	
	# Control Panel
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}")
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\1" @{ Title = "Appearance and Personalization" })
	# shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\4
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\2" @{ Title = "Hardware and Sound" })
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\3" @{ Title = "Network and Internet" })
	# shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\10
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\5" @{ Title = "System and Security" })
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\6" @{ Title = if ($win10_1803) { "Clock and Region" } else { "Clock, Language, and Region" } })
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\7" @{ Title = "Ease of Access" })
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\8" @{ Title = "Programs" })
	Write-Output (newSpecialFolder "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\9" @{ Title = "User Accounts" })
	
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}
	# shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}
	# shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\11
	Write-Output (newSpecialFolder "shell:ControlPanelFolder" @{ Title = "All Control Panel Items" })
	
	# コントロールパネル内の項目はCLSIDだけを指定してもアクセス可能
	# 例えば[電源オプション]なら shell:::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}
	# ただしその場合はアドレスバーからコントロールパネルに移動できない
	
	# DefaultLocation
	# Win7/8のみサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{00C6D95F-329C-409A-81D7-C46C66EA7F33}")
	# Biometrics
	# Win7/8のみサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{0142E4D0-FB7A-11DC-BA4A-000FFE7AB428}")
	# Power Options
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}")
	# Credential Manager
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{1206F5F1-0569-412C-8FEC-3204630DFB70}")
	Write-Output (newSpecialFolder "shell:AddNewProgramsFolder")
	# Set User Defaults
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{E44E5D18-0652-4508-A4E2-8A090067BCB0}
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{17CD9488-1228-4B2F-88CE-4298E93E0966}" @{ Title = "Default Programs" })
	# Workspaces Center
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{241D7C96-F8BF-4F85-B01F-E2B043341A4B}" @{ Title = "RemoteApp and Desktop Connections" })
	# Windows Update
	# Win8.1までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{36EEF7DB-88AD-4E81-AD49-0E313F0C35F8}")
	# if ($win10_1709) { "Windows Defender Firewall" } else { "Windows Firewall" }
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{4026492F-2F69-46B8-B9BF-5654FC07E423}")
	# Speech Recognition
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{58E3C745-D971-4081-9034-86E34B30836A}")
	# User Accounts
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{60632754-C523-4B62-B45C-4172DA012619}")
	# HomeGroup Control Panel
	# Win10 1709までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{67CA7650-96E6-4FDD-BB43-A8E774F73A57}")
	# Performance Information and Tools
	# Win8までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{78F3955E-3B90-4184-BD14-5397C15F1EFC}")
	# Network and Sharing Center
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{8E908FC9-BECC-40F6-915B-F4CA0E70D03D}")
	# Parental Controls
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{96AE8D84-A250-4520-95A5-A47A7E3C548B}")
	# AutoPlay
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{9C60DE1E-E5FC-40F4-A487-460851A8D915}")
	# System Recovery
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{9FE63AFD-59CF-4419-9775-ABCC3849F861}")
	# Device Center
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{A8A91A66-3A7D-4424-8D24-04E180695C7A}" @{ Title = "Devices and Printers" })
	# if ($win10) { "Windows 7 File Recovery" } else { "Backup And Restore" }
	# Win8.1以外でサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{B98A2BEA-7D42-4558-8BD1-832F41BAC6FD}")
	# System
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{BB06C0E4-D293-4F75-8A90-CB05B6477EEE}")
	# if ($win10) { "Security and Maintenance CPL" } else { "Action Center CPL" }
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{BB64F8A7-BEE7-4E1A-AB8D-7D8273F7FDB6}")
	# Microsoft Windows Font Folder
	# shell:Fonts
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{BD84B380-8CA2-1069-AB1D-08000948F534}" @{ Path = "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{BD84B380-8CA2-1069-AB1D-08000948F534}" })
	# Language Settings
	# Win8からWin10 1803までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{BF782CC9-5A52-4A17-806C-2A894FFEEAC5}")
	# Display
	# Win10 1607までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{C555438B-3C23-4769-A71F-B6D3D9B6053A}")
	# Troubleshooting
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{C58C4893-3BE0-4B45-ABB5-A63E4B8C8651}")
	# Getting Started
	# Win7までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{CB1B7F8C-C50A-4176-B604-9E24DEE8D4D1}")
	# Administrative Tools
	# shell:Common Administrative Tools
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{D20EA4E1-3957-11D2-A40B-0C5020524153}" @{ Path = "shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{D20EA4E1-3957-11D2-A40B-0C5020524153}" })
	# Ease of Access
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{D555645E-D4F8-4C29-A827-D93C859C4F2A}")
	# Secure Startup
	# Enterprise/Ultimateで使用可
	# Win8からはProでも使用可
	# Win8.1からはCore/Homeでも使用可
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{D9EF8727-CAC2-4E60-809E-86F80A666C91}" @{ Title = "BitLocker Drive Encryption" })
	# Network Map
	# Win7までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{E7DE9B1A-7533-4556-9484-B26FB486475E}")
	# Windows SideShow
	# Win8までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{E95A4861-D57A-4BE1-AD0F-35267E261739}")
	# Sensors
	# Win8.1までサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{E9950154-C418-419E-A90A-20C5287AE24B}" @{ Title = if ($win81) { "Location Settings" } else { "Location and Other Sensors" } })
	# ECS
	# Win8.1からサポート
	# Win7ではKB2891638をインストールすれば使用可
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{ECDB0924-4208-451E-8EE0-373C0956DE16}" @{ Title = "Work Folders" })
	# Personalization Control Panel
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{ED834ED6-4B5A-4BFE-8F11-A626DCB6A921}")
	# History Vault
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{F6B6E965-E9B2-444B-9286-10C9152EDBC5}" @{ Title = "File History" })
	# Storage Spaces
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:ControlPanelFolder\::{F942C606-0914-47AB-BE56-1321B8035096}")
	
	Write-Output (newSpecialFolder "shell:ChangeRemoveProgramsFolder")
	Write-Output (newSpecialFolder "shell:AppUpdatesFolder")
	
	Write-Output (newSpecialFolder "shell:SyncCenterFolder")
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{2E9E59C0-B437-4981-A647-9C34B9B90891} ([Sync Setup Folder])
	Write-Output (newSpecialFolder "shell:SyncSetupFolder")
	Write-Output (newSpecialFolder "shell:ConflictFolder")
	Write-Output (newSpecialFolder "shell:SyncResultsFolder")
	
	# Taskbar
	Write-Output (newSpecialFolder "shell:$(if ($win10) { "::{21EC2020-3AEA-1069-A2DD-08002B30309D}" } else { "ControlPanelFolder" })\::{05D7B0F4-2121-4EFF-BF6B-ED3F69B894D9}" @{ Title = "Notification Area Icons" })
	# Manage Wireless Networks
	# Win8.1 Update以降ではフォルダーを開けないので非表示に
	if (!$win81) { Write-Output (newSpecialFolder "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{1FA9085F-25A2-489B-85D4-86326EEDCD87}") }
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{863AA9FD-42DF-457B-8E4D-0DE1B8015C60}
	Write-Output (newSpecialFolder "shell:PrintersFolder")
	# Bluetooth Devices
	Write-Output (newSpecialFolder "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{28803F59-3A75-4058-995F-4EE5503B023C}")
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{992CFFA0-F557-101A-88EC-00DD010CCC48}
	Write-Output (newSpecialFolder "shell:ConnectionsFolder")
	# Font Settings
	Write-Output (newSpecialFolder "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{93412589-74D4-4E4E-AD0E-E0CB621440FD}")
	# All Tasks
	Write-Output (newSpecialFolder "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{ED7BA470-8E54-465E-825C-99712043E01C}")
	
	Write-Information "`nCategory: OtherFolders`n"
	
	# Hyper-V Remote File Browsing
	# Win10 1703までサポート
	Write-Output (newSpecialFolder "shell:::{0907616E-F5E6-48D8-9D61-A91C3D28106D}")
	# Cabinet Shell Folder
	Write-Output (newSpecialFolder "shell:::{0CD7A5C0-9F37-11CE-AE65-08002B2E1262}")
	# Network
	Write-Output (newSpecialFolder "shell:::{208D2C60-3AEA-1069-A2D7-08002B30309D}")
	# DLNA Media Servers Data Source
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:::{289AF617-1CC3-42A6-926C-E6A863F0E3BA}")
	# Results Folder
	Write-Output (newSpecialFolder "shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}")
	# Explorer Browser Results Folder
	Write-Output (newSpecialFolder "shell:::{418C8B64-5463-461D-88E0-75E2AFA3C6FA}," @{ Title = "" })
	# Win8からサポート
	Write-Output (newSpecialFolder "shell:AppsFolder")
	# Command Folder
	Write-Output (newSpecialFolder "shell:::{437FF9C0-A07F-4FA0-AF80-84B6C6440A16}")
	# Other Users Folder
	Write-Output (newSpecialFolder "shell:::{6785BFAC-9D2D-4BE5-B7E2-59937E8FB80A}")
	# Programs Folder
	# Win8.1 Update以降ではフォルダーを開けないので非表示に
	if (!$win81) { Write-Output (newSpecialFolder "shell:::{7BE9D83C-A729-4D97-B5A7-1B7313C39E0A}") }
	# Programs Folder and Fast Items
	# Win8.1 Update以降ではフォルダーを開けないので非表示に
	if (!$win81) { Write-Output (newSpecialFolder "shell:::{865E5E76-AD83-4DCA-A109-50DC2113CE9A}") }
	# search:
	# search-ms:
	Write-Output (newSpecialFolder "shell:SearchHomeFolder")
	# Win8.1 UpdateからWin10 1511までサポート
	Write-Output (newSpecialFolder "shell:StartMenuAllPrograms")
	# 企業向けエディションで使用可
	Write-Output (newSpecialFolder "shell:::{AFDB1F70-2A4C-11D2-9039-00C04F8EEB3E}" @{ Title = "Offline Files Folder" })
	# delegate folder that appears in Computer
	Write-Output (newSpecialFolder "shell:::{B155BDF8-02F0-451E-9A26-AE317CFD7779}")
	# AppSuggestedLocations
	Write-Output (newSpecialFolder "shell:::{C57A6066-66A3-4D91-9EB9-41532179F0A5}")
	# Win10 1709までサポート
	Write-Output (newSpecialFolder "shell:Games")
	# Previous Versions Results Folder
	Write-Output (newSpecialFolder "shell:::{F8C2AB3B-17BC-41DA-9758-339D7DBF2D88}")
	
	if ($DebugPreference -eq "SilentlyContinue") { return }
	
	# 通常とは違う名前がエクスプローラーのタイトルバーに表示されるフォルダー
	Write-Information "`nCategory: AnotherName`n"
	
	# (if ($win10_1703) { "UsersFilesFolder" } else { "Public" })
	# shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90} (Win10 1507から1607まで)
	# shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9} (Win10 1703から)
	Write-Output (newSpecialFolder "shell:ThisDeviceFolder")
	# My Documents (Documents)
	# Win8までだと別名にならないので非表示に
	if ($win81) { Write-Output (newSpecialFolder "shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103}" @{ Title = "My Documents" }) }
	# Favorites (Links)
	Write-Output (newSpecialFolder "shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}")
	# Common Places FS Folder (Links)
	Write-Output (newSpecialFolder "shell:::{D34A6CA6-62C2-4C34-8A7C-14709C1AD938}")
	# printhood delegate folder (PrintHood)
	Write-Output (newSpecialFolder "shell:::{ED50FC29-B964-48A9-AFB3-15EBB9B97F36}")
	# Fusion Cache (.NET Framework Assemblies)
	# .NET3.5まで
	# CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
	Write-Output (newSpecialFolder "shell:::{1D2680C9-0E2A-469D-B787-065558BC7D43}")
	# Recent Items Instance Folder (Recent files)
	# Win10から
	Write-Output (newSpecialFolder "shell:::{4564B25E-30CD-4787-82BA-39E73A750B14}")
	# Sync Setup Folder (SyncSetupFolder)
	Write-Output (newSpecialFolder "shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{2E9E59C0-B437-4981-A647-9C34B9B90891}")
	
	# フォルダー以外のshellコマンド
	Write-Information "`nCategory: OtherShellCommands`n"
	
	# if ($win81) { "Taskbar" } else { "Taskbar and Start Menu" }
	Write-Output (newShellCommand "shell:::{0DF44EAA-FF21-4412-828E-260A8728E7F1}")
	# Search
	# Win10 1511まで
	Write-Output (newShellCommand "shell:::{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}")
	# Help and Support
	# Win8.1まで
	Write-Output (newShellCommand "shell:::{2559A1F1-21D7-11D4-BDAF-00C04F60B9F0}")
	# Run...
	Write-Output (newShellCommand "shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}")
	# E-mail
	Write-Output (newShellCommand "shell:::{2559A1F5-21D7-11D4-BDAF-00C04F60B9F0}")
	# Set Program Access and Defaults
	Write-Output (newShellCommand "shell:::{2559A1F7-21D7-11D4-BDAF-00C04F60B9F0}")
	# Win8から
	Write-Output (newShellCommand "shell:::{2559A1F8-21D7-11D4-BDAF-00C04F60B9F0}" $(if ($win10) { "Cortana" } else { "Search" }))
	# Show Desktop
	# Win+Dと同じ
	Write-Output (newShellCommand "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}")
	# Window Switcher
	# Win7ではCtrl+Win+Tab、Win8/8.1ではCtrl+Alt+Tab、Win10 1607以降ではWin+Tabと同じ (Win10 1507/1511では使用不可)
	Write-Output (newShellCommand "shell:::{3080F90E-D7AD-11D9-BD98-0000947B0257}")
	# Windows Sidebar Properties
	# Win7まで
	Write-Output (newShellCommand "shell:::{37EFD44D-EF8D-41B1-940D-96973A50E9E0}")
	Write-Output (newShellCommand "shell:::{38A98528-6CBF-4CA9-8DC0-B1E1D10F7B1B}" "Connect To")
	# Phone and Modem Control Panel
	Write-Output (newShellCommand "shell:::{40419485-C444-4567-851A-2DD7BFA1684D}")
	# Open in new window
	# Win8.1から
	Write-Output (newShellCommand "shell:::{52205FD8-5DFB-447D-801A-D0B52F2E83E1}" "File Explorer")
	# Mobility Center Control Panel
	Write-Output (newShellCommand "shell:::{5EA4F148-308C-46D7-98A9-49041B1DD468}")
	# Region and Language
	Write-Output (newShellCommand "shell:::{62D8ED13-C9D0-4CE8-A914-47DD628FB1B0}")
	# Windows Features
	Write-Output (newShellCommand "shell:::{67718415-C450-4F3C-BF8A-B487642DC39B}")
	# Mouse Control Panel
	Write-Output (newShellCommand "shell:::{6C8EEC18-8D75-41B2-A177-8831D59D2D50}")
	# Folder Options
	Write-Output (newShellCommand "shell:::{6DFD7C5C-2451-11D3-A299-00C04F8EF6AF}")
	# Keyboard Control Panel
	Write-Output (newShellCommand "shell:::{725BE8F7-668E-4C7B-8F90-46BDB0936430}")
	# Device Manager
	Write-Output (newShellCommand "shell:::{74246BFC-4C96-11D0-ABEF-0020AF6B0B7A}")
	# CardSpace
	# Win8まで
	Write-Output (newShellCommand "shell:::{78CB147A-98EA-4AA6-B0DF-C8681F69341C}")
	# User Accounts
	# netplwiz.exe / control.exe userpasswords2
	Write-Output (newShellCommand "shell:::{7A9D77BD-5403-11D2-8785-2E0420524153}")
	# Tablet PC Settings Control Panel
	Write-Output (newShellCommand "shell:::{80F3F1D5-FECA-45F3-BC32-752C152E456E}")
	# Internet Folder
	# Win10以降では開けないので非表示に
	Write-Output $(if (!$win10) { newShellCommand "shell:InternetFolder" })
	# Indexing Options Control Panel
	Write-Output (newShellCommand "shell:::{87D66A43-7B11-4A28-9811-C86EE395ACF7}")
	# Portable Workspace Creator
	# Win8から
	# Enterpriseで使用可
	# Win10 1607以降ではProでも使用可
	Write-Output (newShellCommand "shell:::{8E0C279D-0BD1-43C3-9EBD-31C3DC5B8A77}")
	# Biometrics navigate target object (Welcome to Biometric Devices)
	# Win8まで
	Write-Output (newShellCommand "shell:::{8E35B548-F174-4C7D-81E2-8ED33126F6FD}")
	# Infrared
	# Win10 1607から
	Write-Output (newShellCommand "shell:::{A0275511-0E86-4ECA-97C2-ECD8F1221D08}")
	# Internet Options
	Write-Output (newShellCommand "shell:::{A3DD4F92-658A-410F-84FD-6FBBBEF2FFFE}")
	# Color Management
	Write-Output (newShellCommand "shell:::{B2C761C6-29BC-4F19-9251-E6195265BAF1}")
	# Windows Anytime Upgrade
	# Win8.1まで
	Write-Output (newShellCommand "shell:::{BE122A0E-4503-11DA-8BDE-F66BAD1E3F3A}")
	# Biometrics navigate target object (Biometric Devices\Message)
	# Win8まで
	# shell:::{CCFB7955-B4DC-42CE-893D-884D72DD6B19}
	Write-Output (newShellCommand "shell:::{CBC84B69-69EA-439B-B791-DF15F60333CF}")
	# Text to Speech Control Panel
	Write-Output (newShellCommand "shell:::{D17D1D6D-CC3F-4815-8FE3-607E7D5D10B3}")
	# Add Network Place
	Write-Output (newShellCommand "shell:::{D4480A50-BA28-11D1-8E75-00C04FA31A86}")
	# Windows Defender
	# Win10 1607まで
	Write-Output (newShellCommand "shell:::{D8559EB9-20C0-410E-BEDA-7ED416AECC2A}")
	# Date and Time Control Panel
	Write-Output (newShellCommand "shell:::{E2E7934B-DCE5-43C4-9576-7FE4F75E7480}")
	# Sound Control Panel
	Write-Output (newShellCommand "shell:::{F2DDFC82-8F12-4CDD-B7DC-D4FE1425AA4D}")
	# Pen and Touch Control Panel
	Write-Output (newShellCommand "shell:::{F82DF8F7-8B9F-442E-A48C-818EA735FF9B}")
	
	# フォルダーとして使えないshellコマンド
	Write-Information "`nCategory: Unusable`n"
	
	Write-Output (newSpecialFolder "shell:MAPIFolder")
	Write-Output (newSpecialFolder "shell:RecordedTVLibrary")
	
	Write-Output (newShellCommand "shell:::{00020D75-0000-0000-C000-000000000046}")
	# Desktop
	Write-Output (newShellCommand "shell:::{00021400-0000-0000-C000-000000000046}")
	# Shortcut
	Write-Output (newShellCommand "shell:::{00021401-0000-0000-C000-000000000046}")
	# Win10 1507から1703まで
	Write-Output (newShellCommand "shell:::{047EA9A0-93BB-415F-A1C3-D7AEB3DD5087}")
	# Open With Context Menu Handler
	Write-Output (newShellCommand "shell:::{09799AFB-AD67-11D1-ABCD-00C04FC30936}")
	# Folder Shortcut
	Write-Output (newShellCommand "shell:::{0AFACED1-E828-11D1-9187-B532F1E9575D}")
	Write-Output (newShellCommand "shell:::{0C39A5CF-1A7A-40C8-BA74-8900E6DF5FCD}")
	Write-Output (newShellCommand "shell:::{0D45D530-764B-11D0-A1CA-00AA00C16E65}")
	# Shell File System Folder
	# Win8から
	Write-Output (newShellCommand "shell:::{0E5AAE11-A475-4C5B-AB00-C66DE400274E}")
	# Device Center Print Context Menu Extension
	Write-Output (newShellCommand "shell:::{0E6DAA63-DD4E-47CE-BF9D-FDB72ECE4A0D}")
	# IE History and Feeds Shell Data Source for Windows Search
	Write-Output (newShellCommand "shell:::{11016101-E366-4D22-BC06-4ADA335C892B}")
	# OpenMediaSharing
	Write-Output (newShellCommand "shell:::{17FC1A80-140E-4290-A64F-4A29A951A867}")
	# CLSID_DBFolderBoth
	Write-Output (newShellCommand "shell:::{1BEF2128-2F96-4500-BA7C-098DC0049CB2}")
	# CompatContextMenu Class
	Write-Output (newShellCommand "shell:::{1D27F844-3A1F-4410-85AC-14651078412D}")
	# Windows Security
	Write-Output (newShellCommand "shell:::{2559A1F2-21D7-11D4-BDAF-00C04F60B9F0}")
	# Location Folder
	Write-Output (newShellCommand "shell:::{267CF8A9-F4E3-41E6-95B1-AF881BE130FF}")
	# Enhanced Storage Context Menu Handler Class
	Write-Output (newShellCommand "shell:::{2854F705-3548-414C-A113-93E27C808C85}")
	# System Restore
	Write-Output (newShellCommand "shell:::{3F6BC534-DFA1-4AB4-AE54-EF25A74E0107}")
	# Start Menu Folder
	Write-Output (newShellCommand "shell:::{48E7CAAB-B918-4E58-A94D-505519C795DC}")
	# IGD Property Page
	Write-Output (newShellCommand "shell:::{4A1E5ACD-A108-4100-9E26-D2FAFA1BA486}")
	# LzhCompressedFolder2
	# Win10 1607まで
	Write-Output (newShellCommand "shell:::{4F289A46-2BBB-4AE8-9EDA-E5E034707A71}")
	# This PC
	# Win10から
	Write-Output (newShellCommand "shell:::{5E5F29CE-E0A8-49D3-AF32-7A7BDC173478}")
	Write-Output (newShellCommand "shell:::{62AE1F9A-126A-11D0-A14B-0800361B1103}")
	# Search Connector Folder
	Write-Output (newShellCommand "shell:::{72B36E70-8700-42D6-A7F7-C9AB3323EE51}")
	# CryptPKO Class
	Write-Output (newShellCommand "shell:::{7444C717-39BF-11D1-8CD9-00C04FC29D45}")
	# Temporary Internet Files
	Write-Output (newShellCommand "shell:::{7BD29E00-76C1-11CF-9DD0-00A0C9034933}")
	# Temporary Internet Files
	Write-Output (newShellCommand "shell:::{7BD29E01-76C1-11CF-9DD0-00A0C9034933}")
	# if ($win10_1703) { "" } else { "Briefcase" }
	Write-Output (newShellCommand "shell:::{85BBD920-42A0-1069-A2E4-08002B30309D}")
	# Shortcut
	Write-Output (newShellCommand "shell:::{85CFCCAF-2D14-42B6-80B6-F40F65D016E7}")
	# Mobile Broadband Profile Settings Editor
	Write-Output (newShellCommand "shell:::{87630419-6216-4FF8-A1F0-143562D16D5C}")
	# Compressed (zipped) Folder SendTo Target
	Write-Output (newShellCommand "shell:::{888DCA60-FC0A-11CF-8F0F-00C04FD7D062}")
	# ActiveX Cache Folder
	Write-Output (newShellCommand "shell:::{88C6C381-2E85-11D0-94DE-444553540000}")
	# Libraries delegate folder that appears in Users Files Folder
	Write-Output (newShellCommand "shell:::{896664F7-12E1-490F-8782-C0835AFD98FC}")
	# Windows Search Service Media Center Namespace Extension Handler
	# Win10 1607まで
	Write-Output (newShellCommand "shell:::{98D99750-0B8A-4C59-9151-589053683D73}")
	# MAPI Shell Context Menu
	Write-Output (newShellCommand "shell:::{9D3C0751-A13F-46A6-B833-B46A43C30FE8}")
	# Previous Versions
	Write-Output (newShellCommand "shell:::{9DB7A13C-F208-4981-8353-73CC61AE2783}")
	# Mail Service
	Write-Output (newShellCommand "shell:::{9E56BE60-C50F-11CF-9A2C-00A0C90A90CE}")
	# Desktop Shortcut
	Write-Output (newShellCommand "shell:::{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}")
	# DevicePairingFolder Initialization
	Write-Output (newShellCommand "shell:::{AEE2420F-D50E-405C-8784-363C582BF45A}")
	# CLSID_DBFolder
	Write-Output (newShellCommand "shell:::{B2952B16-0E07-4E5A-B993-58C52CB94CAE}")
	# Device Center Scan Context Menu Extension
	Write-Output (newShellCommand "shell:::{B5A60A9E-A4C7-4A93-AC6E-0B76D1D87DC4}")
	# DeviceCenter Initialization
	Write-Output (newShellCommand "shell:::{C2B136E2-D50E-405C-8784-363C582BF43E}")
	# Win10 1507から1607まで
	Write-Output (newShellCommand "shell:::{D9AC5E73-BB10-467B-B884-AA1E475C51F5}")
	# delegate folder that appears in Users Files Folder
	Write-Output (newShellCommand "shell:::{DFFACDC5-679F-4156-8947-C5C76BC0B67F}")
	# CompressedFolder
	Write-Output (newShellCommand "shell:::{E88DCCE0-B7B3-11D1-A9F0-00AA0060FA31}")
	# MyDocs Drop Target
	Write-Output (newShellCommand "shell:::{ECF03A32-103D-11D2-854D-006008059367}")
	# Shell File System Folder
	Write-Output (newShellCommand "shell:::{F3364BA0-65B9-11CE-A9BA-00AA004AE837}")
	# Sticky Notes Namespace Extension for Windows Desktop Search
	# Win10 1607まで
	Write-Output (newShellCommand "shell:::{F3F5824C-AD58-4728-AF59-A1EBE3392799}")
	# Subscription Folder
	Write-Output (newShellCommand "shell:::{F5175861-2688-11D0-9C5E-00AA00A45957}")
	# Internet Shortcut
	Write-Output (newShellCommand "shell:::{FBF23B40-E3F0-101B-8488-00AA003E56F8}")
	# History
	Write-Output (newShellCommand "shell:::{FF393560-C2A7-11CF-BFF4-444553540000}")
	# Windows Photo Viewer Image Verbs
	Write-Output (newShellCommand "shell:::{FFE2A43C-56B9-4BF5-9A79-CC6D4285608A}")
	
	<#
	Write-Output (newSpecialFolder "shell:")
	#>
}

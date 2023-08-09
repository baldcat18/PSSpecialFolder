using namespace System.Diagnostics
using namespace System.Diagnostics.CodeAnalysis
using namespace System.Drawing
using namespace System.IO
using namespace System.Management.Automation
using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Windows.Input
using namespace System.Windows.Interop
using namespace System.Windows.Markup
using namespace Win32API

param([bool]$Develop)

Set-StrictMode -Version Latest

if ($Develop) { $DebugPreference = 'Inquire' }

if ([Environment]::OSVersion.Platform -ne 'Win32NT') {
	throw [PlatformNotSupportedException]'The PSSpecialFolder module supports Windows only.'
	return
}

class OS {
	static [bool]$Win81_Only # Win8.1のみ
	static [bool]$Win10 # Win10以降
	static [bool]$Win10_1607 # Win10 1607以降
	static [bool]$Win10_1803 # Win10 1803以降
	static [bool]$Win10_20h2_Only # Win10 20H2のみ
	static [bool]$Win10_20h2 # Win10 20H2以降
	static [bool]$Win10_21h1 # Win10 21H1以降
	static [bool]$Win10_21h2 # Win10 21H2以降
	static [bool]$Win11 # Win11以降
	static [bool]$Win11_22h2 # Win11 22H2以降

	static OS() {
		$os = & $PSScriptRoot\Get-OSVersion.ps1
		$version = $os['Version']
		$verString = $os['VersionString']

		[OS]::Win81_Only = $verString -eq '6.3.9600.0'
		[OS]::Win10 = $version -gt '10.0.10240'
		[OS]::Win10_1607 = $version -gt '10.0.14393'
		[OS]::Win10_1803 = $version -gt '10.0.17134'
		[OS]::Win10_20h2_Only = $verString -eq '10.0.19042'
		[OS]::Win10_20h2 = $version -gt '10.0.19042'
		[OS]::Win10_21h1 = $version -gt '10.0.19043'
		[OS]::Win10_21h2 = $version -gt '10.0.19044'
		[OS]::Win11 = $version -gt '10.0.22000'
		[OS]::Win11_22h2 = $version -gt '10.0.22621'
	}
}

do {
	if ([OS]::Win11) { break }
	if ([OS]::Win10_21h1 -and ![OS]::Win11) { break }
	if ([OS]::Win10_20h2_Only) { break }
	if ([OS]::Win81_Only) { break }

	if ([OS]::Win10) {
		Write-Warning 'The PSSpecialFolder module supports Windows 10 Version 20H2+.'
	} else {
		Write-Warning 'The PSSpecialFolder module supports Windows 8.1, 10 and 11.'
	}
} while ($false)


function getItem {
	[OutputType([Microsoft.Win32.RegistryKey])]
	param ([string]$Path)

	return Get-Item -LiteralPath $Path -ErrorAction SilentlyContinue
}

function getItemValue {
	param ($Key, [string]$Name)

	if ($Key -is [string]) { $Key = getItem $Key }

	if (!$Key) { return }
	if ($Key -is [Microsoft.Win32.RegistryKey]) { return $Key.GetValue($Name) }

	throw "TypeError: $Key is $($Key.GetType().Name)"
}

# pwsh.exeがあるフォルダーにパスが通っていない場合もあるのでAPIから取得する
$powershellPath = [Process]::GetCurrentProcess().Path
# ISEなど場合もあるので名前を明示する
if ($powershellPath -notmatch '\\(?:powershell|pwsh)\.exe$') {
	$powershellPath = `
		if ($PSVersionTable['PSVersion'].Major -le 5) { 'powershell.exe' } `
		elseif (!(Get-Command pwsh.exe -ErrorAction SilentlyContinue)) { 'powershell.exe' } `
		else { 'pwsh.exe' }
}

$isWslEnabled = Test-Path 'HKLM:\SOFTWARE\Classes\Directory\shell\WSL\command'
$canFolderBeOpenedAsAdmin = `
	!(getItemValue 'HKLM:\SOFTWARE\Classes\AppID\{CDCBCFCA-3CDC-436f-A4E2-0E02075250C2}' RunAs)

enum PropertyTypes {
	None
	StartProcess
	Verb
	NotSupported
}

class SpecialFolder {
	[string]$Name
	[string]$Path

	hidden [string]$Dir
	hidden [__ComObject]$FolderItem
	hidden [PropertyTypes]$PropertyTypes
	hidden [__ComObject]$PropertiesVerb
	hidden [__ComObject]$FolderItemForProperties
	hidden [string]$Category
	# Save-List.ps1で使用
	hidden [string]$ClassName

	[string]ToString() {
		return "$($this.Name) [$($this.Path)]"
	}

	hidden [void]Open([string]$Verb) {
		Start-Process explorer.exe $(if ($this.Dir) { $this.Dir } else { $this.Path }) -Verb $Verb
	}
	hidden [void]SetProperties() {
		if ($this.PropertyTypes -ne 'None') { return }
		if ($this -is [FileFolder] -and !$this.FolderItemForProperties) {
			$this.PropertyTypes = 'StartProcess'
			return
		}

		$item = $this.FolderItemForProperties
		if ($null -eq $item) { $item = $this.FolderItem }

		$verbs = $item.Verbs()
		if ($verbs -and $verbs.Count) {
			$verb = $verbs.Item($verbs.Count - 1)
			if ($verb.Name -eq $script:propertiesName) {
				$this.PropertiesVerb = $verb
				$this.PropertyTypes = 'Verb'
				return
			}
		}

		$this.PropertyTypes = 'NotSupported'
	}
	hidden [bool]HasProperties() {
		$this.SetProperties()
		return $this.PropertyTypes -ne 'NotSupported'
	}
	hidden [void]Properties() {
		switch ($this.PropertyTypes) {
			StartProcess { Start-Process $this.Dir -Verb properties }
			Verb { $this.PropertiesVerb.DoIt() }
			NotSupported { throw [InvalidOperationException]'The properties of this folder can''t be shown.' }
			None {
				$this.SetProperties()
				$this.Properties()
			}
		}
	}
	hidden [void]Powershell([string]$Verb) {
		throw [InvalidOperationException]'This is not a directory.'
	}
	hidden [void]Cmd([string]$Verb) {
		throw [InvalidOperationException]'This is not a directory.'
	}
}

class FileFolder: SpecialFolder {
	hidden [void]Powershell([string]$Verb) {
		$startArgs = @{
			FilePath = $script:powershellPath
			ArgumentList = "-NoExit -Command `"Push-Location -LiteralPath '$($this.Path)'`""
			Verb = $Verb
		}
		Start-Process @startArgs
	}
	hidden [void]Cmd([string]$Verb) {
		Start-Process cmd.exe "/k pushd `"$($this.Path)`"" -Verb $Verb
	}
	hidden [void]Wsl([string]$Verb) {
		if ([OS]::Win10_21h2) { Start-Process wsl.exe "--cd `"$($this.Path)`"" -Verb $Verb }
		else { Start-Process cmd.exe "/c pushd `"$($this.Path)`" & wsl.exe" -Verb $Verb }
	}
}

Add-Type -ErrorAction Stop `
	-TypeDefinition (Get-Content -LiteralPath "$PSScriptRoot\KnownFolder.cs" -Raw -ErrorAction Stop)


$shell = New-Object -ComObject Shell.Application
$propertiesName = @($shell.NameSpace([Environment+SpecialFolder]::Desktop).Self.Verbs())[-1].Name

# PSSpecialFolder.Tests.ps1を実行するときに変数未定義エラーにならないようにするためのダミー
$categoryName = ''

function newSpecialFolder {
	[OutputType([SpecialFolder])]
	param ([string]$Dir, [string]$Name = '', [string]$Path = '', [__ComObject]$FolderItemForProperties = $null)

	if (!$Dir) { return }
	if ($Dir -match '^\\\\') { $Dir = 'file:' + $Dir }
	elseif ($Dir -notmatch '^(?:file|shell):') { $Dir = "file:\\\$Dir" }

	try { $folder = $shell.NameSpace($Dir) }
	catch { return }

	if (!$folder) { return }
	$folderItem = $shell.NameSpace($Dir).Self

	if (!$Path) { $Path = $folderItem.Path -replace '^::', 'shell:::' }

	$isFileFolder = Test-Path $path -PathType Container
	$className = if ($Dir -match '^shell:.*::(\{\w{8}-\w{4}-\w{4}-\w{4}-\w{12}\})$') {
		getItemValue "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$($Matches[1])"
	}

	$initializer = @{
		Name = $(
			if ($Name) { $Name }
			elseif ($Dir -match '^shell:((?:\w|\s)+)$') { $Matches[1] }
			elseif ($className) { $className }
			else { $Dir -replace '^.+\\(.+?)$', '$1' }
		)
		Path = $Path
		Dir = $Dir
		FolderItem = $folderItem
		FolderItemForProperties = $FolderItemForProperties
		Category = $categoryName
		ClassName = if ($Name) { $className }
	}
	return $(if ($isFileFolder) { [FileFolder]$initializer } else { [SpecialFolder]$initializer })
}

function newShellCommand {
	[OutputType([SpecialFolder])]
	param ([string]$Clsid, [string]$Name = '')

	if (!$Clsid) { return }

	$path = "Microsoft.PowerShell.Core\Registry::HKEY_CLASSES_ROOT\CLSID\$Clsid"
	if (!(Test-Path $path)) { return }
	$className = getItemValue $path

	return [SpecialFolder]@{
		Name = if ($Name) { $Name } else { $className }
		Path = "shell:::$Clsid"
		PropertyTypes = 'NotSupported'
		Category = $categoryName
		ClassName = if ($Name) { $className }
	}
}

function getDirectoryFolderItem {
	[OutputType([__ComObject])]
	param ([string]$path)

	return $shell.NameSpace((Split-Path $path)).Items().Item((Split-Path $path -Leaf))
}

$folderGuids = & $PSScriptRoot\FolderGuids.ps1

function getKnownFolderPath {
	[OutputType([string])]
	param ([string]$Name)

	$folder = [KnownFolder]::new($folderGuids[$Name], 0)
	if ($folder.Result -eq 'OK') { return $folder.Path }
}

function getSpecialFolder {
	# $categoryNameはnewSpecialFolderやnewShellCommand関数で参照する
	[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', 'categoryName')]

	[OutputType([SpecialFolder[]])]
	param ([bool]$IncludeShellCommand)

	$is64bitOS = [Environment]::Is64BitOperatingSystem
	$isWow64 = $is64bitOS -and ![Environment]::Is64BitProcess

	$currentVersionKey = getItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion'
	$appxKey = getItem 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Appx'

	$categoryName = 'User''s Files'

	# shell:Profile
	# shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}
	# shell:ThisDeviceFolder / shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9} (Win10 1703から)
	# %USERPROFILE%
	# %HOMEDRIVE%%HOMEPATH%
	newSpecialFolder 'shell:UsersFilesFolder' -FolderItemForProperties $shell.NameSpace([Environment+SpecialFolder]::UserProfile).Self
	# shell:MyComputerFolder\::{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}
	newSpecialFolder 'shell:ThisPCDesktopFolder' 'DesktopFolder'
	# shell:Local Documents / shell:MyComputerFolder\::{D3162B92-9365-467A-956B-92703ACA08AF} (Win10 1507から)
	# shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103} ([My Documents])
	# shell:MyComputerFolder\::{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}
	newSpecialFolder 'shell:Personal' 'My Documents'
	# shell:Local Downloads / shell:MyComputerFolder\::{088E3905-0323-4B02-9826-5D99428E115F} (Win10 1507から)
	# shell:MyComputerFolder\::{374DE290-123F-4565-9164-39C4925E467B}
	newSpecialFolder 'shell:Downloads'

	# shell:Local Music / shell:MyComputerFolder\::{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE} (Win10 1507から)
	# shell:MyComputerFolder\::{1CF1260C-4DD0-4EBB-811F-33C572699FDE}
	newSpecialFolder 'shell:My Music'
	# WMPやGroove ミュージックで再生リストを作成する時に自動生成される
	newSpecialFolder 'shell:Playlists'

	# shell:Local Pictures / shell:MyComputerFolder\::{24AD3AD4-A569-4530-98E1-AB02F9417AA8} (Win10 1507から)
	# shell:MyComputerFolder\::{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}
	newSpecialFolder 'shell:My Pictures'
	# カメラアプリで写真や動画を撮影する時に自動生成される
	newSpecialFolder 'shell:Camera Roll'
	# Win10 1507からサポート
	newSpecialFolder 'shell:SavedPictures'
	# Win＋PrtScrでスクリーンショットを保存する時に自動生成される
	newSpecialFolder 'shell:Screenshots'
	newSpecialFolder 'shell:PhotoAlbums'

	# shell:Local Videos / shell:MyComputerFolder\::{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A} (Win10 1507から)
	# shell:MyComputerFolder\::{A0953C92-50DC-43BF-BE83-3742FED03C9C}
	newSpecialFolder 'shell:My Video'
	# Win10 1507からサポート
	# ゲームバーで動画やスクリーンショットを保存する時に自動生成される
	newSpecialFolder 'shell:Captures'

	# Win10 1507からサポート
	# shell:::{0DB7E03F-FC29-4DC6-9020-FF41B59E513A} (Win10 1709から)
	# Win10 1507から1703では3D Builderを起動した時に自動生成される
	newSpecialFolder 'shell:3D Objects'
	# Win10 1703からサポート
	newSpecialFolder 'shell:AppMods'
	# shell:UsersFilesFolder\{56784854-C6CB-462B-8169-88E350ACB882}
	newSpecialFolder 'shell:Contacts'
	newSpecialFolder 'shell:Favorites'
	# shell:::{323CA680-C24D-4099-B94D-446DD2D7249E} ([Favorites])
	# shell:::{D34A6CA6-62C2-4C34-8A7C-14709C1AD938} ([Common Places FS Folder])
	# shell:UsersFilesFolder\{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}
	newSpecialFolder 'shell:Links'
	# Win10 1507からサポート
	newSpecialFolder 'shell:Recorded Calls'
	# shell:UsersFilesFolder\{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}
	newSpecialFolder 'shell:SavedGames'
	# shell:UsersFilesFolder\{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}
	newSpecialFolder 'shell:Searches'

	$categoryName = 'OneDrive'

	# Win8.1ではMicrosoftアカウントでサインインする時に自動生成される
	# shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}\::{8E74D236-7F35-4720-B138-1FED0B85EA75} (Win8.1のみ)
	# shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}\::{018D5C66-4533-4307-9B53-224DE2ED1FE6} (Win10 1507から)
	# %OneDrive% (Win10 1607から)
	newSpecialFolder 'shell:OneDrive'
	newSpecialFolder $(if ([OS]::Win10) { 'shell:OneDriveDocuments' } else { 'shell:SkyDriveDocuments' })
	newSpecialFolder $(if ([OS]::Win10) { 'shell:OneDriveMusic' } else { 'shell:SkyDriveMusic' })
	newSpecialFolder $(if ([OS]::Win10) { 'shell:OneDrivePictures' } else { 'shell:SkyDrivePictures' })
	newSpecialFolder $(if ([OS]::Win10) { 'shell:OneDriveCameraRoll' } else { 'shell:SkyDriveCameraRoll' })

	$categoryName = 'AppData'

	# %APPDATA%
	newSpecialFolder 'shell:AppData'
	newSpecialFolder 'shell:CredentialManager'
	newSpecialFolder 'shell:CryptoKeys'
	newSpecialFolder 'shell:DpapiKeys'
	newSpecialFolder 'shell:SystemCertificates'

	newSpecialFolder 'shell:Quick Launch'
	# shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}
	newSpecialFolder 'shell:User Pinned'
	newSpecialFolder 'shell:ImplicitAppShortcuts'

	newSpecialFolder 'shell:AccountPictures'
	# shell:::{B155BDF8-02F0-451E-9A26-AE317CFD7779} ([delegate folder that appears in Computer])
	newSpecialFolder 'shell:NetHood'
	# shell:::{ED50FC29-B964-48A9-AFB3-15EBB9B97F36} ([printhood delegate folder])
	newSpecialFolder 'shell:PrintHood'
	newSpecialFolder 'shell:Recent'
	newSpecialFolder 'shell:SendTo'
	newSpecialFolder 'shell:Templates'

	$categoryName = 'Libraries'

	# shell:UsersLibrariesFolder
	# shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}
	$librariesPath = getKnownFolderPath Libraries
	newSpecialFolder 'shell:Libraries' -Path $librariesPath -FolderItemForProperties (getDirectoryFolderItem $librariesPath)
	# Win10 1507からサポート
	# shell:Libraries\{2B20DF75-1EDA-4039-8097-38798227D5B7}
	newSpecialFolder 'shell:CameraRollLibrary' -Path (getKnownFolderPath CameraRollLibrary)
	# shell:Libraries\{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}
	newSpecialFolder 'shell:DocumentsLibrary' -Path (getKnownFolderPath DocumentsLibrary)
	# shell:Libraries\{2112AB0A-C86A-4FFE-A368-0DE96E47012E}
	newSpecialFolder 'shell:MusicLibrary' -Path (getKnownFolderPath MusicLibrary)
	# shell:Libraries\{A990AE9F-A03B-4E80-94BC-9912D7504104}
	newSpecialFolder 'shell:PicturesLibrary' -Path (getKnownFolderPath PicturesLibrary)
	# Win10 1507からサポート
	# shell:Libraries\{E25B5812-BE88-4BD9-94B0-29233477B6C3}
	newSpecialFolder 'shell:SavedPicturesLibrary' -Path (getKnownFolderPath SavedPicturesLibrary)
	# shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}\{491E922F-5643-4AF4-A7EB-4E7A138D8174}
	newSpecialFolder 'shell:VideosLibrary' -Path (getKnownFolderPath VideosLibrary)

	$categoryName = 'StartMenu'

	newSpecialFolder 'shell:Start Menu'
	newSpecialFolder 'shell:Programs'
	newSpecialFolder 'shell:Administrative Tools'
	newSpecialFolder 'shell:Startup'

	$categoryName = 'LocalAppData'

	# %LOCALAPPDATA%
	newSpecialFolder 'shell:Local AppData'

	# Win10 1709からサポート
	newSpecialFolder 'shell:AppDataDesktop'
	# Win10 1507からサポート
	newSpecialFolder 'shell:Development Files'
	# Win10 1709からサポート
	newSpecialFolder 'shell:AppDataDocuments'
	# Win10 1709からサポート
	newSpecialFolder 'shell:AppDataFavorites'
	# ストアアプリの設定
	newSpecialFolder 'shell:Local AppData\Packages' 'Settings of the Windows Apps'
	# Win10 1709からサポート
	newSpecialFolder 'shell:AppDataProgramData'
	# %TEMP%
	# %TMP%
	newSpecialFolder ([Path]::GetTempPath()) 'Temporary Folder'
	newSpecialFolder 'shell:Local AppData\VirtualStore'

	newSpecialFolder 'shell:Application Shortcuts'
	newSpecialFolder 'shell:CD Burning'
	# Win10 1809からサポート
	# 標準ユーザー権限でフォントをインストールした時に自動生成される
	newSpecialFolder 'shell:Local AppData\Microsoft\Windows\Fonts' 'UserFonts'
	newSpecialFolder 'shell:GameTasks'
	newSpecialFolder 'shell:History'
	newSpecialFolder 'shell:Cache'
	newSpecialFolder 'shell:Cookies'
	newSpecialFolder 'shell:Ringtones'
	newSpecialFolder 'shell:Roamed Tile Images'
	newSpecialFolder 'shell:Roaming Tiles'
	newSpecialFolder 'shell:Local AppData\Microsoft\Windows\WinX'

	newSpecialFolder 'shell:SearchHistoryFolder'
	newSpecialFolder 'shell:SearchTemplatesFolder'

	newSpecialFolder 'shell:Local AppData\Microsoft\Windows Sidebar\Gadgets'
	# フォトギャラリーでファイルを編集する時に自動生成される
	newSpecialFolder 'shell:Original Images'

	newSpecialFolder 'shell:Local AppData\Microsoft\WindowsApps' 'AppExecutionAlias'

	newSpecialFolder 'shell:UserProgramFiles'
	newSpecialFolder 'shell:UserProgramFilesCommon'

	newSpecialFolder 'shell:LocalAppDataLow'

	$categoryName = 'Public'

	# shell:::{4336A54D-038B-4685-AB02-99BB52D3FB8B}
	# shell:ThisDeviceFolder (Win10 1507から1607まで)
	# shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90} (Win10 1507から)
	# %PUBLIC%
	newSpecialFolder 'shell:Public'
	newSpecialFolder 'shell:PublicAccountPictures'
	newSpecialFolder 'shell:Common Desktop'
	newSpecialFolder 'shell:Common Documents'
	newSpecialFolder 'shell:CommonDownloads'
	newSpecialFolder 'shell:PublicLibraries'
	newSpecialFolder (getKnownFolderPath RecordedTVLibrary) 'RecordedTVLibrary'
	newSpecialFolder 'shell:CommonMusic'
	newSpecialFolder 'shell:SampleMusic'
	newSpecialFolder 'shell:CommonPictures'
	newSpecialFolder 'shell:SamplePictures'
	newSpecialFolder 'shell:CommonVideo'
	newSpecialFolder 'shell:SampleVideos'

	$categoryName = 'ProgramData'

	# %ALLUSERSPROFILE%
	# %ProgramData%
	newSpecialFolder 'shell:Common AppData'
	newSpecialFolder 'shell:OEM Links'

	newSpecialFolder (getItemValue $appxKey 'PackageRepositoryRoot') 'State repository of the Windows Apps'
	newSpecialFolder 'shell:Device Metadata Store'
	newSpecialFolder 'shell:PublicGameTasks'
	# Win10 1507からサポート
	# 市販デモ モードで使用される
	newSpecialFolder 'shell:Retail Demo'
	newSpecialFolder 'shell:CommonRingtones'
	newSpecialFolder 'shell:Common Templates'

	$categoryName = 'CommonStartMenu'

	newSpecialFolder 'shell:Common Start Menu'
	newSpecialFolder 'shell:Common Programs'
	# shell:ControlPanelFolder\::{D20EA4E1-3957-11D2-A40B-0C5020524153}
	newSpecialFolder 'shell:Common Administrative Tools'
	newSpecialFolder 'shell:Common Startup'
	# Win10 1507からサポート
	newSpecialFolder 'shell:Common Start Menu Places'

	$categoryName = 'Windows'

	# %SystemRoot%
	# %windir%
	newSpecialFolder 'shell:Windows'
	# shell:::{1D2680C9-0E2A-469D-B787-065558BC7D43} ([Fusion Cache]) (.NET3.5まで)
	# CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
	# Win11では普通のフォルダーと同じ表示になるので非表示に
	if (![OS]::Win11) { newSpecialFolder 'shell:Windows\assembly' '.NET Framework Assemblies' }
	newSpecialFolder (getItemValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings' 'ActiveXCache') 'ActiveX Cache Folder'
	# shell:ControlPanelFolder\::{BD84B380-8CA2-1069-AB1D-08000948F534}
	newSpecialFolder 'shell:Fonts'
	newSpecialFolder 'shell:Windows\Offline Web Pages' 'Subscription Folder'

	newSpecialFolder 'shell:ResourceDir'
	newSpecialFolder 'shell:LocalizedResourcesDir'

	newSpecialFolder $(if (!$isWow64) { 'shell:System' } else { 'shell:SystemX86' } )
	# Win10 1803からサポート
	newSpecialFolder "$Env:DriverData" 'DriverData'
	newSpecialFolder (getItemValue 'HKLM:\SOFTWARE\Microsoft\Wbem' 'Installation Directory') 'WMI'
	if ($is64bitOS) {
		newSpecialFolder $(if (!$isWow64) { 'shell:SystemX86' } else { 'shell:Windows\SysNative' } )
	}

	$categoryName = 'UserProfiles'

	newSpecialFolder 'shell:UserProfiles'
	newSpecialFolder (getItemValue 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' 'Default') 'DefaultUserProfile'

	$categoryName = 'ProgramFiles'

	# shell:ProgramFilesX64 (64ビットアプリのみ)
	# %ProgramFiles%
	newSpecialFolder 'shell:ProgramFiles'
	# shell:ProgramFilesCommonX64 (64ビットアプリのみ)
	# %CommonProgramFiles%
	newSpecialFolder 'shell:ProgramFilesCommon'
	newSpecialFolder (getItemValue $appxKey 'PackageRoot') 'Windows Apps'
	newSpecialFolder 'shell:ProgramFiles\Windows Sidebar\Gadgets' 'Default Gadgets'
	newSpecialFolder 'shell:ProgramFiles\Windows Sidebar\Shared Gadgets'

	if ($is64bitOS) {
		if (!$isWow64) { newSpecialFolder 'shell:ProgramFilesX86' }
		else { newSpecialFolder (getItemValue $currentVersionKey 'ProgramW6432Dir') 'ProgramFilesX64' }
	}
	if ($is64bitOS) {
		if (!$isWow64) { newSpecialFolder 'shell:ProgramFilesCommonX86' }
		else { newSpecialFolder (getItemValue $currentVersionKey 'CommonW6432Dir') 'ProgramFilesCommonX64' }
	}

	$categoryName = 'Desktop / MyComputer'

	newSpecialFolder 'shell:Desktop'
	# shell:MyComputerFolderはWin10 1507/1511だとなぜかデスクトップになってしまう
	newSpecialFolder 'shell:MyComputerFolder'
	# Recent Places Folder
	newSpecialFolder 'shell:::{22877A6D-37A1-461A-91B0-DBDA5AAEBC99}'
	# Win10 1507からサポート
	# shell:::{4564B25E-30CD-4787-82BA-39E73A750B14} ([Recent Items Instance Folder])
	newSpecialFolder 'shell:::{3134EF9C-6B18-4996-AD04-ED5912E00EB5}' 'Recent files'
	# Portable Devices
	newSpecialFolder 'shell:::{35786D3C-B075-49B9-88DD-029876E11C01}'
	# Frequent Places Folder
	# Win10 1507からサポート
	newSpecialFolder 'shell:::{3936E9E4-D92C-4EEE-A85A-BC16D5EA0819}'
	newSpecialFolder 'shell:RecycleBinFolder'
	# (windows.storage.dll)
	# Win10 1507からサポート
	# Win11 22H2からOtherDirs カテゴリに移動するので非表示に
	if (![OS]::Win11_22h2) { newSpecialFolder 'shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}' 'Quick access' }
	# Removable Storage Devices
	# Win8.1では[PC]と同じなので非表示に
	if ([OS]::Win10) { newSpecialFolder 'shell:::{A6482830-08EB-41E2-84C1-73920C2BADB9}' }
	# Win10まで
	newSpecialFolder 'shell:HomeGroupFolder'
	newSpecialFolder 'shell:NetworkPlacesFolder'
	# Removable Drives
	# Win10 1507からサポート
	newSpecialFolder 'shell:::{F5FB2C77-0E2F-4A16-A381-3E560C68BC83}'
	# (windows.storage.dll)
	# Win11 22H2から
	if ([OS]::Win11_22h2) { newSpecialFolder 'shell:::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}' 'Home' }

	$categoryName = 'ControlPanel'

	# Control Panel
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\1' 'Appearance and Personalization'
	# shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\4
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\2' 'Hardware and Sound'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\3' 'Network and Internet'
	# shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\10
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\5' 'System and Security'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\6' $(if ([OS]::Win10_1803) { 'Clock and Region' } else { 'Clock, Language, and Region' })
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\7' 'Ease of Access'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\8' 'Programs'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\9' 'User Accounts'

	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}
	# shell:::{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}
	# shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\11
	newSpecialFolder 'shell:ControlPanelFolder' 'All Control Panel Items'

	# コントロールパネル内の項目はCLSIDだけを指定してもアクセス可能
	# 例えば[電源オプション]なら shell:::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}
	# ただしその場合はアドレスバーからコントロールパネルに移動できない

	# Power Options
	newSpecialFolder 'shell:ControlPanelFolder\::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}'
	# Credential Manager
	newSpecialFolder 'shell:ControlPanelFolder\::{1206F5F1-0569-412C-8FEC-3204630DFB70}'
	newSpecialFolder 'shell:AddNewProgramsFolder'
	# Set User Defaults
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{E44E5D18-0652-4508-A4E2-8A090067BCB0}
	# Win11 22H2でShellCommandsExceptFolders カテゴリに移動するので非表示に
	if (![OS]::Win11_22h2) { newSpecialFolder 'shell:ControlPanelFolder\::{17CD9488-1228-4B2F-88CE-4298E93E0966}' 'Default Programs' }
	# Workspaces Center
	newSpecialFolder 'shell:ControlPanelFolder\::{241D7C96-F8BF-4F85-B01F-E2B043341A4B}' 'RemoteApp and Desktop Connections'
	# Windows Update
	# Win8.1までサポート
	newSpecialFolder 'shell:ControlPanelFolder\::{36EEF7DB-88AD-4E81-AD49-0E313F0C35F8}'
	# Windows Firewall (Win10 1703まで)
	# Windows Defender Firewall (Win10 1709から)
	newSpecialFolder 'shell:ControlPanelFolder\::{4026492F-2F69-46B8-B9BF-5654FC07E423}'
	# Speech Recognition
	newSpecialFolder 'shell:ControlPanelFolder\::{58E3C745-D971-4081-9034-86E34B30836A}'
	# User Accounts
	newSpecialFolder 'shell:ControlPanelFolder\::{60632754-C523-4B62-B45C-4172DA012619}'
	# HomeGroup Control Panel
	# Win10 1709までサポート
	newSpecialFolder 'shell:ControlPanelFolder\::{67CA7650-96E6-4FDD-BB43-A8E774F73A57}'
	# Network and Sharing Center
	newSpecialFolder 'shell:ControlPanelFolder\::{8E908FC9-BECC-40F6-915B-F4CA0E70D03D}'
	# Parental Controls
	# Win8.1まで
	newSpecialFolder 'shell:ControlPanelFolder\::{96AE8D84-A250-4520-95A5-A47A7E3C548B}'
	# AutoPlay
	newSpecialFolder 'shell:ControlPanelFolder\::{9C60DE1E-E5FC-40F4-A487-460851A8D915}'
	# System Recovery
	newSpecialFolder 'shell:ControlPanelFolder\::{9FE63AFD-59CF-4419-9775-ABCC3849F861}'
	# Device Center
	newSpecialFolder 'shell:ControlPanelFolder\::{A8A91A66-3A7D-4424-8D24-04E180695C7A}' 'Devices and Printers'
	# Windows 7 File Recovery
	# Win10 1507からサポート
	newSpecialFolder 'shell:ControlPanelFolder\::{B98A2BEA-7D42-4558-8BD1-832F41BAC6FD}'
	# System
	# Win10 20H2から下部に移動するので非表示に
	if (![OS]::Win10_20h2) { newSpecialFolder 'shell:ControlPanelFolder\::{BB06C0E4-D293-4F75-8A90-CB05B6477EEE}' }
	# Action Center CPL (Win8.1まで)
	# Security and Maintenance CPL (Win10 1507から)
	newSpecialFolder 'shell:ControlPanelFolder\::{BB64F8A7-BEE7-4E1A-AB8D-7D8273F7FDB6}'
	# Microsoft Windows Font Folder
	# shell:Fonts
	newSpecialFolder 'shell:ControlPanelFolder\::{BD84B380-8CA2-1069-AB1D-08000948F534}' -Path 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{BD84B380-8CA2-1069-AB1D-08000948F534}'
	# Language Settings
	# Win10 1803までサポート
	newSpecialFolder 'shell:ControlPanelFolder\::{BF782CC9-5A52-4A17-806C-2A894FFEEAC5}'
	# Display
	# Win10 1607までサポート
	newSpecialFolder 'shell:ControlPanelFolder\::{C555438B-3C23-4769-A71F-B6D3D9B6053A}'
	# Troubleshooting
	newSpecialFolder 'shell:ControlPanelFolder\::{C58C4893-3BE0-4B45-ABB5-A63E4B8C8651}'
	# Administrative Tools (Win10まで)
	# Windows Tools (Win11 21H2から)
	# shell:Common Administrative Tools
	newSpecialFolder 'shell:ControlPanelFolder\::{D20EA4E1-3957-11D2-A40B-0C5020524153}' -Path 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{D20EA4E1-3957-11D2-A40B-0C5020524153}'
	# Ease of Access
	newSpecialFolder 'shell:ControlPanelFolder\::{D555645E-D4F8-4C29-A827-D93C859C4F2A}'
	# Secure Startup
	newSpecialFolder 'shell:ControlPanelFolder\::{D9EF8727-CAC2-4E60-809E-86F80A666C91}' 'BitLocker Drive Encryption'
	# Sensors
	# Win8.1までサポート
	newSpecialFolder 'shell:ControlPanelFolder\::{E9950154-C418-419E-A90A-20C5287AE24B}' 'Location Settings'
	# ECS
	newSpecialFolder 'shell:ControlPanelFolder\::{ECDB0924-4208-451E-8EE0-373C0956DE16}' 'Work Folders'
	# Personalization Control Panel
	newSpecialFolder 'shell:ControlPanelFolder\::{ED834ED6-4B5A-4BFE-8F11-A626DCB6A921}'
	# History Vault
	newSpecialFolder 'shell:ControlPanelFolder\::{F6B6E965-E9B2-444B-9286-10C9152EDBC5}'
	# Storage Spaces
	newSpecialFolder 'shell:ControlPanelFolder\::{F942C606-0914-47AB-BE56-1321B8035096}'

	newSpecialFolder 'shell:ChangeRemoveProgramsFolder'
	# Win11 22H2でShellCommandsExceptFolders カテゴリに移動するので非表示に
	if (![OS]::Win11_22h2) { newSpecialFolder 'shell:AppUpdatesFolder' }

	newSpecialFolder 'shell:SyncCenterFolder'
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{2E9E59C0-B437-4981-A647-9C34B9B90891} ([Sync Setup Folder])
	newSpecialFolder 'shell:SyncSetupFolder'
	newSpecialFolder 'shell:ConflictFolder'
	newSpecialFolder 'shell:SyncResultsFolder'

	# Taskbar
	newSpecialFolder "shell:$(if ([OS]::Win10) { '::{21EC2020-3AEA-1069-A2DD-08002B30309D}' } else { 'ControlPanelFolder' })\::{05D7B0F4-2121-4EFF-BF6B-ED3F69B894D9}" 'Notification Area Icons'
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{863AA9FD-42DF-457B-8E4D-0DE1B8015C60}
	newSpecialFolder 'shell:PrintersFolder'
	# Bluetooth Devices
	newSpecialFolder 'shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{28803F59-3A75-4058-995F-4EE5503B023C}'
	# shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{992CFFA0-F557-101A-88EC-00DD010CCC48}
	newSpecialFolder 'shell:ConnectionsFolder'
	# Font Settings
	newSpecialFolder 'shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{93412589-74D4-4E4E-AD0E-E0CB621440FD}'
	# System
	# Win10 20H2からここに移動
	# Win11 21H2以前まで
	if ([OS]::Win10_20h2 -and ![OS]::Win11) { newSpecialFolder 'shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{BB06C0E4-D293-4F75-8A90-CB05B6477EEE}' }
	# All Tasks
	newSpecialFolder 'shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{ED7BA470-8E54-465E-825C-99712043E01C}'

	$categoryName = 'OtherFolders'

	# Hyper-V Remote File Browsing
	# クライアントHyper-Vを有効にすると利用可
	newSpecialFolder 'shell:::{0907616E-F5E6-48D8-9D61-A91C3D28106D}'
	# Cabinet Shell Folder
	newSpecialFolder 'shell:::{0CD7A5C0-9F37-11CE-AE65-08002B2E1262}'
	# Fusion Cache (.NET Framework Assemblies)
	# .NET3.5まで
	# CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
	# Win11からこのカテゴリに移動
	if ([OS]::Win11) { newSpecialFolder 'shell:::{1D2680C9-0E2A-469D-B787-065558BC7D43}' }
	# Network
	newSpecialFolder 'shell:::{208D2C60-3AEA-1069-A2D7-08002B30309D}'
	# DLNA Media Servers Data Source
	newSpecialFolder 'shell:::{289AF617-1CC3-42A6-926C-E6A863F0E3BA}'
	# Results Folder
	newSpecialFolder 'shell:::{2965E715-EB66-4719-B53F-1672673BBEFA}'
	newSpecialFolder 'shell:AppsFolder'
	# MS Graph Recent File Folder
	# Win11 21H2からサポート
	newSpecialFolder 'shell:::{42254EE9-E625-4065-8F70-775090256F72}'
	# Command Folder
	newSpecialFolder 'shell:::{437FF9C0-A07F-4FA0-AF80-84B6C6440A16}'
	# Other Users Folder
	newSpecialFolder 'shell:::{6785BFAC-9D2D-4BE5-B7E2-59937E8FB80A}'
	# search:
	# search-ms:
	newSpecialFolder 'shell:SearchHomeFolder'
	# Win10 1511までサポート
	newSpecialFolder 'shell:StartMenuAllPrograms'
	# (cscui.dll)
	# 企業向けエディションで使用可
	newSpecialFolder 'shell:::{AFDB1F70-2A4C-11D2-9039-00C04F8EEB3E}' 'Offline Files Folder'
	# AppSuggestedLocations
	newSpecialFolder 'shell:::{C57A6066-66A3-4D91-9EB9-41532179F0A5}'
	# Win10 1709までサポート
	newSpecialFolder 'shell:Games'
	# Previous Versions Results Folder
	newSpecialFolder 'shell:::{F8C2AB3B-17BC-41DA-9758-339D7DBF2D88}'

	if (!$IncludeShellCommand) { return }

	# フォルダー以外のshellコマンド
	$categoryName = 'ShellCommandsExceptFolders'

	# System
	# Win10 20H2から
	# Win11 21H2から下部に移動するので非表示に
	if ([OS]::Win10_20h2 -and ![OS]::Win11) { newSpecialFolder 'shell:ControlPanelFolder\::{BB06C0E4-D293-4F75-8A90-CB05B6477EEE}' }

	# Win11 22H2からここに移動
	if ([OS]::Win11_22h2) { newSpecialFolder 'shell:AppUpdatesFolder' 'Uninstalled updates' }

	# Taskbar
	newShellCommand '{0DF44EAA-FF21-4412-828E-260A8728E7F1}'
	# Set User Defaults
	# Win11 22H2からここに移動
	if ([OS]::Win11_22h2) { newSpecialFolder 'shell:::{17CD9488-1228-4B2F-88CE-4298E93E0966}' 'Default apps' }
	# Search
	# Win10 1511まで
	if (![OS]::Win10_1607) { newShellCommand '{2559A1F0-21D7-11D4-BDAF-00C04F60B9F0}' 'Search Files' }
	# Help and Support
	# Win8.1まで
	newShellCommand '{2559A1F1-21D7-11D4-BDAF-00C04F60B9F0}'
	# Run...
	newShellCommand '{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}'
	# Set Program Access and Defaults
	newShellCommand '{2559A1F7-21D7-11D4-BDAF-00C04F60B9F0}'
	# (shell32.dll#SearchCommand)
	newShellCommand '{2559A1F8-21D7-11D4-BDAF-00C04F60B9F0}' 'Search'
	# Show Desktop
	# Win+Dと同じ
	newShellCommand '{3080F90D-D7AD-11D9-BD98-0000947B0257}'
	# Window Switcher
	# Win8.1ではCtrl+Alt+Tab、Win10 1607以降ではWin+Tabと同じ (Win10 1507/1511では使用不可)
	newShellCommand '{3080F90E-D7AD-11D9-BD98-0000947B0257}'
	# Win8.1まで
	if (![OS]::Win10) { newShellCommand '{38A98528-6CBF-4CA9-8DC0-B1E1D10F7B1B}' 'Connect To' }
	# Phone and Modem Control Panel
	newShellCommand '{40419485-C444-4567-851A-2DD7BFA1684D}'
	# Open in new window (shell32.dll)
	newShellCommand '{52205FD8-5DFB-447D-801A-D0B52F2E83E1}' 'File Explorer'
	# Mobility Center Control Panel
	newShellCommand '{5EA4F148-308C-46D7-98A9-49041B1DD468}'
	# Region and Language
	newShellCommand '{62D8ED13-C9D0-4CE8-A914-47DD628FB1B0}'
	# Windows Features
	newShellCommand '{67718415-C450-4F3C-BF8A-B487642DC39B}'
	# Mouse Control Panel
	newShellCommand '{6C8EEC18-8D75-41B2-A177-8831D59D2D50}'
	# Folder Options
	newShellCommand '{6DFD7C5C-2451-11D3-A299-00C04F8EF6AF}'
	# Keyboard Control Panel
	newShellCommand '{725BE8F7-668E-4C7B-8F90-46BDB0936430}'
	# Device Manager
	newShellCommand '{74246BFC-4C96-11D0-ABEF-0020AF6B0B7A}'
	# User Accounts
	# netplwiz.exe / control.exe userpasswords2
	newShellCommand '{7A9D77BD-5403-11D2-8785-2E0420524153}'
	# Tablet PC Settings Control Panel
	newShellCommand '{80F3F1D5-FECA-45F3-BC32-752C152E456E}'
	# Internet Folder
	# Win10以降では開けないので非表示に
	if (![OS]::Win10) { newSpecialFolder 'shell:InternetFolder' }
	# Indexing Options Control Panel
	newShellCommand '{87D66A43-7B11-4A28-9811-C86EE395ACF7}'
	# Portable Workspace Creator
	# Win10 1909まで
	# Enterpriseで使用可
	# Win10 1607以降ではProでも使用可
	newShellCommand '{8E0C279D-0BD1-43C3-9EBD-31C3DC5B8A77}' 'Windows To Go'
	# Infrared
	# Win10 1607から1809まで
	newShellCommand '{A0275511-0E86-4ECA-97C2-ECD8F1221D08}'
	# Internet Options
	newShellCommand '{A3DD4F92-658A-410F-84FD-6FBBBEF2FFFE}'
	# Color Management
	newShellCommand '{B2C761C6-29BC-4F19-9251-E6195265BAF1}'
	# Win11 21H2から
	if ([OS]::Win11) { newShellCommand '{B4FB3F98-C1EA-428D-A78A-D1F5659CBA93}' 'Media streaming options' }
	# System
	# Win11 21H2からここに移動
	if ([OS]::Win11) { newSpecialFolder 'shell:::{BB06C0E4-D293-4F75-8A90-CB05B6477EEE}' }
	# Windows Anytime Upgrade
	# Win8.1まで
	newShellCommand '{BE122A0E-4503-11DA-8BDE-F66BAD1E3F3A}'
	# Text to Speech Control Panel
	newShellCommand '{D17D1D6D-CC3F-4815-8FE3-607E7D5D10B3}'
	# Add Network Place
	newShellCommand '{D4480A50-BA28-11D1-8E75-00C04FA31A86}'
	# Windows Defender
	# Win10 1607まで
	newShellCommand '{D8559EB9-20C0-410E-BEDA-7ED416AECC2A}'
	# Date and Time Control Panel
	newShellCommand '{E2E7934B-DCE5-43C4-9576-7FE4F75E7480}'
	# Sound Control Panel
	newShellCommand '{F2DDFC82-8F12-4CDD-B7DC-D4FE1425AA4D}'
	# Pen and Touch Control Panel
	newShellCommand '{F82DF8F7-8B9F-442E-A48C-818EA735FF9B}'

	if (!$Develop) { return }

	# 通常とは違う名前がエクスプローラーのタイトルバーに表示されるフォルダー
	$categoryName = 'OtherNames'

	# Fusion Cache (.NET Framework Assemblies)
	# .NET3.5まで
	# CLSIDを使ってアクセスするとエクスプローラーがクラッシュする
	# Win11からOtherFoldersカテゴリに移動するので非表示に
	if (![OS]::Win11) { newSpecialFolder 'shell:::{1D2680C9-0E2A-469D-B787-065558BC7D43}' }
	# Favorites (Links)
	newSpecialFolder 'shell:::{323CA680-C24D-4099-B94D-446DD2D7249E}'
	# My Documents (Documents)
	newSpecialFolder 'shell:::{450D8FBA-AD25-11D0-98A8-0800361B1103}' 'My Documents'
	# Recent Items Instance Folder (Recent files)
	# Win10 1507から
	# Win11 22H2ではすべてのファイルが表示されない
	newSpecialFolder 'shell:::{4564B25E-30CD-4787-82BA-39E73A750B14}'
	# delegate folder that appears in Computer (NetHood)
	newSpecialFolder 'shell:::{B155BDF8-02F0-451E-9A26-AE317CFD7779}'
	# Common Places FS Folder (Links)
	newSpecialFolder 'shell:::{D34A6CA6-62C2-4C34-8A7C-14709C1AD938}'
	# printhood delegate folder (PrintHood)
	newSpecialFolder 'shell:::{ED50FC29-B964-48A9-AFB3-15EBB9B97F36}'
	# Public (Win10 1607まで)
	# UsersFilesFolder (Win10 1703から)
	# shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90} (Win10 1507から1607まで)
	# shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9} (Win10 1703から)
	newSpecialFolder 'shell:ThisDeviceFolder'
	# Sync Setup Folder (SyncSetupFolder)
	newSpecialFolder 'shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}\::{2E9E59C0-B437-4981-A647-9C34B9B90891}'

	# エクスプローラーで開けないフォルダー
	$categoryName = 'CantOpen'

	# CLSID_SearchFolder
	newSpecialFolder 'shell:::{04731B67-D933-450A-90E6-4ACD2E9408FE}'
	# Manage Wireless Networks
	newSpecialFolder 'shell:::{1FA9085F-25A2-489B-85D4-86326EEDCD87}'
	# Sync Center Conflict Folder
	newSpecialFolder 'shell:::{289978AC-A101-4341-A817-21EBA7FD046D}'
	# FileHistoryDataSource
	# ファイル履歴を有効にすると利用可
	newSpecialFolder 'shell:::{2F6CE85C-F9EE-43CA-90C7-8A9BD53A2467}'
	# LayoutFolder
	newSpecialFolder 'shell:::{328B0346-7EAF-4BBE-A479-7CB88A095F5B}'
	# Explorer Browser Results Folder
	newSpecialFolder 'shell:::{418C8B64-5463-461D-88E0-75E2AFA3C6FA}'
	# PC Settings
	newSpecialFolder 'shell:::{5ED4F38C-D3FF-4D61-B506-6820320AEBFE}'
	# Microsoft FTP Folder
	newSpecialFolder 'shell:::{63DA6EC0-2E98-11CF-8D82-444553540000}'
	# CLSID_AppInstanceFolder
	newSpecialFolder 'shell:::{64693913-1C21-4F30-A98F-4E52906D3B56}'
	# Sync Results Folder
	newSpecialFolder 'shell:::{71D99464-3B6B-475C-B241-E15883207529}'
	# Programs Folder
	# Win10 1511まで
	newSpecialFolder 'shell:::{7BE9D83C-A729-4D97-B5A7-1B7313C39E0A}'
	# Programs Folder and Fast Items
	# Win10 1511まで
	newSpecialFolder 'shell:::{865E5E76-AD83-4DCA-A109-50DC2113CE9A}'
	# Win10でこのカテゴリに移動
	if ([OS]::Win10) { newSpecialFolder 'shell:InternetFolder' }
	# File Backup Index
	newSpecialFolder 'shell:::{877CA5AC-CB41-4842-9C69-9136E42D47E2}'
	# (mssvp.dll)
	newSpecialFolder 'shell:::{89D83576-6BD1-4C86-9454-BEB04E94C819}' 'Microsoft Office Outlook'
	# DXP
	newSpecialFolder 'shell:::{8FD8B88D-30E1-4F25-AC2B-553D3D65F0EA}'
	# Enhanced Storage Data Source
	newSpecialFolder 'shell:::{9113A02D-00A3-46B9-BC5F-9C04DADDD5D7}'
	# CLSID_StartMenuLauncherProviderFolder
	newSpecialFolder 'shell:::{98F275B4-4FFF-11E0-89E2-7B86DFD72085}'
	# IE RSS Feeds Folder
	newSpecialFolder 'shell:::{9A096BB5-9DC3-4D1C-8526-C3CBF991EA4E}'
	# CLSID_StartMenuCommandingProviderFolder
	newSpecialFolder 'shell:::{A00EE528-EBD9-48B8-944A-8942113D46AC}'
	# Previous Versions Results Delegate Folder
	newSpecialFolder 'shell:::{A3C3D402-E56C-4033-95F7-4885E80B0111}'
	# Library Folder
	newSpecialFolder 'shell:::{A5A3563A-5755-4A6F-854E-AFA3230B199F}'
	# Win10までサポート
	newSpecialFolder 'shell:HomeGroupCurrentUserFolder'
	# Sync Results Delegate Folder
	newSpecialFolder 'shell:::{BC48B32F-5910-47F5-8570-5074A8A5636A}'
	# (mssvp.dll)
	newSpecialFolder 'shell:::{BD7A2E7B-21CB-41B2-A086-B309680C6B7E}' 'Offline Files'
	# DLNA Content Directory Data Source
	newSpecialFolder 'shell:::{D2035EDF-75CB-4EF1-95A7-410D9EE17170}'
	# CLSID_StartMenuProviderFolder
	newSpecialFolder 'shell:::{DAF95313-E44D-46AF-BE1B-CBACEA2C3065}'
	# CLSID_StartMenuPathCompleteProviderFolder
	newSpecialFolder 'shell:::{E345F35F-9397-435C-8F95-4E922C26259E}'
	# Sync Center Conflict Delegate Folder
	newSpecialFolder 'shell:::{E413D040-6788-4C22-957E-175D1C513A34}'
	# Shell DocObject Viewer
	newSpecialFolder 'shell:::{E7E4BC40-E76A-11CE-A9BB-00AA004AE837}'
	# StreamBackedFolder
	newSpecialFolder 'shell:::{EDC978D6-4D53-4B2F-A265-5805674BE568}'
	# Sync Setup Delegate Folder
	newSpecialFolder 'shell:::{F1390A9A-A3F4-4E5D-9C5F-98F3BD8D935C}'
	newSpecialFolder 'shell:CSCFolder'

	# 上にあるのとは違うデータでフォルダーの情報を取得する
	# CSIDLは扱わない
	$categoryName = 'OtherDirs'

	newSpecialFolder 'shell:Profile'
	newSpecialFolder 'shell:Local Documents'
	newSpecialFolder 'shell:Local Downloads'
	newSpecialFolder 'shell:Local Music'
	newSpecialFolder 'shell:Local Pictures'
	newSpecialFolder 'shell:Local Videos'
	newSpecialFolder 'shell:UsersFilesFolder\{56784854-C6CB-462B-8169-88E350ACB882}' 'Contacts'
	newSpecialFolder 'shell:UsersFilesFolder\{BFB9D5E0-C6A9-404C-B2B2-AE6DB6AF4968}' 'Links'
	newSpecialFolder 'shell:UsersFilesFolder\{4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4}' 'SavedGames'
	newSpecialFolder 'shell:UsersFilesFolder\{7D1D3A04-DEBB-4115-95CF-2F29DA2920DA}' 'Searches'
	newSpecialFolder 'shell:UsersLibrariesFolder'
	newSpecialFolder 'shell:Libraries\{2B20DF75-1EDA-4039-8097-38798227D5B7}' 'CameraRollLibrary'
	newSpecialFolder 'shell:Libraries\{7B0DB17D-9CD2-4A93-9733-46CC89022E7C}' 'DocumentsLibrary'
	newSpecialFolder 'shell:Libraries\{2112AB0A-C86A-4FFE-A368-0DE96E47012E}' 'MusicLibrary'
	newSpecialFolder 'shell:Libraries\{A990AE9F-A03B-4E80-94BC-9912D7504104}' 'PicturesLibrary'
	newSpecialFolder 'shell:Libraries\{E25B5812-BE88-4BD9-94B0-29233477B6C3}' 'SavedPicturesLibrary'
	newSpecialFolder 'shell:Libraries\{491E922F-5643-4AF4-A7EB-4E7A138D8174}' 'VideosLibrary'
	# 64ビットアプリのみ
	newSpecialFolder 'shell:ProgramFilesX64'
	# 64ビットアプリのみ
	newSpecialFolder 'shell:ProgramFilesCommonX64'

	newSpecialFolder "$Env:USERPROFILE" '%USERPROFILE%'
	newSpecialFolder "${Env:HOMEDRIVE}${Env:HOMEPATH}" '%HOMEDRIVE%%HOMEPATH%'
	newSpecialFolder "$Env:OneDrive" '%OneDrive%'
	newSpecialFolder "$Env:APPDATA" '%APPDATA%'
	newSpecialFolder "$Env:LOCALAPPDATA" '%LOCALAPPDATA%'
	newSpecialFolder "$Env:PUBLIC" '%PUBLIC%'
	newSpecialFolder "$Env:ALLUSERSPROFILE" '%ALLUSERSPROFILE%'
	newSpecialFolder "$Env:ProgramData" '%ProgramData%'
	newSpecialFolder "$Env:SystemRoot" '%SystemRoot%'
	newSpecialFolder "$Env:windir" '%windir%'
	newSpecialFolder "$Env:ProgramFiles" '%ProgramFiles%'
	newSpecialFolder "$Env:CommonProgramFiles" '%CommonProgramFiles%'

	# OneDrive
	# Win10 1507から
	newSpecialFolder 'shell:::{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
	# UsersLibraries
	newSpecialFolder 'shell:::{031E4825-7B94-4DC3-B131-E946B44C8DD5}'
	# Win10 1507から
	newSpecialFolder 'shell:::{088E3905-0323-4B02-9826-5D99428E115F}' 'Local Downloads'
	# Win10 1709から
	newSpecialFolder 'shell:::{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}' '3D Object'
	# Install New Programs
	newSpecialFolder 'shell:::{15EAE92E-F17A-4431-9F28-805E482DAFD4}'
	newSpecialFolder 'shell:::{1CF1260C-4DD0-4EBB-811F-33C572699FDE}' 'My Music'
	# User Pinned
	newSpecialFolder 'shell:::{1F3427C8-5C10-4210-AA03-2EE45287D668}'
	# This PC
	newSpecialFolder 'shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}'
	# All Control Panel Items
	newSpecialFolder 'shell:::{21EC2020-3AEA-1069-A2DD-08002B30309D}'
	# Printers
	newSpecialFolder 'shell:::{2227A280-3AEA-1069-A2DE-08002B30309D}'
	# Win10 1507から
	newSpecialFolder 'shell:::{24AD3AD4-A569-4530-98E1-AB02F9417AA8}' 'Local Pictures'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0' 'All Control Panel Items'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\4' 'Hardware and Sound'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\10' 'System and Security'
	newSpecialFolder 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\11' 'All Control Panel Items'
	newSpecialFolder 'shell:::{374DE290-123F-4565-9164-39C4925E467B}' 'Downloads'
	newSpecialFolder 'shell:::{3ADD1653-EB32-4CB0-BBD7-DFA0ABB5ACCA}' 'My Pictures'
	# Win10 1507から
	newSpecialFolder 'shell:::{3DFDF296-DBEC-4FB4-81D1-6A3438BCF4DE}' 'Local Music'
	# Applications
	newSpecialFolder 'shell:::{4234D49B-0245-4DF3-B780-3893943456E1}'
	# Public Folder
	newSpecialFolder 'shell:::{4336A54D-038B-4685-AB02-99BB52D3FB8B}'
	# UsersFiles
	newSpecialFolder 'shell:::{59031A47-3F72-44A7-89C5-5595FE6B30EE}'
	# This Device
	newSpecialFolder 'shell:::{5B934B42-522B-4C34-BBFE-37A3EF7B9C90}'
	# Recycle Bin
	newSpecialFolder 'shell:::{645FF040-5081-101B-9F08-00AA002F954E}'
	# (windows.storage.dll)
	# Win11 22H2では[Pinned files(ピン留めされたファイル)]カテゴリがない旧タイプ
	if ([OS]::Win11_22h2) { newSpecialFolder 'shell:::{679F85CB-0220-4080-B29B-5540CC05AAB6}' 'Home' }
	# Programs and Features
	newSpecialFolder 'shell:::{7B81BE6A-CE2B-4676-A29E-EB907A5126C5}'
	# Network Connections
	newSpecialFolder 'shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}'
	# Remote Printers
	newSpecialFolder 'shell:::{863AA9FD-42DF-457B-8E4D-0DE1B8015C60}'
	# Internet Folder
	newSpecialFolder 'shell:::{871C5380-42A0-1069-A2EA-08002B30309D}'
	# OneDrive
	# Win8.1のみ
	newSpecialFolder 'shell:::{8E74D236-7F35-4720-B138-1FED0B85EA75}'
	# CLSID_SearchHome
	newSpecialFolder 'shell:::{9343812E-1C37-4A49-A12E-4B2D810D956B}'
	# Sync Center Folder
	newSpecialFolder 'shell:::{9C73F5E5-7AE7-4E32-A8E8-8D23B85255BF}'
	# Network Connections
	newSpecialFolder 'shell:::{992CFFA0-F557-101A-88EC-00DD010CCC48}'
	newSpecialFolder 'shell:::{A0953C92-50DC-43BF-BE83-3742FED03C9C}' 'My Video'
	newSpecialFolder 'shell:::{A8CDFF1C-4878-43BE-B5FD-F8091C1C60D0}' 'Personal'
	# Win10 1511まで
	newSpecialFolder 'shell:::{ADFA80E7-9769-4AD9-992C-55DC57E1008C}' 'StartMenuAllPrograms'
	newSpecialFolder 'shell:::{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}' 'ThisPCDesktopFolder'
	# Other Users Folder
	# Win10まで
	newSpecialFolder 'shell:::{B4FB3F98-C1EA-428D-A78A-D1F5659CBA93}'
	# Microsoft Windows Font Folder
	newSpecialFolder 'shell:ControlPanelFolder\::{BD84B380-8CA2-1069-AB1D-08000948F534}'
	# Administrative Tools (Win10まで)
	# Windows Tools (Win11 21H2から)
	newSpecialFolder 'shell:::{D20EA4E1-3957-11D2-A40B-0C5020524153}'
	# Win10 1507から
	newSpecialFolder 'shell:::{D3162B92-9365-467A-956B-92703ACA08AF}' 'Local Documents'
	# Installed Updates
	newSpecialFolder 'shell:::{D450A8A1-9568-45C7-9C0E-B4F9FB4537BD}'
	# Games Explorer
	# Win10 1709までサポート
	newSpecialFolder 'shell:::{ED228FDF-9EA8-4870-83B1-96B02CFE0D52}'
	# Computers and Devices
	newSpecialFolder 'shell:::{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}'
	# This Device
	# Win10 1703から
	newSpecialFolder 'shell:::{F8278C54-A712-415B-B593-B77A2BE0DDA9}'
	# Win10 1507から
	newSpecialFolder 'shell:::{F86FA3AB-70D2-4FC7-9C99-FCBF05467F3A}' 'Local Videos'

	# Control Panel command object for Start menu and desktop
	newShellCommand '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}'
	# Default Programs command object for Start menu
	newShellCommand '{E44E5D18-0652-4508-A4E2-8A090067BCB0}'

	# フォルダーとして使えないshellコマンド
	$categoryName = 'Unusable'

	newSpecialFolder 'shell:MAPIFolder'
	newSpecialFolder 'shell:RecordedTVLibrary'

	newShellCommand '{00020D75-0000-0000-C000-000000000046}'
	# Desktop
	newShellCommand '{00021400-0000-0000-C000-000000000046}'
	# Shortcut
	newShellCommand '{00021401-0000-0000-C000-000000000046}'
	# Win10 1507から
	newShellCommand '{047EA9A0-93BB-415F-A1C3-D7AEB3DD5087}'
	# Open With Context Menu Handler
	newShellCommand '{09799AFB-AD67-11D1-ABCD-00C04FC30936}'
	# Folder Shortcut
	newShellCommand '{0AFACED1-E828-11D1-9187-B532F1E9575D}'
	# (windows.storage.dll)
	newShellCommand '{0C39A5CF-1A7A-40C8-BA74-8900E6DF5FCD}'
	# (dsuiext.dll)
	newShellCommand '{0D45D530-764B-11D0-A1CA-00AA00C16E65}'
	# Shell File System Folder
	newShellCommand '{0E5AAE11-A475-4C5B-AB00-C66DE400274E}'
	# Device Center Print Context Menu Extension
	newShellCommand '{0E6DAA63-DD4E-47CE-BF9D-FDB72ECE4A0D}'
	# IE History and Feeds Shell Data Source for Windows Search
	newShellCommand '{11016101-E366-4D22-BC06-4ADA335C892B}'
	# OpenMediaSharing
	newShellCommand '{17FC1A80-140E-4290-A64F-4A29A951A867}'
	# CLSID_DBFolderBoth
	newShellCommand '{1BEF2128-2F96-4500-BA7C-098DC0049CB2}'
	# CompatContextMenu Class
	newShellCommand '{1D27F844-3A1F-4410-85AC-14651078412D}'
	# Windows Security
	newShellCommand '{2559A1F2-21D7-11D4-BDAF-00C04F60B9F0}'
	# E-mail
	newShellCommand '{2559A1F5-21D7-11D4-BDAF-00C04F60B9F0}'
	# Location Folder
	newShellCommand '{267CF8A9-F4E3-41E6-95B1-AF881BE130FF}'
	# Enhanced Storage Context Menu Handler Class
	newShellCommand '{2854F705-3548-414C-A113-93E27C808C85}'
	# System Restore
	newShellCommand '{3F6BC534-DFA1-4AB4-AE54-EF25A74E0107}'
	# Start Menu Folder
	newShellCommand '{48E7CAAB-B918-4E58-A94D-505519C795DC}'
	# IGD Property Page
	newShellCommand '{4A1E5ACD-A108-4100-9E26-D2FAFA1BA486}'
	# LzhCompressedFolder2
	# Win10 1607まで
	newShellCommand '{4F289A46-2BBB-4AE8-9EDA-E5E034707A71}'
	# Shell File System NetFolder
	# Win11 21H2から
	newShellCommand '{4FE04BFD-85B9-49DD-B914-F4C9556B9DA6}'
	# This PC
	# Win10 1507から
	newShellCommand '{5E5F29CE-E0A8-49D3-AF32-7A7BDC173478}'
	# (dsuiext.dll)
	newShellCommand '{62AE1F9A-126A-11D0-A14B-0800361B1103}'
	# Search Connector Folder
	newShellCommand '{72B36E70-8700-42D6-A7F7-C9AB3323EE51}'
	# CryptPKO Class
	newShellCommand '{7444C717-39BF-11D1-8CD9-00C04FC29D45}'
	# Temporary Internet Files
	newShellCommand '{7BD29E00-76C1-11CF-9DD0-00A0C9034933}'
	# Temporary Internet Files
	newShellCommand '{7BD29E01-76C1-11CF-9DD0-00A0C9034933}'
	# Briefcase (Win10 1607まで)
	#  (Win10 1703から)
	newShellCommand '{85BBD920-42A0-1069-A2E4-08002B30309D}'
	# Shortcut
	newShellCommand '{85CFCCAF-2D14-42B6-80B6-F40F65D016E7}'
	# Mobile Broadband Profile Settings Editor
	# Win11 21H2まで
	newShellCommand '{87630419-6216-4FF8-A1F0-143562D16D5C}'
	# Compressed (zipped) Folder SendTo Target
	newShellCommand '{888DCA60-FC0A-11CF-8F0F-00C04FD7D062}'
	# ActiveX Cache Folder
	newShellCommand '{88C6C381-2E85-11D0-94DE-444553540000}'
	# Libraries delegate folder that appears in Users Files Folder
	newShellCommand '{896664F7-12E1-490F-8782-C0835AFD98FC}'
	# Windows Search Service Media Center Namespace Extension Handler
	# Win10 1607まで
	newShellCommand '{98D99750-0B8A-4C59-9151-589053683D73}'
	# MAPI Shell Context Menu
	newShellCommand '{9D3C0751-A13F-46A6-B833-B46A43C30FE8}'
	# Previous Versions
	newShellCommand '{9DB7A13C-F208-4981-8353-73CC61AE2783}'
	# Mail Service
	newShellCommand '{9E56BE60-C50F-11CF-9A2C-00A0C90A90CE}'
	# Desktop Shortcut
	newShellCommand '{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}'
	# DevicePairingFolder Initialization
	newShellCommand '{AEE2420F-D50E-405C-8784-363C582BF45A}'
	# CLSID_DBFolder
	newShellCommand '{B2952B16-0E07-4E5A-B993-58C52CB94CAE}'
	# Device Center Scan Context Menu Extension
	newShellCommand '{B5A60A9E-A4C7-4A93-AC6E-0B76D1D87DC4}'
	# DeviceCenter Initialization
	newShellCommand '{C2B136E2-D50E-405C-8784-363C582BF43E}'
	# Win10 1507から1607まで
	newShellCommand '{D9AC5E73-BB10-467B-B884-AA1E475C51F5}'
	# delegate folder that appears in Users Files Folder
	newShellCommand '{DFFACDC5-679F-4156-8947-C5C76BC0B67F}'
	# CompressedFolder
	newShellCommand '{E88DCCE0-B7B3-11D1-A9F0-00AA0060FA31}'
	# MyDocs Drop Target
	newShellCommand '{ECF03A32-103D-11D2-854D-006008059367}'
	# Shell File System Folder
	newShellCommand '{F3364BA0-65B9-11CE-A9BA-00AA004AE837}'
	# Sticky Notes Namespace Extension for Windows Desktop Search
	# Win10 1607まで
	newShellCommand '{F3F5824C-AD58-4728-AF59-A1EBE3392799}'
	# Subscription Folder
	newShellCommand '{F5175861-2688-11D0-9C5E-00AA00A45957}'
	# Internet Shortcut
	newShellCommand '{FBF23B40-E3F0-101B-8488-00AA003E56F8}'
	# History
	newShellCommand '{FF393560-C2A7-11CF-BFF4-444553540000}'
	# Windows Photo Viewer Image Verbs
	newShellCommand '{FFE2A43C-56B9-4BF5-9A79-CC6D4285608A}'
}

function Get-SpecialFolder {
	<#
	.SYNOPSIS
	Gets the special folders for Windows. This function supports the virtual folders, e.g. Control Panel and Recycle Bin.
	.OUTPUTS
	SpecialFolder[]
	#>

	[CmdletBinding()]
	[OutputType([SpecialFolder[]])]
	param ([switch]$IncludeShellCommand)

	return getSpecialFolder ($IncludeShellCommand -or $Develop) | Where-Object { $_ }
}

# 1つのリソースは1つのメニューにか設定できないので、必要な項目ごとに使用する
function getShieldImage {
	[OutputType([System.Windows.Controls.Image])]
	param ()

	# System.Drawing.Imageと曖昧になるので完全名で書いている
	$image = [System.Windows.Controls.Image]::new()

	$image.Source = [Imaging]::CreateBitmapSourceFromHIcon(
		[SystemIcons]::Shield.Handle, [Int32Rect]::Empty, $null
	)
	return $image
}

function Show-SpecialFolder {
	<#
	.SYNOPSIS
	Display the special folders for Windows in a dialog.
	.DESCRIPTION
	Display the special folders for Windows in a dialog. Open the folder to double-click on it. Show context menu to right-click on the folder.
	#>

	[CmdletBinding()]
	param ([switch]$IncludeShellCommand)

	# WPFが使えない場合
	if (($PSVersionTable['PSVersion'].Major -eq 6) -or ($Host.Runspace.ApartmentState -ne 'STA')) {
		$PSCmdlet.WriteError(
			[ErrorRecord]::new(
				[NotSupportedException]'Show-SpecialFolder can''t be started because this function needs WPF.',
				$null,
				'OperationStopped',
				$null
			)
		)
		return
	}

	Add-Type -AssemblyName PresentationFramework
	Add-Type -AssemblyName System.Drawing

	# SendKeys用
	# GUIにWPFを使っているのでWinFormsのSendKeysは使ってない
	$wsh = New-Object -ComObject WScript.Shell

	function selectInvokedCommand {
		$item = $dataGrid.SelectedItem
		if (!$item -or $item -isnot [SpecialFolder]) { return }

		$ErrorActionPreference = 'Stop'
		$modifiers = [Keyboard]::Modifiers
		try {
			if ($modifiers -band [ModifierKeys]::Alt) { $item.Properties() }
			elseif ($modifiers -band [ModifierKeys]::Control) { & $startPowershell }
			elseif ($modifiers -band [ModifierKeys]::Shift) { & $startCmd }
			else { & $openFolder }
		} catch { [MessageBox]::Show($_, $_.Exception.GetType().Name, 'OK', 'Warning') > $null }
	}

	function invokeCommandAsAdmin {
		param ([scriptblock]$command)

		# 昇格プロンプトで[いいえ]を選んだときのエラｰを無視する
		$ErrorActionPreference = 'SilentlyContinue'
		& $command
	}

	$openFolder = { $dataGrid.SelectedItem.Open('open') }
	$startPowershell = { $dataGrid.SelectedItem.Powershell('open') }
	$startCmd = { $dataGrid.SelectedItem.Cmd('open') }
	$startWsl = { $dataGrid.SelectedItem.Wsl('open') }

	$window = [Window][XamlReader]::Parse((Get-Content -LiteralPath "$PSScriptRoot\window.xaml" -Raw -ErrorAction Stop))

	$open = [MenuItem]$window.FindName('open')
	$openEx = [MenuItem]$window.FindName('openEx')
	$openAsAdmin = [MenuItem]$window.FindName('openAsAdmin')
	$powershell = [MenuItem]$window.FindName('powershell')
	$powershellEx = [MenuItem]$window.FindName('powershellEx')
	$powershellAsAdmin = [MenuItem]$window.FindName('powershellAsAdmin')
	$cmd = [MenuItem]$window.FindName('cmd')
	$cmdEx = [MenuItem]$window.FindName('cmdEx')
	$cmdAsAdmin = [MenuItem]$window.FindName('cmdAsAdmin')
	$wsl = [MenuItem]$window.FindName('wsl')
	$wslEx = [MenuItem]$window.FindName('wslEx')
	$wslAsAdmin = [MenuItem]$window.FindName('wslAsAdmin')
	$properties = [MenuItem]$window.FindName('properties')

	$openAsAdmin.Icon = getShieldImage
	$powershellAsAdmin.Icon = getShieldImage
	$cmdAsAdmin.Icon = getShieldImage
	$wslAsAdmin.Icon = getShieldImage

	$dataGrid = [DataGrid]($window.FindName('dataGrid'))
	$dataGrid.add_PreviewKeyDown(
		{
			# インテリセンスを効かせるために型を明示している
			$ke = [KeyEventArgs]$_

			# Home/End単独で一番上/一番下に移動できるようにする
			if (!([Keyboard]::Modifiers -band [ModifierKeys]::Control)) {
				switch ($ke.Key) {
					'Home' { $wsh.SendKeys('^{HOME}') }
					'End' { $wsh.SendKeys('^{END}') }
				}
			}

			# $ke.KeyだとAlt単独もAlt+Enterも'System'になるので[Keyboard]::IsKeyDown('Enter')を見ている
			if (![Keyboard]::IsKeyDown('Enter')) { return }

			$source = [Control]$ke.OriginalSource
			if ($source -is [DataGridCell]) { $dataGrid.SelectedItem = $source.DataContext }

			$ke.Handled = $true
			selectInvokedCommand
		}
	)
	$dataGrid.add_MouseDoubleClick(
		{
			$me = [MouseButtonEventArgs]$_

			if ($me.OriginalSource -is [TextBlock]) { selectInvokedCommand }
		}
	)
	$dataGrid.add_ContextMenuOpening(
		{
			$ce = [ContextMenuEventArgs]$_

			$item = $dataGrid.SelectedItem
			if ($item -isnot [SpecialFolder]) {
				$ce.Handled = $true
				return
			}

			$open.Visibility = 'Visible'

			$openEx.Visibility = 'Collapsed'
			$powershell.Visibility = 'Collapsed'
			$powershellEx.Visibility = 'Collapsed'
			$cmd.Visibility = 'Collapsed'
			$cmdEx.Visibility = 'Collapsed'
			$wsl.Visibility = 'Collapsed'
			$wslEx.Visibility = 'Collapsed'
			$properties.Visibility = 'Collapsed'

			if ([Keyboard]::Modifiers -band [ModifierKeys]::Shift) {
				if ($canFolderBeOpenedAsAdmin) {
					$open.Visibility = 'Collapsed'
					$openEx.Visibility = 'Visible'
				}
				if ($item -is [FileFolder]) {
					$cmdEx.Visibility = 'Visible'
					$powershellEx.Visibility = 'Visible'
					if ($isWslEnabled) { $wslEx.Visibility = 'Visible' }
				}
			} else {
				if ($item -is [FileFolder]) {
					$cmd.Visibility = 'Visible'
					$powershell.Visibility = 'Visible'
					if ($isWslEnabled) { $wsl.Visibility = 'Visible' }
				}
			}
			if ($item.HasProperties()) { $properties.Visibility = 'Visible' }
		}
	)
	$open.add_Click($openFolder)
	$window.FindName('openAsInvoker').add_Click($openFolder)
	$openAsAdmin.add_Click({ invokeCommandAsAdmin { $dataGrid.SelectedItem.Open('runas') } })
	$window.FindName('copyAsPath').add_Click({ Set-Clipboard $dataGrid.SelectedItem.Path })
	$powershell.add_Click($startPowershell)
	$window.FindName('powershellAsInvoker').add_Click($startPowershell)
	$powershellAsAdmin.add_Click({ invokeCommandAsAdmin { $dataGrid.SelectedItem.Powershell('runas') } })
	$cmd.add_Click($startCmd)
	$window.FindName('cmdAsInvoker').add_Click($startCmd)
	$cmdAsAdmin.add_Click({ invokeCommandAsAdmin { $dataGrid.SelectedItem.Cmd('runas') } })
	$wsl.add_Click($startWsl)
	$window.FindName('wslAsInvoker').add_Click($startWsl)
	$wslAsAdmin.add_Click({ invokeCommandAsAdmin { $dataGrid.SelectedItem.Wsl('runas') } })
	$properties.add_Click({ $dataGrid.SelectedItem.Properties() })
	$window.add_Loaded({
			$window.Topmost = $true
			$window.Topmost = $false
			$window.Activate()
		})

	$category = ''
	$dataGrid.ItemsSource = Get-SpecialFolder @PSBoundParameters |
		ForEach-Object {
			if ($category -ne $_.Category) {
				$category = $_.Category
				[pscustomobject]@{ Name = "Category: $($_.Category)"; Path = $null }
			}
			$_
		}

	$window.ShowDialog() > $null
}

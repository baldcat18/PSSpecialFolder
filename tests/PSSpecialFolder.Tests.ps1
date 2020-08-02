using namespace System.Diagnostics.CodeAnalysis

#Requires -Module @{ ModuleName = 'Pester'; ModuleVersion = '4.0.0' }

[SuppressMessage('PSUseDeclaredVarsMoreThanAssignments', 'IsDebugging')]
param()

# 開発用以外のバージョンをアンロードする
Get-Module PSSpecialFolder | Remove-Module

$module = Import-Module "$PSScriptRoot/../src/PSSpecialFolder.psm1" -Force -PassThru

InModuleScope PSSpecialFolder {
	Describe 'newSpecialFolder Test' {
		BeforeAll {
			$IsDebugging = $false
		}

		It 'Null' {
			newSpecialFolder $null | Should -BeFalse
		}
		It 'shell:Foo' {
			newSpecialFolder 'shell:Foo' | Should -BeFalse
		}
		It 'shell:::{00000000-0000-0000-0000-000000000000}' {
			newSpecialFolder 'shell:::{00000000-0000-0000-0000-000000000000}' | Should -BeFalse
		}
		It 'shell:Desktop' {
			$folder = newSpecialFolder 'shell:Desktop'

			$folder.Name | Should -Be 'Desktop'
			$folder.Path | Should -Be ([Environment]::GetFolderPath('Desktop'))
			$folder.PropertyTypes | Should -Be 'StartProcess'
			$folder.HasProperties() | Should -Be $true
		}
		It 'shell:MyComputerFolder' {
			$folder = newSpecialFolder 'shell:MyComputerFolder'

			$folder.Name | Should -Be 'MyComputerFolder'
			$folder.Path | Should -Be 'shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}'
			$folder.PropertyTypes | Should -Be 'Verb'
			$folder.HasProperties() | Should -Be $true
			$folder.PropertiesVerb.Name | Should -Be $propertiesName
		}
		$script:recycleBinPath = 'shell:::{645FF040-5081-101B-9F08-00AA002F954E}'
		It "$recycleBinPath (Recycle Bin)" {
			$folder = newSpecialFolder $recycleBinPath

			$folder.Name | Should -Be 'Recycle Bin'
			$folder.Path | Should -Be $recycleBinPath
		}
		$script:powerOptionsPath = `
			'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}'
		It "$powerOptionsPath (Power Options)" {
			$folder = newSpecialFolder $powerOptionsPath

			$folder.Name | Should -Be 'Power Options'
			$folder.Path | Should -Be $powerOptionsPath
			$folder.HasProperties() | Should -Be $false
		}
		$script:controlPanelPath = 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0'
		It "$controlPanelPath (All Control Panel Items)" {
			$folder = newSpecialFolder $controlPanelPath

			$folder.Name | Should -Be '0'
			$folder.Path | Should -Be $controlPanelPath
		}
		It 'B:\FooBar' {
			newSpecialFolder 'B:\FooBar' | Should -BeFalse
		}
		$script:myDocumentsPath = [Environment]::GetFolderPath('MyDocuments')
		It $myDocumentsPath {
			$folder = newSpecialFolder $myDocumentsPath

			$folder.Name | Should -Be (Split-Path -Leaf $myDocumentsPath)
			$folder.Path | Should -Be $myDocumentsPath
			$folder.Dir | Should -Be "file:\\\$myDocumentsPath"
		}
		$script:sharedWinDirPath = "\\$Env:COMPUTERNAME\$($Env:windir -replace ':', '$')"
		It $sharedWinDirPath {
			$folder = newSpecialFolder $sharedWinDirPath

			$folder.Name | Should -Be (Split-Path -Leaf $sharedWinDirPath)
			$folder.Path | Should -Be $sharedWinDirPath
			$folder.Dir | Should -Be "file:$sharedWinDirPath"
		}
		$script:appDataPath = [Environment]::GetFolderPath('ApplicationData')
		It "$appDataPath `"AppData`"" {
			$folder = newSpecialFolder $appDataPath 'AppData'

			$folder.Name | Should -Be "AppData"
			$folder.Path | Should -Be $appDataPath
		}
		$script:librariesPath = (
			Get-Item 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'
		).GetValue('{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}')
		if (!$librariesPath) { $script:librariesPath = "$appDataPath\Microsoft\Windows\Libraries" }
		It "shell:Libraries -Path `"$librariesPath`" -FolderItemForProperties (getDirectoryFolderItem $librariesPath)" {
			$folder = newSpecialFolder shell:Libraries `
				-Path $librariesPath -FolderItemForProperties (getDirectoryFolderItem $librariesPath)

			$folder.Name | Should -Be 'Libraries'
			$folder.Path | Should -Be $librariesPath
			$folder.HasProperties() > $null
			$folder.PropertiesVerb.Name | Should -Be $propertiesName
		}
	}

	Describe 'newShellCommand Test' {
		BeforeAll {
			$IsDebugging = $false
		}

		It 'Null' {
			newShellCommand $null | Should -BeFalse
		}
		It '{00000000-0000-0000-0000-000000000000}' {
			newShellCommand '{00000000-0000-0000-0000-000000000000}' | Should -BeFalse
		}
		$script:runClsid = '{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}'
		It "$runClsid (Run...)" {
			$folder = newShellCommand $runClsid

			$folder.Name | Should -Be 'Run...'
			$folder.Path | Should -Be "shell:::$runClsid"
			$folder.PropertyTypes | Should -Be 'None'
		}
		$script:fileExplorerClsid = '{52205FD8-5DFB-447D-801A-D0B52F2E83E1}'
		It "$fileExplorerClsid `"File Explorer`"" {
			$folder = newShellCommand $fileExplorerClsid 'File Explorer'

			$folder.Name | Should -Be 'File Explorer'
			$folder.Path | Should -Be "shell:::$fileExplorerClsid"
		}
	}

	Describe 'IsDebugging' {
		BeforeAll {
			$IsDebugging = $true
		}

		It 'shell:ThisPCDesktopFolder "DesktopFolder"' {
			(newSpecialFolder shell:ThisPCDesktopFolder 'DesktopFolder').Name | Should -Be 'DesktopFolder'
		}
		It 'shell:::{05D7B0F4-2121-4EFF-BF6B-ED3F69B894D9} "Notification Area Icons"' {
			(newSpecialFolder 'shell:::{05D7B0F4-2121-4EFF-BF6B-ED3F69B894D9}' 'Notification Area Icons').Name |
				Should -Be 'Notification Area Icons (Taskbar)'
		}
		It '{52205FD8-5DFB-447D-801A-D0B52F2E83E1} "File Explorer"' {
			(newShellCommand '{52205FD8-5DFB-447D-801A-D0B52F2E83E1}' 'File Explorer').Name |
				Should -Be 'File Explorer (@C:\Windows\system32\shell32.dll,-22067)'
		}
	}

	Describe 'getKnownFolderPath Test' {
		BeforeAll {
			$IsDebugging = $false
		}

		It 'Desktop' {
			getKnownFolderPath ThisPCDesktopFolder | Should -Be ([Environment]::GetFolderPath('Desktop'))
		}
		It 'FooBar' {
			{ getKnownFolderPath getKnownFolderPath FooBar } |
				Should -Throw -ExceptionType System.Management.Automation.RuntimeException
		}
	}
}

Remove-Module $module

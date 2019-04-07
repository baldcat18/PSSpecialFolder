Import-Module "$PSScriptRoot/../src/GetSpecialFolder.psm1" -Force

Describe "newSpecialFolder Test" {
	Context "Null" {
		It "Should Null" {
			newSpecialFolder $null | Should -BeFalse
		}
	}
	Context "shell:Foo" {
		It "Should Null" {
			newSpecialFolder "shell:Foo" | Should -BeFalse
		}
	}
	Context "shell:::{00000000-0000-0000-0000-000000000000}" {
		It "Should Null" {
			newSpecialFolder "shell:::{00000000-0000-0000-0000-000000000000}" | Should -BeFalse
		}
	}
	Context "shell:Desktop" {
		$folder = newSpecialFolder "shell:Desktop"
		It 'Name should "Desktop"' {
			$folder.Name | Should -Be "Desktop"
		}
		$desktopPath = [Environment]::GetFolderPath("Desktop")
		It "Path should `"$desktopPath`"" {
			$folder.Path | Should -Be $desktopPath
		}
	}
	Context "shell:MyComputerFolder" {
		$folder = newSpecialFolder "shell:MyComputerFolder"
		It 'Name should "MyComputerFolder"' {
			$folder.Name | Should -Be "MyComputerFolder"
		}
		$thisPcPath = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
		It "Path should `"$thisPcPath`"" {
			$folder.Path | Should -Be $thisPcPath
		}
	}
	$recycleBinPath = 'shell:::{645FF040-5081-101B-9F08-00AA002F954E}'
	Context "$recycleBinPath (Recycle Bin)" {
		$folder = newSpecialFolder $recycleBinPath
		It 'Name should "Recycle Bin"' {
			$folder.Name | Should -Be "Recycle Bin"
		}
		It "Path should `"$recycleBinPath`"" {
			$folder.Path | Should -Be $recycleBinPath
		}
	}
	$powerOptionsPath = 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{025A5937-A6BE-4686-A844-36FE4BEC8B6D}'
	Context "$powerOptionsPath (Power Options)" {
		$folder = newSpecialFolder $powerOptionsPath
		It 'Name should "Power Options"' {
			$folder.Name | Should -Be 'Power Options'
		}
		It "Path should `"$powerOptionsPath`"" {
			$folder.Path | Should -Be $powerOptionsPath
		}
	}
	$controlPanelPath = 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0'
	Context "$controlPanelPath (All Control Panel Items)" {
		$folder = newSpecialFolder $controlPanelPath
		It 'Name should "0"' {
			$folder.Name | Should -Be '0'
		}
		It "Path should `"$controlPanelPath`"" {
			$folder.Path | Should -Be $controlPanelPath
		}
	}
	Context "B:\FooBar" {
		It "Should Null" {
			newSpecialFolder "B:\FooBar" | Should -BeFalse
		}
	}
	$myDocumentsPath = [Environment]::GetFolderPath("MyDocuments")
	Context $myDocumentsPath {
		$folder = newSpecialFolder $myDocumentsPath
		$myDocumentsName = Split-Path -Leaf $myDocumentsPath
		It "Name should `"$myDocumentsName`"" {
			$folder.Name | Should -Be $myDocumentsName
		}
		It "Path should `"$myDocumentsPath`"" {
			$folder.Path | Should -Be $myDocumentsPath
		}
		It "Dir should `"file:\\\$myDocumentsPath`"" {
			$folder.Dir | Should -Be "file:\\\$myDocumentsPath"
		}
	}
	$sharedWinDirPath = "\\$Env:COMPUTERNAME\$($Env:windir -replace ':', '$')"
	Context $sharedWinDirPath {
		$folder = newSpecialFolder $sharedWinDirPath
		$sharedWinDirName = Split-Path -Leaf $sharedWinDirPath
		It "Name should `"$sharedWinDirName`"" {
			$folder.Name | Should -Be $sharedWinDirName
		}
		It "Path should `"$sharedWinDirPath`"" {
			$folder.Path | Should -Be $sharedWinDirPath
		}
		It "Dir should `"file:$sharedWinDirPath`"" {
			$folder.Dir | Should -Be "file:$sharedWinDirPath"
		}
	}
	$appDataPath = [Environment]::GetFolderPath("ApplicationData")
	Context "$appDataPath @{ Name = `"AppData`" }" {
		$folder = newSpecialFolder $appDataPath @{ Name = "AppData" }
		It "Name should `"AppData`"" {
			$folder.Name | Should -Be "AppData"
		}
		It "Path should `"$appDataPath`"" {
			$folder.Path | Should -Be $appDataPath
		}
	}
	$librariesPath = (Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders").
		GetValue("{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}")
	if (!$librariesPath) { $librariesPath = "$appDataPath\Microsoft\Windows\Libraries" }
	Context "shell:Libraries @{ Path = `"$librariesPath`" }" {
		$folder = newSpecialFolder shell:Libraries @{ Path = "$librariesPath" }
		It "Name should `"Libraries`"" {
			$folder.Name | Should -Be "Libraries"
		}
		It "Path should `"$librariesPath`"" {
			$folder.Path | Should -Be $librariesPath
		}
	}
}

Describe "newShellCommand Test" {
	Context "Null" {
		It "Should Null" {
			newShellCommand $null | Should -BeFalse
		}
	}
	Context "shell:::{00000000-0000-0000-0000-000000000000}" {
		It "Should Null" {
			newShellCommand "shell:::{00000000-0000-0000-0000-000000000000}" | Should -BeFalse
		}
	}
	$runPath = 'shell:::{2559A1F3-21D7-11D4-BDAF-00C04F60B9F0}'
	Context "$runPath (Run...)" {
		$folder = newShellCommand $runPath
		It 'Name should "Run..."' {
			$folder.Name | Should -Be "Run..."
		}
		It "Path should `"$runPath`"" {
			$folder.Path | Should -Be $runPath
		}
	}
	$fileExplorerPath = 'shell:::{52205FD8-5DFB-447D-801A-D0B52F2E83E1}'
	Context "$fileExplorerPath `"File Explorer`"" {
		$folder = newShellCommand $fileExplorerPath "File Explorer"
		It 'Name should "File Explorer"' {
			$folder.Name | Should -Be "File Explorer"
		}
		It "Path should `"$fileExplorerPath`"" {
			$folder.Path | Should -Be $fileExplorerPath
		}
	}
}

Remove-Module GetSpecialFolder

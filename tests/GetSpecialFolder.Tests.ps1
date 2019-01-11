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
		It 'Title should "Desktop"' {
			$folder.Title | Should -Be "Desktop"
		}
		$desktopPath = [Environment]::GetFolderPath("Desktop")
		It "Path should `"$desktopPath`"" {
			$folder.Path | Should -Be $desktopPath
		}
	}
	Context "shell:MyComputerFolder" {
		$folder = newSpecialFolder "shell:MyComputerFolder"
		It 'Title should "MyComputerFolder"' {
			$folder.Title | Should -Be "MyComputerFolder"
		}
		$thisPcPath = "shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
		It "Path should `"$thisPcPath`"" {
			$folder.Path | Should -Be $thisPcPath
		}
	}
	$recycleBinPath = 'shell:::{645FF040-5081-101B-9F08-00AA002F954E}'
	Context "$recycleBinPath (Recycle Bin)" {
		$folder = newSpecialFolder $recycleBinPath
		It 'Title should "Recycle Bin"' {
			$folder.Title | Should -Be "Recycle Bin"
		}
		It "Path should `"$recycleBinPath`"" {
			$folder.Path | Should -Be $recycleBinPath
		}
	}
	$controlPanelPath = 'shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0'
	Context "$controlPanelPath (All Control Panel Items)" {
		$folder = newSpecialFolder $controlPanelPath
		It 'Title should "0"' {
			$folder.Title | Should -Be '0'
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
		It "Title should `"$myDocumentsName`"" {
			$folder.Title | Should -Be $myDocumentsName
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
		It "Title should `"$sharedWinDirName`"" {
			$folder.Title | Should -Be $sharedWinDirName
		}
		It "Path should `"$sharedWinDirPath`"" {
			$folder.Path | Should -Be $sharedWinDirPath
		}
		It "Dir should `"file:$sharedWinDirPath`"" {
			$folder.Dir | Should -Be "file:$sharedWinDirPath"
		}
	}
	$appDataPath = [Environment]::GetFolderPath("ApplicationData")
	Context "$appDataPath @{ Title = `"AppData`" }" {
		$folder = newSpecialFolder $appDataPath @{ Title = "AppData" }
		It "Title should `"AppData`"" {
			$folder.Title | Should -Be "AppData"
		}
		It "Path should `"$appDataPath`"" {
			$folder.Path | Should -Be $appDataPath
		}
	}
	$LibrariesPath = (Get-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders").
		GetValue("{1B3EA5DC-B587-4786-B4EF-BD1DC332AEAE}")
	if (!$LibrariesPath) { $LibrariesPath = "$appDataPath\Microsoft\Windows\Libraries" }
	Context "shell:Libraries @{ Path = `"$LibrariesPath`" }" {
		$folder = newSpecialFolder shell:Libraries @{ Path = "$LibrariesPath" }
		It "Title should `"Libraries`"" {
			$folder.Title | Should -Be "Libraries"
		}
		It "Path should `"$LibrariesPath`"" {
			$folder.Path | Should -Be $LibrariesPath
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
	$taskbarPath = 'shell:::{0DF44EAA-FF21-4412-828E-260A8728E7F1}'
	Context "$taskbarPath (Taskbar)" {
		$folder = newShellCommand $taskbarPath
		It 'Title should "Taskbar"' {
			$folder.Title | Should -Be "Taskbar"
		}
		It "Path should `"$taskbarPath`"" {
			$folder.Path | Should -Be $taskbarPath
		}
	}
	$fileExplorerPath = 'shell:::{52205FD8-5DFB-447D-801A-D0B52F2E83E1}'
	Context "$fileExplorerPath `"File Explorer`"" {
		$folder = newShellCommand $fileExplorerPath "File Explorer"
		It 'Title should "File Explorer"' {
			$folder.Title | Should -Be "File Explorer"
		}
		It "Path should `"$fileExplorerPath`"" {
			$folder.Path | Should -Be $fileExplorerPath
		}
	}
}

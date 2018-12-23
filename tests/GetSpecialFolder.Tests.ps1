Import-Module "$PSScriptRoot/../src/GetSpecialFolder.psm1" -Force

Describe "newSpecialFolder のテスト" {
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
}

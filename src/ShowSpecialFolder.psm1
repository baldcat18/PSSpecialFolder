using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Windows.Markup

Set-StrictMode -Version Latest

<#
.SYNOPSIS
Display the special folders for Windows in a dialog.
.DESCRIPTION
Display the special folders for Windows in a dialog. Open the folder to double-click on it. Show context menu to right-click on the folder.
#>
function Show-SpecialFolder {
	[CmdletBinding()]
	param ([switch]$IncludeShellCommand)
	
	# WPFが使えない場合
	if (($PSVersionTable['PSVersion'].Major -ge 6) -or ($Host.Runspace.ApartmentState -ne 'STA')) {
		powershell.exe -Sta -Command "Show-SpecialFolder -IncludeShellCommand:`$$IncludeShellCommand"
		return
	}
	
	$isWslEnabled = Test-Path "$([Environment]::GetFolderPath('System'))/wsl.exe"
	
	Add-Type -AssemblyName PresentationFramework
	Add-Type -AssemblyName PresentationCore
	
	$window = [XamlReader]::Parse((Get-Content "$PSScriptRoot/window.xaml" -Raw))
	
	$cmd = $window.FindName('cmd')
	$powershell = $window.FindName('powershell')
	$wsl = $window.FindName('wsl')
	$properties = $window.FindName('properties')
	
	$dataGrid = $window.FindName('dataGrid')
	
	$dataGrid.add_MouseDoubleClick({
		if ($_.OriginalSource.GetType() -eq [TextBlock]) { $dataGrid.SelectedItem.Open() }
	})
	$dataGrid.add_MouseRightButtonUp({
		if ($_.OriginalSource.GetType() -eq [Border]) { $_.Handled = $true }
	})
	$dataGrid.add_ContextMenuOpening({
		$item = $dataGrid.SelectedItem
		if (!$item) { return }
		
		$cmd.Visibility = 'Collapsed'
		$powershell.Visibility = 'Collapsed'
		$wsl.Visibility = 'Collapsed'
		$properties.Visibility = 'Collapsed'
		
		if ($item.IsDirectory) {
			$cmd.Visibility = 'Visible'
			$powershell.Visibility = 'Visible'
			if ($isWslEnabled) { $wsl.Visibility = 'Visible' }
		}
		if ($item.TestProperties()) { $properties.Visibility = 'Visible' }
	})
	$window.FindName('open').add_Click({ $dataGrid.SelectedItem.Open() })
	$window.FindName('copyAsPath').add_Click({ $dataGrid.SelectedItem.CopyAsPath() })
	$cmd.add_Click({ $dataGrid.SelectedItem.StartCmd() })
	$powershell.add_Click({ $dataGrid.SelectedItem.StartPowershell() })
	$wsl.add_Click({ $dataGrid.SelectedItem.StartLinuxShell() })
	$properties.add_Click({
		try {
			$dataGrid.SelectedItem.ShowProperties()
		} catch {
			[MessageBox]::Show($_, $_.Exception.GetType().Name, 'OK', 'Warning') > $null
		}
	})
	
	$getSpecialFolderArgs = @{
		IncludeShellCommand = $IncludeShellCommand
		InformationAction = 'SilentlyContinue' # この関数ではカテゴリ名を表示しない
		Debug = $DebugPreference -ne 'SilentlyContinue'
	}
	$dataGrid.ItemsSource = Get-SpecialFolder @getSpecialFolderArgs
	
	$window.ShowDialog() > $null
}

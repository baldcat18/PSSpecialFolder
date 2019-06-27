using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Windows.Input
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
	
	$openFolderHandler = { $dataGrid.SelectedItem.Open('open') }
	$openCmdHandler = { $dataGrid.SelectedItem.StartCmd('open') }
	$openPowershellHandler = { $dataGrid.SelectedItem.StartPowershell('open') }
	$openWslHandler = { $dataGrid.SelectedItem.StartLinuxShell('open') }
	
	$window = [XamlReader]::Parse((Get-Content "$PSScriptRoot/window.xaml" -Raw))
	
	$openAsAdmin = $window.FindName('openAsAdmin')
	$cmd = $window.FindName('cmd')
	$cmdEx= $window.FindName('cmdEx')
	$powershell = $window.FindName('powershell')
	$powershellEx = $window.FindName('powershellEx')
	$wsl = $window.FindName('wsl')
	$wslEx = $window.FindName('wslEx')
	$properties = $window.FindName('properties')
	
	$dataGrid = $window.FindName('dataGrid')
	
	$dataGrid.add_MouseDoubleClick({
		if ($_.OriginalSource.GetType() -eq [TextBlock]) { & $openFolderHandler }
	})
	$dataGrid.add_MouseRightButtonUp({
		if ($_.OriginalSource.GetType() -eq [Border]) { $_.Handled = $true }
	})
	$dataGrid.add_ContextMenuOpening({
		$item = $dataGrid.SelectedItem
		if (!$item) { return }
		
		$openAsAdmin.Visibility = 'Collapsed'
		$cmd.Visibility = 'Collapsed'
		$cmdEx.Visibility = 'Collapsed'
		$powershell.Visibility = 'Collapsed'
		$powershellEx.Visibility = 'Collapsed'
		$wsl.Visibility = 'Collapsed'
		$wslEx.Visibility = 'Collapsed'
		$properties.Visibility = 'Collapsed'
		
		if ([Keyboard]::Modifiers -band [ModifierKeys]::Shift) {
			$openAsAdmin.Visibility = "Visible"
			if ($item.IsDirectory) {
				$cmdEx.Visibility = 'Visible'
				$powershellEx.Visibility = 'Visible'
				if ($isWslEnabled) { $wslEx.Visibility = 'Visible' }
			}
		} else {
			if ($item.IsDirectory) {
				$cmd.Visibility = 'Visible'
				$powershell.Visibility = 'Visible'
				if ($isWslEnabled) { $wsl.Visibility = 'Visible' }
			}
		}
		if ($item.TestProperties()) { $properties.Visibility = 'Visible' }
	})
	$window.FindName('open').add_Click($openFolderHandler)
	$window.FindName('copyAsPath').add_Click({ $dataGrid.SelectedItem.CopyAsPath() })
	$openAsAdmin.add_Click({ $dataGrid.SelectedItem.Open('runas') })
	$cmd.add_Click($openCmdHandler)
	$window.FindName('cmdAsInvoker').add_Click($openCmdHandler)
	$window.FindName('cmdAsAdmin').add_Click({ $dataGrid.SelectedItem.StartCmd('runas') })
	$powershell.add_Click($openPowershellHandler)
	$window.FindName('powershellAsInvoker').add_Click($openPowershellHandler)
	$window.FindName('powershellAsAdmin').add_Click({ $dataGrid.SelectedItem.StartPowershell('runas') })
	$wsl.add_Click($openWslHandler)
	$window.FindName('wslAsInvoker').add_Click($openWslHandler)
	$window.FindName('wslAsAdmin').add_Click({ $dataGrid.SelectedItem.StartLinuxShell('runas') })
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

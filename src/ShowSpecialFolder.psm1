using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Windows.Input
using namespace System.Windows.Interop
using namespace System.Windows.Markup

Set-StrictMode -Version Latest

function getShieldImage {
	[OutputType([System.Windows.Controls.Image])]
	param ()
	
	$image = [Image]::new()
	$image.Source = [Imaging]::CreateBitmapSourceFromHIcon(
		[System.Drawing.SystemIcons]::Shield.Handle, [Int32Rect]::Empty, $null
	)
	return $image
}

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
	
	$isExplorerRunasLaunchingUser =
		!((Get-Item 'HKLM:/SOFTWARE/Classes/AppID/{CDCBCFCA-3CDC-436f-A4E2-0E02075250C2}').GetValue('RunAs'))
	$isWslEnabled = Test-Path "$([Environment]::GetFolderPath('System'))/wsl.exe"
	
	Add-Type -AssemblyName PresentationFramework
	Add-Type -AssemblyName System.Drawing
	
	$startCommand = {
		param([scriptblock]$command)
		
		$ErrorActionPreference = 'Stop'
		try { & $command }
		catch { [MessageBox]::Show($_, $_.Exception.GetType().Name, 'OK', 'Warning') > $null }
	}
	$startCommandRunas = {
		param([scriptblock]$command)
		
		# 昇格プロンプトで[いいえ]を選んだときのエラｰを無視する
		$ErrorActionPreference = 'SilentlyContinue'
		& $command
	}
	
	$openFolder = { $dataGrid.SelectedItem.Open() }
	$openCmd = { $dataGrid.SelectedItem.StartCmd() }
	$openPowershell = { $dataGrid.SelectedItem.StartPowershell() }
	$openWsl = { $dataGrid.SelectedItem.StartLinuxShell() }
	$showProperties = { & $startCommand { $dataGrid.SelectedItem.ShowProperties() } }
	
	$window = [XamlReader]::Parse((Get-Content "$PSScriptRoot/window.xaml" -Raw))
	
	$openAsAdmin = $window.FindName('openAsAdmin')
	$cmd = $window.FindName('cmd')
	$cmdEx= $window.FindName('cmdEx')
	$cmdAsAdmin = $window.FindName('cmdAsAdmin')
	$powershell = $window.FindName('powershell')
	$powershellEx = $window.FindName('powershellEx')
	$powershellAsAdmin = $window.FindName('powershellAsAdmin')
	$wsl = $window.FindName('wsl')
	$wslEx = $window.FindName('wslEx')
	$wslAsAdmin = $window.FindName('wslAsAdmin')
	$properties = $window.FindName('properties')
	
	$openAsAdmin.Icon = getShieldImage
	$cmdAsAdmin.Icon = getShieldImage
	$powershellAsAdmin.Icon = getShieldImage
	$wslAsAdmin.Icon = getShieldImage
	
	$dataGrid = $window.FindName('dataGrid')
	$dataGrid.add_MouseDoubleClick({
		if ($_.OriginalSource.GetType() -ne [TextBlock]) { return }
		
		$item = $dataGrid.SelectedItem
		if (!$item -or $item.GetType().FullName -ne 'SpecialFolder') { return }
		
		$modifiers = [Keyboard]::Modifiers
		if ($modifiers -band [ModifierKeys]::Alt) { & $showProperties }
		elseif ($modifiers -band [ModifierKeys]::Control) { & $startCommand $openPowershell }
		elseif ($modifiers -band [ModifierKeys]::Shift) { & $startCommand $openCmd }
		else { & $openFolder }
	})
	$dataGrid.add_ContextMenuOpening({
		if ($_.OriginalSource.GetType() -ne [TextBlock]) {
			$_.Handled = $true
			return
		}
		
		$item = $dataGrid.SelectedItem
		if ($item.GetType().FullName -ne 'SpecialFolder') {
			$_.Handled = $true
			return
		}
		
		$openAsAdmin.Visibility = 'Collapsed'
		$cmd.Visibility = 'Collapsed'
		$cmdEx.Visibility = 'Collapsed'
		$powershell.Visibility = 'Collapsed'
		$powershellEx.Visibility = 'Collapsed'
		$wsl.Visibility = 'Collapsed'
		$wslEx.Visibility = 'Collapsed'
		$properties.Visibility = 'Collapsed'
		
		if ([Keyboard]::Modifiers -band [ModifierKeys]::Shift) {
			if ($isExplorerRunasLaunchingUser) { $openAsAdmin.Visibility = 'Visible' }
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
	$dataGrid.add_PreviewKeyDown({
		if ($_.Key -ne 'Enter') { return }
		
		$_.Handled = $true
		$item = $dataGrid.SelectedItem
		if (!$item -or $item.GetType().FullName -ne 'SpecialFolder') { return }
		
		$modifiers = [Keyboard]::Modifiers
		if ($modifiers -band [ModifierKeys]::Control) { & $startCommand $openPowershell }
		elseif ($modifiers -band [ModifierKeys]::Shift) { & $startCommand $openCmd }
		else { & $openFolder }
	})
	$window.FindName('open').add_Click($openFolder)
	$window.FindName('copyAsPath').add_Click({ $dataGrid.SelectedItem.CopyAsPath() })
	$openAsAdmin.add_Click({ & $startCommandRunas { $dataGrid.SelectedItem.Open('runas') } })
	$cmd.add_Click($openCmd)
	$window.FindName('cmdAsInvoker').add_Click($openCmd)
	$cmdAsAdmin.add_Click({ & $startCommandRunas { $dataGrid.SelectedItem.StartCmd('runas') } })
	$powershell.add_Click($openPowershell)
	$window.FindName('powershellAsInvoker').add_Click($openPowershell)
	$powershellAsAdmin.add_Click({ & $startCommandRunas { $dataGrid.SelectedItem.StartPowershell('runas') } })
	$wsl.add_Click($openWsl)
	$window.FindName('wslAsInvoker').add_Click($openWsl)
	$wslAsAdmin.add_Click({ & $startCommandRunas { $dataGrid.SelectedItem.StartLinuxShell('runas') } })
	$properties.add_Click($showProperties)
	
	$getSpecialFolderArgs = @{
		IncludeShellCommand = $IncludeShellCommand
		Debug = $DebugPreference -ne 'SilentlyContinue'
	}
	$isShowCategory = $InformationPreference -ne 'Ignore' -and $InformationPreference -ne 'SilentlyContinue'
	$dataGrid.ItemsSource = Get-SpecialFolder @getSpecialFolderArgs 6>&1 |
		ForEach-Object {
			if ($_.GetType().FullName -eq 'SpecialFolder') { $_ }
			elseif ($isShowCategory) { [pscustomobject]@{ Name = $_.ToString().Replace("`n", ''); Path = $null } }
		}
	
	$window.ShowDialog() > $null
}

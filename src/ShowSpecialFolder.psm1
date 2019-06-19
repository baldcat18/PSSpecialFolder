using namespace System.Windows.Markup

Set-StrictMode -Version Latest

function Show-SpecialFolder {
	[CmdletBinding()]
	param ([switch]$IncludeShellCommand)
	
	# WPFが使えない場合
	if (($PSVersionTable['PSVersion'].Major -ge 6) -or ($Host.Runspace.ApartmentState -ne "STA")) {
		powershell.exe -Sta -Command "Show-SpecialFolder -IncludeShellCommand:`$$IncludeShellCommand"
		return
	}
	
	Add-Type -AssemblyName PresentationFramework
	Add-Type -AssemblyName PresentationCore
	
	$window = [XamlReader]::Parse((Get-Content "$PSScriptRoot/window.xaml" -Raw))
	
	$dataGrid = $window.FindName('dataGrid')
	
	$getSpecialFolderArgs = @{
		IncludeShellCommand = $IncludeShellCommand
		InformationAction = 'SilentlyContinue' # この関数ではカテゴリ名を表示しない
		Debug = $DebugPreference -ne 'SilentlyContinue'
	}
	$dataGrid.ItemsSource = Get-SpecialFolder @getSpecialFolderArgs
	
	$window.ShowDialog() > $null
}

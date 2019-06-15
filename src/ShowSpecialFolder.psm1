Set-StrictMode -Version Latest

function Show-SpecialFolder {
	[CmdletBinding()]
	param ([switch]$IncludeShellCommand)
	
	# WPFが使えない場合
	if (($PSVersionTable['PSVersion'].Major -ge 6) -or ($Host.Runspace.ApartmentState -ne "STA")) {
		powershell.exe -Sta -Command "Show-SpecialFolder -IncludeShellCommand:`$$IncludeShellCommand"
		return
	}
	
	# 今のところは単にGet-SpecialFolderを呼ぶだけ
	# Show-SpecialFolderでは情報ストリームを使わないので非表示に
	Get-SpecialFolder -IncludeShellCommand:$IncludeShellCommand -InformationAction SilentlyContinue
}

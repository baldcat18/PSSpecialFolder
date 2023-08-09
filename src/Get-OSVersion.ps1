<#
.SYNOPSIS
OSのバージョンやエディション情報を返す
#>

[CmdletBinding()]
[OutputType([hashtable])]
param()

$osVersionString = (Get-CimInstance Win32_OperatingSystem).Version

$versionKey = Get-Item 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

$revison = $versionKey.GetValue('UBR')
if ($null -eq $revison) { $revison = ($versionKey.GetValue('BuildLabEx') -split '\.', 3)[1] }

$displayVersion = $versionKey.GetValue('DisplayVersion') # Win10 20H2から
if ($null -eq $displayVersion) { $displayVersion = $versionKey.GetValue('ReleaseId') } # Win10 1511から

return @{
	Version = [Version]"$osVersionString.$revison"
	VersionString = $osVersionString
	DisplayVersion = $displayVersion
	Edition = $versionKey.GetValue('EditionID')
}

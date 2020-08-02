<#
.SYNOPSIS
既知のフォルダーをまとめて作成する
#>

#Requires -Version 4.0

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$verbose = $VerbosePreference -eq 'Continue'

$source = Get-Content -LiteralPath "$PSScriptRoot\..\src\KnownFolder.cs" -Raw
Add-Type -TypeDefinition $source -ErrorAction Stop

$KF_FLAG_DEFAULT = [uint32]0x00000000
$KF_FLAG_CREATE = [uint32]0x00008000

(& $PSScriptRoot\..\src\FolderGuids.ps1).GetEnumerator() |
	ForEach-Object {
		$result = ''
		$folder = New-Object Win32API.KnownFolder $_.Value, $KF_FLAG_DEFAULT
		if ($folder.Result -eq 'OK') {
			if (!$verbose) { return }
			$result = $folder.Result
		} else {
			$folder = New-Object Win32API.KnownFolder $_.Value, $KF_FLAG_CREATE
			if ($folder.Result -eq 'NotFound' -and !$verbose) { return }
			$result = if ($folder.Result -eq 'OK') { 'New' } else { $folder.Result }
		}

		Write-Output ([pscustomobject]@{ Name = $_.Name; Result = $result; Path = $folder.Path })
	}

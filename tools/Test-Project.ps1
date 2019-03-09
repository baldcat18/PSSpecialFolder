<#
.SYNOPSIS
PSSpecialFolderプロジェクトに対するテストを実行する
#>

[CmdletBinding()]
param()

Push-Location $PSScriptRoot/..

Write-Output "Test-ModuleManifest:"
Test-ModuleManifest src/PSSpecialFolder.psd1
Write-Output "`nInvoke-ScriptAnalyzer:"
Invoke-ScriptAnalyzer -Path . -Recurse
Write-Output "`nInvoke-Pester:"
Invoke-Pester tests

Pop-Location

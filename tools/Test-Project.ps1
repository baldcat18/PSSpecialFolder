<#
.SYNOPSIS
PSSpecialFolderプロジェクトに対するテストを実行する
#>

[CmdletBinding()]
param()

$projPath = (Resolve-Path $PSScriptRoot\..).Path

Write-Output 'Test-ModuleManifest:'
Test-ModuleManifest $projPath\src\PSSpecialFolder.psd1

Write-Output "`nInvoke-ScriptAnalyzer:"
Invoke-ScriptAnalyzer -Path $projPath -Recurse | Format-List RuleName, Severity, Message, ScriptPath, Line, Extent

Write-Output "`nInvoke-Pester:"
Invoke-Pester $projPath\tests

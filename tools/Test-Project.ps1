<#
.SYNOPSIS
PSSpecialFolderプロジェクトに対するテストを実行する
#>

#Requires -Module @{ ModuleName = 'Pester'; ModuleVersion = '4.0.0' }

[CmdletBinding()]
param()

$projPath = (Resolve-Path $PSScriptRoot\..).Path

Write-Output 'Test-ModuleManifest:'
Test-ModuleManifest $projPath\src\PSSpecialFolder.psd1

Write-Output "`nInvoke-ScriptAnalyzer:"
$analyzeResult = @(Invoke-ScriptAnalyzer -Path $projPath -Recurse)
if ($analyzeResult.Length) {
	$analyzeResult | Format-List RuleName, Severity, Message, ScriptPath, Line, Extent
	return
}

Write-Output "`nInvoke-Pester:"
Invoke-Pester $projPath\tests

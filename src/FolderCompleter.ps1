<#
.SYNOPSIS
Get-SpecialFolderPath や New-SpecialFolder に渡す引数を補完する
.NOTES
関数にすると .psd1 ファイルでエクスポートする必要があるのでスクリプトファイルにしている
#>

[OutputType([string[]])]
param([string]$WordToComplete, [switch]$Get)

& $PSScriptRoot\FolderNames.ps1 -Get:$Get |
	Where-Object { $_ -like "${WordToComplete}*" } |
	ForEach-Object { "'$_'" }

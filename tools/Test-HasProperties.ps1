<#
.SYNOPSIS
特殊フォルダーでプロパティを使えるか調べる
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest

$shell = New-Object -ComObject Shell.Application
$propertiesName = @($shell.NameSpace([Environment+SpecialFolder]::Desktop).Self.Verbs())[-1].Name

@(
	# True になるはず
	[Environment+SpecialFolder]::UserProfile
	'shell:InternetFolder'

	# False になるはず
	'shell:UsersFilesFolder'
	'shell:Profile'
	'shell:Libraries'
	'shell:UsersLibrariesFolder'
) |
	ForEach-Object {
		return [pscustomobject]@{
			Folder = $_
			HasProperties = @($shell.NameSpace($_).Self.Verbs())[-1].Name -eq $propertiesName
		}
	}

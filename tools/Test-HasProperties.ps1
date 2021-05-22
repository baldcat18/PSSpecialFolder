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
	[ordered]@{ Folder = [Environment+SpecialFolder]::UserProfile; Expected = $true }
	[ordered]@{ Folder = 'shell:InternetFolder'; Expected = $true }

	[ordered]@{ Folder = 'shell:UsersFilesFolder'; Expected = $false }
	[ordered]@{ Folder = 'shell:Profile'; Expected = $false }
	[ordered]@{ Folder = 'shell:Libraries'; Expected = $false }
	[ordered]@{ Folder = 'shell:UsersLibrariesFolder'; Expected = $false }
) |
	ForEach-Object {
		$_['Actual'] = @($shell.NameSpace($_['Folder']).Self.Verbs())[-1].Name -eq $propertiesName
		return [pscustomobject]$_
	}

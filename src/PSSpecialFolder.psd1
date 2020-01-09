@{
	ModuleVersion = '2.4.0'
	GUID = '958cb972-fb7e-4222-92b1-152d5b95a2a9'
	Author = 'BaldCat'
	Copyright = '(c) 2019 BaldCat. All rights reserved.'
	Description = 'PowerShell Module gets the special folders for Windows 8.1 and 10 (Version 1709 or later).'
	PowerShellVersion = '5.1'
	CompatiblePSEditions = @('Core', 'Desktop')
	RootModule = 'PSSpecialFolder.psm1'
	FunctionsToExport = @('Get-SpecialFolder', 'Show-SpecialFolder')
	CmdletsToExport = @()
	AliasesToExport = @()
	PrivateData = @{
		LicenseUri = 'https://github.com/baldcat18/PSSpecialFolder/blob/master/LICENSE'
		PSData = @{ Prerelease = 'alpha'}
		Tags = @('Folder', 'Windows')
	}
}

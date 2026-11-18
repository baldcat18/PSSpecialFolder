@{
	ModuleVersion = '2.6.2'
	GUID = '958cb972-fb7e-4222-92b1-152d5b95a2a9'
	Author = 'BaldCat'
	Copyright = '(c) 2019 BaldCat. All rights reserved.'
	Description = 'PowerShell Module gets the special folders for Windows 10 (Version 22H2) and 11 (Version 25H2 to 26H2).'
	PowerShellVersion = '5.1'
	CompatiblePSEditions = @('Core', 'Desktop')
	RootModule = 'PSSpecialFolder.psm1'
	FunctionsToExport = @('Get-SpecialFolder', 'Get-SpecialFolderPath', 'New-SpecialFolder', 'Show-SpecialFolder')
	CmdletsToExport = @()
	AliasesToExport = @()
	PrivateData = @{
		PSData = @{
			LicenseUri = 'https://github.com/baldcat18/PSSpecialFolder/blob/main/LICENSE'
			ProjectUri = 'https://github.com/baldcat18/PSSpecialFolder'
			Tags = @('Folder', 'Windows')
		}
	}
}

class OS {
	static [bool]$Win10 # Win10以降
	static [bool]$Win10_20h2 # Win10 20H2以降
	static [bool]$Win10_22h2_Only # Win10 22H2のみ
	static [bool]$Win11 # Win11以降
	static [bool]$Win11_22h2 # Win11 22H2以降
	static [bool]$Win11_22h2_moment4 # Win11 22H2 Moment4以降
	static [bool]$Win11_24h2_Only # Win11 24H2のみ
	static [bool]$Win11_24h2_3624 # Win11 24H2 26100.3624以降
	static [bool]$Win11_25h2 # Win11 25H2以降


	static OS() {
		$os = & $PSScriptRoot\Get-OSVersion.ps1
		$version = $os['Version']
		$verString = $os['VersionString']

		[OS]::Win10 = $version -gt '10.0.10240'
		[OS]::Win10_20h2 = $version -gt '10.0.19042'
		[OS]::Win10_22h2_Only = $verString -eq '10.0.19045'
		[OS]::Win11 = $version -gt '10.0.22000'
		[OS]::Win11_22h2 = $version -gt '10.0.22621'
		[OS]::Win11_22h2_moment4 = $version -ge '10.0.22621.2361'
		[OS]::Win11_24h2_Only = $verString -eq '10.0.26100'
		[OS]::Win11_24h2_3624 = $version -ge '10.0.26100.3624'
		[OS]::Win11_25h2 = $version -gt '10.0.26200'
	}
}

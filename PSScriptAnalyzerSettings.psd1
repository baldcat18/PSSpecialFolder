@{
	Rules = @{
		PSUseCompatibleSyntax = @{
			Enable = $true
			TargetVersions = @('5.1', '6.2', '7.0')
		}
	}
	ExcludeRules = @(
		'PSAvoidTrailingWhitespace'
	)
}

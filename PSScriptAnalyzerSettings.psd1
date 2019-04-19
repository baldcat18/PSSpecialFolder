@{
	Rules = @{
		PSUseCompatibleSyntax = @{
			Enable = $true
			TargetVersions = @('5.0', '5.1', '6.1', '6.2')
		}
	}
	ExcludeRules = @(
		'PSAvoidTrailingWhitespace'
		'PSAvoidUsingPositionalParameters'
		'PSPossibleIncorrectComparisonWithNull'
	)
}

@{
	Rules = @{
		PSAvoidUsingDoubleQuotesForConstantString = @{ Enable = $true }
		PSUseCompatibleSyntax = @{
			Enable = $true
			TargetVersions = @('5.1', '6.2', '7.0')
		}
	}
}

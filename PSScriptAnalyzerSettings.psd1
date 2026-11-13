@{
    IncludeRules = @(
        'PSAvoidAssignmentToAutomaticVariable'
        'PSAvoidUsingEmptyCatchBlock'
        'PSPossibleIncorrectComparisonWithNull'
        'PSPossibleIncorrectUsageOfAssignmentOperator'
        'PSUseCompatibleSyntax'
    )

    Rules = @{
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @('5.1')
        }
    }
}

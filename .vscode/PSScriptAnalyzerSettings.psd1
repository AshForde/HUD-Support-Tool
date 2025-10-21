@{
    # Only diagnostic records of the specified severity will be generated.
    # Uncomment the following line if you only want Errors and Warnings but
    # not Information diagnostic records.
    # Severity = @('Error', 'Warning')

    # Analyze **only** the following rules. Use IncludeRules when you want
    # to invoke only a small subset of the default rules.
    IncludeRules = @(
        'PSAvoidDefaultValueSwitchParameter',
        'PSMisleadingBacktick',
        'PSMissingModuleManifestField',
        'PSReservedCmdletChar',
        'PSReservedParams',
        'PSShouldProcess',
        'PSUseApprovedVerbs',
        'PSAvoidUsingCmdletAliases',
        'PSUseDeclaredVarsMoreThanAssignments'
    );
    Severity = @(
        'Error',
        'Warning'
    );
    ExcludeRules = @(
        'PSAlignAssignmentStatement'
    );
    Rules = @{
        'PSAvoidUsingCmdletAliases' = @{
            AllowList = @('cd', 'clear', 'ls')
        }
        'PSAvoidUsingWriteHost' = @{
            AllowForegroundColor = @('Yellow', 'Green', 'Cyan', 'Magenta')
        }
        'PSUseCompatibleCmdlets' = @{
            compatibility = @('desktop-5.1.14393.206-windows')
        }
    }
}

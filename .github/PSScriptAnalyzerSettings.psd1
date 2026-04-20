@{
    # Only treat these rules as errors (build-breaking); everything else is a warning.
    Severity = @('Error', 'Warning')

    ExcludeRules = @(
        # Internal script-scope functions intentionally use non-standard verbs
        # (Apply-Telemetry, Apply-Firewall, etc.) — these are not exported cmdlets.
        'PSUseApprovedVerbs',

        # Write-Host is intentional throughout — this is a console-first admin tool,
        # not a library, so suppressing output via Write-Verbose is not appropriate.
        'PSAvoidUsingWriteHost',

        # Some admin operations legitimately require WMI over CIM for compatibility
        # with PowerShell 5.1 on older Windows builds.
        'PSUseShouldProcessForStateChangingFunctions'
    )
}

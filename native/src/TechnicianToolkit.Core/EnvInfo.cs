namespace TechnicianToolkit.Core;

/// <summary>
/// Thin wrappers over the environment values the PowerShell scripts read from
/// <c>$env:COMPUTERNAME</c>, <c>$env:USERNAME</c> and <c>$env:USERDOMAIN</c>.
/// Centralised so reports and logs render the same identity strings the scripts
/// produced.
/// </summary>
public static class EnvInfo
{
    /// <summary>Machine / computer name (<c>$env:COMPUTERNAME</c>).</summary>
    public static string MachineName => Environment.MachineName;

    /// <summary>Current user name (<c>$env:USERNAME</c>).</summary>
    public static string UserName => Environment.UserName;

    /// <summary>Current user domain (<c>$env:USERDOMAIN</c>).</summary>
    public static string UserDomain => Environment.UserDomainName;

    /// <summary>The <c>DOMAIN\User</c> form used in report "Run As" fields.</summary>
    public static string UserDomainQualified => $"{UserDomain}\\{UserName}";
}

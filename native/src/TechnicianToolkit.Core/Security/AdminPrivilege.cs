using System.Diagnostics;
using System.Runtime.Versioning;
using System.Security.Principal;

namespace TechnicianToolkit.Core.Security;

/// <summary>
/// Elevation helpers — the native counterpart to <c>Test-IsAdmin</c>,
/// <c>Assert-AdminPrivilege</c> and <c>Invoke-AdminElevation</c>.
/// </summary>
public static class AdminPrivilege
{
    /// <summary>
    /// True when the current process is running elevated (member of the local
    /// Administrators role). Always false on non-Windows platforms.
    /// </summary>
    public static bool IsAdmin()
    {
        if (!OperatingSystem.IsWindows())
        {
            return false;
        }

        return IsAdminWindows();
    }

    [SupportedOSPlatform("windows")]
    private static bool IsAdminWindows()
    {
        using var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    /// <summary>
    /// Relaunches the current executable elevated (UAC prompt) and signals the
    /// caller to exit. Returns false when already elevated (nothing to do) or
    /// when the user cancelled the UAC prompt. This is the GUI-friendly analogue
    /// of <c>Invoke-AdminElevation</c>.
    /// </summary>
    /// <returns>
    /// True if an elevated instance was started and the current process should
    /// exit; false if already elevated or elevation was declined/failed.
    /// </returns>
    public static bool RelaunchElevated()
    {
        if (IsAdmin())
        {
            return false;
        }

        var exePath = Environment.ProcessPath;
        if (string.IsNullOrEmpty(exePath))
        {
            return false;
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            UseShellExecute = true,
            Verb = "runas",
            WorkingDirectory = AppContext.BaseDirectory,
        };

        try
        {
            Process.Start(startInfo);
            return true;
        }
        catch
        {
            // User declined the UAC prompt, or elevation is unavailable.
            return false;
        }
    }
}

using TechnicianToolkit.Core.Config;

namespace TechnicianToolkit.Core;

/// <summary>
/// Path resolution helpers. Ports <c>Resolve-LogDirectory</c> from the module.
/// </summary>
public static class TkPaths
{
    /// <summary>
    /// Fallback directory used when nothing else resolves. The PowerShell module
    /// used <c>$PSScriptRoot</c>; the native default is the executable folder.
    /// </summary>
    public static string BaseDirectory { get; set; } = AppContext.BaseDirectory;

    /// <summary>
    /// Returns the configured <c>LogDirectory</c> (creating it if needed), or the
    /// supplied fallback path when no log directory is configured. Mirrors
    /// <c>Resolve-LogDirectory</c>.
    /// </summary>
    public static string ResolveLogDirectory(string fallbackPath)
    {
        var cfg = TkConfig.Get();
        if (!string.IsNullOrWhiteSpace(cfg.LogDirectory))
        {
            try
            {
                if (!Directory.Exists(cfg.LogDirectory))
                {
                    Directory.CreateDirectory(cfg.LogDirectory);
                }
            }
            catch
            {
                // Best effort — fall through to returning the configured path.
            }

            return cfg.LogDirectory;
        }

        return fallbackPath;
    }
}

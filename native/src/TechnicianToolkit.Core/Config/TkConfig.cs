using System.Text.Json;
using System.Text.Json.Nodes;

namespace TechnicianToolkit.Core.Config;

/// <summary>
/// Reads and writes <c>config.json</c>, porting <c>Get-TKConfig</c> and
/// <c>Set-TKConfig</c> from TechnicianToolkit.psm1.
/// </summary>
/// <remarks>
/// The PowerShell module keeps <c>config.json</c> next to the module file
/// (<c>$PSScriptRoot</c>). The native app has no such fixed anchor, so the path
/// is exposed via <see cref="ConfigPath"/> and defaults to a file beside the
/// executable. <see cref="Set"/> preserves unknown keys the same way the
/// PowerShell version does (it round-trips the whole object, only touching the
/// one key), so hand-added config is never dropped.
/// </remarks>
public static class TkConfig
{
    /// <summary>
    /// Location of <c>config.json</c>. Defaults to a file next to the running
    /// executable; the host app may repoint it (e.g. to a per-user profile dir).
    /// </summary>
    public static string ConfigPath { get; set; } =
        Path.Combine(AppContext.BaseDirectory, "config.json");

    private static readonly JsonSerializerOptions WriteOptions = new()
    {
        WriteIndented = true,
    };

    /// <summary>
    /// Returns the toolkit configuration, filling missing keys with empty-string
    /// defaults and migrating the legacy <c>Phantom</c> section forward to
    /// <c>Revenant</c> (v2.x → v3.0 rename), exactly as <c>Get-TKConfig</c> does.
    /// A missing or unparseable file yields all-default values.
    /// </summary>
    public static TkConfigData Get()
    {
        var defaults = new TkConfigData();

        if (!File.Exists(ConfigPath))
        {
            return defaults;
        }

        JsonObject? root;
        try
        {
            var raw = File.ReadAllText(ConfigPath);
            root = JsonNode.Parse(raw) as JsonObject;
        }
        catch
        {
            return defaults;
        }

        if (root is null)
        {
            return defaults;
        }

        var data = new TkConfigData
        {
            OrgName = ReadString(root, "OrgName"),
            LogDirectory = ReadString(root, "LogDirectory"),
            TeamsWebhook = ReadString(root, "TeamsWebhook"),
        };

        data.Archive.DefaultDestination = ReadSectionString(root, "Archive", "DefaultDestination");
        data.Revenant.DefaultDestination = ReadSectionString(root, "Revenant", "DefaultDestination");
        data.Covenant.DefaultTimezone = ReadSectionString(root, "Covenant", "DefaultTimezone");
        data.Covenant.DefaultLocalAdminUser = ReadSectionString(root, "Covenant", "DefaultLocalAdminUser");

        // Legacy migration: v2.x used a 'Phantom' section; v3.0 renamed it to
        // 'Revenant'. Carry a populated Phantom.DefaultDestination forward when
        // Revenant's is still blank, matching Get-TKConfig. Read via the safe
        // helper so a malformed Phantom value can never make Get() throw.
        var phantomDest = ReadSectionString(root, "Phantom", "DefaultDestination");
        if (!string.IsNullOrWhiteSpace(phantomDest) &&
            string.IsNullOrWhiteSpace(data.Revenant.DefaultDestination))
        {
            data.Revenant.DefaultDestination = phantomDest;
        }

        return data;
    }

    /// <summary>
    /// Writes a single value into <c>config.json</c>, porting <c>Set-TKConfig</c>.
    /// When <paramref name="section"/> is supplied the value is nested under that
    /// section (created if missing). All other existing keys are preserved.
    /// </summary>
    public static void Set(string key, string value, string? section = null)
    {
        if (string.IsNullOrEmpty(key))
        {
            throw new ArgumentException("Key is required.", nameof(key));
        }

        JsonObject root;
        if (File.Exists(ConfigPath))
        {
            try
            {
                root = JsonNode.Parse(File.ReadAllText(ConfigPath)) as JsonObject ?? new JsonObject();
            }
            catch
            {
                root = new JsonObject();
            }
        }
        else
        {
            root = new JsonObject();
        }

        if (!string.IsNullOrEmpty(section))
        {
            if (root[section] is not JsonObject sectionObj)
            {
                sectionObj = new JsonObject();
                root[section] = sectionObj;
            }

            sectionObj[key] = value;
        }
        else
        {
            root[key] = value;
        }

        var dir = Path.GetDirectoryName(ConfigPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        File.WriteAllText(ConfigPath, root.ToJsonString(WriteOptions));
    }

    private static string ReadString(JsonObject root, string key)
    {
        var node = root[key];
        if (node is null)
        {
            return string.Empty;
        }

        try
        {
            return node.GetValue<string?>() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string ReadSectionString(JsonObject root, string section, string key)
    {
        if (root[section] is not JsonObject sectionObj)
        {
            return string.Empty;
        }

        var node = sectionObj[key];
        if (node is null)
        {
            return string.Empty;
        }

        try
        {
            return node.GetValue<string?>() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }
}

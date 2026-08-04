using System.Management;

namespace TechnicianToolkit.Tools.Collectors;

/// <summary>
/// Small conveniences over <see cref="System.Management"/> for reading WMI/CIM
/// property values without repeating null/type guards at every call site. These
/// replace the loose property access PowerShell gives for free (e.g.
/// <c>$obj.FreeSpace</c>).
/// </summary>
internal static class WmiUtil
{
    /// <summary>Reads a property as a string, or null if absent.</summary>
    public static string? Str(ManagementBaseObject mo, string prop)
    {
        try
        {
            return mo[prop]?.ToString();
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Reads a numeric property as ulong, or null if absent/unparsable.</summary>
    public static ulong? U64(ManagementBaseObject mo, string prop)
    {
        try
        {
            var v = mo[prop];
            return v is null ? null : Convert.ToUInt64(v);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Reads a numeric property as uint, or null if absent/unparsable.</summary>
    public static uint? U32(ManagementBaseObject mo, string prop)
    {
        try
        {
            var v = mo[prop];
            return v is null ? null : Convert.ToUInt32(v);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Reads a bool property, or null if absent.</summary>
    public static bool? Bool(ManagementBaseObject mo, string prop)
    {
        try
        {
            var v = mo[prop];
            return v is null ? null : Convert.ToBoolean(v);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Converts a CIM DATETIME string (e.g. <c>20260804153000.000000+000</c>) to
    /// a local <see cref="DateTime"/>, or null if absent/invalid.
    /// </summary>
    public static DateTime? CimDate(ManagementBaseObject mo, string prop)
    {
        try
        {
            var raw = mo[prop]?.ToString();
            if (string.IsNullOrEmpty(raw))
            {
                return null;
            }

            return ManagementDateTimeConverter.ToDateTime(raw);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Runs a WMI query against a scope path and yields each object.</summary>
    public static IEnumerable<ManagementObject> Query(string scopePath, string query)
    {
        var scope = new ManagementScope(scopePath);
        scope.Connect();
        using var searcher = new ManagementObjectSearcher(scope, new ObjectQuery(query));
        foreach (ManagementObject mo in searcher.Get())
        {
            yield return mo;
        }
    }

    /// <summary>Runs a query against the default <c>root\cimv2</c> namespace.</summary>
    public static IEnumerable<ManagementObject> QueryCimv2(string query) =>
        Query(@"\\.\root\cimv2", query);
}

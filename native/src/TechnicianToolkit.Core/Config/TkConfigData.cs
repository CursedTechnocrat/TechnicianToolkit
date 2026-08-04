namespace TechnicianToolkit.Core.Config;

/// <summary>
/// Strongly-typed view of <c>config.json</c>. Mirrors the shape returned by
/// <c>Get-TKConfig</c>: every field defaults to an empty string so callers never
/// receive null.
/// </summary>
public sealed class TkConfigData
{
    public string OrgName { get; set; } = string.Empty;
    public string LogDirectory { get; set; } = string.Empty;
    public string TeamsWebhook { get; set; } = string.Empty;

    public ArchiveSection Archive { get; set; } = new();
    public RevenantSection Revenant { get; set; } = new();
    public CovenantSection Covenant { get; set; } = new();

    public sealed class ArchiveSection
    {
        public string DefaultDestination { get; set; } = string.Empty;
    }

    public sealed class RevenantSection
    {
        public string DefaultDestination { get; set; } = string.Empty;
    }

    public sealed class CovenantSection
    {
        public string DefaultTimezone { get; set; } = string.Empty;
        public string DefaultLocalAdminUser { get; set; } = string.Empty;
    }
}

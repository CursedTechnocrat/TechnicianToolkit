using System.Globalization;
using System.Text;
using TechnicianToolkit.Core;
using TechnicianToolkit.Core.Config;
using TechnicianToolkit.Core.Html;
using TechnicianToolkit.Tools.Collectors;

namespace TechnicianToolkit.Tools.Diagnostics;

/// <summary>
/// A.U.G.U.R. — Analyzes, Uncovers &amp; Gauges Unit Reliability.
/// Native port of augur.ps1: inspects every physical disk (health, operational
/// state, SMART failure prediction, bus/media type, serial/firmware) and every
/// lettered/labelled volume, then writes a dark-themed HTML report.
/// </summary>
public sealed class AugurTool : ITool
{
    public string Key => "14";
    public string Name => "A.U.G.U.R.";
    public string Title => "Disk Health & SMART";
    public string Category => "Diagnostics & Reporting";
    public string Description =>
        "Physical disk health & SMART failure prediction, plus per-volume integrity and usage.";
    public bool RequiresAdmin => true;
    public bool SupportsWhatIf => false;

    private static string Esc(string? s) => HtmlUtil.Esc(s);

    public ToolResult Run(ToolContext ctx)
    {
        try
        {
            ctx.Report(ProgressLevel.Step, "Reading physical disks...");
            IReadOnlyList<PhysicalDiskInfo> disks;
            try
            {
                disks = SmartCollector.CollectPhysicalDisks();
            }
            catch (Exception ex)
            {
                ctx.Report(ProgressLevel.Fail, $"Error reading physical disks: {ex.Message}");
                disks = Array.Empty<PhysicalDiskInfo>();
            }

            ctx.Report(ProgressLevel.Ok, $"Read {disks.Count} physical disk(s).");

            ctx.Report(ProgressLevel.Step, "Reading volumes...");
            IReadOnlyList<SmartVolumeInfo> volumes;
            try
            {
                volumes = SmartCollector.CollectVolumes();
            }
            catch (Exception ex)
            {
                ctx.Report(ProgressLevel.Fail, $"Error reading volumes: {ex.Message}");
                volumes = Array.Empty<SmartVolumeInfo>();
            }

            ctx.Report(ProgressLevel.Ok, $"Read {volumes.Count} volume(s).");

            var healthy = disks.Count(d => d.HealthStatus == "Healthy" && d.SmartPrediction != "FAILING");
            var warning = disks.Count(d => d.HealthStatus == "Warning");
            var critical = disks.Count(d => d.HealthStatus != "Healthy" && d.HealthStatus != "Warning");
            var smartFail = disks.Count(d => d.SmartPrediction == "FAILING");

            ctx.Report(ProgressLevel.Step, "Building HTML report...");
            var reportDir = string.IsNullOrWhiteSpace(ctx.OutputPath) ? TkPaths.BaseDirectory : ctx.OutputPath!;
            Directory.CreateDirectory(reportDir);
            var reportFile = Path.Combine(reportDir, $"AUGUR_{DateTime.Now:yyyyMMdd_HHmmss}.html");

            var html = BuildHtml(disks, volumes, healthy, warning, critical, smartFail);
            File.WriteAllText(reportFile, html, new UTF8Encoding(false));
            ctx.Report(ProgressLevel.Ok, $"Report saved: {reportFile}");

            var summary = new List<KeyValuePair<string, string>>
            {
                new("Disks", disks.Count.ToString()),
                new("Healthy", healthy.ToString()),
                new("Warning", warning.ToString()),
                new("Critical", critical.ToString()),
                new("SMART Fail", smartFail.ToString()),
            };

            return ToolResult.Ok(reportFile, summary);
        }
        catch (Exception ex)
        {
            TechnicianToolkit.Core.Diagnostics.TkErrorLog.Write("augur", ex.Message, "Run");
            ctx.Report(ProgressLevel.Fail, ex.Message);
            return ToolResult.Fail(ex.Message);
        }
    }

    private static string DiskHealthBadge(string health) => health switch
    {
        "Healthy" => "tk-badge-ok",
        "Warning" => "tk-badge-warn",
        _ => "tk-badge-err",
    };

    private static string SmartBadge(string prediction) => prediction switch
    {
        "FAILING" => "tk-badge-err",
        "OK" => "tk-badge-ok",
        _ => "tk-badge-info",
    };

    private static string UsageBadge(double pct) =>
        pct >= 90 ? "tk-badge-err" : pct >= 75 ? "tk-badge-warn" : "tk-badge-ok";

    private static string BuildHtml(
        IReadOnlyList<PhysicalDiskInfo> disks,
        IReadOnlyList<SmartVolumeInfo> volumes,
        int healthy, int warning, int critical, int smartFail)
    {
        var cfg = TkConfig.Get();
        var machine = EnvInfo.MachineName;
        var orgPrefix = string.IsNullOrWhiteSpace(cfg.OrgName) ? "" : $"{cfg.OrgName} -- ";
        var generated = DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

        // The original passes an HTML badge as the "Overall" meta value, but the
        // shared head helper HTML-escapes meta values, so it rendered as literal
        // tags. We surface the same verdict as clean plain text instead.
        var overall = (critical > 0 || smartFail > 0)
            ? $"{critical + smartFail} critical issue(s)"
            : warning > 0
                ? $"{warning} warning(s)"
                : "All Healthy";

        var sb = new StringBuilder();
        sb.Append(TkHtml.Head(
            title: "Disk Health Assessment",
            scriptName: "A.U.G.U.R.",
            subtitle: $"{orgPrefix}{machine}",
            metaItems: new[]
            {
                new KeyValuePair<string, string>("Machine", machine),
                new KeyValuePair<string, string>("Run As", EnvInfo.UserDomainQualified),
                new KeyValuePair<string, string>("Generated", generated),
                new KeyValuePair<string, string>("Overall", overall),
            },
            navItems: new[] { "Physical Disks", "Volumes" }));

        sb.Append($@"<div class=""tk-summary-row"">
  <div class=""tk-summary-card ok""><div class=""tk-summary-num"">{healthy}</div><div class=""tk-summary-lbl"">Healthy</div></div>
  <div class=""tk-summary-card warn""><div class=""tk-summary-num"">{warning}</div><div class=""tk-summary-lbl"">Warning</div></div>
  <div class=""tk-summary-card err""><div class=""tk-summary-num"">{critical}</div><div class=""tk-summary-lbl"">Critical</div></div>
  <div class=""tk-summary-card err""><div class=""tk-summary-num"">{smartFail}</div><div class=""tk-summary-lbl"">SMART Fail</div></div>
</div>
");

        // Section 1 — Physical Disks
        string diskBody;
        if (disks.Count == 0)
        {
            diskBody = "<p class='tk-info-box'>No physical disk data available.</p>";
        }
        else
        {
            var rows = new StringBuilder();
            foreach (var d in disks)
            {
                rows.Append($@"
        <tr>
          <td>{Esc(d.DeviceId)}</td>
          <td>{Esc(d.FriendlyName)}</td>
          <td>{Esc(d.Serial)}</td>
          <td>{Esc(d.MediaType)}</td>
          <td>{Esc(d.BusType)}</td>
          <td>{d.SizeGb} GB</td>
          <td><span class='{DiskHealthBadge(d.HealthStatus)}'>{Esc(d.HealthStatus)}</span></td>
          <td>{Esc(d.OperationalStatus)}</td>
          <td><span class='{SmartBadge(d.SmartPrediction)}'>{Esc(d.SmartPrediction)}</span></td>
          <td><code>{Esc(d.SmartReason)}</code></td>
          <td>{Esc(d.Firmware)}</td>
        </tr>");
            }

            diskBody = $@"
    <div class=""tk-table-wrap"">
    <table class=""tk-table"">
      <thead>
        <tr>
          <th>ID</th><th>Name</th><th>Serial</th><th>Type</th><th>Bus</th><th>Size</th>
          <th>Health</th><th>Status</th><th>SMART</th><th>SMART Reason</th><th>Firmware</th>
        </tr>
      </thead>
      <tbody>
{rows}
      </tbody>
    </table>
    </div>";
        }

        sb.Append($@"
<div class=""tk-section"" id=""physical-disks"">
  <div class=""tk-card"">
    <div class=""tk-card-header""><span class=""tk-card-label"">Physical Disks ({disks.Count})</span></div>
{diskBody}
  </div>
</div>

<div class=""tk-divider""></div>
");

        // Section 2 — Volumes
        string volBody;
        if (volumes.Count == 0)
        {
            volBody = "<p class='tk-info-box'>No volume data available.</p>";
        }
        else
        {
            var rows = new StringBuilder();
            foreach (var v in volumes)
            {
                rows.Append($@"
        <tr>
          <td><code>{Esc(v.Drive)}</code></td>
          <td>{Esc(v.Label)}</td>
          <td>{Esc(v.FileSystem)}</td>
          <td>{v.TotalGb} GB</td>
          <td>{v.FreeGb} GB</td>
          <td><span class='{UsageBadge(v.PctUsed)}'>{v.PctUsed}%</span></td>
          <td><span class='{DiskHealthBadge(v.Health)}'>{Esc(v.Health)}</span></td>
        </tr>");
            }

            volBody = $@"
    <div class=""tk-table-wrap"">
    <table class=""tk-table"">
      <thead>
        <tr><th>Drive</th><th>Label</th><th>File System</th><th>Total</th><th>Free</th><th>Usage</th><th>Health</th></tr>
      </thead>
      <tbody>
{rows}
      </tbody>
    </table>
    </div>";
        }

        sb.Append($@"
<div class=""tk-section"" id=""volumes"">
  <div class=""tk-card"">
    <div class=""tk-card-header""><span class=""tk-card-label"">Volumes ({volumes.Count})</span></div>
{volBody}
  </div>
</div>
");

        sb.Append(TkHtml.Foot("A.U.G.U.R. v3.6"));
        return sb.ToString();
    }
}

using System.Globalization;
using System.Text;
using TechnicianToolkit.Core;
using TechnicianToolkit.Core.Config;
using TechnicianToolkit.Core.Html;
using TechnicianToolkit.Tools.Collectors;

namespace TechnicianToolkit.Tools.Diagnostics;

/// <summary>
/// S.C.R.Y.E.R. — System Consolidated Report Yielding Exhaustive Results.
/// Native port of scryer.ps1: one HTML report rolling up system overview, local
/// users, disk space, disk (SMART) health and service/scheduled-task status.
/// </summary>
public sealed class ScryerTool : ITool
{
    public string Key => "16";
    public string Name => "S.C.R.Y.E.R.";
    public string Title => "Unified Diagnostic Report";
    public string Category => "Diagnostics & Reporting";
    public string Description =>
        "One consolidated HTML report: system, users, disk space, SMART health, services & tasks.";
    public bool RequiresAdmin => true;
    public bool SupportsWhatIf => false;

    private static string Esc(string? s) => HtmlUtil.Esc(s);

    public ToolResult Run(ToolContext ctx)
    {
        try
        {
            // 1/5 — system overview
            ctx.Report(ProgressLevel.Step, "[1/5] Collecting system information...");
            var sys = SystemInfoCollector.Collect();
            ctx.Report(ProgressLevel.Ok, "System info collected.");

            // 2/5 — user accounts
            ctx.Report(ProgressLevel.Step, "[2/5] Auditing local user accounts...");
            var admins = UserCollector.GetAdminMembers();
            var users = UserCollector.Collect(admins);
            // Faithful to SCRYER's Sort-Object { -not IsAdmin }, { LastLogon } -Descending.
            var usersSorted = users
                .OrderByDescending(u => !u.IsAdmin)
                .ThenByDescending(u => u.LastLogon ?? DateTime.MinValue)
                .ToList();
            ctx.Report(ProgressLevel.Ok, $"User accounts audited ({usersSorted.Count} users).");

            // 3/5 — disk space
            ctx.Report(ProgressLevel.Step, "[3/5] Checking disk space...");
            var volumes = VolumeCollector.Collect();
            var volWarn = volumes.Count(v => v.Health != "ok");
            ctx.Report(ProgressLevel.Ok, $"Disk space checked ({volumes.Count} volumes, {volWarn} warnings).");

            // 4/5 — disk health (SMART)
            ctx.Report(ProgressLevel.Step, "[4/5] Assessing disk health...");
            IReadOnlyList<PhysicalDiskInfo> disks;
            bool smartAvailable;
            try
            {
                disks = SmartCollector.CollectPhysicalDisks();
                smartAvailable = true;
            }
            catch
            {
                disks = Array.Empty<PhysicalDiskInfo>();
                smartAvailable = false;
            }

            var diskWarn = disks.Count(d => d.HealthClass != "ok");
            ctx.Report(ProgressLevel.Ok, $"Disk health assessed ({disks.Count} disks).");

            // 5/5 — services & scheduled tasks
            ctx.Report(ProgressLevel.Step, "[5/5] Checking services and scheduled tasks...");
            var stoppedSvcs = ServiceCollector.CollectStoppedAutomatic();
            var failedTasks = ScheduledTaskCollector.CollectFailed();
            ctx.Report(ProgressLevel.Ok,
                $"Services/tasks checked ({stoppedSvcs.Count} svc issues, {failedTasks.Count} task failures).");

            // Build report
            ctx.Report(ProgressLevel.Step, "Building HTML report...");
            var reportDir = ResolveReportDir(ctx.OutputPath);
            Directory.CreateDirectory(reportDir);
            var reportFile = Path.Combine(reportDir,
                $"SCRYER_Report_{DateTime.Now:yyyyMMdd_HHmmss}.html");

            var html = BuildHtml(sys, usersSorted, volumes, disks, smartAvailable, stoppedSvcs, failedTasks);
            File.WriteAllText(reportFile, html, new UTF8Encoding(false));
            ctx.Report(ProgressLevel.Ok, $"Report saved: {reportFile}");

            var summary = new List<KeyValuePair<string, string>>
            {
                new("Users", usersSorted.Count.ToString()),
                new("Volumes", $"{volumes.Count} ({volWarn} with warnings)"),
                new("Disks", $"{disks.Count} ({diskWarn} with issues)"),
                new("Svc Issues", stoppedSvcs.Count.ToString()),
                new("Task Fails", failedTasks.Count.ToString()),
            };

            return ToolResult.Ok(reportFile, summary);
        }
        catch (Exception ex)
        {
            TechnicianToolkit.Core.Diagnostics.TkErrorLog.Write("scryer", ex.Message, "Run");
            ctx.Report(ProgressLevel.Fail, ex.Message);
            return ToolResult.Fail(ex.Message);
        }
    }

    private static string ResolveReportDir(string? outputPath)
    {
        if (!string.IsNullOrWhiteSpace(outputPath))
        {
            return outputPath;
        }

        var cfg = TkConfig.Get();
        if (!string.IsNullOrWhiteSpace(cfg.LogDirectory))
        {
            return cfg.LogDirectory;
        }

        return TkPaths.BaseDirectory;
    }

    private static string BuildHtml(
        SystemInfo sys,
        IReadOnlyList<UserAccount> users,
        IReadOnlyList<VolumeInfo> volumes,
        IReadOnlyList<PhysicalDiskInfo> disks,
        bool smartAvailable,
        IReadOnlyList<StoppedServiceInfo> stoppedSvcs,
        IReadOnlyList<FailedTaskInfo> failedTasks)
    {
        var cfg = TkConfig.Get();
        var orgPrefix = string.IsNullOrWhiteSpace(cfg.OrgName) ? "" : $"{Esc(cfg.OrgName)} -- ";

        var sb = new StringBuilder();
        sb.Append(TkHtml.Head(
            title: "SCRYER Unified Report",
            scriptName: "S.C.R.Y.E.R.",
            subtitle: $"{orgPrefix}{EnvInfo.MachineName}",
            metaItems: new[]
            {
                new KeyValuePair<string, string>("Generated", DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)),
                new KeyValuePair<string, string>("OS", $"{sys.OsCaption} (Build {sys.OsBuild})"),
                new KeyValuePair<string, string>("Run As", EnvInfo.UserDomainQualified),
            },
            navItems: new[] { "System Overview", "User Accounts", "Disk Space", "Disk Health", "Services and Tasks" }));

        // 01 — System Overview
        sb.Append($@"
<div class=""tk-section"" id=""system-overview"">
  <div class=""tk-section-title""><span class=""tk-section-num"">01</span> System Overview</div>
  <div class=""tk-card"">
    <table class=""tk-table"">
      <tbody>
        <tr><td class=""tk-card-label"">Hostname</td><td>{Esc(EnvInfo.MachineName)}</td></tr>
        <tr><td class=""tk-card-label"">Operating System</td><td>{Esc(sys.OsCaption)} (Build {sys.OsBuild})</td></tr>
        <tr><td class=""tk-card-label"">Manufacturer / Model</td><td>{Esc(sys.Manufacturer)} {Esc(sys.Model)}</td></tr>
        <tr><td class=""tk-card-label"">CPU</td><td>{Esc(sys.CpuName)} ({sys.CpuCores} logical cores)</td></tr>
        <tr><td class=""tk-card-label"">RAM</td><td>{sys.TotalRamGb} GB total / {sys.FreeRamGb} GB free</td></tr>
        <tr><td class=""tk-card-label"">Last Boot</td><td>{Esc(sys.LastBoot)}</td></tr>
        <tr><td class=""tk-card-label"">Uptime</td><td>{Esc(sys.Uptime)}</td></tr>
        <tr><td class=""tk-card-label"">.NET Runtime</td><td>{Esc(sys.RuntimeVersion)}</td></tr>
      </tbody>
    </table>
  </div>
</div>
");

        // 02 — User Accounts
        var userRows = new StringBuilder();
        foreach (var u in users)
        {
            var statusBadge = u.Enabled
                ? "<span class='tk-badge-ok'>Enabled</span>"
                : "<span class='tk-badge-err'>Disabled</span>";
            var adminCell = u.IsAdmin ? "<span class='tk-badge-err'>Yes</span>" : "";
            var pwdCell = u.PasswordRequired ? "Yes" : "No";
            var lastLogon = u.LastLogon?.ToString("yyyy-MM-dd") ?? "Never";
            userRows.Append($@"
        <tr>
          <td>{Esc(u.Name)}</td>
          <td>{Esc(u.FullName)}</td>
          <td>{statusBadge}</td>
          <td>{Esc(lastLogon)}</td>
          <td>{adminCell}</td>
          <td>{pwdCell}</td>
        </tr>");
        }

        sb.Append($@"
<div class=""tk-section"" id=""user-accounts"">
  <div class=""tk-section-title""><span class=""tk-section-num"">02</span> User Accounts</div>
  <div class=""tk-card"">
    <table class=""tk-table"">
      <thead>
        <tr>
          <th>User</th>
          <th>Full Name</th>
          <th>Status</th>
          <th>Last Logon</th>
          <th>Admin</th>
          <th>Pwd Required</th>
        </tr>
      </thead>
      <tbody>
{userRows}
      </tbody>
    </table>
  </div>
</div>
");

        // 03 — Disk Space
        var diskSpaceCards = new StringBuilder();
        foreach (var v in volumes)
        {
            var letterLabel = string.IsNullOrEmpty(v.Label)
                ? Esc(v.Letter)
                : $"{Esc(v.Letter)} -- {Esc(v.Label)}";
            diskSpaceCards.Append($@"
  <div class=""tk-card"">
    <div class=""tk-card-header"">{letterLabel}</div>
    <div class=""tk-progress-wrap"">
      <div class=""tk-progress-bar {v.Health}"" style=""width: {v.PctUsed}%""></div>
    </div>
    <div style=""font-size:0.85em; color:var(--tk-text-dim); margin-top:4px"">
      {v.PctUsed}% used &nbsp;|&nbsp; {v.UsedGb} GB used / {v.TotalGb} GB total &nbsp;|&nbsp; {v.FreeGb} GB free
    </div>
  </div>");
        }

        if (diskSpaceCards.Length == 0)
        {
            diskSpaceCards.Append("<div class='tk-card'><div class='tk-info-box'>No fixed volumes found.</div></div>");
        }

        sb.Append($@"
<div class=""tk-section"" id=""disk-space"">
  <div class=""tk-section-title""><span class=""tk-section-num"">03</span> Disk Space</div>
{diskSpaceCards}
</div>
");

        // 04 — Disk Health
        string diskHealthBody;
        if (!smartAvailable)
        {
            diskHealthBody = @"
  <div class=""tk-card"">
    <div class=""tk-info-box"">SMART data unavailable -- Storage module not accessible on this system.</div>
  </div>";
        }
        else
        {
            var diskHealthRows = new StringBuilder();
            foreach (var d in disks)
            {
                var healthBadge = $"<span class='tk-badge-{d.HealthClass}'>{Esc(d.HealthStatus)}</span>";
                var tempVal = d.Temperature?.ToString() ?? "--";
                var wearVal = d.Wear is not null ? $"{d.Wear}%" : "--";
                diskHealthRows.Append($@"
        <tr>
          <td>{Esc(d.FriendlyName)}</td>
          <td>{Esc(d.MediaType)}</td>
          <td>{d.SizeGb}</td>
          <td>{healthBadge}</td>
          <td>{tempVal}</td>
          <td>{wearVal}</td>
        </tr>");
            }

            diskHealthBody = $@"
  <div class=""tk-card"">
    <table class=""tk-table"">
      <thead>
        <tr>
          <th>Drive</th>
          <th>Type</th>
          <th>Size (GB)</th>
          <th>Health</th>
          <th>Temp (C)</th>
          <th>Wear (%)</th>
        </tr>
      </thead>
      <tbody>
{diskHealthRows}
      </tbody>
    </table>
  </div>";
        }

        sb.Append($@"
<div class=""tk-section"" id=""disk-health"">
  <div class=""tk-section-title""><span class=""tk-section-num"">04</span> Disk Health</div>
{diskHealthBody}
</div>
");

        // 05 — Services and Tasks
        string svcCardBody;
        if (stoppedSvcs.Count == 0)
        {
            svcCardBody = "    <div class='tk-info-box'><span class='tk-badge-ok'>OK</span> No stopped automatic services found.</div>";
        }
        else
        {
            var svcRows = new StringBuilder();
            foreach (var s in stoppedSvcs)
            {
                svcRows.Append($@"
        <tr>
          <td>{Esc(s.Name)}</td>
          <td>{Esc(s.DisplayName)}</td>
          <td><span class='tk-badge-warn'>Stopped</span></td>
          <td>{Esc(s.StartType)}</td>
        </tr>");
            }

            svcCardBody = $@"
    <table class=""tk-table"">
      <thead>
        <tr>
          <th>Name</th>
          <th>Display Name</th>
          <th>Status</th>
          <th>Start Type</th>
        </tr>
      </thead>
      <tbody>
{svcRows}
      </tbody>
    </table>";
        }

        string taskCardBody;
        if (failedTasks.Count == 0)
        {
            taskCardBody = "    <div class='tk-info-box'><span class='tk-badge-ok'>OK</span> No failed scheduled tasks found.</div>";
        }
        else
        {
            var taskRows = new StringBuilder();
            foreach (var t in failedTasks)
            {
                taskRows.Append($@"
        <tr>
          <td>{Esc(t.TaskName)}</td>
          <td>{Esc(t.TaskPath)}</td>
          <td>{Esc(t.LastRunTime)}</td>
          <td><span class='tk-badge-warn'>{Esc(t.LastResult)}</span></td>
        </tr>");
            }

            taskCardBody = $@"
    <table class=""tk-table"">
      <thead>
        <tr>
          <th>Task Name</th>
          <th>Path</th>
          <th>Last Run</th>
          <th>Last Result</th>
        </tr>
      </thead>
      <tbody>
{taskRows}
      </tbody>
    </table>";
        }

        sb.Append($@"
<div class=""tk-section"" id=""services-and-tasks"">
  <div class=""tk-section-title""><span class=""tk-section-num"">05</span> Services and Tasks</div>
  <div class=""tk-card"">
    <div class=""tk-card-header"">Stopped Automatic Services</div>
{svcCardBody}
  </div>
  <div class=""tk-card"">
    <div class=""tk-card-header"">Failed Scheduled Tasks</div>
{taskCardBody}
  </div>
</div>
");

        sb.Append(TkHtml.Foot("S.C.R.Y.E.R. v3.6"));
        return sb.ToString();
    }
}

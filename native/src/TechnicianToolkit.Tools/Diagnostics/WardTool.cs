using System.Globalization;
using System.Text;
using TechnicianToolkit.Core;
using TechnicianToolkit.Core.Config;
using TechnicianToolkit.Core.Html;
using TechnicianToolkit.Tools.Collectors;

namespace TechnicianToolkit.Tools.Diagnostics;

/// <summary>
/// W.A.R.D. — Watches Accounts, Reviews Roles &amp; Detects anomalies.
/// Native port of ward.ps1: audits local user accounts (status, last logon,
/// password config, admin role) and flags anomalies (no password, never set,
/// stale &gt;90 days, disabled).
/// </summary>
public sealed class WardTool : ITool
{
    private const int StaleDays = 90;

    public string Key => "11";
    public string Name => "W.A.R.D.";
    public string Title => "Account Audit";
    public string Category => "Diagnostics & Reporting";
    public string Description =>
        "Local user account & security audit: roles, last logon, password posture, anomaly flags.";
    public bool RequiresAdmin => true;
    public bool SupportsWhatIf => false;

    private static string Esc(string? s) => HtmlUtil.Esc(s);

    public ToolResult Run(ToolContext ctx)
    {
        try
        {
            ctx.Report(ProgressLevel.Step, "Enumerating Administrators group...");
            var admins = UserCollector.GetAdminMembers();

            ctx.Report(ProgressLevel.Step, "Collecting local user accounts...");
            var accounts = UserCollector.Collect(admins);
            ctx.Report(ProgressLevel.Ok, $"Collected {accounts.Count} accounts.");

            var staleDate = DateTime.Now.AddDays(-StaleDays);

            // Compute the anomaly-flag string per account (ward.ps1 flag logic).
            var rows = accounts
                .Select(a => new AuditRow(a, BuildFlags(a, staleDate)))
                .OrderByDescending(r => r.Account.IsAdmin) // admins first (WARD sort)
                .ToList();

            var total = rows.Count;
            var enabled = rows.Count(r => r.Account.Enabled);
            var disabled = rows.Count(r => !r.Account.Enabled);
            var adminCount = rows.Count(r => r.Account.IsAdmin);
            var flagged = rows.Count(r => !string.IsNullOrEmpty(r.Flags));

            ctx.Report(ProgressLevel.Step, "Building HTML report...");
            var reportDir = TkPaths.ResolveLogDirectory(TkPaths.BaseDirectory);
            Directory.CreateDirectory(reportDir);
            var reportFile = Path.Combine(reportDir, $"WARD_{DateTime.Now:yyyyMMdd_HHmmss}.html");

            var html = BuildHtml(rows, total, enabled, disabled, adminCount, flagged);
            File.WriteAllText(reportFile, html, new UTF8Encoding(false));
            ctx.Report(ProgressLevel.Ok, $"Report saved: {reportFile}");

            var summary = new List<KeyValuePair<string, string>>
            {
                new("Accounts", total.ToString()),
                new("Enabled", enabled.ToString()),
                new("Disabled", disabled.ToString()),
                new("Administrators", adminCount.ToString()),
                new("Flagged", flagged.ToString()),
            };

            return ToolResult.Ok(reportFile, summary);
        }
        catch (Exception ex)
        {
            TechnicianToolkit.Core.Diagnostics.TkErrorLog.Write("ward", ex.Message, "Run");
            ctx.Report(ProgressLevel.Fail, ex.Message);
            return ToolResult.Fail(ex.Message);
        }
    }

    private static string BuildFlags(UserAccount a, DateTime staleDate)
    {
        var flags = new List<string>();
        if (a.Enabled && !a.PasswordRequired)
        {
            flags.Add("No password required");
        }

        if (a.Enabled && a.PasswordLastSet is null)
        {
            flags.Add("Password never set");
        }

        if (a.Enabled && (a.LastLogon is null || a.LastLogon < staleDate))
        {
            flags.Add($"Stale (>{StaleDays} days)");
        }

        if (!a.Enabled)
        {
            flags.Add("Disabled");
        }

        return string.Join("; ", flags);
    }

    private sealed record AuditRow(UserAccount Account, string Flags);

    private static string BuildHtml(
        IReadOnlyList<AuditRow> rows,
        int total, int enabled, int disabled, int adminCount, int flagged)
    {
        var cfg = TkConfig.Get();
        var machine = EnvInfo.MachineName;
        var subtitle = string.IsNullOrWhiteSpace(cfg.OrgName) ? machine : $"{cfg.OrgName} -- {machine}";
        var generated = DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

        var adminClass = adminCount > 1 ? "warn" : "info";
        var flaggedClass = flagged > 0 ? "err" : "ok";

        var sb = new StringBuilder();
        sb.Append(TkHtml.Head(
            title: "Account Audit Report",
            scriptName: "W.A.R.D.",
            subtitle: subtitle,
            metaItems: new[]
            {
                new KeyValuePair<string, string>("Machine", machine),
                new KeyValuePair<string, string>("Generated", generated),
                new KeyValuePair<string, string>("Accounts", total.ToString()),
                new KeyValuePair<string, string>("Flagged", flagged.ToString()),
            },
            navItems: new[] { "Local User Accounts" }));

        sb.Append($@"<div class=""tk-summary-row"">
  <div class=""tk-summary-card info""><div class=""tk-summary-num"">{total}</div><div class=""tk-summary-lbl"">Total Accounts</div></div>
  <div class=""tk-summary-card ok""><div class=""tk-summary-num"">{enabled}</div><div class=""tk-summary-lbl"">Enabled</div></div>
  <div class=""tk-summary-card""><div class=""tk-summary-num"">{disabled}</div><div class=""tk-summary-lbl"">Disabled</div></div>
  <div class=""tk-summary-card {adminClass}""><div class=""tk-summary-num"">{adminCount}</div><div class=""tk-summary-lbl"">Administrators</div></div>
  <div class=""tk-summary-card {flaggedClass}""><div class=""tk-summary-num"">{flagged}</div><div class=""tk-summary-lbl"">Flagged</div></div>
</div>
");

        var body = new StringBuilder();
        foreach (var r in rows)
        {
            var a = r.Account;
            var statusBadge = a.Enabled
                ? "<span class='tk-badge-ok'>Enabled</span>"
                : "<span class='tk-badge-warn'>Disabled</span>";
            var roleBadge = a.IsAdmin
                ? "<span class='tk-badge-err'>Admin</span>"
                : "<span class='tk-badge-info'>Standard</span>";
            var lastLogon = a.LastLogon?.ToString("yyyy-MM-dd HH:mm") ?? "Never";
            var pwdSet = a.PasswordLastSet?.ToString("yyyy-MM-dd") ?? "Never";
            var pwdExpires = a.PasswordExpires?.ToString("yyyy-MM-dd") ?? "Never / No Expiry";
            var flagCell = string.IsNullOrEmpty(r.Flags)
                ? ""
                : $"<span class='tk-badge-warn'>{Esc(r.Flags)}</span>";

            body.Append($@"
        <tr>
          <td><strong>{Esc(a.Name)}</strong></td>
          <td>{Esc(a.FullName)}</td>
          <td>{statusBadge}</td>
          <td>{roleBadge}</td>
          <td>{Esc(lastLogon)}</td>
          <td>{Esc(pwdSet)}</td>
          <td>{Esc(pwdExpires)}</td>
          <td>{flagCell}</td>
        </tr>");
        }

        sb.Append($@"
<div class=""tk-section"" id=""local-user-accounts"">
  <div class=""tk-section-tag"">PART 1</div>
  <div class=""tk-section-title"">Local User Accounts</div>
  <div class=""tk-card"">
    <div class=""tk-table-wrap"">
    <table class=""tk-table"">
      <thead>
        <tr>
          <th>Username</th>
          <th>Full Name</th>
          <th>Status</th>
          <th>Role</th>
          <th>Last Logon</th>
          <th>Password Set</th>
          <th>Password Expires</th>
          <th>Flags</th>
        </tr>
      </thead>
      <tbody>
{body}
      </tbody>
    </table>
    </div>
    <div class=""tk-info-box""><div class=""tk-info-label"">Note</div>Stale threshold: {StaleDays} days without logon</div>
  </div>
</div>
");

        sb.Append(TkHtml.Foot("W.A.R.D. v3.6"));
        return sb.ToString();
    }
}

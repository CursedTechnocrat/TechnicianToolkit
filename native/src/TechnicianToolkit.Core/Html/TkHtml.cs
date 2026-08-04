using System.Globalization;
using System.Text;

namespace TechnicianToolkit.Core.Html;

/// <summary>
/// Report head/foot markup — a native port of <c>Get-TKHtmlHead</c> and
/// <c>Get-TKHtmlFoot</c> from TechnicianToolkit.psm1. The output is intended to
/// be byte-compatible with the PowerShell templates so a native report is
/// indistinguishable from a script-generated one.
/// </summary>
public static class TkHtml
{
    /// <summary>
    /// Returns the opening HTML, head, page-header and (optional) nav markup,
    /// through the opening <c>&lt;div class="tk-main"&gt;</c>.
    /// </summary>
    /// <param name="title">Report title shown as the page heading.</param>
    /// <param name="scriptName">Acronym label (e.g. "S.C.R.Y.E.R.").</param>
    /// <param name="subtitle">Optional subtitle / machine line beneath the title.</param>
    /// <param name="metaItems">Ordered label→value pairs for the metadata bar.</param>
    /// <param name="navItems">Section labels for the sticky nav bar.</param>
    public static string Head(
        string title = "Technician Toolkit Report",
        string scriptName = "T.K.",
        string subtitle = "",
        IEnumerable<KeyValuePair<string, string>>? metaItems = null,
        IEnumerable<string>? navItems = null)
    {
        string Esc(string? s) => HtmlUtil.Esc(s);

        var metaHtml = string.Empty;
        var metaList = metaItems?.ToList() ?? new List<KeyValuePair<string, string>>();
        if (metaList.Count > 0)
        {
            var parts = new StringBuilder();
            foreach (var kv in metaList)
            {
                parts.Append(
                    $"<div class='tk-meta-item'><div class='tk-meta-label'>{Esc(kv.Key)}</div>" +
                    $"<div class='tk-meta-value'>{Esc(kv.Value)}</div></div>");
            }

            metaHtml = $"<div class='tk-meta-bar'>{parts}</div>";
        }

        var subtitleHtml = string.IsNullOrEmpty(subtitle)
            ? string.Empty
            : $"<div class='tk-page-subtitle'>{Esc(subtitle)}</div>";

        var navHtml = string.Empty;
        var navList = navItems?.ToList() ?? new List<string>();
        if (navList.Count > 0)
        {
            var links = new StringBuilder();
            for (var i = 0; i < navList.Count; i++)
            {
                var n = (i + 1).ToString("D2", CultureInfo.InvariantCulture);
                links.Append($"<a href='#s{n}'><span class='tk-nav-num'>{n}</span> &middot; {Esc(navList[i])}</a>");
            }

            navHtml = $"<nav class='tk-nav'>{links}</nav>";
        }

        return
$@"<!DOCTYPE html>
<html lang=""en"">
<head>
<meta charset=""UTF-8""/>
<meta name=""viewport"" content=""width=device-width,initial-scale=1""/>
<title>{title}</title>
{TkHtmlTheme.Css}
</head>
<body>
<div class=""tk-page-header"">
  <div class=""tk-report-label"">{Esc(scriptName)} REPORT</div>
  <div class=""tk-page-title"">{Esc(title)}</div>
  {subtitleHtml}
  {metaHtml}
</div>
{navHtml}
<div class=""tk-main"">
";
    }

    /// <summary>
    /// Returns the closing report markup, a native port of <c>Get-TKHtmlFoot</c>.
    /// </summary>
    /// <param name="scriptName">Shown in the footer right (e.g. "S.C.R.Y.E.R. v3.6").</param>
    public static string Foot(string scriptName = "TechnicianToolkit")
    {
        var ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        var hostName = EnvInfo.MachineName;
        return
$@"</div>
<div class=""tk-footer"">
  <span>Generated {ts} on {hostName}</span>
  <span>{HtmlUtil.Esc(scriptName)}</span>
</div>
</body>
</html>
";
    }
}

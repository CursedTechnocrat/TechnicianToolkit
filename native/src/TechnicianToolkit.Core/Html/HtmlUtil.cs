using System.Globalization;

namespace TechnicianToolkit.Core.Html;

/// <summary>
/// String helpers that mirror the <c>EscHtml</c> and <c>Format-Bytes</c>
/// functions from TechnicianToolkit.psm1.
/// </summary>
public static class HtmlUtil
{
    /// <summary>
    /// HTML-escapes a string for safe use inside report templates.
    /// Faithful to the module's <c>EscHtml</c>: escapes &amp;, &lt;, &gt; and
    /// double-quote, in that order, and leaves single quotes untouched.
    /// A null/empty input returns an empty string.
    /// </summary>
    public static string Esc(string? s)
    {
        if (string.IsNullOrEmpty(s))
        {
            return string.Empty;
        }

        return s
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;");
    }

    private const long Kb = 1024L;
    private const long Mb = 1024L * 1024L;
    private const long Gb = 1024L * 1024L * 1024L;
    private const long Tb = 1024L * 1024L * 1024L * 1024L;

    /// <summary>
    /// Formats a byte count as a human-readable string (B/KB/MB/GB/TB), matching
    /// the module's <c>Format-Bytes</c> (two decimal places above 1 KB).
    /// </summary>
    public static string FormatBytes(long bytes)
    {
        if (bytes < Kb)
        {
            return $"{bytes} B";
        }

        if (bytes < Mb)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:N2} KB", (double)bytes / Kb);
        }

        if (bytes < Gb)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:N2} MB", (double)bytes / Mb);
        }

        if (bytes < Tb)
        {
            return string.Format(CultureInfo.InvariantCulture, "{0:N2} GB", (double)bytes / Gb);
        }

        return string.Format(CultureInfo.InvariantCulture, "{0:N2} TB", (double)bytes / Tb);
    }
}

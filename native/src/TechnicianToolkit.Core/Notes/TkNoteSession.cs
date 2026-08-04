using System.Globalization;
using System.Text;
using TechnicianToolkit.Core.Html;

namespace TechnicianToolkit.Core.Notes;

/// <summary>
/// A buffer of technician notes and the ticket-ready report export — the native
/// port of the module's <c>Add-TKNote</c> / <c>Get-TKNote</c> / <c>Clear-TKNote</c>
/// / <c>Export-TKNoteReport</c> functions.
/// </summary>
/// <remarks>
/// The PowerShell module kept a single module-scoped note buffer that lived for
/// one tool run. In a GUI where several tools run in one process, global state
/// would leak between runs, so notes are held per-instance here. Create one
/// session per tool run.
/// </remarks>
public sealed class TkNoteSession
{
    private readonly List<TkNote> _notes = new();

    /// <summary>Records a timestamped note. Mirrors <c>Add-TKNote</c>.</summary>
    public void Add(string text, NoteCategory category = NoteCategory.Info, string scriptName = "")
    {
        _notes.Add(new TkNote
        {
            Timestamp = DateTime.Now,
            Category = category,
            Script = scriptName,
            Text = text,
        });
    }

    /// <summary>All notes recorded this session, oldest first (<c>Get-TKNote</c>).</summary>
    public IReadOnlyList<TkNote> GetAll() => _notes.AsReadOnly();

    /// <summary>Discards all recorded notes (<c>Clear-TKNote</c>).</summary>
    public void Clear() => _notes.Clear();

    private static readonly IReadOnlyDictionary<NoteCategory, string> BadgeClass =
        new Dictionary<NoteCategory, string>
        {
            [NoteCategory.Info] = "tk-badge-info",
            [NoteCategory.Action] = "tk-badge-blue",
            [NoteCategory.Warning] = "tk-badge-warn",
            [NoteCategory.Issue] = "tk-badge-err",
            [NoteCategory.Resolution] = "tk-badge-ok",
        };

    /// <summary>
    /// Writes the session's notes to a ticket-ready HTML report, porting
    /// <c>Export-TKNoteReport</c> (including the plain-text paste block).
    /// Returns the path written.
    /// </summary>
    public string ExportReport(
        string path,
        string title = "Technician Notes",
        string scriptName = "T.K.",
        string ticket = "",
        string? technician = null,
        string summary = "")
    {
        technician ??= EnvInfo.UserName;
        var generated = DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        var hostName = EnvInfo.MachineName;

        var countAction = _notes.Count(n => n.Category == NoteCategory.Action);
        var countIssue = _notes.Count(n => n.Category == NoteCategory.Issue);
        var countResolution = _notes.Count(n => n.Category == NoteCategory.Resolution);

        var meta = new List<KeyValuePair<string, string>>
        {
            new("Generated", generated),
            new("Machine", hostName),
            new("Technician", technician),
        };
        if (!string.IsNullOrWhiteSpace(ticket))
        {
            meta.Add(new KeyValuePair<string, string>("Ticket", ticket));
        }

        var sb = new StringBuilder();
        sb.Append(TkHtml.Head(
            title: title,
            scriptName: scriptName,
            subtitle: hostName,
            metaItems: meta,
            navItems: new[] { "Notes", "Ticket Summary" }));

        sb.Append(
$@"<div class=""tk-summary-row"">
  <div class=""tk-summary-card info""><div class=""tk-summary-num"">{_notes.Count}</div><div class=""tk-summary-lbl"">Total Notes</div></div>
  <div class=""tk-summary-card""><div class=""tk-summary-num"">{countAction}</div><div class=""tk-summary-lbl"">Actions</div></div>
  <div class=""tk-summary-card err""><div class=""tk-summary-num"">{countIssue}</div><div class=""tk-summary-lbl"">Issues</div></div>
  <div class=""tk-summary-card ok""><div class=""tk-summary-num"">{countResolution}</div><div class=""tk-summary-lbl"">Resolutions</div></div>
</div>
");

        if (!string.IsNullOrWhiteSpace(summary))
        {
            sb.Append($"<div class=\"tk-info-box\"><div class=\"tk-info-label\">Summary</div>{HtmlUtil.Esc(summary)}</div>");
        }

        var rows = new StringBuilder();
        if (_notes.Count == 0)
        {
            rows.Append("<tr><td colspan='4'>No notes were recorded this session.</td></tr>");
        }
        else
        {
            foreach (var n in _notes)
            {
                var cls = BadgeClass.TryGetValue(n.Category, out var c) ? c : "tk-badge-info";
                var src = string.IsNullOrEmpty(n.Script) ? "&mdash;" : HtmlUtil.Esc(n.Script);
                rows.Append(
                    $"<tr><td class='tk-mono'>{n.Timestamp:HH:mm:ss}</td>" +
                    $"<td><span class='tk-badge {cls}'>{HtmlUtil.Esc(n.Category.ToString())}</span></td>" +
                    $"<td>{src}</td><td>{HtmlUtil.Esc(n.Text)}</td></tr>");
            }
        }

        sb.Append(
$@"<div class=""tk-section"" id=""s01"">
  <div class=""tk-section-title""><span class=""tk-section-num"">01</span> Notes</div>
  <div class=""tk-card"">
    <div class=""tk-table-wrap"">
    <table class=""tk-table""><thead><tr><th>Time</th><th>Category</th><th>Source</th><th>Note</th></tr></thead>
    <tbody>{rows}</tbody></table>
    </div>
  </div>
</div>
");

        // Plain-text block — copy/paste straight into a ticket comment field.
        var plain = new StringBuilder();
        plain.AppendLine($"TECHNICIAN NOTES - {title}");
        plain.AppendLine($"Machine    : {hostName}");
        plain.AppendLine($"Technician : {technician}");
        plain.AppendLine($"Generated  : {generated}");
        if (!string.IsNullOrWhiteSpace(ticket))
        {
            plain.AppendLine($"Ticket     : {ticket}");
        }

        if (!string.IsNullOrWhiteSpace(summary))
        {
            plain.AppendLine();
            plain.AppendLine($"Summary: {summary}");
        }

        plain.AppendLine();
        if (_notes.Count == 0)
        {
            plain.AppendLine("(no notes recorded)");
        }
        else
        {
            foreach (var n in _notes)
            {
                plain.AppendLine(string.Format(
                    CultureInfo.InvariantCulture,
                    "[{0}] [{1,-10}] {2}",
                    n.Timestamp.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
                    n.Category.ToString().ToUpperInvariant(),
                    n.Text));
            }
        }

        sb.Append(
$@"<div class=""tk-section"" id=""s02"">
  <div class=""tk-section-title""><span class=""tk-section-num"">02</span> Ticket Summary</div>
  <div class=""tk-section-subtitle"">Copy the block below straight into the ticket comment field.</div>
  <div class=""tk-card"">
    <pre class=""tk-mono"" style=""white-space:pre-wrap;display:block;padding:14px;line-height:1.5"">{HtmlUtil.Esc(plain.ToString())}</pre>
  </div>
</div>
");

        sb.Append(TkHtml.Foot(scriptName));

        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
        return path;
    }
}

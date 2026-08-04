using TechnicianToolkit.Core.Notes;
using Xunit;

namespace TechnicianToolkit.Core.Tests;

public class TkNoteSessionTests : IDisposable
{
    private readonly string _tempDir;

    public TkNoteSessionTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "tknotes_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // best effort
        }
    }

    [Fact]
    public void AddAndGet_AccumulatesInOrder()
    {
        var session = new TkNoteSession();
        session.Add("first", NoteCategory.Info);
        session.Add("second", NoteCategory.Action);

        var notes = session.GetAll();
        Assert.Equal(2, notes.Count);
        Assert.Equal("first", notes[0].Text);
        Assert.Equal("second", notes[1].Text);
    }

    [Fact]
    public void Clear_EmptiesBuffer()
    {
        var session = new TkNoteSession();
        session.Add("x");
        session.Clear();
        Assert.Empty(session.GetAll());
    }

    [Fact]
    public void ExportReport_WritesFileWithCountsAndPlainTextBlock()
    {
        var session = new TkNoteSession();
        session.Add("Renamed computer", NoteCategory.Action, "covenant");
        session.Add("Disk failing", NoteCategory.Issue, "augur");
        session.Add("Replaced disk", NoteCategory.Resolution, "augur");

        var path = Path.Combine(_tempDir, "notes.html");
        var written = session.ExportReport(path, title: "Session", ticket: "INC001");

        Assert.Equal(path, written);
        Assert.True(File.Exists(path));

        var html = File.ReadAllText(path);
        // Summary counts
        Assert.Contains(">3<", html);                 // total notes
        Assert.Contains("Resolutions", html);
        // Ticket appears in the metadata
        Assert.Contains("INC001", html);
        // Plain-text ticket block includes an uppercased category and the note.
        Assert.Contains("[ACTION", html);
        Assert.Contains("Renamed computer", html);
    }

    [Fact]
    public void ExportReport_EmptySession_StillProducesValidReport()
    {
        var session = new TkNoteSession();
        var path = Path.Combine(_tempDir, "empty.html");
        session.ExportReport(path);

        var html = File.ReadAllText(path);
        Assert.Contains("No notes were recorded this session.", html);
        Assert.Contains("</html>", html);
    }

    [Fact]
    public void ExportReport_EscapesNoteText()
    {
        var session = new TkNoteSession();
        session.Add("value <b> & stuff", NoteCategory.Info);
        var path = Path.Combine(_tempDir, "esc.html");
        session.ExportReport(path);

        var html = File.ReadAllText(path);
        Assert.Contains("&lt;b&gt;", html);
        Assert.Contains("&amp;", html);
    }
}

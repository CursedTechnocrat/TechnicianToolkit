namespace TechnicianToolkit.Core.Notes;

/// <summary>Category for a technician note (matches the module's ValidateSet).</summary>
public enum NoteCategory
{
    Info,
    Action,
    Warning,
    Issue,
    Resolution,
}

/// <summary>A single timestamped technician note.</summary>
public sealed class TkNote
{
    public DateTime Timestamp { get; init; } = DateTime.Now;
    public NoteCategory Category { get; init; } = NoteCategory.Info;
    public string Script { get; init; } = string.Empty;
    public string Text { get; init; } = string.Empty;
}

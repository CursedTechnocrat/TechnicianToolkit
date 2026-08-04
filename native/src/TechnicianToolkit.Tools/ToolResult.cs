namespace TechnicianToolkit.Tools;

/// <summary>
/// The outcome of a tool run: whether it succeeded, where the report landed, an
/// optional error message, and the key/value summary the console version printed
/// at the end (surfaced in the app's completion panel).
/// </summary>
public sealed class ToolResult
{
    public bool Success { get; init; }

    /// <summary>Path to the generated HTML report, if one was produced.</summary>
    public string? ReportPath { get; init; }

    /// <summary>Error message when <see cref="Success"/> is false.</summary>
    public string? Error { get; init; }

    /// <summary>End-of-run summary lines (label → value).</summary>
    public IReadOnlyList<KeyValuePair<string, string>> Summary { get; init; } =
        Array.Empty<KeyValuePair<string, string>>();

    public static ToolResult Ok(string reportPath, IReadOnlyList<KeyValuePair<string, string>> summary) =>
        new() { Success = true, ReportPath = reportPath, Summary = summary };

    public static ToolResult Fail(string error) =>
        new() { Success = false, Error = error };
}

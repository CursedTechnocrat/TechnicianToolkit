using TechnicianToolkit.Core.Notes;

namespace TechnicianToolkit.Tools;

/// <summary>Severity of a progress line — maps to the console color schema.</summary>
public enum ProgressLevel
{
    /// <summary>Indented detail line (Gray / Write-Info).</summary>
    Info,

    /// <summary>A step in progress (Magenta / Write-Step).</summary>
    Step,

    /// <summary>Success (Green / Write-Ok).</summary>
    Ok,

    /// <summary>Warning (Yellow / Write-Warn).</summary>
    Warn,

    /// <summary>Failure (Red / Write-Fail).</summary>
    Fail,
}

/// <summary>A single progress message emitted while a tool runs.</summary>
public readonly record struct ToolProgress(ProgressLevel Level, string Message);

/// <summary>
/// Inputs and side-channels for a tool run. Replaces the per-script switch
/// parameters (<c>-Unattended</c>, <c>-WhatIf</c>, <c>-OutputPath</c>) and gives
/// the tool a structured way to stream progress and record technician notes.
/// </summary>
public sealed class ToolContext
{
    /// <summary>Skip interactive prompts / run with defaults (<c>-Unattended</c>).</summary>
    public bool Unattended { get; init; } = true;

    /// <summary>Preview mode for state-changing tools (<c>-WhatIf</c>).</summary>
    public bool WhatIf { get; init; }

    /// <summary>Explicit report output directory; null uses the tool's default.</summary>
    public string? OutputPath { get; init; }

    /// <summary>Progress sink; the WPF app routes this to the run log pane.</summary>
    public IProgress<ToolProgress>? Progress { get; init; }

    /// <summary>Technician-note buffer for this run.</summary>
    public TkNoteSession Notes { get; init; } = new();

    /// <summary>Cancellation for long-running collectors.</summary>
    public CancellationToken CancellationToken { get; init; } = CancellationToken.None;

    /// <summary>Emits a progress line to the sink (no-op if none is attached).</summary>
    public void Report(ProgressLevel level, string message) =>
        Progress?.Report(new ToolProgress(level, message));
}

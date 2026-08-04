using TechnicianToolkit.Tools.Diagnostics;

namespace TechnicianToolkit.Tools;

/// <summary>
/// The set of tools reimplemented natively so far. This intentionally lists only
/// the ported tools rather than mirroring GRIMOIRE's full 41-entry registry —
/// the project guide warns that a duplicated tool list drifts out of sync, so
/// the canonical roster stays in grimoire.ps1 and the app surfaces the remainder
/// as an informational note.
/// </summary>
public static class ToolCatalog
{
    /// <summary>Total tools in the PowerShell suite (for the "N of M ported" note).</summary>
    public const int TotalSuiteToolCount = 41;

    /// <summary>All natively-ported, runnable tools.</summary>
    public static IReadOnlyList<ITool> Tools { get; } = new ITool[]
    {
        new ScryerTool(),
        new WardTool(),
        new AugurTool(),
    };

    /// <summary>Tools grouped by their GRIMOIRE category, category order preserved.</summary>
    public static IReadOnlyList<IGrouping<string, ITool>> ByCategory() =>
        Tools.GroupBy(t => t.Category).ToList();
}

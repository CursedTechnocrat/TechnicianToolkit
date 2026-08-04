namespace TechnicianToolkit.Tools;

/// <summary>
/// A single native toolkit tool. Metadata mirrors the fields GRIMOIRE keeps in
/// its <c>$Tools</c> registry (Key, Name, Category), plus the capability flags
/// the suite documents (<c>-WhatIf</c> support, admin requirement).
/// </summary>
public interface ITool
{
    /// <summary>Numeric key, matching the tool's GRIMOIRE registry key.</summary>
    string Key { get; }

    /// <summary>Acronym name, e.g. "S.C.R.Y.E.R.".</summary>
    string Name { get; }

    /// <summary>Short human title, e.g. "Unified Diagnostic Report".</summary>
    string Title { get; }

    /// <summary>GRIMOIRE category, e.g. "Diagnostics &amp; Reporting".</summary>
    string Category { get; }

    /// <summary>One-line description for the tool list.</summary>
    string Description { get; }

    /// <summary>True if the tool needs Administrator rights to produce full data.</summary>
    bool RequiresAdmin { get; }

    /// <summary>True if the tool honours <see cref="ToolContext.WhatIf"/>.</summary>
    bool SupportsWhatIf { get; }

    /// <summary>Runs the tool and returns its result (report path + summary).</summary>
    ToolResult Run(ToolContext context);
}

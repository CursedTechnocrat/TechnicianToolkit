# TechnicianToolkit — Native Edition

A **truly native Windows program** that reimplements the TechnicianToolkit in
C# / .NET, with a WPF desktop UI — no PowerShell required at runtime.

This is distinct from `../launcher/`, which merely *embeds and runs* the
PowerShell scripts inside a single `.exe`. The native edition doesn't shell out
to PowerShell at all: it talks to Windows directly (WMI/CIM, the Service Control
Manager, the Task Scheduler COM API, the WinNT ADSI provider) and renders the
same dark-themed HTML reports the scripts produce.

> **Status: foundation + 3 reference tools.** The shared module is fully ported,
> and 3 of the suite's 41 tools are reimplemented end-to-end (SCRYER, WARD,
> AUGUR) to establish the pattern the rest follow. See
> [Porting the remaining tools](#porting-the-remaining-tools).

## Layout

```
native/
├── TechnicianToolkit.Native.sln
├── Directory.Build.props            # shared version / lang settings
├── src/
│   ├── TechnicianToolkit.Core/      # port of TechnicianToolkit.psm1 (platform-neutral)
│   │   ├── Html/                    # EscHtml, Format-Bytes, CSS theme, Head/Foot
│   │   ├── Config/                  # Get-TKConfig / Set-TKConfig (+ Phantom→Revenant)
│   │   ├── Diagnostics/             # Write-TKError (JSONL + Teams webhook)
│   │   ├── Notes/                   # Add-TKNote / Export-TKNoteReport
│   │   └── Security/                # Test-IsAdmin / elevation relaunch
│   ├── TechnicianToolkit.Tools/     # native tools + Windows data collectors (Windows-only)
│   │   ├── Collectors/              # System / User / Volume / SMART / Service / Task
│   │   └── Diagnostics/             # ScryerTool, WardTool, AugurTool
│   └── TechnicianToolkit.App/       # WPF desktop UI
└── tests/
    └── TechnicianToolkit.Core.Tests/  # xUnit tests for the platform-neutral Core
```

## Project responsibilities

| Project | TFM | Runs on | Purpose |
|---------|-----|---------|---------|
| **Core** | `net8.0` | any OS | The shared module, ported. Pure logic — unit-tested in CI. Only the admin check is Windows-gated. |
| **Tools** | `net8.0-windows` | Windows | Tool logic + the Windows data collectors. Compiles anywhere as reference; must **run** on Windows. |
| **App** | `net8.0-windows` | Windows | The WPF program. Sidebar of tools, run-with-live-log, open-report. |
| **Core.Tests** | `net8.0` | any OS | xUnit tests over Core. |

## Building

Requires the **.NET SDK 8.0+** (<https://dotnet.microsoft.com/download>) on a
**Windows** machine (the App and Tools projects need the Windows Desktop
workload).

```powershell
# From this folder:
dotnet build TechnicianToolkit.Native.sln -c Release

# Run the desktop app:
dotnet run --project src/TechnicianToolkit.App -c Release

# Run the Core unit tests (works on any OS, including CI containers):
dotnet test tests/TechnicianToolkit.Core.Tests
```

To produce a distributable single-file `.exe` of the app:

```powershell
dotnet publish src/TechnicianToolkit.App -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o dist
```

## Using the app

1. Launch `TechnicianToolkit.App`. If it isn't elevated, a banner offers
   **Restart as Administrator** (the diagnostic tools need admin for complete
   data).
2. Pick a tool from the sidebar. The detail pane shows what it does and whether
   it needs admin.
3. Optionally set an output folder, then click **Run tool**. Progress streams
   into the log pane exactly like the console tools' `[*]/[+]/[!]` lines.
4. On completion, **Open report** launches the generated HTML in your browser.

## What's ported

| Tool | Source | What it does |
|------|--------|--------------|
| **S.C.R.Y.E.R.** | `scryer.ps1` | Consolidated report: system overview, users, disk space, SMART health, services & tasks. |
| **W.A.R.D.** | `ward.ps1` | Local user account & security audit with anomaly flags. |
| **A.U.G.U.R.** | `augur.ps1` | Physical disk health / SMART failure prediction + per-volume integrity. |

All three are read-only diagnostics — a safe first set to port because they
change no system state.

The shared module (`TechnicianToolkit.psm1`) is ported in full: HTML theme +
Head/Foot templating, `EscHtml`, `Format-Bytes`, config read/write (including
the legacy `Phantom`→`Revenant` migration), the JSONL error log + Teams webhook,
technician notes + the ticket-ready note report, and the elevation helpers.

## Known deviations from the PowerShell

These are deliberate and documented so the port's fidelity is auditable:

- **`UserAccount.PasswordExpires`** is computed as `PasswordLastSet +
  MaxPasswordAge` (machine policy) via the WinNT ADSI provider, rather than read
  from a dedicated `LocalUser` field. With no maximum-age policy it is null
  ("Never / No Expiry"), matching how WARD renders a null.
- **AUGUR's "Overall" metadata value** is surfaced as clean plain text (e.g.
  "2 critical issue(s)"). The original passed an HTML badge, but the shared head
  helper HTML-escapes metadata values, so the script actually rendered literal
  `<span>` tags there — the port shows the verdict correctly instead.
- **Notes are per-run**, held on a `TkNoteSession` instance rather than in a
  single module-global buffer, because the GUI can run several tools in one
  process and global state would leak between runs.
- **`config.json` location** defaults to next to the app executable (the module
  anchored it to `$PSScriptRoot`). Repoint via `TkConfig.ConfigPath`.

## Porting the remaining tools

The pattern is fixed by the three reference tools:

1. Add any new Windows data access as a collector under
   `TechnicianToolkit.Tools/Collectors/` (reuse the existing ones where the data
   overlaps — SCRYER, WARD and AUGUR already share the user/volume/SMART
   collectors).
2. Implement `ITool` under `Diagnostics/` (or a new category folder): fill in the
   metadata (mirror the tool's GRIMOIRE key/category), stream progress via
   `ctx.Report(...)`, and build the report with `TkHtml.Head/Foot` + the shared
   CSS classes so it matches the script output.
3. Register it in `ToolCatalog.Tools`.
4. Add Core-level tests where logic is platform-neutral.

State-changing tools (REVENANT, SIGIL, CLEANSE, …) additionally set
`SupportsWhatIf => true` and must honour `ctx.WhatIf` — the app already exposes a
**Preview only (-WhatIf)** toggle that lights up for those tools.

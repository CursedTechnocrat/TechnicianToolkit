# TechnicianToolkit — Portable Launcher

A self-contained, **single-file `.exe`** that carries the entire TechnicianToolkit
suite inside it. Drop it on a USB stick, run it on any Windows machine, and the
GRIMOIRE hub menu comes up with every tool available — **fully offline, with no
update checks and nothing downloaded from GitHub**.

This project does not modify any existing toolkit script. It *embeds* them
verbatim from the repository root at build time.

## Why this exists

Normally each toolkit script bootstraps itself by downloading
`TechnicianToolkit.psm1` (and any missing tool) from GitHub on first run. That is
great for staying current, but it assumes internet access and means the suite is
spread across many files. For a field technician working from a USB stick on an
isolated or air-gapped machine, you want the opposite: one file, everything
present, zero network dependency.

The launcher delivers that. Because every script is extracted locally before
GRIMOIRE starts, the "download if missing" paths in `grimoire.ps1` and in each
tool's module bootstrap never trigger.

## How it works

1. **Build time** — `TechnicianToolkit.Launcher.csproj` embeds every top-level
   `.ps1`, the shared module `TechnicianToolkit.psm1`, and `config.json` as
   resources inside the `.exe`. The .NET runtime is bundled too (self-contained
   single-file publish).
2. **Run time** — `Program.cs` extracts the embedded scripts to a working folder
   (next to the `.exe`, or `%TEMP%` if the stick is read-only), then launches
   `grimoire.ps1` with Windows PowerShell, passing through any arguments you gave
   (e.g. `-WhatIf`).

## Building

Requires the **.NET SDK 8.0+** on the build machine
(<https://dotnet.microsoft.com/download>). The target machines do **not** need
.NET installed — it is bundled into the `.exe`. They only need Windows
PowerShell, which ships with every Windows install.

```powershell
# From the repository root:
.\launcher\build.ps1

# Options:
.\launcher\build.ps1 -Runtime win-x64 -Output .\dist
.\launcher\build.ps1 -Runtime win-arm64
```

The resulting `dist\TechnicianToolkit.exe` is the whole suite in one file.

Equivalent raw command if you prefer not to use the script:

```powershell
dotnet publish .\launcher\TechnicianToolkit.Launcher.csproj `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -o .\dist
```

## Using it

```text
TechnicianToolkit.exe            # launches the GRIMOIRE hub menu
TechnicianToolkit.exe -WhatIf    # hub in dry-run mode (passed through to tools)
```

GRIMOIRE will request Administrator elevation just as it does normally.

## Notes & limitations

- **Windows only.** The toolkit scripts use Windows-specific APIs (WMI/CIM,
  registry, Defender, `netsh`, …), so the launcher targets Windows.
- **Runtime dependency:** the only thing the `.exe` needs on the target box is
  Windows PowerShell 5.1, which is part of Windows. .NET is bundled.
- **Single-file size** is on the order of tens of MB because it carries the .NET
  runtime. `EnableCompressionInSingleFile` keeps it as small as practical.
- **Refreshing the suite:** the scripts are baked in at build time. To ship newer
  tool versions, rebuild — that re-embeds the current repository scripts.

## Roadmap: native logic port

This launcher embeds and runs the existing PowerShell. A future phase can
reimplement individual tools' logic natively in C# (WMI/CIM, registry, Defender,
etc.) so the suite no longer depends on PowerShell at all. The launcher's
embed-and-extract structure is a clean place to introduce native tools
incrementally — a tool could run as native C# when available and fall back to its
embedded `.ps1` otherwise — without disturbing the existing scripts.

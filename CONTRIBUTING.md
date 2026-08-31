# Contributing to the Technician Toolkit

This toolkit exists because the same problems land on the same helpdesk queues every week.
It is meant to be more than one person's scripts — if you fix something at 2am on a client
site, that fix is worth more in here than in your local copy.

Contributions are welcome from technicians at every level. You do not need to be a PowerShell
expert. A bug report that says "AUGUR crashes on a machine with no physical disks" is a real
contribution.

## Licensing of contributions

The toolkit is licensed **GPL-3.0-or-later**. Contributions are accepted under that same
license — inbound matches outbound.

- **You keep the copyright in what you write.** There is no copyright assignment and no CLA.
- By opening a pull request you are licensing your contribution under GPL-3.0-or-later, so it
  can ship with the rest of the project.
- Only submit code you have the right to license. Do not paste in a client's proprietary
  scripts, vendor sample code with restrictive terms, or code carrying an incompatible
  license.

New files need the standard notice header. Copy it verbatim from any existing script and
change only the first two lines:

```powershell
# <filename> - <A.C.R.O.N.Y.M.> — <what it does>
# Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
#
# Copyright (C) 2026 CursedTechnocrat and the Technician Toolkit contributors
#
# This program is free software: you can redistribute it and/or modify
# ... (rest of the notice, unchanged)
#
# SPDX-License-Identifier: GPL-3.0-or-later
```

The Pester suite enforces the presence of this header, so a new tool without it will fail CI.

## Ways to contribute

| | |
|---|---|
| **Report a bug** | Open an issue. Include the tool name, the Windows build, and the error text. A screenshot of the console output is fine. |
| **Fix a bug** | Small, focused pull requests are easiest to review and land fastest. |
| **Add a tool** | Follow *Adding a New Tool* in [CLAUDE.md](CLAUDE.md) — header block, shared-module bootstrap, GRIMOIRE registration, README entry. |
| **Improve the docs** | The README is the front door for technicians who have never seen the toolkit. Clarity fixes are genuinely valuable. |
| **Report what broke in the field** | Even without a patch. Real-world failure modes are the hardest thing to find from a dev machine. |
| **Test the ARM64 build** | **Actively wanted.** The `win-arm64` binary is built and published but has never run on a real ARM device — there is none here to test on. If you have a Snapdragon X, Surface Pro, or any Windows-on-ARM machine, running it and saying what happened is one of the most useful things you can contribute. A report that it simply worked is as valuable as a bug. |

The toolkit is GPL-3.0 and collaborative by design — it was never meant to be one
person's. If you use it in the field, you know things about how it behaves that
cannot be discovered from a dev box, and that knowledge is welcome here.

## Ground rules for code

These are the conventions the existing suite follows; matching them keeps the toolkit
predictable for the person who has to run it under pressure.

- **PowerShell 5.1 compatible.** Windows PowerShell ships with every Windows install; that is
  the floor. Do not require PowerShell 7.
- **Every script is self-contained.** A technician must be able to drop one `.ps1` on a machine
  and have it work. Use the shared-module bootstrap block so the module is fetched if missing.
- **Interactive tools expose `-Unattended`.** Anything destructive or state-changing also
  exposes `-WhatIf` and must honour it.
- **Fail loudly, never silently.** `-ErrorAction Stop` on the module import; handle errors and
  tell the operator what happened.
- **Read-only tools stay read-only.** If a tool is documented as an audit, it must not write
  system state.
- **Keep the UTF-8 BOM.** Windows PowerShell 5.1 reads a BOM-less file as ANSI and mangles the
  console glyphs. The test suite enforces this.

## Before you open a pull request

```powershell
# Install Pester 5 if needed
Install-Module -Name Pester -MinimumVersion 5.0 -Force -SkipPublisherCheck

# Run the suite
Invoke-Pester -Path .\tests\TechnicianToolkit.Tests.ps1 -Output Detailed
```

CI also runs PSScriptAnalyzer. The tests run without Administrator rights and without
Windows-only APIs, so they pass on any machine.

Explain in the pull request what broke and how you hit it. "Fixes the null comparison in
GARGOYLE that threw when a service had no dependent services" tells a reviewer everything
they need.

## A note on scope

The toolkit deliberately covers tools a technician runs *at a machine*. If your contribution
is parameter-only and meant for a remote shell such as Kaseya VSA LiveConnect, it likely
belongs in the companion repository,
[TechnicianToolkit-LiveConnect](https://github.com/CursedTechnocrat/TechnicianToolkit-LiveConnect),
which is licensed the same way.

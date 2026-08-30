# Security Policy

The toolkit runs with Administrator rights on machines that are not the operator's
own, and several tools write reports containing material a client would consider
sensitive. That combination is worth taking seriously, so this is how to report a
problem and what is in scope.

## Reporting a vulnerability

**Use GitHub's private vulnerability reporting** — on the repository, go to
**Security → Report a vulnerability**. That opens a private advisory visible only
to the maintainers. It is the preferred route because it does not disclose the
issue publicly before there is a fix.

Please do **not** open a public issue for a security problem.

What helps most in a report:

| | |
|---|---|
| **Which tool** | The script name, and its `.NOTES Version` line |
| **Windows build** | `winver`, and whether Windows PowerShell 5.1 or PowerShell 7 |
| **What happens** | The behaviour, and what an attacker gains from it |
| **How to reproduce** | The command line you ran, including any parameters |

You will get an acknowledgement. This is a small project maintained alongside a
day job, so a fix may not be immediate — but you will be told what is happening
rather than left waiting, and you will be credited in the advisory and the
CHANGELOG unless you would rather not be.

## Supported versions

The most recent release is the only one that receives fixes. The toolkit is
distributed as scripts a technician copies onto a machine, so old copies circulate
indefinitely — if you are carrying a copy on a USB stick, refresh it rather than
assume it is patched.

## In scope

- **Privilege escalation** beyond what a tool documents — a read-only audit tool
  that writes system state, or any path that grants rights the operator did not
  already hold.
- **Credential handling.** Several tools touch secrets: `beacon.ps1` reads saved
  WLAN profiles including key material, `cipher.ps1` handles BitLocker recovery
  keys, and `covenant.ps1` accepts a local administrator password. Leaking any of
  those into a log, a transcript, an HTML report, or the console when it should
  not be there is a vulnerability.
- **Report contents.** HTML reports and CSVs are written to the configured log
  directory and routinely attached to tickets. Anything written there that a
  technician would not expect to be sharing counts.
- **The module bootstrap.** Every tool downloads `TechnicianToolkit.psm1` from
  GitHub over HTTPS when it is missing. Anything that could cause a tool to fetch,
  trust, or execute code from somewhere else is in scope.
- **Injection** through parameters — a computer name, path, or account name that
  escapes into a command it should not.
- **Webhook delivery.** `Write-TKError` can POST to a configured Teams webhook.
  Leaking the webhook URL, or sending more than the error to it, is in scope.

## Out of scope

- **Tools requiring Administrator.** The toolkit is administrative software. That
  it can change system state when run by an administrator is the purpose, not a
  flaw.
- **Findings that need an attacker who already has Administrator** on the machine,
  since that attacker can do anything the toolkit does without it.
- **PSScriptAnalyzer or scanner output on its own.** A rule hit is not a
  vulnerability without a path to exploit it. Two rules are known false positives
  in this repo — see `CLAUDE.md`.
- **SmartScreen or antivirus warnings** on downloaded scripts and, in future, the
  packaged application. Signed or not, a tool that unpacks scripts and runs them
  elevated is structurally what a dropper looks like.
- **Social engineering** of maintainers or contributors.

## A note on running it safely

The toolkit is GPL-3.0-or-later and carries no warranty — read sections 15 and 16
of the [LICENSE](LICENSE). Practical advice regardless:

- Get the scripts from this repository, not from a copy someone sent you.
- Reports land in the configured log directory and may contain client data. Treat
  that directory as sensitive, and clear it between sites.
- Run destructive tools with `-WhatIf` first. Every tool that changes state
  supports it.

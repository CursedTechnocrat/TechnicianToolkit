// IHostSink.cs - Everything the hosted engine pushes at a front end.
// Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
//
// Copyright (C) 2026 CursedTechnocrat and the Technician Toolkit contributors
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <https://www.gnu.org/licenses/>.
//
// SPDX-License-Identifier: GPL-3.0-or-later

using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Security;

namespace TechnicianToolkit.Engine
{
    /// <summary>
    /// The single seam between the PowerShell host and whatever is displaying it.
    /// The console harness implements this directly; the phase 02 window will
    /// implement it and marshal each call onto the dispatcher thread.
    ///
    /// Every one of the suite's 3,565 Write-Host calls and 222 prompts arrives
    /// through here. That is the whole reason no tool script needs rewriting to
    /// gain a GUI.
    /// </summary>
    public interface IHostSink
    {
        void Write(string text, ConsoleColor? foreground, ConsoleColor? background);

        /// <summary>Clear-Host reached the host.</summary>
        void Clear();

        void Progress(string activity, string status, int percentComplete, bool completed);

        string? ReadLine(string? prompt);

        /// <summary>
        /// Masked input. Returns a <see cref="SecureString"/> rather than a string
        /// so a password is never materialised in a managed string the GC may
        /// leave lying around — covenant.ps1 takes a [securestring] for exactly
        /// this reason.
        /// </summary>
        SecureString ReadLineAsSecureString(string? prompt);

        PSCredential? PromptForCredential(string caption, string message, string userName, string targetName);

        int PromptForChoice(string caption, string message, IReadOnlyList<string> choices, int defaultChoice);

        /// <summary>
        /// Raised when a script calls <c>exit</c>. The application must absorb
        /// this — it is a script ending, not the process shutting down.
        /// Assert-AdminPrivilege takes this path on its failure branch.
        /// </summary>
        void ScriptRequestedExit(int exitCode);
    }
}

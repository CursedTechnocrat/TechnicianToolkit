// SpikeHost.cs - Minimal PSHost implementation for the Phase 00 spike.
// Part of the Technician Toolkit - https://github.com/CursedTechnocrat/TechnicianToolkit
//
// Copyright (C) 2026 John Joseph Bejarana (CursedTechnocrat) and the Technician Toolkit contributors
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
using System.Collections.ObjectModel;
using System.Globalization;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Security;

namespace TechnicianToolkit.Spike
{
    /// <summary>
    /// Everything the host pushes at the UI. The window implements this and
    /// marshals each call onto the dispatcher thread.
    /// </summary>
    internal interface IHostSink
    {
        void Write(string text, ConsoleColor? foreground, ConsoleColor? background);
        void Clear();
        void Progress(string activity, string status, int percentComplete, bool completed);
        string? ReadLine(string? prompt);
        int PromptForChoice(string caption, string message, IReadOnlyList<string> choices, int defaultChoice);

        /// <summary>
        /// Raised when the script calls <c>exit</c>. The app must absorb this —
        /// it is a script exiting, not the application shutting down.
        /// </summary>
        void ScriptRequestedExit(int exitCode);
    }

    /// <summary>
    /// The raw UI surface. Console geometry is faked with sane values; what
    /// actually matters here is <see cref="SetBufferContents(Rectangle, BufferCell)"/>,
    /// which is how PowerShell 7's Clear-Host reaches the host.
    /// </summary>
    internal sealed class SpikeRawUi : PSHostRawUserInterface
    {
        private readonly IHostSink _sink;

        internal SpikeRawUi(IHostSink sink) => _sink = sink;

        public override ConsoleColor BackgroundColor { get; set; } = ConsoleColor.Black;
        public override ConsoleColor ForegroundColor { get; set; } = ConsoleColor.Gray;

        // 120x9999 is a deliberate choice: several toolkit scripts pad banners
        // and separator rules to the buffer width, so a realistic width keeps
        // their output looking the way it does in a console.
        public override Size BufferSize { get; set; } = new Size(120, 9999);
        public override Size WindowSize { get; set; } = new Size(120, 50);
        public override Coordinates CursorPosition { get; set; } = new Coordinates(0, 0);
        public override Coordinates WindowPosition { get; set; } = new Coordinates(0, 0);
        public override int CursorSize { get; set; } = 25;
        public override string WindowTitle { get; set; } = "Technician Toolkit (spike)";

        public override Size MaxPhysicalWindowSize => new Size(240, 100);
        public override Size MaxWindowSize => new Size(240, 100);

        // No keyboard is attached to a GUI host. Scripts that poll this — see
        // covenant.ps1 and restoration.ps1 — must be guarded on it.
        public override bool KeyAvailable => false;

        public override void FlushInputBuffer() { }

        public override KeyInfo ReadKey(ReadKeyOptions options) =>
            new KeyInfo(0, '\0', default(ControlKeyStates), keyDown: true);

        public override BufferCell[,] GetBufferContents(Rectangle rectangle) =>
            new BufferCell[0, 0];

        public override void ScrollBufferContents(
            Rectangle source, Coordinates destination, Rectangle clip, BufferCell fill) { }

        public override void SetBufferContents(Coordinates origin, BufferCell[,] contents) { }

        /// <summary>
        /// PowerShell 7 implements Clear-Host as a call to SetBufferContents
        /// over the whole buffer (Top/Bottom/Left/Right all -1). That is the
        /// signal to clear the output pane.
        /// </summary>
        public override void SetBufferContents(Rectangle rectangle, BufferCell fill)
        {
            bool wholeBuffer = rectangle.Left < 0 && rectangle.Top < 0
                            && rectangle.Right < 0 && rectangle.Bottom < 0;
            if (wholeBuffer)
            {
                _sink.Clear();
            }
        }
    }

    /// <summary>
    /// Routes every Write-Host, prompt, and stream message into the sink.
    /// </summary>
    internal sealed class SpikeHostUi : PSHostUserInterface
    {
        private readonly IHostSink _sink;
        private readonly SpikeRawUi _rawUi;

        internal SpikeHostUi(IHostSink sink)
        {
            _sink = sink;
            _rawUi = new SpikeRawUi(sink);
        }

        public override PSHostRawUserInterface RawUI => _rawUi;

        public override void Write(string value) =>
            _sink.Write(value, null, null);

        public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value) =>
            _sink.Write(value, foregroundColor, backgroundColor);

        public override void WriteLine(string value) =>
            _sink.Write(value + Environment.NewLine, null, null);

        public override void WriteErrorLine(string value) =>
            _sink.Write(value + Environment.NewLine, ConsoleColor.Red, null);

        public override void WriteWarningLine(string message) =>
            _sink.Write("WARNING: " + message + Environment.NewLine, ConsoleColor.Yellow, null);

        public override void WriteVerboseLine(string message) =>
            _sink.Write("VERBOSE: " + message + Environment.NewLine, ConsoleColor.Cyan, null);

        public override void WriteDebugLine(string message) =>
            _sink.Write("DEBUG: " + message + Environment.NewLine, ConsoleColor.DarkGray, null);

        public override void WriteProgress(long sourceId, ProgressRecord record) =>
            _sink.Progress(
                record.Activity ?? string.Empty,
                record.StatusDescription ?? string.Empty,
                record.PercentComplete,
                record.RecordType == ProgressRecordType.Completed);

        public override string ReadLine() =>
            _sink.ReadLine(null) ?? string.Empty;

        public override SecureString ReadLineAsSecureString()
        {
            var secure = new SecureString();
            foreach (char c in _sink.ReadLine("(masked input)") ?? string.Empty)
            {
                secure.AppendChar(c);
            }
            secure.MakeReadOnly();
            return secure;
        }

        public override Dictionary<string, PSObject> Prompt(
            string caption, string message, Collection<FieldDescription> descriptions)
        {
            var results = new Dictionary<string, PSObject>();
            foreach (FieldDescription field in descriptions)
            {
                string label = string.IsNullOrEmpty(field.Label) ? field.Name : field.Label;
                string? answer = _sink.ReadLine(label);
                results[field.Name] = PSObject.AsPSObject(answer ?? string.Empty);
            }
            return results;
        }

        public override int PromptForChoice(
            string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice)
        {
            var labels = new List<string>(choices.Count);
            foreach (ChoiceDescription choice in choices)
            {
                labels.Add(choice.Label);
            }
            return _sink.PromptForChoice(caption, message, labels, defaultChoice);
        }

        public override PSCredential PromptForCredential(
            string caption, string message, string userName, string targetName) =>
            PromptForCredential(caption, message, userName, targetName,
                PSCredentialTypes.Default, PSCredentialUIOptions.Default);

        public override PSCredential PromptForCredential(
            string caption, string message, string userName, string targetName,
            PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options)
        {
            // Out of scope for the spike — the real app gets a proper dialog.
            string user = string.IsNullOrEmpty(userName)
                ? _sink.ReadLine("Username") ?? string.Empty
                : userName;
            return new PSCredential(
                string.IsNullOrEmpty(user) ? "unknown" : user,
                ReadLineAsSecureString());
        }
    }

    /// <summary>
    /// The host itself. Note <see cref="SetShouldExit"/>: toolkit scripts call
    /// <c>exit</c> on their failure paths (Assert-AdminPrivilege does), and in a
    /// hosted runspace that arrives here rather than ending the process.
    /// </summary>
    internal sealed class SpikeHost : PSHost
    {
        private readonly IHostSink _sink;
        private readonly SpikeHostUi _ui;
        private readonly Guid _instanceId = Guid.NewGuid();

        internal SpikeHost(IHostSink sink)
        {
            _sink = sink;
            _ui = new SpikeHostUi(sink);
        }

        public override string Name => "TechnicianToolkit.Spike";
        public override Version Version => new Version(0, 1);
        public override Guid InstanceId => _instanceId;
        public override PSHostUserInterface UI => _ui;
        public override CultureInfo CurrentCulture => CultureInfo.CurrentCulture;
        public override CultureInfo CurrentUICulture => CultureInfo.CurrentUICulture;

        public override void EnterNestedPrompt() =>
            throw new NotSupportedException("Nested prompts are not supported by this host.");

        public override void ExitNestedPrompt() =>
            throw new NotSupportedException("Nested prompts are not supported by this host.");

        public override void NotifyBeginApplication() { }
        public override void NotifyEndApplication() { }

        public override void SetShouldExit(int exitCode) => _sink.ScriptRequestedExit(exitCode);
    }
}

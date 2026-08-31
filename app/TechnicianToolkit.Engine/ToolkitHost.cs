// ToolkitHost.cs - PSHost implementation carrying toolkit console output into any front end.
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

namespace TechnicianToolkit.Engine
{
    /// <summary>
    /// The raw UI surface. Console geometry is reported with realistic values
    /// because several tools pad banners and separator rules to the buffer
    /// width; what actually carries behaviour here is SetBufferContents.
    /// </summary>
    internal sealed class ToolkitRawUi : PSHostRawUserInterface
    {
        private readonly IHostSink _sink;

        internal ToolkitRawUi(IHostSink sink) => _sink = sink;

        public override ConsoleColor BackgroundColor { get; set; } = ConsoleColor.Black;
        public override ConsoleColor ForegroundColor { get; set; } = ConsoleColor.Gray;

        public override Size BufferSize { get; set; } = new Size(120, 9999);
        public override Size WindowSize { get; set; } = new Size(120, 50);
        public override Coordinates CursorPosition { get; set; } = new Coordinates(0, 0);
        public override Coordinates WindowPosition { get; set; } = new Coordinates(0, 0);
        public override int CursorSize { get; set; } = 25;
        public override string WindowTitle { get; set; } = "Technician Toolkit";

        public override Size MaxPhysicalWindowSize => new Size(240, 100);
        public override Size MaxWindowSize => new Size(240, 100);

        /// <summary>
        /// No keyboard is attached to a hosted runspace. covenant.ps1 and
        /// restoration.ps1 poll [Console]::KeyAvailable directly rather than
        /// this property, which is why they carry their own probe.
        /// </summary>
        public override bool KeyAvailable => false;

        public override void FlushInputBuffer() { }

        public override KeyInfo ReadKey(ReadKeyOptions options) =>
            new KeyInfo(0, (char)0, default(ControlKeyStates), keyDown: true);

        public override BufferCell[,] GetBufferContents(Rectangle rectangle) => new BufferCell[0, 0];

        public override void ScrollBufferContents(
            Rectangle source, Coordinates destination, Rectangle clip, BufferCell fill) { }

        public override void SetBufferContents(Coordinates origin, BufferCell[,] contents) { }

        /// <summary>
        /// PowerShell 7 has no RawUI.Clear(). Clear-Host is implemented as a
        /// SetBufferContents over the whole buffer, with every edge set to -1 --
        /// that rectangle is the signal, and it is how all 37 Clear-Host sites
        /// in the suite arrive.
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
    /// Routes Write-Host, the prompts, and the non-success streams into the sink.
    /// </summary>
    internal sealed class ToolkitHostUi : PSHostUserInterface
    {
        private readonly IHostSink _sink;
        private readonly ToolkitRawUi _rawUi;

        internal ToolkitHostUi(IHostSink sink)
        {
            _sink = sink;
            _rawUi = new ToolkitRawUi(sink);
        }

        public override PSHostRawUserInterface RawUI => _rawUi;

        public override void Write(string value) => _sink.Write(value, null, null);

        public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value) =>
            _sink.Write(value, foregroundColor, backgroundColor);

        public override void WriteLine(string value) => _sink.Write(value + Environment.NewLine, null, null);

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

        public override string ReadLine() => _sink.ReadLine(null) ?? string.Empty;

        public override SecureString ReadLineAsSecureString() => _sink.ReadLineAsSecureString(null);

        public override Dictionary<string, PSObject> Prompt(
            string caption, string message, Collection<FieldDescription> descriptions)
        {
            var results = new Dictionary<string, PSObject>();
            foreach (FieldDescription field in descriptions)
            {
                string label = string.IsNullOrEmpty(field.Label) ? field.Name : field.Label;

                // A securestring field must never round-trip through a managed string.
                bool isSecure = field.ParameterTypeFullName != null
                    && field.ParameterTypeFullName.IndexOf("SecureString", StringComparison.Ordinal) >= 0;

                results[field.Name] = isSecure
                    ? PSObject.AsPSObject(_sink.ReadLineAsSecureString(label))
                    : PSObject.AsPSObject(_sink.ReadLine(label) ?? string.Empty);
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
            PSCredential? credential = _sink.PromptForCredential(caption, message, userName, targetName);
            if (credential != null)
            {
                return credential;
            }

            // A cancelled credential prompt has no null-shaped answer in this API,
            // so hand back an empty one and let the cmdlet own validation fail.
            return new PSCredential(
                string.IsNullOrEmpty(userName) ? "unknown" : userName,
                new SecureString());
        }
    }

    /// <summary>
    /// The host itself.
    /// </summary>
    public sealed class ToolkitHost : PSHost
    {
        private readonly IHostSink _sink;
        private readonly ToolkitHostUi _ui;
        private readonly Guid _instanceId = Guid.NewGuid();

        public ToolkitHost(IHostSink sink)
        {
            _sink = sink;
            _ui = new ToolkitHostUi(sink);
        }

        public override string Name => "TechnicianToolkit";
        public override Version Version => new Version(5, 0);
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

        /// <summary>
        /// Toolkit scripts call exit on their failure paths - Assert-AdminPrivilege
        /// does. In a hosted runspace that arrives here instead of ending the
        /// process, and the application absorbs it.
        /// </summary>
        public override void SetShouldExit(int exitCode) => _sink.ScriptRequestedExit(exitCode);
    }
}

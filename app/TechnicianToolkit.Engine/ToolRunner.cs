// ToolRunner.cs - Runs one toolkit script in a hosted runspace, streaming and cancellable.
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
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace TechnicianToolkit.Engine
{
    /// <summary>What a completed run leaves behind.</summary>
    public sealed class ToolRunResult
    {
        public bool Cancelled { get; init; }

        /// <summary>The code a script passed to exit, if it called it at all.</summary>
        public int? ExitCode { get; init; }

        public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

        public bool Succeeded => !Cancelled && Errors.Count == 0 && (ExitCode ?? 0) == 0;
    }

    /// <summary>
    /// Runs a tool in a runspace bound to <see cref="ToolkitHost"/>.
    ///
    /// A fresh runspace per run is deliberate: the tools set script-scoped state
    /// and import modules, and one tool must never inherit another leftovers.
    /// </summary>
    public sealed class ToolRunner
    {
        private readonly IHostSink _sink;
        private readonly string _workDir;

        public ToolRunner(IHostSink sink, string workDir)
        {
            _sink = sink;
            _workDir = workDir;
        }

        /// <summary>
        /// Run a script by path, streaming everything through the host as it
        /// happens. Cancelling the token calls PowerShell.Stop(), which is the
        /// real reason the engine is hosted rather than shelled out to.
        /// </summary>
        public async Task<ToolRunResult> RunAsync(
            string scriptPath,
            IDictionary<string, object?> parameters,
            CancellationToken cancellationToken = default)
        {
            if (!File.Exists(scriptPath))
            {
                throw new FileNotFoundException("Tool script not found.", scriptPath);
            }

            var recorder = new ExitCodeRecordingSink(_sink);
            var errors = new List<string>();

            var initialState = InitialSessionState.CreateDefault2();

            // The extracted scripts are unsigned files on disk. Without this the
            // machine execution policy decides whether the app works at all,
            // which is not a decision a portable field tool can leave to chance.
            initialState.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;

            using Runspace runspace = RunspaceFactory.CreateRunspace(new ToolkitHost(recorder), initialState);
            runspace.Open();
            runspace.SessionStateProxy.Path.SetLocation(_workDir);

            using var shell = PowerShell.Create();
            shell.Runspace = runspace;

            // By path, never from a string: the tool bootstrap blocks depend on
            // $PSScriptRoot and $PSCommandPath resolving to a real location.
            shell.AddCommand(scriptPath, useLocalScope: false);
            foreach (KeyValuePair<string, object?> parameter in parameters)
            {
                if (parameter.Value is bool flag)
                {
                    // A switch is passed as a SwitchParameter, and omitted entirely
                    // when false so the script sees its own default.
                    if (flag)
                    {
                        shell.AddParameter(parameter.Key, new SwitchParameter(true));
                    }
                    continue;
                }

                shell.AddParameter(parameter.Key, parameter.Value);
            }

            // Anything reaching the success stream without Write-Host still has to
            // be visible, formatted the way a console would render it.
            shell.AddCommand("Out-String").AddParameter("Stream", true);

            var output = new PSDataCollection<PSObject>();
            output.DataAdded += (_, _) =>
            {
                foreach (PSObject item in output.ReadAll())
                {
                    _sink.Write(item?.ToString() + Environment.NewLine, null, null);
                }
            };

            shell.Streams.Error.DataAdded += (_, _) =>
            {
                foreach (ErrorRecord record in shell.Streams.Error.ReadAll())
                {
                    errors.Add(record.ToString());
                    _sink.Write("ERROR: " + record + Environment.NewLine, ConsoleColor.Red, null);
                }
            };

            // Warning, verbose, debug and information already reach the host UI
            // through Write*Line, so wiring them here too would double them up.

            bool cancelled = false;
            using CancellationTokenRegistration registration = cancellationToken.Register(() =>
            {
                cancelled = true;
                try
                {
                    shell.Stop();
                }
                catch
                {
                    // Stopping a pipeline that already finished is not an error.
                }
            });

            try
            {
                await Task.Factory.FromAsync(
                    shell.BeginInvoke<PSObject, PSObject>(null, output),
                    shell.EndInvoke).ConfigureAwait(false);
            }
            catch (PipelineStoppedException)
            {
                cancelled = true;
            }
            catch (Exception ex)
            {
                errors.Add(ex.Message);
                _sink.Write("EXCEPTION: " + ex.Message + Environment.NewLine, ConsoleColor.Red, null);
            }

            return new ToolRunResult
            {
                Cancelled = cancelled,
                ExitCode = recorder.ExitCode,
                Errors = errors,
            };
        }

        /// <summary>
        /// Wraps the real sink so the runner learns the code a script passed to
        /// exit, while the front end still sees the call.
        /// </summary>
        private sealed class ExitCodeRecordingSink : IHostSink
        {
            private readonly IHostSink _inner;

            internal ExitCodeRecordingSink(IHostSink inner) => _inner = inner;

            internal int? ExitCode { get; private set; }

            public void Write(string text, ConsoleColor? foreground, ConsoleColor? background) =>
                _inner.Write(text, foreground, background);

            public void Clear() => _inner.Clear();

            public void Progress(string activity, string status, int percentComplete, bool completed) =>
                _inner.Progress(activity, status, percentComplete, completed);

            public string? ReadLine(string? prompt) => _inner.ReadLine(prompt);

            public SecureString ReadLineAsSecureString(string? prompt) => _inner.ReadLineAsSecureString(prompt);

            public PSCredential? PromptForCredential(string caption, string message, string userName, string targetName) =>
                _inner.PromptForCredential(caption, message, userName, targetName);

            public int PromptForChoice(string caption, string message, IReadOnlyList<string> choices, int defaultChoice) =>
                _inner.PromptForChoice(caption, message, choices, defaultChoice);

            public void ScriptRequestedExit(int exitCode)
            {
                ExitCode = exitCode;
                _inner.ScriptRequestedExit(exitCode);
            }
        }
    }
}

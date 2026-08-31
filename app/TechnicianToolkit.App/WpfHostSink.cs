// WpfHostSink.cs - Carries the hosted engine's output onto the UI thread.
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
using System.Management.Automation;
using System.Security;
using System.Text;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;
using TechnicianToolkit.Engine;

namespace TechnicianToolkit.App
{
    /// <summary>Where a Write-Progress record lands.</summary>
    public sealed class ProgressState
    {
        public string Activity { get; init; } = string.Empty;
        public string Status { get; init; } = string.Empty;
        public int PercentComplete { get; init; }
        public bool Completed { get; init; }
    }

    /// <summary>
    /// The window's implementation of the engine sink.
    ///
    /// Every call arrives on the runspace thread, so everything here marshals.
    /// Writes are batched rather than marshalled one at a time: a tool like
    /// LANTERN sweeping a subnet emits thousands of lines in a burst, and one
    /// dispatcher operation per Write-Host would starve the UI thread of the
    /// time it needs to actually render them.
    /// </summary>
    public sealed class WpfHostSink : IHostSink
    {
        /// <summary>
        /// Past this many lines the oldest are dropped. A long LANTERN or
        /// THRESHOLD run is otherwise unbounded, and nobody scrolls back
        /// twenty thousand lines -- that is what saving the output is for.
        /// </summary>
        private const int MaxLines = 20000;

        private readonly Dispatcher _dispatcher;
        private readonly ObservableCollection<OutputLine> _lines;
        private readonly Action<ProgressState> _onProgress;
        private readonly Func<Window?> _owner;

        private readonly object _gate = new object();
        private readonly List<PendingWrite> _pending = new List<PendingWrite>();
        private bool _flushScheduled;

        // Touched only on the UI thread.
        private OutputLine? _current;
        private OutputLine? _notice;
        private int _dropped;

        public WpfHostSink(
            Dispatcher dispatcher,
            ObservableCollection<OutputLine> lines,
            Action<ProgressState> onProgress,
            Func<Window?> owner)
        {
            _dispatcher = dispatcher;
            _lines = lines;
            _onProgress = onProgress;
            _owner = owner;
        }

        /// <summary>Set when a script calls exit.</summary>
        public int? LastExitCode { get; private set; }

        private readonly struct PendingWrite
        {
            internal PendingWrite(string text, ConsoleColor? foreground)
            {
                Text = text;
                Foreground = foreground;
            }

            internal string Text { get; }
            internal ConsoleColor? Foreground { get; }
        }

        public void Write(string text, ConsoleColor? foreground, ConsoleColor? background)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            lock (_gate)
            {
                _pending.Add(new PendingWrite(text, foreground));

                if (_flushScheduled)
                {
                    return;
                }

                _flushScheduled = true;
            }

            _dispatcher.InvokeAsync(Flush, DispatcherPriority.Background);
        }

        /// <summary>
        /// Apply everything buffered. Safe to call from the UI thread at any
        /// time; the run calls it once more when it finishes so the last few
        /// lines are never left sitting in the buffer.
        /// </summary>
        public void Flush()
        {
            List<PendingWrite> batch;
            lock (_gate)
            {
                _flushScheduled = false;
                if (_pending.Count == 0)
                {
                    return;
                }

                batch = new List<PendingWrite>(_pending);
                _pending.Clear();
            }

            foreach (PendingWrite write in batch)
            {
                Emit(write.Text, ConsolePalette.For(write.Foreground));
            }

            Trim();
        }

        /// <summary>
        /// Enforce the cap, but say so. Dropping the oldest lines silently is fine
        /// for one chatty tool and wrong for a queue of them: the first tool's
        /// output would vanish, and Save would then write an incomplete log with
        /// nothing to indicate it. The notice line keeps the loss visible and
        /// keeps saved output honest.
        /// </summary>
        private void Trim()
        {
            if (_lines.Count <= MaxLines)
            {
                return;
            }

            while (_lines.Count > MaxLines)
            {
                // Never drop the notice itself, or the count would reset to zero.
                _lines.RemoveAt(_dropped > 0 ? 1 : 0);
                _dropped++;
            }

            if (_notice == null)
            {
                _notice = new OutputLine();
                _lines.Insert(0, _notice);
            }

            _notice.Reset();
            _notice.Append(
                "-- " + _dropped + " earlier line(s) dropped; the buffer holds the last "
                + MaxLines + " --" + Environment.NewLine,
                ConsolePalette.For(ConsoleColor.Yellow));
        }

        /// <summary>
        /// Split a chunk into lines. CRLF is normalised first so that a lone CR
        /// can keep its console meaning: return to column zero and redraw, which
        /// is how the reboot countdowns animate.
        /// </summary>
        private void Emit(string text, Brush brush)
        {
            text = text.Replace("\r\n", "\n");

            var buffer = new StringBuilder();

            foreach (char c in text)
            {
                switch (c)
                {
                    case '\n':
                        Commit(buffer, brush);
                        EnsureCurrent();
                        _current = null;
                        break;

                    case '\r':
                        Commit(buffer, brush);
                        EnsureCurrent().Reset();
                        break;

                    default:
                        buffer.Append(c);
                        break;
                }
            }

            Commit(buffer, brush);
        }

        private void Commit(StringBuilder buffer, Brush brush)
        {
            if (buffer.Length == 0)
            {
                return;
            }

            EnsureCurrent().Append(buffer.ToString(), brush);
            buffer.Clear();
        }

        private OutputLine EnsureCurrent()
        {
            if (_current != null)
            {
                return _current;
            }

            _current = new OutputLine();
            _lines.Add(_current);
            return _current;
        }

        public void Clear() => _dispatcher.Invoke(() =>
        {
            Flush();
            _lines.Clear();
            _current = null;
            _notice = null;
            _dropped = 0;
        });

        public void Progress(string activity, string status, int percentComplete, bool completed) =>
            _dispatcher.InvokeAsync(() => _onProgress(new ProgressState
            {
                Activity = activity,
                Status = status,
                PercentComplete = percentComplete,
                Completed = completed,
            }));

        // The prompts block the runspace thread until the dialog closes, which
        // is exactly what Read-Host does at a console.
        public string? ReadLine(string? prompt) =>
            _dispatcher.Invoke(() =>
            {
                Flush();
                return PromptWindow.AskLine(_owner(), prompt);
            });

        public SecureString ReadLineAsSecureString(string? prompt) =>
            _dispatcher.Invoke(() =>
            {
                Flush();
                return PromptWindow.AskSecure(_owner(), prompt);
            });

        public PSCredential? PromptForCredential(string caption, string message, string userName, string targetName) =>
            _dispatcher.Invoke(() =>
            {
                Flush();
                return PromptWindow.AskCredential(_owner(), caption, message, userName, targetName);
            });

        public int PromptForChoice(string caption, string message, IReadOnlyList<string> choices, int defaultChoice) =>
            _dispatcher.Invoke(() =>
            {
                Flush();
                return PromptWindow.AskChoice(_owner(), caption, message, choices, defaultChoice);
            });

        /// <summary>
        /// Absorbed. A script calling exit is a script ending, not the
        /// application shutting down.
        /// </summary>
        public void ScriptRequestedExit(int exitCode) => LastExitCode = exitCode;
    }
}

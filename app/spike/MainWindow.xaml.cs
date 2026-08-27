// MainWindow.xaml.cs - Spike window: output pane, prompts, and the two probes.
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
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace TechnicianToolkit.Spike
{
    public partial class MainWindow : Window, IHostSink
    {
        private readonly Paragraph _paragraph;

        public MainWindow()
        {
            InitializeComponent();
            _paragraph = new Paragraph { Margin = new Thickness(0) };
            OutputBox.Document.Blocks.Clear();
            OutputBox.Document.Blocks.Add(_paragraph);
        }

        // ── IHostSink ────────────────────────────────────────────────────────
        // Every member marshals to the dispatcher: the pipeline runs on a
        // background thread, and WPF elements are thread-affine.

        public void Write(string text, ConsoleColor? foreground, ConsoleColor? background)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            Dispatcher.Invoke(() =>
            {
                var run = new Run(text);
                if (foreground.HasValue)
                {
                    run.Foreground = new SolidColorBrush(MapConsoleColor(foreground.Value));
                }
                if (background.HasValue && background.Value != ConsoleColor.Black)
                {
                    run.Background = new SolidColorBrush(MapConsoleColor(background.Value));
                }
                _paragraph.Inlines.Add(run);
                OutputScroller.ScrollToEnd();
            });
        }

        public void Clear() => Dispatcher.Invoke(() => _paragraph.Inlines.Clear());

        public void Progress(string activity, string status, int percentComplete, bool completed) =>
            Dispatcher.Invoke(() =>
                ProgressText.Text = completed
                    ? string.Empty
                    : $"{activity} — {status} ({Math.Max(percentComplete, 0)}%)");

        public string? ReadLine(string? prompt)
        {
            // A real dialog is phase 02 work. The spike only has to prove the
            // call arrives on the UI thread and the pipeline waits for it.
            return Dispatcher.Invoke(() =>
            {
                MessageBox.Show(this,
                    $"The script asked for input{(prompt is null ? string.Empty : $": {prompt}")}.\n\n" +
                    "The spike answers with an empty string.",
                    "Read-Host reached the host", MessageBoxButton.OK, MessageBoxImage.Information);
                return string.Empty;
            });
        }

        public int PromptForChoice(
            string caption, string message, IReadOnlyList<string> choices, int defaultChoice) =>
            Dispatcher.Invoke(() =>
            {
                MessageBox.Show(this,
                    $"{caption}\n{message}\n\nChoices: {string.Join(", ", choices)}\n\n" +
                    "The spike takes the default.",
                    "PromptForChoice reached the host", MessageBoxButton.OK, MessageBoxImage.Question);
                return defaultChoice < 0 ? 0 : defaultChoice;
            });

        public void ScriptRequestedExit(int exitCode) =>
            Write($"[host] script called exit {exitCode} — absorbed, app still running{Environment.NewLine}",
                  ConsoleColor.Magenta, null);

        // ── Buttons ──────────────────────────────────────────────────────────

        private async void OnCheckRuntime(object sender, RoutedEventArgs e)
        {
            SetBusy(true, "Loading the PowerShell engine…");
            Write($"── Runtime check ──{Environment.NewLine}", ConsoleColor.Cyan, null);

            try
            {
                string report = await Task.Run(() => new ToolRunner(this).RuntimeReport());
                Write(report + Environment.NewLine + Environment.NewLine, ConsoleColor.Green, null);
                SetBusy(false, "Runtime check complete");
            }
            catch (Exception ex)
            {
                Write($"FAILED: {ex.GetType().Name}: {ex.Message}{Environment.NewLine}", ConsoleColor.Red, null);
                SetBusy(false, "Runtime check FAILED");
            }
        }

        private async void OnRunWard(object sender, RoutedEventArgs e)
        {
            SetBusy(true, "Running WARD…");
            Write($"── WARD ──{Environment.NewLine}", ConsoleColor.Cyan, null);

            try
            {
                IReadOnlyList<string> errors = await Task.Run(() =>
                {
                    var runner = new ToolRunner(this);
                    string workDir = ToolRunner.ExtractScripts();
                    string ward = Path.Combine(workDir, "ward.ps1");
                    return runner.RunScript(ward, new Dictionary<string, object>
                    {
                        ["Unattended"] = true
                    });
                });

                Write(Environment.NewLine +
                      $"[host] WARD finished with {errors.Count} error record(s).{Environment.NewLine}",
                      ConsoleColor.Magenta, null);
                SetBusy(false, errors.Count == 0 ? "WARD completed" : $"WARD completed with {errors.Count} error(s)");
            }
            catch (Exception ex)
            {
                Write($"FAILED: {ex.GetType().Name}: {ex.Message}{Environment.NewLine}", ConsoleColor.Red, null);
                SetBusy(false, "WARD FAILED");
            }
        }

        private void OnClear(object sender, RoutedEventArgs e) => Clear();

        private void SetBusy(bool busy, string status)
        {
            CheckRuntimeButton.IsEnabled = !busy;
            RunWardButton.IsEnabled = !busy;
            StatusText.Text = status;
        }

        private static Color MapConsoleColor(ConsoleColor color) => color switch
        {
            ConsoleColor.Black => Color.FromRgb(0x0B, 0x0F, 0x0E),
            ConsoleColor.DarkBlue => Color.FromRgb(0x2E, 0x61, 0x80),
            ConsoleColor.DarkGreen => Color.FromRgb(0x2F, 0x70, 0x42),
            ConsoleColor.DarkCyan => Color.FromRgb(0x1F, 0x6C, 0x60),
            ConsoleColor.DarkRed => Color.FromRgb(0xA3, 0x3A, 0x2A),
            ConsoleColor.DarkMagenta => Color.FromRgb(0x8A, 0x4F, 0x9E),
            ConsoleColor.DarkYellow => Color.FromRgb(0x8F, 0x63, 0x29),
            ConsoleColor.Gray => Color.FromRgb(0xB7, 0xC2, 0xBD),
            ConsoleColor.DarkGray => Color.FromRgb(0x7F, 0x8C, 0x87),
            ConsoleColor.Blue => Color.FromRgb(0x77, 0xAE, 0xCD),
            ConsoleColor.Green => Color.FromRgb(0x6F, 0xBE, 0x84),
            ConsoleColor.Cyan => Color.FromRgb(0x5B, 0xC0, 0xAE),
            ConsoleColor.Red => Color.FromRgb(0xE0, 0x84, 0x72),
            ConsoleColor.Magenta => Color.FromRgb(0xC9, 0x8F, 0xD8),
            ConsoleColor.Yellow => Color.FromRgb(0xD9, 0xA9, 0x4A),
            ConsoleColor.White => Color.FromRgb(0xE7, 0xEC, 0xE9),
            _ => Color.FromRgb(0xB7, 0xC2, 0xBD),
        };
    }
}

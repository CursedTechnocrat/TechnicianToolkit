// App.xaml.cs - Startup, plus the headless render used to verify the window.
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
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using TechnicianToolkit.Engine;

namespace TechnicianToolkit.App
{
    public partial class App : Application
    {
        private const int RenderWidth = 1280;
        private const int RenderHeight = 820;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            string[] args = e.Args;

            // --screenshot <path> [--tool NAME] [--run]
            //
            // Renders the window to a PNG without ever showing it, following the
            // precedent the phase 00 spike set with --probe: a GUI has no console
            // to report into, so the result goes to a file. It is how the layout
            // and theme get reviewed in CI, and with --run it captures the window
            // after a real tool has actually run through the real engine.
            int flag = Array.FindIndex(args, a =>
                string.Equals(a, "--screenshot", StringComparison.OrdinalIgnoreCase));

            if (flag >= 0 && flag + 1 < args.Length)
            {
                int exitCode = Screenshot(
                    args[flag + 1],
                    ValueAfter(args, "--tool"),
                    args.Any(a => string.Equals(a, "--run", StringComparison.OrdinalIgnoreCase)),
                    ValueAfter(args, "--pane"),
                    ValueAfter(args, "--queue"));

                Shutdown(exitCode);
                return;
            }

            var window = new MainWindow();
            MainWindow = window;
            window.Show();
        }

        private static string? ValueAfter(string[] args, string flag)
        {
            int index = Array.FindIndex(args, a => string.Equals(a, flag, StringComparison.OrdinalIgnoreCase));
            return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
        }

        private int Screenshot(string path, string? tool, bool run, string? pane, string? queue)
        {
            try
            {
                // The settings screen is a dialog, so it cannot be reached by
                // switching panes. Render it directly instead -- it is a whole
                // window's worth of layout that would otherwise never be looked at
                // without a person clicking Settings.
                if (string.Equals(pane, "settings", StringComparison.OrdinalIgnoreCase))
                {
                    return ScreenshotSettings(path);
                }

                var window = new MainWindow();
                window.InitializeForRender(tool);

                bool queued = false;
                if (!string.IsNullOrWhiteSpace(queue))
                {
                    window.QueueForRender(queue!.Split(',', StringSplitOptions.RemoveEmptyEntries));
                    queued = true;
                }

                var root = (FrameworkElement)window.Content;
                Layout(root);

                if (run)
                {
                    // With a queue staged, running means running the queue -- the
                    // phase 03 exit criterion, driven end to end.
                    RunToCompletion(window, queued);

                    // Lay out first so the output list has a real viewport, then
                    // scroll, then lay out again so the new offset is applied.
                    Layout(root);
                    window.ScrollOutputToEnd();
                    Layout(root);
                }

                if (!string.IsNullOrWhiteSpace(pane))
                {
                    window.ShowPane(pane!);
                    Layout(root);
                }

                // Rendered at 2x and tagged 192 DPI, so text is judged on shape
                // rather than on hinting artefacts at 96.
                var bitmap = new RenderTargetBitmap(
                    RenderWidth * 2, RenderHeight * 2, 192, 192, PixelFormats.Pbgra32);
                bitmap.Render(root);

                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));

                string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                using FileStream file = File.Create(path);
                encoder.Save(file);

                return 0;
            }
            catch (Exception ex)
            {
                // No console to complain to, so leave the reason beside the target.
                File.WriteAllText(path + ".error.txt", ex.ToString());
                return 1;
            }
        }

        private int ScreenshotSettings(string path)
        {
            ToolkitLayout layout = ScriptExtractor.Prepare();

            var window = new SettingsWindow(layout.SuiteDirectory, layout.ReportDirectory);
            window.PopulateForRender();

            var root = (FrameworkElement)window.Content;
            root.Width = 620;
            root.Height = 580;
            root.Measure(new Size(620, 580));
            root.Arrange(new Rect(0, 0, 620, 580));
            root.UpdateLayout();
            Drain();

            var bitmap = new RenderTargetBitmap(620 * 2, 580 * 2, 192, 192, PixelFormats.Pbgra32);
            bitmap.Render(root);

            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bitmap));

            using FileStream file = File.Create(path);
            encoder.Save(file);

            return 0;
        }

        private static void Layout(FrameworkElement root)
        {
            root.Width = RenderWidth;
            root.Height = RenderHeight;
            root.Measure(new Size(RenderWidth, RenderHeight));
            root.Arrange(new Rect(0, 0, RenderWidth, RenderHeight));
            root.UpdateLayout();

            // Let queued dispatcher work (batched output writes, bindings) settle
            // before anything is committed to pixels.
            Drain();
        }

        /// <summary>
        /// Run the selected tool to completion while pumping the dispatcher,
        /// since a render-only process never starts a message loop of its own.
        /// </summary>
        private static void RunToCompletion(MainWindow window, bool queue)
        {
            var frame = new DispatcherFrame();

            Task run = queue ? window.RunQueueAsync() : window.RunSelectedAsync();
            run.ContinueWith(_ => frame.Continue = false, TaskScheduler.Default);

            Dispatcher.PushFrame(frame);
            Drain();
        }

        private static void Drain() =>
            Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.ContextIdle);
    }
}

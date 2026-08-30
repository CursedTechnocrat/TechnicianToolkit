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
                    args.Any(a => string.Equals(a, "--run", StringComparison.OrdinalIgnoreCase)));

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

        private int Screenshot(string path, string? tool, bool run)
        {
            try
            {
                var window = new MainWindow();
                window.InitializeForRender(tool);

                var root = (FrameworkElement)window.Content;
                Layout(root);

                if (run)
                {
                    RunToCompletion(window);

                    // Lay out first so the output list has a real viewport, then
                    // scroll, then lay out again so the new offset is applied.
                    Layout(root);
                    window.ScrollOutputToEnd();
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
        private static void RunToCompletion(MainWindow window)
        {
            var frame = new DispatcherFrame();

            Task run = window.RunSelectedAsync();
            run.ContinueWith(_ => frame.Continue = false, TaskScheduler.Default);

            Dispatcher.PushFrame(frame);
            Drain();
        }

        private static void Drain() =>
            Dispatcher.CurrentDispatcher.Invoke(() => { }, DispatcherPriority.ContextIdle);
    }
}

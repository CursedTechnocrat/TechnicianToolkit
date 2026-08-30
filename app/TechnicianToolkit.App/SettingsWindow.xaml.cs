// SettingsWindow.xaml.cs - Edits config.json via Get-TKConfig / Set-TKConfig.
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
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using TechnicianToolkit.Engine;

namespace TechnicianToolkit.App
{
    /// <summary>
    /// A form over config.json.
    ///
    /// It reads through Get-TKConfig and writes through Set-TKConfig rather than
    /// touching the file, so the window and hearth.ps1 stay two doors onto one
    /// room. hearth.ps1 is not replaced: it remains the way to do this from a
    /// console, which is still the primary documented path.
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private readonly string _suiteDirectory;
        private readonly List<(ConfigField Field, TextBox Box, string Original)> _rows =
            new List<(ConfigField, TextBox, string)>();

        public SettingsWindow(string suiteDirectory, string currentReportDirectory)
        {
            InitializeComponent();

            _suiteDirectory = suiteDirectory;
            ReportDirectory = currentReportDirectory;

            Loaded += (_, _) => Populate();
        }

        /// <summary>
        /// Build the form without showing the window, for the screenshot render.
        /// Loaded never fires on a window that is never shown.
        /// </summary>
        internal void PopulateForRender() => Populate();

        /// <summary>
        /// The report directory after saving, so the caller can re-point its
        /// watcher when the technician moves it.
        /// </summary>
        public string? ReportDirectory { get; private set; }

        private void Populate()
        {
            IReadOnlyList<ConfigField> fields;

            try
            {
                fields = ToolkitConfig.Read(_suiteDirectory);
            }
            catch (Exception ex)
            {
                StatusText.Text = "Could not read the configuration: " + ex.Message;
                StatusText.Foreground = (Brush)FindResource("Red");
                SaveButton.IsEnabled = false;
                return;
            }

            foreach (ConfigField field in fields)
            {
                FieldHost.Children.Add(BuildRow(field));
            }
        }

        private UIElement BuildRow(ConfigField field)
        {
            var stack = new StackPanel { Margin = new Thickness(0, 0, 0, 20) };

            stack.Children.Add(new TextBlock
            {
                Text = field.Label,
                Style = (Style)FindResource("DimLabel"),
                Margin = new Thickness(0, 0, 0, 5),
            });

            var box = new TextBox { Text = field.Value };

            if (field.IsPath)
            {
                var row = new Grid();
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                var browse = new Button
                {
                    Content = "...",
                    Style = (Style)FindResource("GhostButton"),
                    Padding = new Thickness(11, 7, 11, 7),
                    Margin = new Thickness(7, 0, 0, 0),
                };

                browse.Click += (_, _) =>
                {
                    // A folder picker without dragging in a shell COM dependency:
                    // the file dialog is pointed at the folder and its directory
                    // is taken. Clumsy, and it keeps the runtime footprint honest.
                    var dialog = new OpenFileDialog
                    {
                        CheckFileExists = false,
                        ValidateNames = false,
                        FileName = "Select this folder",
                        Title = field.Label,
                        InitialDirectory = Directory.Exists(box.Text) ? box.Text : null,
                    };

                    if (dialog.ShowDialog(this) == true)
                    {
                        box.Text = Path.GetDirectoryName(dialog.FileName) ?? box.Text;
                    }
                };

                Grid.SetColumn(box, 0);
                Grid.SetColumn(browse, 1);
                row.Children.Add(box);
                row.Children.Add(browse);
                stack.Children.Add(row);
            }
            else
            {
                stack.Children.Add(box);
            }

            stack.Children.Add(new TextBlock
            {
                Text = field.Description,
                Style = (Style)FindResource("BodyText"),
                FontSize = 11,
                Foreground = (Brush)FindResource("TextDim"),
                Margin = new Thickness(0, 6, 0, 0),
            });

            _rows.Add((field, box, field.Value));
            return stack;
        }

        private void OnSave(object sender, RoutedEventArgs e)
        {
            List<ConfigField> changed = _rows
                .Where(r => !string.Equals(r.Box.Text ?? string.Empty, r.Original, StringComparison.Ordinal))
                .Select(r =>
                {
                    r.Field.Value = r.Box.Text ?? string.Empty;
                    return r.Field;
                })
                .ToList();

            if (changed.Count == 0)
            {
                DialogResult = false;
                Close();
                return;
            }

            // A report directory that does not exist yet is a reasonable thing to
            // type; creating it here means the watcher has something to watch.
            ConfigField? logDirectory = changed.FirstOrDefault(f =>
                f.Key.Equals("LogDirectory", StringComparison.OrdinalIgnoreCase));

            if (logDirectory != null && !string.IsNullOrWhiteSpace(logDirectory.Value))
            {
                try
                {
                    Directory.CreateDirectory(logDirectory.Value);
                }
                catch (Exception ex)
                {
                    StatusText.Text = "That report directory cannot be created: " + ex.Message;
                    StatusText.Foreground = (Brush)FindResource("Red");
                    return;
                }
            }

            try
            {
                ToolkitConfig.Write(_suiteDirectory, changed);
            }
            catch (Exception ex)
            {
                StatusText.Text = "Save failed: " + ex.Message;
                StatusText.Foreground = (Brush)FindResource("Red");
                return;
            }

            if (logDirectory != null && !string.IsNullOrWhiteSpace(logDirectory.Value))
            {
                ReportDirectory = logDirectory.Value;
            }

            DialogResult = true;
            Close();
        }

        private void OnCancel(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

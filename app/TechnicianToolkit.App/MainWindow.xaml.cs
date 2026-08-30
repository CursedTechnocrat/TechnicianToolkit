// MainWindow.xaml.cs - The window: catalog, generated form, live output.
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
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Win32;
using TechnicianToolkit.Engine;

namespace TechnicianToolkit.App
{
    /// <summary>One row in the catalog.</summary>
    public sealed class ToolItem
    {
        public ToolEntry Entry { get; init; } = new ToolEntry();
        public ToolTrait Traits { get; init; } = new ToolTrait();
        public string ScriptPath { get; init; } = string.Empty;

        public string Key => Entry.Key;
        public string ShortName => Entry.ShortName;
        public string Description => Entry.Description;
        public string Category => Entry.Category;

        public Visibility AdminVisibility => Traits.RequiresAdmin ? Visibility.Visible : Visibility.Collapsed;
        public Visibility DestructiveVisibility => Traits.IsDestructive ? Visibility.Visible : Visibility.Collapsed;
    }

    /// <summary>
    /// One generated control, paired with the parameter it came from and a way
    /// to read its value back out.
    /// </summary>
    internal sealed class ParameterField
    {
        internal ToolParameter Parameter { get; init; } = new ToolParameter();
        internal Func<object?> ReadValue { get; init; } = () => null;
        internal Func<bool> IsValid { get; init; } = () => true;
    }

    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<OutputLine> _lines = new ObservableCollection<OutputLine>();
        private readonly List<ToolItem> _catalog = new List<ToolItem>();
        private readonly List<ParameterField> _fields = new List<ParameterField>();

        private string _workDir = string.Empty;

        /// <summary>
        /// Where the suite and the reports live. Phase 03 watches
        /// <see cref="ToolkitLayout.ReportDirectory"/> for new artifacts.
        /// </summary>
        private ToolkitLayout _layout = new ToolkitLayout();
        private WpfHostSink? _sink;
        private CancellationTokenSource? _cancellation;
        private ToolItem? _selected;
        private bool _running;

        public MainWindow()
        {
            InitializeComponent();

            OutputList.ItemsSource = _lines;
            _lines.CollectionChanged += (_, _) =>
            {
                OutputCount.Text = _lines.Count == 0 ? string.Empty : _lines.Count + " lines";
                OutputEmpty.Visibility = _lines.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
                ScrollOutputToEnd();
            };

            _sink = new WpfHostSink(Dispatcher, _lines, ApplyProgress, () => this);

            Loaded += OnLoaded;
        }

        internal WpfHostSink Sink => _sink!;

        // ─── Startup ──────────────────────────────────────────────────────────

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ApplyDarkTitleBar();
            ShowElevationState();

            try
            {
                LoadCatalog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "The toolkit could not be prepared:" + Environment.NewLine + Environment.NewLine + ex.Message,
                    "Technician Toolkit", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Bring the window up without showing it, for the --screenshot render.
        /// Loaded never fires on a window that is never shown, so the same setup
        /// is driven directly. The title-bar call is skipped because there is no
        /// window handle to apply it to.
        /// </summary>
        internal void InitializeForRender(string? selectTool)
        {
            ShowElevationState();
            LoadCatalog();

            ToolItem? pick = selectTool == null
                ? _catalog.FirstOrDefault()
                : _catalog.FirstOrDefault(t =>
                      string.Equals(t.ShortName, selectTool.Replace(".", string.Empty), StringComparison.OrdinalIgnoreCase)
                   || string.Equals(t.Key, selectTool, StringComparison.OrdinalIgnoreCase));

            if (pick == null)
            {
                return;
            }

            CatalogList.SelectedItem = pick;

            // SelectionChanged does not fire on a list that has never rendered.
            SelectTool(pick);
        }

        /// <summary>
        /// Windows paints the title bar itself, and it defaults to light. Without
        /// this the app wears a white cap over a near-black window.
        /// </summary>
        private void ApplyDarkTitleBar()
        {
            try
            {
                IntPtr handle = new WindowInteropHelper(this).Handle;
                int enabled = 1;
                // DWMWA_USE_IMMERSIVE_DARK_MODE
                DwmSetWindowAttribute(handle, 20, ref enabled, sizeof(int));
            }
            catch
            {
                // Older Windows builds do not know the attribute. Cosmetic only.
            }
        }

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int size);

        private void ShowElevationState()
        {
            if (Elevation.IsElevated())
            {
                return;
            }

            // The shipping manifest requests Administrator, so this state should
            // not occur in the field. It does when the app is built with the
            // asInvoker manifest for testing, and a tool refusing to start is
            // otherwise silent -- a script exit never reaches SetShouldExit.
            ElevationBadge.Background = (Brush)FindResource("YellowDim");
            ElevationText.Text = "NOT ELEVATED";
            ElevationText.Foreground = (Brush)FindResource("Yellow");
        }

        private void LoadCatalog()
        {
            _layout = ScriptExtractor.Prepare();
            _workDir = _layout.SuiteDirectory;

            string grimoire = Path.Combine(_workDir, "grimoire.ps1");
            IReadOnlyList<ToolEntry> entries = ToolCatalog.Load(grimoire);
            IReadOnlyList<string> order = ToolCatalog.LoadCategoryOrder(grimoire);

            foreach (ToolEntry entry in entries)
            {
                string scriptPath = Path.Combine(_workDir, entry.File);
                if (!File.Exists(scriptPath))
                {
                    continue;
                }

                _catalog.Add(new ToolItem
                {
                    Entry = entry,
                    Traits = ToolTraits.Inspect(scriptPath),
                    ScriptPath = scriptPath,
                });
            }

            var view = (CollectionViewSource)FindResource("CatalogView");
            view.Source = _catalog;

            // Group in the hub's declared order rather than alphabetically, so
            // the window reads like the console menu it replaces.
            view.View.SortDescriptions.Clear();
            if (view.View is ListCollectionView list)
            {
                list.CustomSort = new CategoryThenKeyComparer(order);
            }

            CatalogCount.Text = _catalog.Count + " tools";
        }

        /// <summary>
        /// Orders by the hub's own $CategoryOrder, then by registry key. The keys
        /// are numeric strings with deliberate gaps, so they sort as numbers.
        /// </summary>
        private sealed class CategoryThenKeyComparer : System.Collections.IComparer
        {
            private readonly IReadOnlyList<string> _order;

            internal CategoryThenKeyComparer(IReadOnlyList<string> order) => _order = order;

            public int Compare(object? x, object? y)
            {
                if (x is not ToolItem a || y is not ToolItem b)
                {
                    return 0;
                }

                int categoryA = _order.ToList().IndexOf(a.Category);
                int categoryB = _order.ToList().IndexOf(b.Category);
                if (categoryA < 0) { categoryA = int.MaxValue; }
                if (categoryB < 0) { categoryB = int.MaxValue; }

                if (categoryA != categoryB)
                {
                    return categoryA.CompareTo(categoryB);
                }

                bool parsedA = int.TryParse(a.Key, out int keyA);
                bool parsedB = int.TryParse(b.Key, out int keyB);
                if (parsedA && parsedB)
                {
                    return keyA.CompareTo(keyB);
                }

                return string.Compare(a.Key, b.Key, StringComparison.OrdinalIgnoreCase);
            }
        }

        // ─── Search ───────────────────────────────────────────────────────────

        private void OnSearchChanged(object sender, TextChangedEventArgs e)
        {
            string term = SearchBox.Text.Trim();
            SearchPlaceholder.Visibility = term.Length == 0 ? Visibility.Visible : Visibility.Collapsed;

            var view = (CollectionViewSource)FindResource("CatalogView");
            if (view.View == null)
            {
                return;
            }

            view.View.Filter = term.Length == 0
                ? null
                : item => item is ToolItem tool
                    && (tool.ShortName.Contains(term, StringComparison.OrdinalIgnoreCase)
                     || tool.Description.Contains(term, StringComparison.OrdinalIgnoreCase)
                     || tool.Entry.Name.Contains(term, StringComparison.OrdinalIgnoreCase));

            CatalogCount.Text = view.View.Cast<object>().Count() + " of " + _catalog.Count + " tools";
        }

        // ─── Selection and the generated form ─────────────────────────────────

        private void OnToolSelected(object sender, SelectionChangedEventArgs e)
        {
            if (CatalogList.SelectedItem is ToolItem tool)
            {
                SelectTool(tool);
            }
        }

        private void SelectTool(ToolItem tool)
        {
            _selected = tool;
            ToolName.Text = tool.Entry.Name;
            ToolVersion.Text = "v" + tool.Entry.Version;
            ToolDescription.Text = tool.Description;

            BuildBadges(tool);
            BuildForm(tool);

            if (_lines.Count == 0)
            {
                OutputEmpty.Text = "Press Run to start " + tool.ShortName + ".";
            }

            UpdateRunEnabled();
        }

        private void BuildBadges(ToolItem tool)
        {
            ToolBadges.Items.Clear();

            if (tool.Traits.RequiresAdmin)
            {
                ToolBadges.Items.Add(Badge("REQUIRES ADMIN", "Yellow", "YellowDim"));
            }
            if (tool.Traits.IsDestructive)
            {
                ToolBadges.Items.Add(Badge("CHANGES THIS MACHINE", "Red", "RedDim"));
            }
            if (!tool.Traits.IsDestructive && !tool.Traits.RequiresAdmin)
            {
                ToolBadges.Items.Add(Badge("READ ONLY", "Green", "GreenDim"));
            }
        }

        private Border Badge(string text, string foreground, string background) => new Border
        {
            Style = (Style)FindResource("Badge"),
            Background = (Brush)FindResource(background),
            Child = new TextBlock
            {
                Text = text,
                Style = (Style)FindResource("BadgeText"),
                Foreground = (Brush)FindResource(foreground),
            },
        };

        /// <summary>
        /// The form builds itself from the script's param() block, so it can
        /// never drift from the tool it drives.
        /// </summary>
        private void BuildForm(ToolItem tool)
        {
            FormHost.Items.Clear();
            _fields.Clear();

            IReadOnlyList<ToolParameter> parameters;
            try
            {
                parameters = ToolParameters.Load(tool.ScriptPath);
            }
            catch (Exception ex)
            {
                FormHost.Items.Add(new TextBlock
                {
                    Text = "Could not read this tool's parameters: " + ex.Message,
                    Style = (Style)FindResource("BodyText"),
                    Foreground = (Brush)FindResource("Red"),
                });
                return;
            }

            foreach (ToolParameter parameter in parameters)
            {
                FormHost.Items.Add(BuildField(parameter));
            }
        }

        private UIElement BuildField(ToolParameter parameter)
        {
            var stack = new StackPanel { Margin = new Thickness(0, 0, 22, 12), MinWidth = 200 };

            if (parameter.Control == ParameterControl.Checkbox)
            {
                // -Unattended arrives ticked: the form supplies what the prompts
                // would otherwise ask for, which is the whole point of it.
                bool unattended = parameter.Name.Equals("Unattended", StringComparison.OrdinalIgnoreCase);

                // A checkbox carries its own label, but the fields beside it have
                // one stacked above; without a matching spacer the row of
                // controls sits at two different heights.
                stack.Children.Add(new TextBlock
                {
                    Text = " ",
                    Style = (Style)FindResource("DimLabel"),
                    Margin = new Thickness(0, 0, 0, 5),
                });

                var box = new CheckBox
                {
                    Content = "-" + parameter.Name,
                    IsChecked = unattended,
                    Height = 33,
                    VerticalContentAlignment = VerticalAlignment.Center,
                };

                if (parameter.Name.Equals("WhatIf", StringComparison.OrdinalIgnoreCase))
                {
                    box.Foreground = (Brush)FindResource("Yellow");
                    box.ToolTip = "Dry run: report what would change without changing it.";
                }

                stack.Children.Add(box);
                _fields.Add(new ParameterField
                {
                    Parameter = parameter,
                    ReadValue = () => box.IsChecked == true ? (object)true : null,
                });
                return stack;
            }

            stack.Children.Add(new TextBlock
            {
                Text = "-" + parameter.Name + (parameter.IsMandatory ? " *" : string.Empty),
                Style = (Style)FindResource("DimLabel"),
                Margin = new Thickness(0, 0, 0, 5),
            });

            switch (parameter.Control)
            {
                case ParameterControl.Dropdown:
                {
                    var combo = new ComboBox { Width = 210 };
                    foreach (string value in parameter.ValidValues)
                    {
                        combo.Items.Add(value);
                    }

                    string? preset = StripLiteral(parameter.DefaultValue);
                    combo.SelectedItem = preset != null && parameter.ValidValues.Contains(preset)
                        ? preset
                        : parameter.ValidValues.FirstOrDefault();

                    stack.Children.Add(combo);
                    _fields.Add(new ParameterField
                    {
                        Parameter = parameter,
                        ReadValue = () => combo.SelectedItem,
                    });
                    break;
                }

                case ParameterControl.MaskedText:
                {
                    var box = new PasswordBox { Width = 210 };
                    stack.Children.Add(box);
                    _fields.Add(new ParameterField
                    {
                        Parameter = parameter,
                        ReadValue = () => box.SecurePassword.Length == 0 ? null : box.SecurePassword,
                    });
                    break;
                }

                case ParameterControl.PathPicker:
                {
                    var row = new StackPanel { Orientation = Orientation.Horizontal };
                    var box = new TextBox { Width = 210, Text = StripLiteral(parameter.DefaultValue) ?? string.Empty };
                    var browse = new Button
                    {
                        Content = "...",
                        Style = (Style)FindResource("GhostButton"),
                        Padding = new Thickness(11, 7, 11, 7),
                        Margin = new Thickness(7, 0, 0, 0),
                    };
                    browse.Click += (_, _) =>
                    {
                        var dialog = new OpenFileDialog
                        {
                            CheckFileExists = false,
                            ValidateNames = false,
                            FileName = "Select this folder",
                            Title = "Choose a path for -" + parameter.Name,
                        };
                        if (dialog.ShowDialog(this) == true)
                        {
                            box.Text = Path.GetDirectoryName(dialog.FileName) ?? dialog.FileName;
                        }
                    };

                    row.Children.Add(box);
                    row.Children.Add(browse);
                    stack.Children.Add(row);

                    _fields.Add(new ParameterField
                    {
                        Parameter = parameter,
                        ReadValue = () => string.IsNullOrWhiteSpace(box.Text) ? null : box.Text,
                    });
                    break;
                }

                default:
                {
                    var box = new TextBox { Width = 210, Text = StripLiteral(parameter.DefaultValue) ?? string.Empty };

                    // A ValidatePattern is a live check, not a post-hoc failure:
                    // the script would reject the value anyway, so say so first.
                    Regex? pattern = null;
                    if (parameter.ValidationPattern != null)
                    {
                        try { pattern = new Regex(parameter.ValidationPattern); }
                        catch (ArgumentException) { pattern = null; }
                    }

                    Func<bool> valid = () =>
                        pattern == null
                        || string.IsNullOrEmpty(box.Text)
                        || pattern.IsMatch(box.Text);

                    if (pattern != null)
                    {
                        box.TextChanged += (_, _) =>
                        {
                            box.BorderBrush = valid()
                                ? (Brush)FindResource("BorderBrushDim")
                                : (Brush)FindResource("Red");
                            UpdateRunEnabled();
                        };

                        stack.Children.Add(box);
                        stack.Children.Add(new TextBlock
                        {
                            Text = parameter.ValidationPattern,
                            Style = (Style)FindResource("DimLabel"),
                            FontSize = 9.5,
                            Margin = new Thickness(0, 4, 0, 0),
                        });
                    }
                    else
                    {
                        stack.Children.Add(box);
                    }

                    _fields.Add(new ParameterField
                    {
                        Parameter = parameter,
                        ReadValue = () => string.IsNullOrWhiteSpace(box.Text) ? null : box.Text,
                        IsValid = valid,
                    });
                    break;
                }
            }

            return stack;
        }

        /// <summary>
        /// Defaults arrive as the literal source text -- "Status", 'C', @().
        /// Only simple quoted strings are worth pre-filling; anything else is an
        /// expression the script should evaluate itself.
        /// </summary>
        private static string? StripLiteral(string? literal)
        {
            if (string.IsNullOrWhiteSpace(literal))
            {
                return null;
            }

            literal = literal.Trim();

            if (literal.Length >= 2
                && ((literal[0] == '"' && literal[^1] == '"') || (literal[0] == '\'' && literal[^1] == '\'')))
            {
                return literal.Substring(1, literal.Length - 2);
            }

            return null;
        }

        private void UpdateRunEnabled() =>
            RunButton.IsEnabled = !_running && _selected != null && _fields.All(f => f.IsValid());

        // ─── Running ──────────────────────────────────────────────────────────

        private async void OnRun(object sender, RoutedEventArgs e) => await RunSelectedAsync();

        /// <summary>
        /// Run the selected tool. Exposed so the --screenshot render can drive a
        /// real run through the real engine rather than staging fake output.
        /// </summary>
        internal async Task RunSelectedAsync()
        {
            if (_selected == null || _running)
            {
                return;
            }

            var parameters = new Dictionary<string, object?>();
            foreach (ParameterField field in _fields)
            {
                object? value = field.ReadValue();
                if (value != null)
                {
                    parameters[field.Parameter.Name] = value;
                }
            }

            _running = true;
            RunButton.IsEnabled = false;
            CancelButton.IsEnabled = true;
            CatalogList.IsEnabled = false;

            _cancellation = new CancellationTokenSource();
            var runner = new ToolRunner(Sink, _workDir);

            try
            {
                ToolRunResult result = await runner
                    .RunAsync(_selected.ScriptPath, parameters, _cancellation.Token)
                    .ConfigureAwait(true);

                Sink.Flush();
                ReportResult(result);
            }
            catch (Exception ex)
            {
                Sink.Write("EXCEPTION: " + ex.Message + Environment.NewLine, ConsoleColor.Red, null);
                Sink.Flush();
            }
            finally
            {
                _running = false;
                CancelButton.IsEnabled = false;
                CatalogList.IsEnabled = true;
                ProgressPanel.Visibility = Visibility.Collapsed;
                _cancellation?.Dispose();
                _cancellation = null;
                UpdateRunEnabled();
            }
        }

        private void ReportResult(ToolRunResult result)
        {
            if (result.Cancelled)
            {
                Sink.Write(Environment.NewLine + "-- cancelled --" + Environment.NewLine,
                    ConsoleColor.Yellow, null);
            }
            else if (result.RefusedNeedsAdmin)
            {
                // Without this the pane shows one line of refusal and then a green
                // "finished", which reads as success for a run that did nothing.
                Sink.Write(Environment.NewLine
                    + "-- refused: this tool requires Administrator and did no work --"
                    + Environment.NewLine, ConsoleColor.Yellow, null);
            }
            else if (result.Errors.Count > 0)
            {
                Sink.Write(Environment.NewLine + "-- finished with " + result.Errors.Count
                    + " error(s) --" + Environment.NewLine, ConsoleColor.Red, null);
            }
            else
            {
                Sink.Write(Environment.NewLine + "-- finished --" + Environment.NewLine,
                    ConsoleColor.Green, null);
            }

            Sink.Flush();
        }

        private void OnCancel(object sender, RoutedEventArgs e) => _cancellation?.Cancel();

        private void ApplyProgress(ProgressState progress)
        {
            if (progress.Completed)
            {
                ProgressPanel.Visibility = Visibility.Collapsed;
                return;
            }

            ProgressPanel.Visibility = Visibility.Visible;
            ProgressText.Text = string.IsNullOrEmpty(progress.Status)
                ? progress.Activity
                : progress.Activity + " - " + progress.Status;

            // A negative percentage means the tool is working without knowing how
            // far along it is; show the bar full rather than empty so it does not
            // read as stalled at zero.
            double fraction = progress.PercentComplete < 0 ? 1.0 : progress.PercentComplete / 100.0;
            ProgressBarFill.Width = 260 * Math.Clamp(fraction, 0, 1);
        }

        // ─── Output pane actions ──────────────────────────────────────────────

        /// <summary>
        /// Follow the tail of the output. Internal because the screenshot render
        /// has to re-issue it after laying the window out: scrolling a list that
        /// has no viewport yet computes an offset against a zero-height panel,
        /// and the pane ends up parked past the end of its own content.
        /// </summary>
        internal void ScrollOutputToEnd()
        {
            if (_lines.Count > 0)
            {
                OutputList.ScrollIntoView(_lines[^1]);
            }
        }

        private string OutputAsText()
        {
            var builder = new StringBuilder();
            foreach (OutputLine line in _lines)
            {
                builder.AppendLine(line.Text);
            }
            return builder.ToString();
        }

        private void OnCopyOutput(object sender, RoutedEventArgs e)
        {
            if (_lines.Count == 0)
            {
                return;
            }

            try
            {
                Clipboard.SetText(OutputAsText());
            }
            catch (Exception ex)
            {
                // The clipboard is a shared resource and another process can hold
                // it open; that is not worth taking the app down for.
                MessageBox.Show(this, "Could not copy: " + ex.Message, "Technician Toolkit",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OnSaveOutput(object sender, RoutedEventArgs e)
        {
            if (_lines.Count == 0)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "Text file (*.txt)|*.txt|All files (*.*)|*.*",
                FileName = (_selected?.ShortName ?? "toolkit") + "_"
                           + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt",
            };

            if (dialog.ShowDialog(this) == true)
            {
                File.WriteAllText(dialog.FileName, OutputAsText(), new UTF8Encoding(true));
            }
        }

        private void OnClearOutput(object sender, RoutedEventArgs e) => _lines.Clear();

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_running)
            {
                MessageBoxResult answer = MessageBox.Show(this,
                    "A tool is still running. Cancel it and close?",
                    "Technician Toolkit", MessageBoxButton.YesNo, MessageBoxImage.Warning);

                if (answer != MessageBoxResult.Yes)
                {
                    e.Cancel = true;
                    return;
                }

                _cancellation?.Cancel();
            }

            base.OnClosing(e);
        }
    }
}

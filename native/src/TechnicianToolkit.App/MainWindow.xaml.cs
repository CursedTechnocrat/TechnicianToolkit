using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;
using TechnicianToolkit.Core.Security;
using TechnicianToolkit.Tools;

namespace TechnicianToolkit.App;

public partial class MainWindow : Window
{
    private string? _lastReportPath;
    private bool _running;

    private static readonly Brush StepBrush = new SolidColorBrush(Color.FromRgb(0xD0, 0x6B, 0xFF));

    public MainWindow()
    {
        InitializeComponent();

        ToolList.ItemsSource = ToolCatalog.Tools;
        SuiteNote.Text =
            $"{ToolCatalog.Tools.Count} of {ToolCatalog.TotalSuiteToolCount} suite tools are ported to " +
            "native code so far. The rest remain available as PowerShell scripts in the toolkit.";

        UpdateAdminBanner();

        if (ToolList.Items.Count > 0)
        {
            ToolList.SelectedIndex = 0;
        }
    }

    private void UpdateAdminBanner()
    {
        AdminBanner.Visibility = AdminPrivilege.IsAdmin() ? Visibility.Collapsed : Visibility.Visible;
    }

    private ITool? SelectedTool => ToolList.SelectedItem as ITool;

    private void ToolList_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        var tool = SelectedTool;
        if (tool is null)
        {
            return;
        }

        DetailName.Text = tool.Name;
        DetailTitle.Text = tool.Title;
        DetailDescription.Text = tool.Description;

        BadgeRow.Children.Clear();
        BadgeRow.Children.Add(MakeBadge(tool.Category, (Brush)FindResource("CyanBrush")));
        if (tool.RequiresAdmin)
        {
            BadgeRow.Children.Add(MakeBadge("Requires Administrator", (Brush)FindResource("YellowBrush")));
        }

        BadgeRow.Children.Add(MakeBadge(
            tool.SupportsWhatIf ? "Supports preview" : "Read-only",
            (Brush)FindResource("GreenBrush")));

        WhatIfCheck.IsEnabled = tool.SupportsWhatIf;
        if (!tool.SupportsWhatIf)
        {
            WhatIfCheck.IsChecked = false;
        }

        // Reset the run surface for the newly selected tool.
        ResultPanel.Visibility = Visibility.Collapsed;
        StatusText.Text = string.Empty;
    }

    private static Border MakeBadge(string text, Brush color)
    {
        return new Border
        {
            BorderBrush = color,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(20),
            Padding = new Thickness(10, 2, 10, 2),
            Margin = new Thickness(0, 0, 8, 0),
            Child = new TextBlock
            {
                Text = text,
                Foreground = color,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 11,
            },
        };
    }

    private async void RunButton_Click(object sender, RoutedEventArgs e)
    {
        var tool = SelectedTool;
        if (tool is null || _running)
        {
            return;
        }

        _running = true;
        RunButton.IsEnabled = false;
        ResultPanel.Visibility = Visibility.Collapsed;
        LogDoc.Blocks.Clear();
        StatusText.Text = "Running…";

        var progress = new Progress<ToolProgress>(AppendLog);
        var outputPath = string.IsNullOrWhiteSpace(OutputBox.Text) ? null : OutputBox.Text.Trim();
        var whatIf = tool.SupportsWhatIf && WhatIfCheck.IsChecked == true;

        var ctx = new ToolContext
        {
            Unattended = true,
            WhatIf = whatIf,
            OutputPath = outputPath,
            Progress = progress,
        };

        ToolResult result;
        try
        {
            result = await Task.Run(() => tool.Run(ctx));
        }
        catch (Exception ex)
        {
            result = ToolResult.Fail(ex.Message);
        }

        _running = false;
        RunButton.IsEnabled = true;

        if (result.Success)
        {
            _lastReportPath = result.ReportPath;
            StatusText.Text = "Done.";
            ResultPanel.BorderBrush = (Brush)FindResource("GreenBrush");
            ResultSummary.Text = "Report generated" +
                (result.Summary.Count > 0
                    ? "  —  " + string.Join("   ·   ", result.Summary.Select(kv => $"{kv.Key}: {kv.Value}"))
                    : ".");
            ResultPath.Text = result.ReportPath;
            OpenReportButton.IsEnabled = true;
            OpenFolderButton.IsEnabled = true;
            ResultPanel.Visibility = Visibility.Visible;
        }
        else
        {
            _lastReportPath = null;
            StatusText.Text = "Failed.";
            ResultPanel.BorderBrush = (Brush)FindResource("RedBrush");
            ResultSummary.Text = "The tool reported an error:";
            ResultPath.Text = result.Error ?? "(unknown error)";
            OpenReportButton.IsEnabled = false;
            OpenFolderButton.IsEnabled = false;
            ResultPanel.Visibility = Visibility.Visible;
        }
    }

    private void AppendLog(ToolProgress p)
    {
        var (prefix, brush) = p.Level switch
        {
            ProgressLevel.Ok => ("  [+] ", (Brush)FindResource("GreenBrush")),
            ProgressLevel.Warn => ("  [!] ", (Brush)FindResource("YellowBrush")),
            ProgressLevel.Fail => ("  [-] ", (Brush)FindResource("RedBrush")),
            ProgressLevel.Step => ("  [*] ", StepBrush),
            _ => ("      ", (Brush)FindResource("TextDimBrush")),
        };

        var para = new Paragraph { Margin = new Thickness(0) };
        para.Inlines.Add(new Run(prefix + p.Message) { Foreground = brush });
        LogDoc.Blocks.Add(para);
        LogBox.ScrollToEnd();
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "Choose a report output folder",
        };

        if (dialog.ShowDialog(this) == true)
        {
            OutputBox.Text = dialog.FolderName;
        }
    }

    private void ElevateButton_Click(object sender, RoutedEventArgs e)
    {
        if (AdminPrivilege.RelaunchElevated())
        {
            Application.Current.Shutdown();
        }
    }

    private void OpenReportButton_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_lastReportPath) && System.IO.File.Exists(_lastReportPath))
        {
            OpenPath(_lastReportPath);
        }
    }

    private void OpenFolderButton_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(_lastReportPath))
        {
            var dir = System.IO.Path.GetDirectoryName(_lastReportPath);
            if (!string.IsNullOrEmpty(dir) && System.IO.Directory.Exists(dir))
            {
                OpenPath(dir);
            }
        }
    }

    private static void OpenPath(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not open:\n{path}\n\n{ex.Message}", "TechnicianToolkit",
                MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}

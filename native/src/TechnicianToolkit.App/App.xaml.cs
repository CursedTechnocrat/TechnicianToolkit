using System.Windows;
using TechnicianToolkit.Core;
using TechnicianToolkit.Core.Config;

namespace TechnicianToolkit.App;

public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        // Anchor config.json and the default report/log fallback next to the
        // executable, so a portable copy of the app keeps its settings with it.
        // Done before base.OnStartup so it is in place before the first window.
        var baseDir = AppContext.BaseDirectory;
        TkPaths.BaseDirectory = baseDir;
        TkConfig.ConfigPath = System.IO.Path.Combine(baseDir, "config.json");

        base.OnStartup(e);
    }
}

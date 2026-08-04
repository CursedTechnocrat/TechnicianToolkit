using System.ServiceProcess;

namespace TechnicianToolkit.Tools.Collectors;

/// <summary>A stopped automatic service (SCRYER section 5).</summary>
public sealed class StoppedServiceInfo
{
    public string Name { get; init; } = "";
    public string DisplayName { get; init; } = "";
    public string StartType { get; init; } = "Automatic";
}

/// <summary>
/// Finds automatic-start services that are currently stopped, excluding the
/// trigger-started / on-demand services SCRYER deliberately ignores. Uses
/// <see cref="ServiceController"/> in place of <c>Get-Service</c>.
/// </summary>
public static class ServiceCollector
{
    // Verbatim from SCRYER: services that are legitimately stopped much of the
    // time (trigger-started, on-demand) and would be noise if reported.
    private static readonly HashSet<string> TriggerExclusions = new(StringComparer.OrdinalIgnoreCase)
    {
        "gupdate", "gupdatem", "edgeupdate", "edgeupdatem", "MapsBroker",
        "RemoteRegistry", "SharedAccess", "TabletInputService", "WbioSrvc", "lfsvc",
        "SCardSvr", "SensrSvc", "WSearch", "wuauserv", "BITS", "DoSvc", "UsoSvc", "WerSvc",
        "AppReadiness", "tiledatamodelsvc", "CDPSvc", "OneSyncSvc", "PimIndexMaintenanceSvc",
        "MessagingService", "cbdhsvc", "DevicesFlowUserSvc",
    };

    public static IReadOnlyList<StoppedServiceInfo> CollectStoppedAutomatic()
    {
        var rows = new List<StoppedServiceInfo>();

        foreach (var svc in ServiceController.GetServices())
        {
            using (svc)
            {
                ServiceStartMode startMode;
                try
                {
                    startMode = svc.StartType;
                }
                catch
                {
                    continue; // some services deny StartType queries
                }

                if (startMode != ServiceStartMode.Automatic)
                {
                    continue;
                }

                if (svc.Status != ServiceControllerStatus.Stopped)
                {
                    continue;
                }

                if (TriggerExclusions.Contains(svc.ServiceName))
                {
                    continue;
                }

                rows.Add(new StoppedServiceInfo
                {
                    Name = svc.ServiceName,
                    DisplayName = svc.DisplayName,
                    StartType = "Automatic",
                });
            }
        }

        // Sort by display name to match the PowerShell (Sort-Object DisplayName).
        return rows.OrderBy(r => r.DisplayName, StringComparer.OrdinalIgnoreCase).ToList();
    }
}

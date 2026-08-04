using System.Globalization;

namespace TechnicianToolkit.Tools.Collectors;

/// <summary>A scheduled task whose last run failed (SCRYER section 5).</summary>
public sealed class FailedTaskInfo
{
    public string TaskName { get; init; } = "";
    public string TaskPath { get; init; } = "";
    public string LastRunTime { get; init; } = "";
    public string LastResult { get; init; } = "";
}

/// <summary>
/// Enumerates non-disabled, non-Microsoft scheduled tasks whose last run
/// returned a non-zero result — SCRYER's failed-task pass. Uses the Task
/// Scheduler 2.0 COM API (<c>Schedule.Service</c>) via late binding, so no
/// third-party Task Scheduler library is needed. The whole pass is best-effort:
/// any COM failure yields an empty list, mirroring the PowerShell try/catch.
/// </summary>
public static class ScheduledTaskCollector
{
    private const int TaskStateDisabled = 1; // TASK_STATE_DISABLED
    private const int MaxResults = 20;

    public static IReadOnlyList<FailedTaskInfo> CollectFailed()
    {
        var results = new List<FailedTaskInfo>();

        var progId = Type.GetTypeFromProgID("Schedule.Service");
        if (progId is null)
        {
            return results;
        }

        object? serviceObj = null;
        try
        {
            serviceObj = Activator.CreateInstance(progId);
            dynamic service = serviceObj!;
            service.Connect();
            dynamic rootFolder = service.GetFolder("\\");
            Walk(rootFolder, results);
        }
        catch
        {
            // Task Scheduler unavailable / access denied — return what we have.
        }
        finally
        {
            if (serviceObj is not null && System.Runtime.InteropServices.Marshal.IsComObject(serviceObj))
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(serviceObj);
            }
        }

        return results;
    }

    private static void Walk(dynamic folder, List<FailedTaskInfo> results)
    {
        if (results.Count >= MaxResults)
        {
            return;
        }

        // Tasks in this folder (1 => include hidden).
        dynamic tasks = folder.GetTasks(1);
        foreach (dynamic task in tasks)
        {
            if (results.Count >= MaxResults)
            {
                return;
            }

            try
            {
                int state = task.State;
                if (state == TaskStateDisabled)
                {
                    continue;
                }

                string fullPath = task.Path;                 // e.g. "\Foo\Bar"
                var lastSlash = fullPath.LastIndexOf('\\');
                var folderPath = lastSlash <= 0 ? "\\" : fullPath[..(lastSlash + 1)];

                if (folderPath.StartsWith(@"\Microsoft\", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                int lastResult = task.LastTaskResult;
                DateTime lastRun = task.LastRunTime;

                if (lastResult != 0 && lastRun > DateTime.MinValue)
                {
                    results.Add(new FailedTaskInfo
                    {
                        TaskName = task.Name,
                        TaskPath = folderPath,
                        LastRunTime = lastRun.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
                        LastResult = string.Format(CultureInfo.InvariantCulture, "0x{0:X8}", (uint)lastResult),
                    });
                }
            }
            catch
            {
                // Skip tasks that deny inspection.
            }
        }

        // Recurse into subfolders (0 => no flags).
        dynamic subFolders = folder.GetFolders(0);
        foreach (dynamic sub in subFolders)
        {
            if (results.Count >= MaxResults)
            {
                return;
            }

            Walk(sub, results);
        }
    }
}

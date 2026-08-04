using System.Management;

namespace TechnicianToolkit.Tools.Collectors;

/// <summary>System overview, as gathered by SCRYER's section 1.</summary>
public sealed class SystemInfo
{
    public string Hostname { get; init; } = EnvInfoName();
    public string OsCaption { get; init; } = "Unknown";
    public string OsBuild { get; init; } = "";
    public string Manufacturer { get; init; } = "";
    public string Model { get; init; } = "";
    public string CpuName { get; init; } = "Unknown";
    public string CpuCores { get; init; } = "";
    public double TotalRamGb { get; init; }
    public double FreeRamGb { get; init; }
    public string LastBoot { get; init; } = "";
    public string Uptime { get; init; } = "";
    public string RuntimeVersion { get; init; } = Environment.Version.ToString();

    private static string EnvInfoName() => Environment.MachineName;
}

/// <summary>
/// Collects OS / hardware overview from <c>Win32_OperatingSystem</c>,
/// <c>Win32_ComputerSystem</c> and <c>Win32_Processor</c> — the native
/// equivalent of SCRYER's "system overview" pass. Each field degrades to the
/// same fallback the PowerShell used ("Unknown" / empty) when WMI is silent.
/// </summary>
public static class SystemInfoCollector
{
    private const double Gb = 1024d * 1024d * 1024d;

    public static SystemInfo Collect()
    {
        var os = WmiUtil.QueryCimv2("SELECT * FROM Win32_OperatingSystem").FirstOrDefault();
        var cs = WmiUtil.QueryCimv2("SELECT * FROM Win32_ComputerSystem").FirstOrDefault();
        var cpu = WmiUtil.QueryCimv2("SELECT * FROM Win32_Processor").FirstOrDefault();

        var caption = os is null ? "Unknown" : WmiUtil.Str(os, "Caption") ?? "Unknown";
        var build = os is null ? "" : WmiUtil.Str(os, "BuildNumber") ?? "";
        var manufacturer = cs is null ? "" : WmiUtil.Str(cs, "Manufacturer") ?? "";
        var model = cs is null ? "" : WmiUtil.Str(cs, "Model") ?? "";
        var cpuName = cpu is null ? "Unknown" : (WmiUtil.Str(cpu, "Name") ?? "Unknown").Trim();
        var cores = cs is null ? "" : (WmiUtil.U32(cs, "NumberOfLogicalProcessors")?.ToString() ?? "");

        // TotalPhysicalMemory is bytes; FreePhysicalMemory is KB (WMI quirk,
        // matching the PowerShell which divides by 1MB after /1KB semantics).
        var totalRam = cs is null ? 0 : Math.Round((WmiUtil.U64(cs, "TotalPhysicalMemory") ?? 0) / Gb, 1);
        var freeKb = os is null ? 0 : WmiUtil.U64(os, "FreePhysicalMemory") ?? 0; // in kilobytes
        var freeRam = Math.Round(freeKb / (1024d * 1024d), 1);

        var lastBootDt = os is null ? null : WmiUtil.CimDate(os, "LastBootUpTime");
        var lastBoot = lastBootDt?.ToString("yyyy-MM-dd HH:mm") ?? "";
        var uptime = "";
        if (lastBootDt is not null)
        {
            var span = DateTime.Now - lastBootDt.Value;
            uptime = $"{(int)span.TotalDays}d {span.Hours}h";
        }

        os?.Dispose();
        cs?.Dispose();
        cpu?.Dispose();

        return new SystemInfo
        {
            Hostname = Environment.MachineName,
            OsCaption = caption,
            OsBuild = build,
            Manufacturer = manufacturer,
            Model = model,
            CpuName = cpuName,
            CpuCores = cores,
            TotalRamGb = totalRam,
            FreeRamGb = freeRam,
            LastBoot = lastBoot,
            Uptime = uptime,
            RuntimeVersion = Environment.Version.ToString(),
        };
    }
}

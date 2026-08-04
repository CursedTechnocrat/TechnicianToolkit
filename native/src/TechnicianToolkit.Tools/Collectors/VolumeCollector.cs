namespace TechnicianToolkit.Tools.Collectors;

/// <summary>A fixed volume's space usage, as SCRYER's disk-space pass records it.</summary>
public sealed class VolumeInfo
{
    public string Letter { get; init; } = "";
    public string Label { get; init; } = "";
    public double TotalGb { get; init; }
    public double UsedGb { get; init; }
    public double FreeGb { get; init; }
    public int PctUsed { get; init; }

    /// <summary>Health class for the progress bar: "ok", "warn" or "err".</summary>
    public string Health { get; init; } = "ok";
}

/// <summary>
/// Collects fixed-disk (DriveType=3) space usage from <c>Win32_LogicalDisk</c>,
/// applying SCRYER's thresholds: &gt;95% used → err, &gt;85% → warn, else ok.
/// </summary>
public static class VolumeCollector
{
    private const double Gb = 1024d * 1024d * 1024d;

    public static IReadOnlyList<VolumeInfo> Collect()
    {
        var rows = new List<VolumeInfo>();

        foreach (var disk in WmiUtil.QueryCimv2("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3"))
        {
            using (disk)
            {
                var size = (double)(WmiUtil.U64(disk, "Size") ?? 0);
                var free = (double)(WmiUtil.U64(disk, "FreeSpace") ?? 0);

                var totalGb = Math.Round(size / Gb, 1);
                var freeGb = Math.Round(free / Gb, 1);
                var usedGb = Math.Round((size - free) / Gb, 1);
                var pctUsed = size > 0 ? (int)Math.Round((size - free) / size * 100, MidpointRounding.AwayFromZero) : 0;
                var health = pctUsed > 95 ? "err" : pctUsed > 85 ? "warn" : "ok";

                rows.Add(new VolumeInfo
                {
                    Letter = WmiUtil.Str(disk, "DeviceID") ?? "",
                    Label = WmiUtil.Str(disk, "VolumeName") ?? "",
                    TotalGb = totalGb,
                    UsedGb = usedGb,
                    FreeGb = freeGb,
                    PctUsed = pctUsed,
                    Health = health,
                });
            }
        }

        return rows;
    }
}

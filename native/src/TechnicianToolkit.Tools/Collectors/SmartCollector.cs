using System.Management;

namespace TechnicianToolkit.Tools.Collectors;

/// <summary>A physical disk with health, SMART and identity fields (AUGUR/SCRYER).</summary>
public sealed class PhysicalDiskInfo
{
    public string DeviceId { get; init; } = "";
    public string FriendlyName { get; init; } = "";
    public string Serial { get; init; } = "N/A";
    public string Firmware { get; init; } = "N/A";
    public string MediaType { get; init; } = "";
    public string BusType { get; init; } = "";
    public double SizeGb { get; init; }
    public string HealthStatus { get; init; } = "";
    public string OperationalStatus { get; init; } = "";
    public int? Temperature { get; init; }
    public int? Wear { get; init; }

    /// <summary>Three-state SMART verdict: "N/A", "OK" or "FAILING".</summary>
    public string SmartPrediction { get; init; } = "N/A";

    /// <summary>SMART reason code (8-digit hex) or "N/A".</summary>
    public string SmartReason { get; init; } = "N/A";

    /// <summary>Health CSS class ("ok"/"warn"/"err") derived from HealthStatus.</summary>
    public string HealthClass =>
        HealthStatus.Contains("Healthy", StringComparison.OrdinalIgnoreCase) ? "ok"
        : HealthStatus.Contains("Warning", StringComparison.OrdinalIgnoreCase) ? "warn"
        : "err";
}

/// <summary>A volume as AUGUR's <c>Get-Volume</c> pass records it.</summary>
public sealed class SmartVolumeInfo
{
    public string Drive { get; init; } = "";
    public string Label { get; init; } = "";
    public string FileSystem { get; init; } = "";
    public double TotalGb { get; init; }
    public double FreeGb { get; init; }
    public double PctUsed { get; init; }
    public string Health { get; init; } = "";
    public string DriveType { get; init; } = "";
}

/// <summary>
/// Physical-disk health via the Storage WMI provider (<c>MSFT_PhysicalDisk</c>,
/// <c>MSFT_StorageReliabilityCounter</c>, <c>MSFT_Volume</c>), SMART failure
/// prediction via <c>root\wmi\MSStorageDriver_FailurePredictStatus</c>, and
/// serial/firmware via <c>Win32_DiskDrive</c>. Native equivalent of
/// <c>Get-PhysicalDisk</c> + <c>Get-StorageReliabilityCounter</c> + <c>Get-Volume</c>.
/// </summary>
/// <remarks>
/// SMART and reliability data are frequently unavailable on USB / NVMe / RAID /
/// virtual buses; every supplementary lookup is best-effort and degrades to
/// "N/A"/null exactly like the PowerShell try/catch blocks. The SMART↔disk match
/// is the same fragile <c>InstanceName contains DeviceId</c> heuristic AUGUR
/// uses, reproduced faithfully.
/// </remarks>
public static class SmartCollector
{
    private const string StorageScope = @"\\.\root\Microsoft\Windows\Storage";
    private const string WmiScope = @"\\.\root\wmi";
    private const double Gb = 1024d * 1024d * 1024d;

    private static readonly Dictionary<uint, string> MediaTypeMap = new()
    {
        [0] = "Unspecified",
        [3] = "HDD",
        [4] = "SSD",
        [5] = "SCM",
    };

    private static readonly Dictionary<uint, string> BusTypeMap = new()
    {
        [0] = "Unknown", [1] = "SCSI", [2] = "ATAPI", [3] = "ATA", [4] = "1394",
        [5] = "SSA", [6] = "Fibre Channel", [7] = "USB", [8] = "RAID", [9] = "iSCSI",
        [10] = "SAS", [11] = "SATA", [12] = "SD", [13] = "MMC", [15] = "File Backed Virtual",
        [16] = "Storage Spaces", [17] = "NVMe",
    };

    private static readonly Dictionary<uint, string> DiskHealthMap = new()
    {
        [0] = "Healthy", [1] = "Warning", [2] = "Unhealthy", [5] = "Unknown",
    };

    private static readonly Dictionary<uint, string> VolumeHealthMap = new()
    {
        [0] = "Healthy", [1] = "Warning", [2] = "Unhealthy",
    };

    private static readonly Dictionary<uint, string> DriveTypeMap = new()
    {
        [0] = "Unknown", [1] = "Invalid Root Path", [2] = "Removable", [3] = "Fixed",
        [4] = "Network", [5] = "CD-ROM", [6] = "RAM Disk",
    };

    /// <summary>
    /// Collects all physical disks with health, media/bus type, size, wear,
    /// temperature, SMART prediction and (best-effort) serial/firmware.
    /// </summary>
    public static IReadOnlyList<PhysicalDiskInfo> CollectPhysicalDisks()
    {
        var smart = LoadSmartPredictions();     // best-effort, may be empty
        var driveIdentity = LoadDiskDriveIdentity(); // Win32_DiskDrive by Index

        var result = new List<PhysicalDiskInfo>();

        foreach (var disk in WmiUtil.Query(StorageScope, "SELECT * FROM MSFT_PhysicalDisk"))
        {
            using (disk)
            {
                var deviceId = WmiUtil.Str(disk, "DeviceId") ?? "";
                var size = (double)(WmiUtil.U64(disk, "Size") ?? 0);
                var sizeGb = size > 0 ? Math.Round(size / Gb, 1) : 0;

                var mediaType = MapOrRaw(MediaTypeMap, WmiUtil.U32(disk, "MediaType"));
                var busType = MapOrRaw(BusTypeMap, WmiUtil.U32(disk, "BusType"));
                var health = MapOrRaw(DiskHealthMap, WmiUtil.U32(disk, "HealthStatus"));
                var opStatus = ReadOperationalStatus(disk);

                var (wear, temp) = ReadReliability(disk);

                // SMART match: first prediction whose InstanceName contains the
                // disk DeviceId (faithful to AUGUR's heuristic).
                var smartEntry = smart.FirstOrDefault(s =>
                    !string.IsNullOrEmpty(deviceId) &&
                    s.InstanceName.Contains(deviceId, StringComparison.OrdinalIgnoreCase));

                string smartLabel;
                string smartReason;
                if (smartEntry is null)
                {
                    smartLabel = "N/A";
                    smartReason = "N/A";
                }
                else
                {
                    smartLabel = smartEntry.PredictFailure ? "FAILING" : "OK";
                    smartReason = $"0x{smartEntry.Reason:X8}";
                }

                driveIdentity.TryGetValue(deviceId, out var identity);

                result.Add(new PhysicalDiskInfo
                {
                    DeviceId = deviceId,
                    FriendlyName = WmiUtil.Str(disk, "FriendlyName") ?? "",
                    Serial = identity.Serial ?? "N/A",
                    Firmware = identity.Firmware ?? "N/A",
                    MediaType = mediaType,
                    BusType = busType,
                    SizeGb = sizeGb,
                    HealthStatus = health,
                    OperationalStatus = opStatus,
                    Wear = wear,
                    Temperature = temp,
                    SmartPrediction = smartLabel,
                    SmartReason = smartReason,
                });
            }
        }

        return result;
    }

    /// <summary>Collects volumes with a drive letter or label (AUGUR's Get-Volume pass).</summary>
    public static IReadOnlyList<SmartVolumeInfo> CollectVolumes()
    {
        var result = new List<SmartVolumeInfo>();

        foreach (var vol in WmiUtil.Query(StorageScope, "SELECT * FROM MSFT_Volume"))
        {
            using (vol)
            {
                var letter = ReadDriveLetter(vol);
                var label = WmiUtil.Str(vol, "FileSystemLabel") ?? "";
                if (letter is null && string.IsNullOrEmpty(label))
                {
                    continue; // only volumes with a drive letter OR a label
                }

                var size = (double)(WmiUtil.U64(vol, "Size") ?? 0);
                var remaining = (double)(WmiUtil.U64(vol, "SizeRemaining") ?? 0);
                var totalGb = size > 0 ? Math.Round(size / Gb, 1) : 0;
                var freeGb = remaining > 0 ? Math.Round(remaining / Gb, 1) : 0;
                var pct = size > 0 ? Math.Round((size - remaining) / size * 100, 1) : 0;

                result.Add(new SmartVolumeInfo
                {
                    Drive = letter is not null ? $"{letter}:" : "(no letter)",
                    Label = label,
                    FileSystem = WmiUtil.Str(vol, "FileSystem") ?? "",
                    TotalGb = totalGb,
                    FreeGb = freeGb,
                    PctUsed = pct,
                    Health = MapOrRaw(VolumeHealthMap, WmiUtil.U32(vol, "HealthStatus")),
                    DriveType = MapOrRaw(DriveTypeMap, WmiUtil.U32(vol, "DriveType")),
                });
            }
        }

        return result;
    }

    private sealed record SmartEntry(string InstanceName, bool PredictFailure, uint Reason);

    private static List<SmartEntry> LoadSmartPredictions()
    {
        var list = new List<SmartEntry>();
        try
        {
            foreach (var s in WmiUtil.Query(WmiScope, "SELECT * FROM MSStorageDriver_FailurePredictStatus"))
            {
                using (s)
                {
                    list.Add(new SmartEntry(
                        WmiUtil.Str(s, "InstanceName") ?? "",
                        WmiUtil.Bool(s, "PredictFailure") ?? false,
                        WmiUtil.U32(s, "Reason") ?? 0));
                }
            }
        }
        catch
        {
            // SMART namespace unavailable on this bus — expected, not an error.
        }

        return list;
    }

    private static Dictionary<string, (string? Serial, string? Firmware)> LoadDiskDriveIdentity()
    {
        var map = new Dictionary<string, (string?, string?)>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (var d in WmiUtil.QueryCimv2("SELECT Index, SerialNumber, FirmwareRevision FROM Win32_DiskDrive"))
            {
                using (d)
                {
                    var index = WmiUtil.U32(d, "Index")?.ToString();
                    if (index is null)
                    {
                        continue;
                    }

                    var serial = WmiUtil.Str(d, "SerialNumber")?.Trim();
                    var firmware = WmiUtil.Str(d, "FirmwareRevision")?.Trim();
                    map[index] = (
                        string.IsNullOrEmpty(serial) ? null : serial,
                        string.IsNullOrEmpty(firmware) ? null : firmware);
                }
            }
        }
        catch
        {
            // Win32_DiskDrive unavailable — serial/firmware fall back to N/A.
        }

        return map;
    }

    private static (int? Wear, int? Temp) ReadReliability(ManagementObject disk)
    {
        try
        {
            foreach (ManagementObject rel in disk.GetRelated("MSFT_StorageReliabilityCounter"))
            {
                using (rel)
                {
                    var wear = WmiUtil.U32(rel, "Wear");
                    var temp = WmiUtil.U32(rel, "Temperature");
                    return ((int?)wear, (int?)temp);
                }
            }
        }
        catch
        {
            // No reliability counter for this disk — leave both null.
        }

        return (null, null);
    }

    private static string ReadOperationalStatus(ManagementObject disk)
    {
        try
        {
            if (disk["OperationalStatus"] is ushort[] codes && codes.Length > 0)
            {
                // The storage OperationalStatus value map is large and version
                // dependent; surface the common "2 = OK" and otherwise the raw
                // codes rather than risk a misleading friendly string.
                var parts = codes.Select(c => c == 2 ? "OK" : c.ToString());
                return string.Join(", ", parts);
            }
        }
        catch
        {
            // fall through
        }

        return "";
    }

    private static char? ReadDriveLetter(ManagementObject vol)
    {
        try
        {
            var raw = vol["DriveLetter"];
            switch (raw)
            {
                case null:
                    return null;
                case char c:
                    return c == '\0' ? null : c;
                default:
                    var num = Convert.ToUInt16(raw);
                    return num == 0 ? null : (char)num;
            }
        }
        catch
        {
            return null;
        }
    }

    private static string MapOrRaw(IReadOnlyDictionary<uint, string> map, uint? value)
    {
        if (value is null)
        {
            return "";
        }

        return map.TryGetValue(value.Value, out var s) ? s : value.Value.ToString();
    }
}

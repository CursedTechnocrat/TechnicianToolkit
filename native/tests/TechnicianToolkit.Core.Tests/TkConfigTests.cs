using TechnicianToolkit.Core.Config;
using Xunit;

namespace TechnicianToolkit.Core.Tests;

public class TkConfigTests : IDisposable
{
    private readonly string _tempDir;

    public TkConfigTests()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), "tktests_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);
        TkConfig.ConfigPath = Path.Combine(_tempDir, "config.json");
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_tempDir, recursive: true);
        }
        catch
        {
            // best effort
        }
    }

    [Fact]
    public void Get_MissingFile_ReturnsDefaults()
    {
        var cfg = TkConfig.Get();

        Assert.Equal(string.Empty, cfg.OrgName);
        Assert.Equal(string.Empty, cfg.LogDirectory);
        Assert.Equal(string.Empty, cfg.TeamsWebhook);
        Assert.Equal(string.Empty, cfg.Archive.DefaultDestination);
        Assert.Equal(string.Empty, cfg.Revenant.DefaultDestination);
        Assert.Equal(string.Empty, cfg.Covenant.DefaultTimezone);
    }

    [Fact]
    public void Set_TopLevel_RoundTrips()
    {
        TkConfig.Set("OrgName", "Contoso");
        Assert.Equal("Contoso", TkConfig.Get().OrgName);
    }

    [Fact]
    public void Set_Section_RoundTrips()
    {
        TkConfig.Set("DefaultDestination", @"\\srv\backups", "Archive");
        Assert.Equal(@"\\srv\backups", TkConfig.Get().Archive.DefaultDestination);
    }

    [Fact]
    public void Set_PreservesUnknownKeys()
    {
        File.WriteAllText(TkConfig.ConfigPath, "{\"OrgName\":\"Old\",\"CustomKey\":\"keepme\"}");

        TkConfig.Set("OrgName", "New");

        var raw = File.ReadAllText(TkConfig.ConfigPath);
        Assert.Contains("keepme", raw);
        Assert.Equal("New", TkConfig.Get().OrgName);
    }

    [Fact]
    public void Get_MigratesLegacyPhantomSectionToRevenant()
    {
        File.WriteAllText(TkConfig.ConfigPath,
            "{\"Phantom\":{\"DefaultDestination\":\"\\\\\\\\legacy\\\\path\"}}");

        var cfg = TkConfig.Get();

        Assert.Equal(@"\\legacy\path", cfg.Revenant.DefaultDestination);
    }

    [Fact]
    public void Get_RevenantWins_WhenBothPhantomAndRevenantPresent()
    {
        File.WriteAllText(TkConfig.ConfigPath,
            "{\"Phantom\":{\"DefaultDestination\":\"old\"},\"Revenant\":{\"DefaultDestination\":\"current\"}}");

        Assert.Equal("current", TkConfig.Get().Revenant.DefaultDestination);
    }

    [Fact]
    public void Get_UnparseableFile_ReturnsDefaults()
    {
        File.WriteAllText(TkConfig.ConfigPath, "{ this is not json ");
        Assert.Equal(string.Empty, TkConfig.Get().OrgName);
    }
}

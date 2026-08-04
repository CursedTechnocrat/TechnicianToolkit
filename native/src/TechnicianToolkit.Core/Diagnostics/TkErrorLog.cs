using System.Net.Http;
using System.Text;
using System.Text.Json;
using TechnicianToolkit.Core.Config;

namespace TechnicianToolkit.Core.Diagnostics;

/// <summary>
/// Structured error telemetry — a port of <c>Write-TKError</c>. Appends a JSONL
/// entry to a monthly log file in the configured log directory and, when a
/// <c>TeamsWebhook</c> is configured, POSTs a MessageCard to it. Never throws:
/// a logging helper must not interrupt its caller.
/// </summary>
public static class TkErrorLog
{
    private static readonly HttpClient Http = new();

    /// <summary>
    /// Logs a structured error and optionally notifies the configured Teams
    /// webhook. Any failure (disk, network, serialization) is swallowed.
    /// </summary>
    /// <param name="scriptName">Originating tool name (e.g. "scryer").</param>
    /// <param name="message">The error message.</param>
    /// <param name="category">Free-form category (default "General").</param>
    public static void Write(string scriptName, string message, string category = "General")
    {
        var entry = new ErrorEntry
        {
            Timestamp = DateTimeOffset.Now.ToString("o"),
            Script = scriptName,
            Category = category,
            Message = message,
            Host = EnvInfo.MachineName,
            User = EnvInfo.UserName,
        };

        TkConfigData cfg;
        try
        {
            cfg = TkConfig.Get();
        }
        catch
        {
            cfg = new TkConfigData();
        }

        // Append to the monthly JSONL error log.
        try
        {
            var logRoot = !string.IsNullOrWhiteSpace(cfg.LogDirectory) && Directory.Exists(cfg.LogDirectory)
                ? cfg.LogDirectory
                : TkPaths.BaseDirectory;

            var logFile = Path.Combine(logRoot, $"TK_Errors_{DateTime.Now:yyyyMM}.jsonl");
            var line = JsonSerializer.Serialize(entry);
            File.AppendAllText(logFile, line + Environment.NewLine, Encoding.UTF8);
        }
        catch
        {
            // Never throw from a logging helper.
        }

        // Optional Teams webhook notification.
        try
        {
            if (!string.IsNullOrWhiteSpace(cfg.TeamsWebhook))
            {
                var card = new
                {
                    type = "MessageCard",
                    context = "http://schema.org/extensions",
                    themeColor = "FF0000",
                    summary = $"TechnicianToolkit error in {scriptName}",
                    sections = new[]
                    {
                        new
                        {
                            activityTitle = $"TechnicianToolkit - {scriptName} [{category}]",
                            activitySubtitle = $"{entry.Host} / {entry.User}  |  {entry.Timestamp}",
                            activityText = message,
                        },
                    },
                };

                // Teams expects the '@type'/'@context' keys; serialize then fix
                // the two reserved names that C# can't express as identifiers.
                var json = JsonSerializer.Serialize(card)
                    .Replace("\"type\":", "\"@type\":")
                    .Replace("\"context\":", "\"@context\":");

                using var content = new StringContent(json, Encoding.UTF8, "application/json");
                Http.PostAsync(cfg.TeamsWebhook, content).GetAwaiter().GetResult();
            }
        }
        catch
        {
            // Webhook failures are silent — never interrupt the caller.
        }
    }

    private sealed class ErrorEntry
    {
        public string Timestamp { get; set; } = string.Empty;
        public string Script { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string Host { get; set; } = string.Empty;
        public string User { get; set; } = string.Empty;
    }
}

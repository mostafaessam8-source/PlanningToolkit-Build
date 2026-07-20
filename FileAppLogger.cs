using System.Text;

namespace PlanningToolkit.Infrastructure.Logging;

public sealed class FileAppLogger : IAppLogger
{
    private readonly object _syncRoot = new();
    private readonly string _logDirectory;
    private readonly Func<string> _minimumLevel;

    public FileAppLogger(string logDirectory, Func<string>? minimumLevel = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(logDirectory);
        _logDirectory = logDirectory;
        _minimumLevel = minimumLevel ?? (() => "Information");
    }

    public string CurrentLogFile => Path.Combine(_logDirectory, $"planning-toolkit-{DateTime.Now:yyyyMMdd}.log");

    public void Debug(string message) => Write("DEBUG", message, null);
    public void Information(string message) => Write("INFO", message, null);
    public void Warning(string message) => Write("WARN", message, null);
    public void Error(string message, Exception? exception = null) => Write("ERROR", message, exception);

    private void Write(string level, string message, Exception? exception)
    {
        if (!ShouldWrite(level))
            return;

        try
        {
            Directory.CreateDirectory(_logDirectory);
            var builder = new StringBuilder()
                .Append(DateTimeOffset.Now.ToString("O"))
                .Append(" [")
                .Append(level)
                .Append("] ")
                .AppendLine(Sanitize(message));

            if (exception is not null)
                builder.AppendLine(Sanitize(exception.ToString()));

            lock (_syncRoot)
                File.AppendAllText(CurrentLogFile, builder.ToString(), Encoding.UTF8);
        }
        catch
        {
            // Logging must never crash Excel or hide the original operation result.
        }
    }

    private bool ShouldWrite(string requestedLevel)
    {
        var ranks = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["DEBUG"] = 0,
            ["INFORMATION"] = 1,
            ["INFO"] = 1,
            ["WARNING"] = 2,
            ["WARN"] = 2,
            ["ERROR"] = 3
        };

        var configured = _minimumLevel();
        var configuredRank = ranks.TryGetValue(configured, out var value) ? value : 1;
        var requestedRank = ranks.TryGetValue(requestedLevel, out value) ? value : 1;
        return requestedRank >= configuredRank;
    }

    private static string Sanitize(string value) => value
        .Replace("\r", " ", StringComparison.Ordinal)
        .Replace("\n", Environment.NewLine, StringComparison.Ordinal);
}


namespace PlanningToolkit.Infrastructure.Logging;

public interface IAppLogger
{
    string CurrentLogFile { get; }
    void Debug(string message);
    void Information(string message);
    void Warning(string message);
    void Error(string message, Exception? exception = null);
}


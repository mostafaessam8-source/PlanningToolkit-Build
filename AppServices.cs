using PlanningToolkit.Core;
using PlanningToolkit.Infrastructure;
using PlanningToolkit.Infrastructure.Logging;
using PlanningToolkit.Infrastructure.Settings;

namespace PlanningToolkit.Excel;

internal static class AppServices
{
    private static AppSettings _currentSettings = new();

    public static IAppLogger Logger { get; } = new FileAppLogger(
        AppPaths.LogDirectory,
        () => _currentSettings.LoggingLevel);

    public static ISettingsStore SettingsStore { get; } = new JsonSettingsStore(
        AppPaths.SettingsFile,
        Logger);

    public static AppSettings CurrentSettings => _currentSettings;

    public static void Initialize() => ReloadSettings();

    public static void ReloadSettings() => _currentSettings = SettingsStore.Load();
}


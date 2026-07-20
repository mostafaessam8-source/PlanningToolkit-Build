namespace PlanningToolkit.Infrastructure;

public static class AppPaths
{
    public static string RoamingDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "PlanningToolkit");

    public static string LocalDirectory => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "PlanningToolkit");

    public static string SettingsFile => Path.Combine(RoamingDirectory, "settings.json");
    public static string LogDirectory => Path.Combine(LocalDirectory, "Logs");
}


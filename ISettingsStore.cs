using PlanningToolkit.Core;

namespace PlanningToolkit.Infrastructure.Settings;

public interface ISettingsStore
{
    AppSettings Load();
    void Save(AppSettings settings);
}


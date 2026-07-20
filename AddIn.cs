using ExcelDna.Integration;

namespace PlanningToolkit.Excel;

public sealed class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        AppServices.Initialize();
        AppServices.Logger.Information("Planning Toolkit 0.1.0 loaded.");
    }

    public void AutoClose()
    {
        AppServices.Logger.Information("Planning Toolkit 0.1.0 unloaded.");
    }
}


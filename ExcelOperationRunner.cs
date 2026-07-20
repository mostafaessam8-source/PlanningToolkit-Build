using System.Windows.Forms;
using ExcelDna.Integration;

namespace PlanningToolkit.Excel.Excel;

internal static class ExcelOperationRunner
{
    public static void Run(string commandName, Action<dynamic> operation)
    {
        try
        {
            dynamic application = ExcelDnaUtil.Application
                ?? throw new InvalidOperationException("Microsoft Excel is not available.");

            using var state = new ExcelApplicationState(
                application,
                $"Planning Toolkit: {commandName}...",
                AppServices.CurrentSettings.CalculationBehavior);
            AppServices.Logger.Information($"Command started: {commandName}.");
            operation(application);
            AppServices.Logger.Information($"Command completed: {commandName}.");
            application.StatusBar = $"Planning Toolkit: {commandName} completed.";
        }
        catch (Exception exception)
        {
            AppServices.Logger.Error($"Command failed: {commandName}.", exception);
            MessageBox.Show(
                $"{commandName} could not be completed.\n\n{exception.Message}\n\nTechnical details were written to the log.",
                "Planning Toolkit",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }
}

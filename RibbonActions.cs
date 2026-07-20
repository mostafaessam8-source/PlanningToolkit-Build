using System.Diagnostics;
using System.Windows.Forms;
using PlanningToolkit.Core.Text;
using PlanningToolkit.Excel.Excel;
using PlanningToolkit.Excel.UI;
using PlanningToolkit.Infrastructure;

namespace PlanningToolkit.Excel;

internal static class RibbonActions
{
    public static void FillDown() => ExcelOperationRunner.Run("Fill Down", application =>
    {
        ExcelSelectionTools.FillDown(application);
        ShowCompleted("Fill Down completed.");
    });

    public static void TrimSpaces() => RunTextTransform("Trim", TextTransforms.TrimWhitespace);
    public static void CleanText() => RunTextTransform("Clean Text", TextTransforms.RemoveNonPrintingCharacters);
    public static void UpperCase() => RunTextTransform("Upper Case", value => TextTransforms.ToUpper(value));
    public static void LowerCase() => RunTextTransform("Lower Case", value => TextTransforms.ToLower(value));
    public static void ProperCase() => RunTextTransform("Proper Case", value => TextTransforms.ToProperCase(value));
    public static void SentenceCase() => RunTextTransform("Sentence Case", value => TextTransforms.ToSentenceCase(value));

    public static void TextToDate() => ExcelOperationRunner.Run("Text to Date", application =>
    {
        var count = ExcelSelectionTools.ConvertTextToDates(application, AppServices.CurrentSettings.DateFormat);
        ShowCompleted($"Converted {count:N0} cell(s) to dates.");
    });

    public static void SplitText() => ExcelOperationRunner.Run("Split Text", application =>
    {
        var count = ExcelSelectionTools.SplitText(application);
        if (count > 0)
            ShowCompleted($"Split {count:N0} row(s).");
    });

    public static void MergeText() => ExcelOperationRunner.Run("Merge Text", application =>
    {
        var count = ExcelSelectionTools.MergeText(application);
        if (count > 0)
            ShowCompleted($"Merged {count:N0} row(s).");
    });

    public static void UniqueValues() => ExcelOperationRunner.Run("Unique Values", application =>
    {
        var count = ExcelSelectionTools.CreateUniqueValuesSheet(application);
        ShowCompleted($"Created a new sheet with {count:N0} unique value(s).");
    });

    public static void RemoveBlankRows() => ExcelOperationRunner.Run("Remove Blank Rows", application =>
    {
        var count = ExcelSelectionTools.RemoveBlankRows(application);
        ShowCompleted(count == 0 ? "No blank rows were removed." : $"Removed {count:N0} blank row(s).");
    });

    public static void ShowSettings()
    {
        try
        {
            using var form = new SettingsForm(AppServices.CurrentSettings, AppServices.SettingsStore);
            if (form.ShowDialog() == DialogResult.OK)
            {
                AppServices.ReloadSettings();
                AppServices.Logger.Information("Settings reloaded after user update.");
            }
        }
        catch (Exception exception)
        {
            AppServices.Logger.Error("Settings window failed.", exception);
            ShowError(exception.Message);
        }
    }

    public static void ViewLogs()
    {
        try
        {
            Directory.CreateDirectory(AppPaths.LogDirectory);
            Process.Start(new ProcessStartInfo
            {
                FileName = AppPaths.LogDirectory,
                UseShellExecute = true
            });
        }
        catch (Exception exception)
        {
            AppServices.Logger.Error("Log folder could not be opened.", exception);
            ShowError(exception.Message);
        }
    }

    public static void OpenUserGuide()
    {
        try
        {
            var candidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "docs", "UserGuide.md"),
                Path.Combine(AppContext.BaseDirectory, "UserGuide.md"),
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "docs", "UserGuide.md"))
            };
            var path = candidates.FirstOrDefault(File.Exists)
                ?? throw new FileNotFoundException("UserGuide.md was not found beside the add-in.");

            Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
        catch (Exception exception)
        {
            AppServices.Logger.Error("User guide could not be opened.", exception);
            ShowError(exception.Message);
        }
    }

    public static void ShowAbout() => MessageBox.Show(
        "Planning Toolkit\nVersion 0.1.0 — Phase 1\n\n" +
        "Independent Excel productivity foundation for project controls.\n" +
        "XER, WBS, Gantt, comparison and PMS modules are scheduled for later phases.",
        "About Planning Toolkit",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information);

    private static void RunTextTransform(string commandName, Func<string, string> transform) =>
        ExcelOperationRunner.Run(commandName, application =>
        {
            var count = ExcelSelectionTools.TransformText(application, transform);
            ShowCompleted($"Updated {count:N0} cell(s).");
        });

    private static void ShowCompleted(string message) => MessageBox.Show(
        message,
        "Planning Toolkit",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information);

    private static void ShowError(string message) => MessageBox.Show(
        message,
        "Planning Toolkit",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error);
}


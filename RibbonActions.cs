using System.Diagnostics;
using System.Windows.Forms;
using PlanningToolkit.Core.Text;
using PlanningToolkit.Excel.Excel;
using PlanningToolkit.Excel.UI;
using PlanningToolkit.Infrastructure;

namespace PlanningToolkit.Excel;

internal static class RibbonActions
{
    public static void CreatePms() => ExcelOperationRunner.Run("Create PMS Dashboard", application =>
    {
        var baselinePath = XerExcelTools.SelectXerFile("Select BASELINE XER for PMS (1 of 2)");
        if (baselinePath is null)
            return;
        var updatePath = XerExcelTools.SelectXerFile("Select UPDATE XER for PMS (2 of 2)");
        if (updatePath is null)
            return;
        if (string.Equals(baselinePath, updatePath, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Select two different XER files: baseline first, then update.");
        var lookAheadWeeks = SelectLookAheadWeeks();
        if (!lookAheadWeeks.HasValue)
            return;
        var result = XerPmsReport.Create((object)application, baselinePath, updatePath, lookAheadWeeks.Value);
        ShowCompleted($"PMS Dashboard created. Planned: {result.Planned:P1}, Actual: {result.Actual:P1}, SPI: {result.Spi:N2}, {lookAheadWeeks}-week Look Ahead: {result.LookAhead:N0}, Critical: {result.Critical:N0}.");
    });

    public static void CompareXer() => ExcelOperationRunner.Run("Compare Baseline and Update XER", application =>
    {
        var baselinePath = XerExcelTools.SelectXerFile("Select BASELINE XER file (1 of 2)");
        if (baselinePath is null)
            return;
        var updatePath = XerExcelTools.SelectXerFile("Select UPDATE XER file (2 of 2)");
        if (updatePath is null)
            return;
        if (string.Equals(baselinePath, updatePath, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Select two different XER files: baseline first, then update.");
        var result = XerComparisonReport.Create((object)application, baselinePath, updatePath);
        ShowCompleted($"Compared {result.Activities:N0} activities. Delayed: {result.Delayed:N0}, added: {result.Added:N0}, deleted: {result.Deleted:N0}. Project finish variance: {result.FinishVariance:N1} day(s).");
    });

    public static void BuildSchedule() => ExcelOperationRunner.Run("Build WBS Schedule", application =>
    {
        var path = XerExcelTools.SelectXerFile();
        if (path is null)
            return;
        var result = XerScheduleReport.Create((object)application, path);
        ShowCompleted($"Created WBS Schedule and Gantt Chart for {result.Activities:N0} activities, {result.WbsLevels:N0} WBS level(s), including {result.CriticalActivities:N0} critical activities.");
    });

    public static void ImportXer() => ExcelOperationRunner.Run("Import XER", application =>
    {
        var path = XerExcelTools.SelectXerFile();
        if (path is null)
            return;
        var result = XerExcelTools.Import((object)application, path);
        ShowCompleted($"Imported {result.Tables:N0} table(s) and {result.Rows:N0} row(s) into a new workbook.");
    });

    public static void ExportXer() => ExcelOperationRunner.Run("Export XER", application =>
    {
        var result = XerExportTools.ExportActiveWorkbook(application);
        if (result is null)
            return;
        var editedTables = result.EditedTables.Count == 0
            ? "No editable values changed"
            : $"Edited tables: {string.Join(", ", result.EditedTables)}";
        ShowCompleted(
            $"Validated XER exported successfully.\n\n" +
            $"File: {result.OutputPath}\n" +
            $"Tables: {result.Tables:N0}, rows: {result.Rows:N0}, edited cells: {result.EditedCells:N0}\n" +
            $"{editedTables}\n" +
            $"Automatic backup(s): {result.BackupPaths.Count:N0}");
    });

    public static void ExportReportPdf() => ExcelOperationRunner.Run("Export Report as PDF", application =>
    {
        var outputPath = XerReportExport.ExportActiveWorkbookToPdf((object)application);
        if (outputPath is null)
            return;
        ShowCompleted($"Report exported to PDF.\n\nFile: {outputPath}");
    });

    public static void ValidateXer() => ExcelOperationRunner.Run("Validate XER", application =>
    {
        var path = XerExcelTools.SelectXerFile();
        if (path is null)
            return;
        var result = XerExcelTools.ValidateToWorkbook((object)application, path);
        ShowCompleted(result.IsValid
            ? $"XER is valid. Found {result.WarningCount:N0} warning(s)."
            : $"XER is invalid. Found {result.ErrorCount:N0} error(s) and {result.WarningCount:N0} warning(s). See the validation workbook.");
    });

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
        "Planning Toolkit\nVersion 0.7.0 — Phase 6 + Professional Report Theme\n\n" +
        "Primavera XER import/edit/export, WBS/Gantt, comparison, PMS dashboard, S-curve, look-ahead and critical path toolkit.\n" +
        "Export XER includes read-only ID protection, backups, validation and exact round-trip verification.\n" +
        "All reports share a single branded theme (ReportTheme.cs): consistent fonts, colors, borders, print setup and one-click PDF export.",
        "About Planning Toolkit",
        MessageBoxButtons.OK,
        MessageBoxIcon.Information);

    private static int? SelectLookAheadWeeks()
    {
        while (true)
        {
            var value = PromptDialog.Show(
                "Look Ahead Period",
                "Enter the Look Ahead duration in weeks (1 to 52):",
                "2");
            if (value is null)
                return null;
            if (int.TryParse(value.Trim(), out var weeks) && weeks is >= 1 and <= 52)
                return weeks;

            MessageBox.Show(
                "Enter a whole number between 1 and 52.",
                "Planning Toolkit",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

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

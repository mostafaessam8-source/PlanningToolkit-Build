using System.Windows.Forms;

namespace PlanningToolkit.Excel.Excel;

/// <summary>
/// Exports the active Planning Toolkit report workbook (WBS Schedule, Comparison, PMS Dashboard,
/// etc.) to a single branded PDF using the print setup already applied by <see cref="ReportTheme"/>
/// to every sheet.
/// </summary>
internal static class XerReportExport
{
    public static string? ExportActiveWorkbookToPdf(dynamic application)
    {
        dynamic workbook = application.ActiveWorkbook
            ?? throw new InvalidOperationException("Open a Planning Toolkit report workbook first.");

        var suggestedName = SanitizeFileName((string)workbook.Name) + ".pdf";
        var outputPath = SelectOutputPath(suggestedName);
        if (outputPath is null)
            return null;

        ReportTheme.ExportSheetToPdf(workbook, outputPath);
        return outputPath;
    }

    private static string SanitizeFileName(string workbookName)
    {
        var name = Path.GetFileNameWithoutExtension(workbookName);
        foreach (var invalid in Path.GetInvalidFileNameChars())
            name = name.Replace(invalid, '_');
        return name;
    }

    private static string? SelectOutputPath(string suggestedName)
    {
        using var dialog = new SaveFileDialog
        {
            Title = "Export report as PDF",
            Filter = "PDF files (*.pdf)|*.pdf",
            DefaultExt = "pdf",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = suggestedName,
            InitialDirectory = Directory.Exists(AppServices.CurrentSettings.DefaultOutputFolder)
                ? AppServices.CurrentSettings.DefaultOutputFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };
        return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
    }
}

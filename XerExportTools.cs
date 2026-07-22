using System.Globalization;
using System.Windows.Forms;
using PlanningToolkit.Core.Xer;

namespace PlanningToolkit.Excel.Excel;

internal sealed record XerWorkbookExportResult(
    string OutputPath,
    IReadOnlyList<string> BackupPaths,
    int Tables,
    int Rows,
    int EditedCells,
    IReadOnlyList<string> EditedTables);

internal static class XerExportTools
{
    public static XerWorkbookExportResult? ExportActiveWorkbook(dynamic application)
    {
        if ((int)application.Workbooks.Count == 0)
            throw new InvalidOperationException("Open a workbook created by Import XER first.");

        dynamic workbook = application.ActiveWorkbook
            ?? throw new InvalidOperationException("No active workbook is available.");
        dynamic summary = FindWorksheet(workbook, "XER Summary")
            ?? throw new InvalidOperationException("The active workbook does not contain the 'XER Summary' sheet created by Import XER.");

        var sourcePath = Convert.ToString(summary.Cells[2, 2].Value2, CultureInfo.InvariantCulture)?.Trim();
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
        {
            sourcePath = XerExcelTools.SelectXerFile("Select the ORIGINAL XER used to create this workbook");
            if (sourcePath is null)
                return null;
        }

        var document = XerParser.Parse(sourcePath);
        var sourceValidation = XerValidator.Validate(document);
        if (!sourceValidation.IsValid)
            throw new InvalidOperationException(
                $"The original XER contains {sourceValidation.ErrorCount:N0} critical validation error(s). Validate it before export.");

        var snapshots = new List<XerTableSnapshot>(document.Tables.Count);
        foreach (var table in document.Tables)
        {
            var sheetName = FindMappedSheetName(summary, table.Name) ?? table.Name;
            dynamic sheet = FindWorksheet(workbook, sheetName)
                ?? throw new InvalidOperationException($"Worksheet '{sheetName}' for XER table '{table.Name}' is missing.");
            snapshots.Add(ReadSnapshot(sheet, table));
        }

        var editResult = XerEditService.Apply(document, snapshots);
        var outputPath = SelectOutputPath(sourcePath);
        if (outputPath is null)
            return null;

        var fileResult = XerRoundTripExporter.Export(document, sourcePath, outputPath);
        return new XerWorkbookExportResult(
            fileResult.OutputPath,
            fileResult.BackupPaths,
            fileResult.Tables,
            fileResult.Rows,
            editResult.EditedCells,
            editResult.EditedTables);
    }

    private static XerTableSnapshot ReadSnapshot(dynamic sheet, XerTable sourceTable)
    {
        var expectedRows = sourceTable.Rows.Count + 1;
        var expectedColumns = sourceTable.Fields.Count;
        dynamic usedRange = sheet.UsedRange;
        var usedRow = (int)usedRange.Row;
        var usedColumn = (int)usedRange.Column;
        var usedRows = (int)usedRange.Rows.Count;
        var usedColumns = (int)usedRange.Columns.Count;
        if (usedRow != 1 || usedColumn != 1 || usedRows != expectedRows || usedColumns != expectedColumns)
            throw new InvalidOperationException(
                $"Worksheet '{sheet.Name}' structure changed. Expected {expectedRows:N0} row(s) and {expectedColumns:N0} column(s), but found {usedRows:N0} row(s) and {usedColumns:N0} column(s). Do not insert or delete XER rows/columns in v0.6.0.");

        dynamic range = sheet.Cells[1, 1].Resize[expectedRows, expectedColumns];
        object? raw = range.Value2;
        var fields = new string[expectedColumns];
        for (var column = 0; column < expectedColumns; column++)
            fields[column] = ToXerText(ReadValue(raw, 0, column, expectedRows, expectedColumns), null);

        var rows = new List<IReadOnlyList<string>>(sourceTable.Rows.Count);
        for (var row = 1; row < expectedRows; row++)
        {
            var values = new string[expectedColumns];
            for (var column = 0; column < expectedColumns; column++)
                values[column] = ToXerText(ReadValue(raw, row, column, expectedRows, expectedColumns), fields[column]);
            rows.Add(values);
        }
        return new XerTableSnapshot(sourceTable.Name, fields, rows);
    }

    private static object? ReadValue(object? raw, int row, int column, int rowCount, int columnCount)
    {
        if (rowCount == 1 && columnCount == 1)
            return raw;
        return ((object[,])raw!)[row + 1, column + 1];
    }

    private static string ToXerText(object? value, string? fieldName)
    {
        if (value is null)
            return string.Empty;
        if (value is string text)
            return text;
        if (value is DateTime date)
            return date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        if (value is double number && fieldName?.EndsWith("_date", StringComparison.OrdinalIgnoreCase) == true)
        {
            try
            {
                return DateTime.FromOADate(number).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            }
            catch (ArgumentException)
            {
                return number.ToString(CultureInfo.InvariantCulture);
            }
        }
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    private static string? SelectOutputPath(string sourcePath)
    {
        var sourceDirectory = Path.GetDirectoryName(sourcePath);
        using var dialog = new SaveFileDialog
        {
            Title = "Export validated Primavera XER",
            Filter = "Primavera XER files (*.xer)|*.xer",
            DefaultExt = "xer",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = Path.GetFileNameWithoutExtension(sourcePath) + "-Edited-v0.6.0.xer",
            InitialDirectory = Directory.Exists(sourceDirectory)
                ? sourceDirectory
                : AppServices.CurrentSettings.DefaultOutputFolder
        };
        return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
    }

    private static string? FindMappedSheetName(dynamic summary, string tableName)
    {
        var usedRows = (int)summary.UsedRange.Rows.Count;
        for (var row = 7; row <= usedRows; row++)
        {
            var mappedTable = Convert.ToString(summary.Cells[row, 1].Value2, CultureInfo.InvariantCulture);
            if (!string.Equals(mappedTable, tableName, StringComparison.OrdinalIgnoreCase))
                continue;
            var mappedSheet = Convert.ToString(summary.Cells[row, 4].Value2, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(mappedSheet) ? null : mappedSheet.Trim();
        }
        return null;
    }

    private static dynamic? FindWorksheet(dynamic workbook, string name)
    {
        foreach (dynamic sheet in workbook.Worksheets)
        {
            if (string.Equals((string)sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                return sheet;
        }
        return null;
    }
}

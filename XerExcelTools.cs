using System.Globalization;
using System.Windows.Forms;
using PlanningToolkit.Core.Xer;

namespace PlanningToolkit.Excel.Excel;

internal static class XerExcelTools
{
    private const int ExcelMaximumRows = 1_048_576;
    private const int ExcelMaximumColumns = 16_384;

    public static string? SelectXerFile(string title = "Select Primavera XER file")
    {
        using var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = "Primavera XER files (*.xer)|*.xer|All files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };
        return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
    }

    public static (int Tables, int Rows) Import(dynamic application, string path)
    {
        var document = XerParser.Parse(path);
        var validation = XerValidator.Validate(document);
        if (!validation.IsValid)
            throw new InvalidOperationException($"XER validation found {validation.ErrorCount:N0} error(s). Use Validate XER for details.");

        dynamic workbook = application.Workbooks.Add();
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var sheetNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var importedRows = 0;
        var firstSheet = true;

        foreach (var table in document.Tables)
        {
            if (table.Fields.Count > ExcelMaximumColumns)
                throw new InvalidOperationException($"Table '{table.Name}' exceeds Excel's column limit.");
            if (table.Rows.Count + 1 > ExcelMaximumRows)
                throw new InvalidOperationException($"Table '{table.Name}' exceeds Excel's row limit.");

            dynamic sheet = firstSheet ? workbook.Worksheets[1] : workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            firstSheet = false;
            sheet.Name = CreateUniqueSheetName(table.Name, usedNames);
            sheetNames[table.Name] = sheet.Name;
            WriteTable(sheet, table);
            importedRows += table.Rows.Count;
        }

        dynamic summary = workbook.Worksheets.Add(Before: workbook.Worksheets[1]);
        summary.Name = CreateUniqueSheetName("XER Summary", usedNames);
        WriteSummary(summary, path, document, validation, sheetNames);
        summary.Activate();
        return (document.Tables.Count, importedRows);
    }

    public static XerValidationResult ValidateToWorkbook(dynamic application, string path)
    {
        var document = XerParser.Parse(path);
        var result = XerValidator.Validate(document);
        dynamic workbook = application.Workbooks.Add();
        dynamic sheet = workbook.Worksheets[1];
        sheet.Name = "XER Validation";
        WriteValidation(sheet, path, document, result);
        return result;
    }

    private static void WriteTable(dynamic sheet, XerTable table)
    {
        var rowCount = table.Rows.Count + 1;
        var columnCount = table.Fields.Count;
        var values = new object?[rowCount, columnCount];
        for (var column = 0; column < columnCount; column++)
            values[0, column] = table.Fields[column];
        for (var row = 0; row < table.Rows.Count; row++)
        for (var column = 0; column < columnCount; column++)
            values[row + 1, column] = column < table.Rows[row].Values.Count ? table.Rows[row].Values[column] : null;

        dynamic target = sheet.Cells[1, 1].Resize[rowCount, columnCount];
        target.NumberFormat = "@";
        target.Value2 = values;
        sheet.Rows[1].Font.Bold = true;
        for (var column = 0; column < table.Fields.Count; column++)
        {
            dynamic header = sheet.Cells[1, column + 1];
            var editable = XerEditPolicy.IsEditable(table.Name, table.Fields[column]);
            header.Interior.Color = editable ? 0xCCF2FF : 0xD9D9D9;
            if (editable && table.Rows.Count > 0)
                sheet.Cells[2, column + 1].Resize[table.Rows.Count, 1].Interior.Color = 0xE6F4FF;
        }
        sheet.Rows[1].AutoFilter();
        sheet.Application.ActiveWindow.SplitRow = 1;
        sheet.Application.ActiveWindow.FreezePanes = true;
        sheet.UsedRange.Columns.AutoFit();
    }

    private static void WriteSummary(
        dynamic sheet,
        string path,
        XerDocument document,
        XerValidationResult validation,
        IReadOnlyDictionary<string, string> sheetNames)
    {
        var values = new object?[document.Tables.Count + 6, 4];
        values[0, 0] = "Planning Toolkit — XER Import Summary";
        values[1, 0] = "Source file";
        values[1, 1] = path;
        values[2, 0] = "Imported at";
        values[2, 1] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        values[3, 0] = "Tables";
        values[3, 1] = document.Tables.Count;
        values[4, 0] = "Validation warnings";
        values[4, 1] = validation.WarningCount;
        values[5, 0] = "Table";
        values[5, 1] = "Rows";
        values[5, 2] = "Fields";
        values[5, 3] = "Worksheet";
        for (var index = 0; index < document.Tables.Count; index++)
        {
            values[index + 6, 0] = document.Tables[index].Name;
            values[index + 6, 1] = document.Tables[index].Rows.Count;
            values[index + 6, 2] = document.Tables[index].Fields.Count;
            values[index + 6, 3] = sheetNames[document.Tables[index].Name];
        }
        dynamic target = sheet.Cells[1, 1].Resize[document.Tables.Count + 6, 4];
        target.Value2 = values;
        sheet.Rows[1].Font.Bold = true;
        sheet.Rows[6].Font.Bold = true;
        sheet.UsedRange.Columns.AutoFit();
    }

    private static void WriteValidation(dynamic sheet, string path, XerDocument document, XerValidationResult result)
    {
        var issueRows = Math.Max(1, result.Issues.Count);
        var values = new object?[issueRows + 5, 5];
        values[0, 0] = "Planning Toolkit — XER Validation";
        values[1, 0] = "Source file";
        values[1, 1] = path;
        values[2, 0] = "Result";
        values[2, 1] = result.IsValid ? "VALID" : "INVALID";
        values[2, 2] = $"{result.ErrorCount} error(s), {result.WarningCount} warning(s), {document.Tables.Count} table(s)";
        values[4, 0] = "Severity";
        values[4, 1] = "Code";
        values[4, 2] = "Table";
        values[4, 3] = "Line";
        values[4, 4] = "Message";
        if (result.Issues.Count == 0)
            values[5, 0] = "No issues found";
        for (var index = 0; index < result.Issues.Count; index++)
        {
            var issue = result.Issues[index];
            values[index + 5, 0] = issue.Severity.ToString();
            values[index + 5, 1] = issue.Code;
            values[index + 5, 2] = issue.Table;
            values[index + 5, 3] = issue.LineNumber;
            values[index + 5, 4] = issue.Message;
        }
        dynamic target = sheet.Cells[1, 1].Resize[issueRows + 5, 5];
        target.Value2 = values;
        sheet.Rows[1].Font.Bold = true;
        sheet.Rows[5].Font.Bold = true;
        sheet.UsedRange.Columns.AutoFit();
    }

    private static string CreateUniqueSheetName(string value, ISet<string> used)
    {
        var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        var baseName = new string(value.Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim();
        if (baseName.Length == 0)
            baseName = "XER Table";
        if (baseName.Length > 31)
            baseName = baseName[..31];
        var candidate = baseName;
        var suffix = 2;
        while (!used.Add(candidate))
        {
            var ending = $"_{suffix++}";
            candidate = baseName[..Math.Min(baseName.Length, 31 - ending.Length)] + ending;
        }
        return candidate;
    }
}

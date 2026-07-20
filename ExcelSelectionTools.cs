using System.Globalization;
using System.Windows.Forms;
using PlanningToolkit.Core.Text;
using PlanningToolkit.Excel.UI;

namespace PlanningToolkit.Excel.Excel;

internal static class ExcelSelectionTools
{
    private const long MaximumCellsPerOperation = 1_000_000;

    public static void FillDown(dynamic application)
    {
        dynamic range = RequireSingleAreaRange(application);
        if ((int)range.Rows.Count < 2)
            throw new InvalidOperationException("Select at least two rows before using Fill Down.");

        range.FillDown();
    }

    public static int TransformText(dynamic application, Func<string, string> transform)
    {
        dynamic range = RequireSingleAreaRange(application);
        var (data, rowCount, columnCount) = ReadRange(range);
        var changed = 0;

        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                if (data[row, column] is not string text || IsFormula(text))
                    continue;

                var transformed = transform(text);
                if (string.Equals(text, transformed, StringComparison.Ordinal))
                    continue;

                data[row, column] = transformed;
                changed++;
            }
        }

        WriteRange(range, data, rowCount, columnCount);
        return changed;
    }

    public static int ConvertTextToDates(dynamic application, string numberFormat)
    {
        dynamic range = RequireSingleAreaRange(application);
        var (data, rowCount, columnCount) = ReadRange(range);
        var converted = 0;

        for (var row = 0; row < rowCount; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                if (data[row, column] is not string text || IsFormula(text))
                    continue;

                if (!DateParser.TryParse(text, out var date))
                    continue;

                data[row, column] = date.ToOADate();
                converted++;
            }
        }

        WriteRange(range, data, rowCount, columnCount);
        if (converted > 0)
            range.NumberFormat = numberFormat;

        return converted;
    }

    public static int SplitText(dynamic application)
    {
        dynamic range = RequireSingleAreaRange(application);
        var columnCount = (int)range.Columns.Count;
        if (columnCount != 1)
            throw new InvalidOperationException("Split Text requires a single selected column.");

        var delimiterText = PromptDialog.Show("Delimiter", "Enter the delimiter. Use \\t for a tab:", ",");
        if (delimiterText is null)
            return 0;

        var delimiter = delimiterText == "\\t" ? "\t" : delimiterText;
        if (delimiter.Length == 0)
            throw new InvalidOperationException("Delimiter cannot be empty.");

        var (source, rowCount, _) = ReadValues(range);
        var rows = new string[rowCount][];
        var maximumParts = 1;
        for (var row = 0; row < rowCount; row++)
        {
            var value = Convert.ToString(source[row, 0], CultureInfo.CurrentCulture) ?? string.Empty;
            rows[row] = value.Split(new[] { delimiter }, StringSplitOptions.None);
            maximumParts = Math.Max(maximumParts, rows[row].Length);
        }

        dynamic outputRange = range.Cells[1, 1].Resize[rowCount, maximumParts];
        if (maximumParts > 1 && MessageBox.Show(
                $"This will write into {maximumParts} columns starting at the selected column. Continue?",
                "Planning Toolkit",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) != DialogResult.Yes)
            return 0;

        var output = new object?[rowCount, maximumParts];
        for (var row = 0; row < rowCount; row++)
        for (var column = 0; column < rows[row].Length; column++)
            output[row, column] = rows[row][column];

        WriteValues(outputRange, output, rowCount, maximumParts);
        return rowCount;
    }

    public static int MergeText(dynamic application)
    {
        dynamic range = RequireSingleAreaRange(application);
        var delimiterText = PromptDialog.Show("Separator", "Enter the separator for merged values:", " - ");
        if (delimiterText is null)
            return 0;

        var (source, rowCount, columnCount) = ReadValues(range);
        if (columnCount < 2)
            throw new InvalidOperationException("Select at least two columns before using Merge Text.");

        if (MessageBox.Show(
                "Merged results will be written in the first column immediately to the right of the selection. Continue?",
                "Planning Toolkit",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) != DialogResult.Yes)
            return 0;

        var output = new object?[rowCount, 1];
        for (var row = 0; row < rowCount; row++)
        {
            var values = new List<string>();
            for (var column = 0; column < columnCount; column++)
            {
                var text = Convert.ToString(source[row, column], CultureInfo.CurrentCulture);
                if (!string.IsNullOrEmpty(text))
                    values.Add(text);
            }

            output[row, 0] = string.Join(delimiterText, values);
        }

        dynamic outputRange = range.Offset[0, columnCount].Resize[rowCount, 1];
        WriteValues(outputRange, output, rowCount, 1);
        return rowCount;
    }

    public static int CreateUniqueValuesSheet(dynamic application)
    {
        dynamic range = RequireSingleAreaRange(application);
        var (source, rowCount, columnCount) = ReadValues(range);
        var values = new List<string>();
        var seen = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

        for (var row = 0; row < rowCount; row++)
        for (var column = 0; column < columnCount; column++)
        {
            var value = Convert.ToString(source[row, column], CultureInfo.CurrentCulture)?.Trim();
            if (!string.IsNullOrEmpty(value) && seen.Add(value))
                values.Add(value);
        }

        if (values.Count == 0)
            throw new InvalidOperationException("The selection does not contain any non-empty values.");

        dynamic workbook = application.ActiveWorkbook
            ?? throw new InvalidOperationException("Open a workbook before running this command.");
        dynamic sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
        sheet.Name = $"Unique_{DateTime.Now:HHmmssfff}";

        var output = new object?[values.Count + 1, 1];
        output[0, 0] = "Unique Value";
        for (var index = 0; index < values.Count; index++)
            output[index + 1, 0] = values[index];

        dynamic outputRange = sheet.Cells[1, 1].Resize[values.Count + 1, 1];
        WriteValues(outputRange, output, values.Count + 1, 1);
        sheet.Cells[1, 1].Font.Bold = true;
        sheet.Columns[1].AutoFit();
        return values.Count;
    }

    public static int RemoveBlankRows(dynamic application)
    {
        dynamic range = RequireSingleAreaRange(application);
        var (source, rowCount, columnCount) = ReadRange(range);
        var blankRows = new List<int>();

        for (var row = 0; row < rowCount; row++)
        {
            var blank = true;
            for (var column = 0; column < columnCount; column++)
            {
                if (!string.IsNullOrWhiteSpace(Convert.ToString(source[row, column], CultureInfo.CurrentCulture)))
                {
                    blank = false;
                    break;
                }
            }

            if (blank)
                blankRows.Add(row + 1);
        }

        if (blankRows.Count == 0)
            return 0;

        if (MessageBox.Show(
                $"Delete {blankRows.Count} entire worksheet row(s) that are blank within the selected columns?",
                "Planning Toolkit",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) != DialogResult.Yes)
            return 0;

        for (var index = blankRows.Count - 1; index >= 0; index--)
            range.Rows[blankRows[index]].EntireRow.Delete();

        return blankRows.Count;
    }

    private static dynamic RequireSingleAreaRange(dynamic application)
    {
        dynamic range = application.Selection
            ?? throw new InvalidOperationException("Select a worksheet range before running this command.");

        try
        {
            _ = range.Address;
            _ = range.Areas.Count;
            _ = range.Rows.Count;
        }
        catch
        {
            throw new InvalidOperationException("The current Excel selection is not a cell range.");
        }

        if ((int)range.Areas.Count != 1)
            throw new InvalidOperationException("Select one continuous cell range.");

        var cellCount = Convert.ToInt64(range.CountLarge, CultureInfo.InvariantCulture);
        if (cellCount > MaximumCellsPerOperation)
            throw new InvalidOperationException($"The selection contains {cellCount:N0} cells. Select no more than {MaximumCellsPerOperation:N0} cells per operation.");

        return range;
    }

    private static (object?[,] Data, int Rows, int Columns) ReadRange(dynamic range)
    {
        var rowCount = (int)range.Rows.Count;
        var columnCount = (int)range.Columns.Count;
        object? raw = range.Formula;

        if (raw is object[,] matrix)
        {
            var result = new object?[rowCount, columnCount];
            var rowLowerBound = matrix.GetLowerBound(0);
            var columnLowerBound = matrix.GetLowerBound(1);
            for (var row = 0; row < rowCount; row++)
            for (var column = 0; column < columnCount; column++)
                result[row, column] = matrix[row + rowLowerBound, column + columnLowerBound];

            return (result, rowCount, columnCount);
        }

        var single = new object?[1, 1];
        single[0, 0] = raw;
        return (single, 1, 1);
    }

    private static (object?[,] Data, int Rows, int Columns) ReadValues(dynamic range)
    {
        var rowCount = (int)range.Rows.Count;
        var columnCount = (int)range.Columns.Count;
        object? raw = range.Value2;

        if (raw is object[,] matrix)
        {
            var result = new object?[rowCount, columnCount];
            var rowLowerBound = matrix.GetLowerBound(0);
            var columnLowerBound = matrix.GetLowerBound(1);
            for (var row = 0; row < rowCount; row++)
            for (var column = 0; column < columnCount; column++)
                result[row, column] = matrix[row + rowLowerBound, column + columnLowerBound];

            return (result, rowCount, columnCount);
        }

        var single = new object?[1, 1];
        single[0, 0] = raw;
        return (single, 1, 1);
    }

    private static void WriteRange(dynamic range, object?[,] data, int rowCount, int columnCount)
    {
        range.Formula = rowCount == 1 && columnCount == 1 ? data[0, 0] : data;
    }

    private static void WriteValues(dynamic range, object?[,] data, int rowCount, int columnCount)
    {
        range.Value2 = rowCount == 1 && columnCount == 1 ? data[0, 0] : data;
    }

    private static bool IsFormula(string value) => value.StartsWith('=');
}

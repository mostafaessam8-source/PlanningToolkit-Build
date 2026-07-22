using System.Globalization;
using PlanningToolkit.Core.Text;

namespace PlanningToolkit.Core.Xer;

public sealed record XerTableSnapshot(
    string Name,
    IReadOnlyList<string> Fields,
    IReadOnlyList<IReadOnlyList<string>> Rows);

public sealed record XerEditResult(int EditedCells, IReadOnlyList<string> EditedTables);

public static class XerEditPolicy
{
    private static readonly IReadOnlyDictionary<string, HashSet<string>> EditableFields =
        new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["PROJECT"] = Fields(
                "proj_short_name", "proj_name", "plan_start_date", "plan_end_date", "scd_end_date",
                "anticip_start_date", "anticip_end_date", "last_recalc_date"),
            ["PROJWBS"] = Fields("wbs_short_name", "wbs_name", "seq_num"),
            ["TASK"] = Fields(
                "task_code", "task_name", "status_code", "task_type", "duration_type", "complete_pct_type",
                "target_drtn_hr_cnt", "remain_drtn_hr_cnt", "phys_complete_pct", "complete_pct",
                "target_start_date", "target_end_date", "act_start_date", "act_end_date", "expect_end_date",
                "early_start_date", "early_end_date", "late_start_date", "late_end_date", "restart_date", "reend_date",
                "total_float_hr_cnt", "free_float_hr_cnt", "cstr_type", "cstr_date", "cstr_type2", "cstr_date2",
                "priority_type", "suspend_date", "resume_date"),
            ["TASKPRED"] = Fields("pred_type", "lag_hr_cnt"),
            ["TASKRSRC"] = Fields(
                "target_qty", "remain_qty", "act_reg_qty", "act_ot_qty", "target_cost", "remain_cost",
                "act_reg_cost", "act_ot_cost", "cost_per_qty", "remain_cost_per_qty"),
            ["RSRC"] = Fields(
                "rsrc_short_name", "rsrc_name", "email_addr", "office_phone", "other_phone", "active_flag",
                "cost_qty_type", "ot_factor", "def_qty_per_hr"),
            ["CALENDAR"] = Fields(
                "clndr_name", "default_flag", "clndr_type", "day_hr_cnt", "week_hr_cnt", "month_hr_cnt", "year_hr_cnt"),
            ["ACTVTYPE"] = Fields("actv_code_type", "actv_short_len", "seq_num"),
            ["ACTVCODE"] = Fields("actv_code_name", "short_name", "seq_num", "color"),
            ["UDFVALUE"] = Fields("udf_text", "udf_number", "udf_date")
        };

    public static bool IsEditable(string tableName, string fieldName) =>
        EditableFields.TryGetValue(tableName, out var fields) && fields.Contains(fieldName);

    public static IReadOnlyList<string> GetEditableFields(string tableName) =>
        EditableFields.TryGetValue(tableName, out var fields)
            ? fields.OrderBy(field => field, StringComparer.OrdinalIgnoreCase).ToArray()
            : Array.Empty<string>();

    private static HashSet<string> Fields(params string[] names) => new(names, StringComparer.OrdinalIgnoreCase);
}

public static class XerEditService
{
    public static XerEditResult Apply(XerDocument original, IEnumerable<XerTableSnapshot> snapshots)
    {
        ArgumentNullException.ThrowIfNull(original);
        ArgumentNullException.ThrowIfNull(snapshots);

        var snapshotMap = snapshots.ToDictionary(snapshot => snapshot.Name, StringComparer.OrdinalIgnoreCase);
        var editedCells = 0;
        var editedTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var table in original.Tables)
        {
            if (!snapshotMap.TryGetValue(table.Name, out var snapshot))
                throw new InvalidOperationException($"Worksheet for XER table '{table.Name}' is missing.");

            ValidateStructure(table, snapshot);
            for (var rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var sourceRow = table.Rows[rowIndex];
                var editedRow = snapshot.Rows[rowIndex];
                var values = sourceRow.Values.ToArray();

                for (var columnIndex = 0; columnIndex < table.Fields.Count; columnIndex++)
                {
                    var originalValue = values[columnIndex] ?? string.Empty;
                    var editedValue = editedRow[columnIndex] ?? string.Empty;
                    if (string.Equals(originalValue, editedValue, StringComparison.Ordinal))
                        continue;

                    var field = table.Fields[columnIndex];
                    if (!XerEditPolicy.IsEditable(table.Name, field))
                        throw new InvalidOperationException(
                            $"{table.Name}.{field} is read-only. Restore row {rowIndex + 2} to '{originalValue}'.");

                    ValidateEditedValue(table.Name, field, editedValue, rowIndex + 2);
                    values[columnIndex] = editedValue;
                    editedCells++;
                    editedTables.Add(table.Name);
                }

                table.Rows[rowIndex] = new XerRow(sourceRow.LineNumber, values);
            }
        }

        var validation = XerValidator.Validate(original);
        if (!validation.IsValid)
            throw new InvalidOperationException(BuildValidationMessage(validation));

        return new XerEditResult(
            editedCells,
            editedTables.OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray());
    }

    private static void ValidateStructure(XerTable table, XerTableSnapshot snapshot)
    {
        if (snapshot.Fields.Count != table.Fields.Count ||
            !snapshot.Fields.SequenceEqual(table.Fields, StringComparer.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"Worksheet '{snapshot.Name}' column headers were changed. Restore the original XER field names and order.");

        if (snapshot.Rows.Count != table.Rows.Count)
            throw new InvalidOperationException(
                $"Worksheet '{snapshot.Name}' contains {snapshot.Rows.Count:N0} data row(s), but the source XER contains {table.Rows.Count:N0}. Adding or deleting XER rows is not supported in v0.6.0.");

        var badRow = snapshot.Rows.Select((row, index) => (row, index)).FirstOrDefault(item => item.row.Count != table.Fields.Count);
        if (badRow.row is not null)
            throw new InvalidOperationException(
                $"Worksheet '{snapshot.Name}' row {badRow.index + 2} contains {badRow.row.Count:N0} value(s); {table.Fields.Count:N0} are required.");
    }

    private static void ValidateEditedValue(string tableName, string fieldName, string value, int worksheetRow)
    {
        if (value.IndexOfAny(['\t', '\r', '\n']) >= 0)
            throw new InvalidOperationException(
                $"{tableName}.{fieldName} at worksheet row {worksheetRow} contains a tab or line break, which is not valid in an XER cell.");

        if (string.IsNullOrWhiteSpace(value))
            return;

        if (IsDateField(fieldName) && !DateParser.TryParse(value, out _))
            throw new InvalidOperationException(
                $"{tableName}.{fieldName} at worksheet row {worksheetRow} is not a valid date: '{value}'.");

        if (IsNumericField(fieldName))
        {
            if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
                throw new InvalidOperationException(
                    $"{tableName}.{fieldName} at worksheet row {worksheetRow} must be a number using a dot as the decimal separator.");

            if (fieldName.EndsWith("_pct", StringComparison.OrdinalIgnoreCase) && (number < 0 || number > 100))
                throw new InvalidOperationException(
                    $"{tableName}.{fieldName} at worksheet row {worksheetRow} must be between 0 and 100.");
        }
    }

    private static bool IsDateField(string fieldName) =>
        fieldName.EndsWith("_date", StringComparison.OrdinalIgnoreCase);

    private static bool IsNumericField(string fieldName) =>
        fieldName.EndsWith("_cnt", StringComparison.OrdinalIgnoreCase) ||
        fieldName.EndsWith("_qty", StringComparison.OrdinalIgnoreCase) ||
        fieldName.EndsWith("_cost", StringComparison.OrdinalIgnoreCase) ||
        fieldName.EndsWith("_pct", StringComparison.OrdinalIgnoreCase) ||
        fieldName.EndsWith("_num", StringComparison.OrdinalIgnoreCase) ||
        fieldName is "seq_num" or "ot_factor" or "def_qty_per_hr" or "cost_per_qty" or "remain_cost_per_qty";

    private static string BuildValidationMessage(XerValidationResult validation)
    {
        var details = validation.Issues
            .Where(issue => issue.Severity == XerIssueSeverity.Error)
            .Take(5)
            .Select(issue => $"{issue.Code}: {issue.Message}");
        return $"The edited XER contains {validation.ErrorCount:N0} critical validation error(s):\n" + string.Join("\n", details);
    }
}

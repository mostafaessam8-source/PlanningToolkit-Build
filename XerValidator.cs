namespace PlanningToolkit.Core.Xer;

public static class XerValidator
{
    private static readonly string[] RecommendedTables = ["PROJECT", "PROJWBS", "TASK"];

    public static XerValidationResult Validate(XerDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        var result = new XerValidationResult();

        if (string.IsNullOrWhiteSpace(document.Header) || !document.Header.StartsWith("ERMHDR\t", StringComparison.Ordinal))
            Add(result, XerIssueSeverity.Error, "HEADER_MISSING", "The file does not contain a valid ERMHDR header.");
        if (document.Tables.Count == 0)
            Add(result, XerIssueSeverity.Error, "NO_TABLES", "The file does not contain any XER tables.");

        foreach (var duplicate in document.Tables.GroupBy(t => t.Name, StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
            Add(result, XerIssueSeverity.Error, "DUPLICATE_TABLE", $"Table '{duplicate.Key}' appears more than once.", duplicate.Key);

        foreach (var table in document.Tables)
        {
            if (table.Fields.Count == 0)
            {
                Add(result, XerIssueSeverity.Error, "FIELDS_MISSING", $"Table '{table.Name}' has no %F field definition.", table.Name, table.LineNumber);
                continue;
            }

            foreach (var duplicate in table.Fields.GroupBy(f => f, StringComparer.OrdinalIgnoreCase).Where(g => !string.IsNullOrWhiteSpace(g.Key) && g.Count() > 1))
                Add(result, XerIssueSeverity.Error, "DUPLICATE_FIELD", $"Field '{duplicate.Key}' appears more than once.", table.Name, table.LineNumber);

            foreach (var row in table.Rows.Where(row => row.Values.Count != table.Fields.Count))
                Add(result, XerIssueSeverity.Error, "COLUMN_COUNT", $"Row has {row.Values.Count} value(s), but table '{table.Name}' defines {table.Fields.Count} field(s).", table.Name, row.LineNumber);

            foreach (var row in table.Rows)
            foreach (var value in row.Values.Where(value => value.IndexOfAny(['\t', '\r', '\n']) >= 0))
                Add(result, XerIssueSeverity.Error, "INVALID_CELL_TEXT", "A value contains a tab or line break that cannot be serialized safely.", table.Name, row.LineNumber);

            if (table.Rows.Count == 0)
                Add(result, XerIssueSeverity.Warning, "EMPTY_TABLE", $"Table '{table.Name}' contains no rows.", table.Name, table.LineNumber);
        }

        foreach (var name in RecommendedTables.Where(name => document.FindTable(name) is null))
            Add(result, XerIssueSeverity.Warning, "CORE_TABLE_MISSING", $"Common scheduling table '{name}' is not present.", name);

        ValidateUniqueId(document, result, "PROJECT", "proj_id");
        ValidateUniqueId(document, result, "PROJWBS", "wbs_id");
        ValidateUniqueId(document, result, "TASK", "task_id");
        ValidateUniqueId(document, result, "CALENDAR", "clndr_id");
        ValidateUniqueId(document, result, "RSRC", "rsrc_id");

        ValidateReference(document, result, "PROJWBS", "proj_id", "PROJECT", "proj_id", allowEmpty: false);
        ValidateReference(document, result, "PROJWBS", "parent_wbs_id", "PROJWBS", "wbs_id", allowEmpty: true, missingSeverity: XerIssueSeverity.Warning);
        ValidateReference(document, result, "TASK", "proj_id", "PROJECT", "proj_id", allowEmpty: false);
        ValidateReference(document, result, "TASK", "wbs_id", "PROJWBS", "wbs_id", allowEmpty: true, missingSeverity: XerIssueSeverity.Warning);
        ValidateReference(document, result, "TASK", "clndr_id", "CALENDAR", "clndr_id", allowEmpty: true);
        ValidateReference(document, result, "TASKPRED", "task_id", "TASK", "task_id", allowEmpty: false);
        ValidateReference(document, result, "TASKPRED", "pred_task_id", "TASK", "task_id", allowEmpty: false);
        ValidateReference(document, result, "TASKRSRC", "task_id", "TASK", "task_id", allowEmpty: false);
        ValidateReference(document, result, "TASKRSRC", "rsrc_id", "RSRC", "rsrc_id", allowEmpty: true);
        ValidateSelfPredecessors(document, result);
        return result;
    }

    private static void ValidateReference(
        XerDocument document,
        XerValidationResult result,
        string childTableName,
        string childFieldName,
        string parentTableName,
        string parentFieldName,
        bool allowEmpty,
        XerIssueSeverity missingSeverity = XerIssueSeverity.Error)
    {
        var child = document.FindTable(childTableName);
        var parent = document.FindTable(parentTableName);
        if (child is null || parent is null)
            return;

        var childIndex = FindField(child, childFieldName);
        var parentIndex = FindField(parent, parentFieldName);
        if (childIndex < 0 || parentIndex < 0)
            return;

        var parentValues = parent.Rows
            .Where(row => parentIndex < row.Values.Count && !string.IsNullOrWhiteSpace(row.Values[parentIndex]))
            .Select(row => row.Values[parentIndex])
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var row in child.Rows)
        {
            if (childIndex >= row.Values.Count)
                continue;
            var value = row.Values[childIndex];
            if (string.IsNullOrWhiteSpace(value))
            {
                if (!allowEmpty)
                    Add(result, XerIssueSeverity.Error, "REFERENCE_EMPTY", $"Field '{childFieldName}' is empty.", childTableName, row.LineNumber);
                continue;
            }
            if (!parentValues.Contains(value))
                Add(result, missingSeverity, "REFERENCE_MISSING", $"{childFieldName} value '{value}' does not exist in {parentTableName}.{parentFieldName}.", childTableName, row.LineNumber);
        }
    }

    private static void ValidateSelfPredecessors(XerDocument document, XerValidationResult result)
    {
        var table = document.FindTable("TASKPRED");
        if (table is null)
            return;
        var taskIndex = FindField(table, "task_id");
        var predecessorIndex = FindField(table, "pred_task_id");
        if (taskIndex < 0 || predecessorIndex < 0)
            return;

        foreach (var row in table.Rows)
        {
            if (taskIndex >= row.Values.Count || predecessorIndex >= row.Values.Count)
                continue;
            if (string.Equals(row.Values[taskIndex], row.Values[predecessorIndex], StringComparison.OrdinalIgnoreCase))
                Add(result, XerIssueSeverity.Error, "SELF_PREDECESSOR", $"Task '{row.Values[taskIndex]}' cannot be its own predecessor.", table.Name, row.LineNumber);
        }
    }

    private static int FindField(XerTable table, string fieldName) =>
        table.Fields.ToList().FindIndex(field => string.Equals(field, fieldName, StringComparison.OrdinalIgnoreCase));

    private static void ValidateUniqueId(XerDocument document, XerValidationResult result, string tableName, string fieldName)
    {
        var table = document.FindTable(tableName);
        if (table is null)
            return;
        var index = table.Fields.ToList().FindIndex(field => string.Equals(field, fieldName, StringComparison.OrdinalIgnoreCase));
        if (index < 0)
        {
            Add(result, XerIssueSeverity.Warning, "ID_FIELD_MISSING", $"Table '{tableName}' does not contain expected field '{fieldName}'.", tableName, table.LineNumber);
            return;
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in table.Rows)
        {
            if (index >= row.Values.Count || string.IsNullOrWhiteSpace(row.Values[index]))
            {
                Add(result, XerIssueSeverity.Error, "ID_EMPTY", $"Field '{fieldName}' is empty.", tableName, row.LineNumber);
                continue;
            }
            if (!seen.Add(row.Values[index]))
                Add(result, XerIssueSeverity.Error, "ID_DUPLICATE", $"Duplicate {fieldName} value '{row.Values[index]}'.", tableName, row.LineNumber);
        }
    }

    private static void Add(XerValidationResult result, XerIssueSeverity severity, string code, string message, string? table = null, int? line = null) =>
        result.Issues.Add(new XerValidationIssue(severity, code, message, table, line));
}

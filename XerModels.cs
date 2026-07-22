namespace PlanningToolkit.Core.Xer;

public sealed record XerRow(int LineNumber, IReadOnlyList<string> Values);

public sealed class XerTable
{
    public XerTable(string name, int lineNumber)
    {
        Name = name;
        LineNumber = lineNumber;
    }

    public string Name { get; }
    public int LineNumber { get; }
    public IReadOnlyList<string> Fields { get; internal set; } = Array.Empty<string>();
    public List<XerRow> Rows { get; } = new();
}

public sealed class XerDocument
{
    public string? Header { get; internal set; }
    public List<XerTable> Tables { get; } = new();

    public XerTable? FindTable(string name) => Tables.FirstOrDefault(
        table => string.Equals(table.Name, name, StringComparison.OrdinalIgnoreCase));
}

public enum XerIssueSeverity
{
    Warning,
    Error
}

public sealed record XerValidationIssue(
    XerIssueSeverity Severity,
    string Code,
    string Message,
    string? Table = null,
    int? LineNumber = null);

public sealed class XerValidationResult
{
    public List<XerValidationIssue> Issues { get; } = new();
    public int ErrorCount => Issues.Count(issue => issue.Severity == XerIssueSeverity.Error);
    public int WarningCount => Issues.Count(issue => issue.Severity == XerIssueSeverity.Warning);
    public bool IsValid => ErrorCount == 0;
}

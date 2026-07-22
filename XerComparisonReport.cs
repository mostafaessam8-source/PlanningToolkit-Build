using PlanningToolkit.Core.Xer;

namespace PlanningToolkit.Excel.Excel;

internal static class XerComparisonReport
{
    private const int HeaderRow = 5;
    private const int FirstDataRow = 6;

    public static (int Activities, int Delayed, int Added, int Deleted, double FinishVariance) Create(dynamic application, string baselinePath, string updatePath)
    {
        var baselineDocument = XerParser.Parse(baselinePath);
        var updateDocument = XerParser.Parse(updatePath);
        EnsureValid(baselineDocument, "Baseline");
        EnsureValid(updateDocument, "Update");
        var hours = AppServices.CurrentSettings.WorkingHoursPerDay;
        var baseline = XerScheduleBuilder.Build(baselineDocument, hours, XerScheduleDateMode.Baseline);
        var update = XerScheduleBuilder.Build(updateDocument, hours, XerScheduleDateMode.Current);
        var comparison = XerScheduleComparer.Compare(baseline, update);

        dynamic workbook = application.Workbooks.Add();
        dynamic summary = workbook.Worksheets[1];
        summary.Name = "Comparison Summary";
        WriteSummary(summary, baselinePath, updatePath, comparison);
        dynamic details = workbook.Worksheets.Add(After: summary);
        details.Name = "Activity Comparison";
        WriteDetails(details, comparison, comparison.Activities);
        dynamic delayed = workbook.Worksheets.Add(After: details);
        delayed.Name = "Delayed Activities";
        WriteDetails(delayed, comparison, comparison.Activities.Where(item => item.FinishVarianceDays > 0.001 || item.UpdateCritical).ToList());
        summary.Activate();
        return (comparison.Activities.Count, comparison.DelayedCount, comparison.AddedCount, comparison.DeletedCount, comparison.ProjectFinishVarianceDays);
    }

    private static void EnsureValid(XerDocument document, string label)
    {
        var result = XerValidator.Validate(document);
        if (!result.IsValid)
            throw new InvalidOperationException($"{label} XER validation found {result.ErrorCount:N0} error(s). Validate the file first.");
    }

    private static void WriteSummary(dynamic sheet, string baselinePath, string updatePath, XerScheduleComparison comparison)
    {
        var values = new object?[16, 4];
        values[0, 0] = "Planning Toolkit — Baseline vs Update Comparison";
        values[2, 0] = "Baseline file";
        values[2, 1] = baselinePath;
        values[3, 0] = "Update file";
        values[3, 1] = updatePath;
        values[5, 0] = "Metric";
        values[5, 1] = "Baseline";
        values[5, 2] = "Update";
        values[5, 3] = "Variance / Count";
        values[6, 0] = "Project";
        values[6, 1] = comparison.Baseline.ProjectName;
        values[6, 2] = comparison.Update.ProjectName;
        values[7, 0] = "Data Date";
        values[7, 1] = comparison.Baseline.DataDate?.ToOADate();
        values[7, 2] = comparison.Update.DataDate?.ToOADate();
        values[8, 0] = "Project Finish";
        values[8, 1] = comparison.BaselineProjectFinish?.ToOADate();
        values[8, 2] = comparison.UpdateProjectFinish?.ToOADate();
        values[8, 3] = comparison.ProjectFinishVarianceDays;
        values[9, 0] = "Activities";
        values[9, 1] = comparison.Baseline.Activities.Count;
        values[9, 2] = comparison.Update.Activities.Count;
        values[10, 0] = "Delayed activities";
        values[10, 3] = comparison.DelayedCount;
        values[11, 0] = "Added activities";
        values[11, 3] = comparison.AddedCount;
        values[12, 0] = "Deleted activities";
        values[12, 3] = comparison.DeletedCount;
        values[13, 0] = "Current critical activities";
        values[13, 3] = comparison.CriticalCount;
        var costWeighted = XerProgressWeighting.HasBaselineCost(comparison.Activities);
        values[1, 0] = "Weight basis";
        values[1, 1] = costWeighted
            ? $"Baseline Budgeted Cost | BAC: {XerProgressWeighting.TotalBaselineCost(comparison.Activities):N2}"
            : "Duration fallback — no Baseline TASKRSRC/PROJCOST cost found";
        var overallPlanned = WeightedProgress(comparison.Activities, item => item.PlannedPercent, costWeighted);
        var overallActual = WeightedProgress(comparison.Activities, item => item.ActualPercent, costWeighted);
        values[14, 0] = "Overall Progress";
        values[14, 1] = overallPlanned;
        values[14, 2] = overallActual;
        values[14, 3] = overallActual - overallPlanned;
        values[15, 0] = comparison.ProjectFinishVarianceDays > 0
            ? $"Update finish is {comparison.ProjectFinishVarianceDays:N1} calendar day(s) later than baseline."
            : comparison.ProjectFinishVarianceDays < 0
                ? $"Update finish is {Math.Abs(comparison.ProjectFinishVarianceDays):N1} calendar day(s) earlier than baseline."
                : "Update project finish matches baseline.";

        dynamic target = sheet.Cells[1, 1].Resize[16, 4];
        target.Value2 = values;
        ReportTheme.ApplyBodyFont(sheet.Range["A1:D16"]);
        sheet.Range["A1:D1"].Merge();
        ReportTheme.ApplyTitle(sheet.Range["A1:D1"]);
        ReportTheme.ApplySubtitle(sheet.Range["A2:D3"]);
        sheet.Range["A6:D6"].Font.Bold = true;
        ReportTheme.ApplyHeaderRow(sheet.Range["A6:D6"]);
        sheet.Range["B8:C9"].NumberFormat = "dd-mmm-yyyy";
        sheet.Range["D9"].NumberFormat = "0.0 \"days\"";
        sheet.Range["B15:D15"].NumberFormat = "0.0%";
        ReportTheme.ApplyGridBorders(sheet.Range["A6:D14"]);
        sheet.Range["A16:D16"].Merge();
        sheet.Range["A16:D16"].Font.Bold = true;
        sheet.Range["A16:D16"].Interior.Color = comparison.ProjectFinishVarianceDays > 0 ? ReportTheme.NegativeFill : ReportTheme.PositiveFill;
        sheet.Columns[1].ColumnWidth = 32;
        sheet.Columns[2].ColumnWidth = 55;
        sheet.Columns[3].ColumnWidth = 28;
        sheet.Columns[4].ColumnWidth = 20;
        ReportTheme.ApplyPrintSetup(sheet, "Comparison Summary", landscape: false);
    }

    private static void WriteDetails(dynamic sheet, XerScheduleComparison comparison, IReadOnlyList<XerActivityComparison> activities)
    {
        string[] headers = ["Change", "WBS Path", "Activity ID", "Activity Name", "Baseline Status", "Update Status", "Baseline Start", "Update Start", "Start Variance", "Baseline Finish", "Update Finish", "Finish Variance", "Baseline Duration", "Update Duration", "Duration Variance", "Baseline Float", "Update Float", "Float Variance", "Planned %", "Actual %", "Progress Variance", "Baseline Critical", "Update Critical", "Assessment"];
        var layout = BuildComparisonLayout(comparison, activities);
        var costWeighted = XerProgressWeighting.HasBaselineCost(comparison.Activities);
        sheet.Cells[1, 1] = "Activity Comparison — WBS Roll-up";
        sheet.Cells[2, 1] = costWeighted
            ? $"Planned/Actual roll-up basis: Baseline Budgeted Cost | BAC: {XerProgressWeighting.TotalBaselineCost(comparison.Activities):N2}"
            : "Planned/Actual roll-up basis: Duration fallback — no Baseline TASKRSRC/PROJCOST cost found";
        var rows = layout.Rows;
        var values = new object?[rows.Count + 1, headers.Length];
        for (var column = 0; column < headers.Length; column++) values[0, column] = headers[column];
        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            if (row.IsSummary)
            {
                var summaryItems = row.SummaryItems!;
                values[index + 1, 0] = "WBS Summary";
                values[index + 1, 1] = row.WbsPath;
                values[index + 1, 3] = new string(' ', Math.Max(0, row.Level - 1) * 3) + row.WbsName;
                values[index + 1, 11] = summaryItems.Where(item => item.FinishVarianceDays > 0).Select(item => item.FinishVarianceDays).DefaultIfEmpty(0d).Max();
                values[index + 1, 18] = WeightedProgress(summaryItems, item => item.PlannedPercent, costWeighted);
                values[index + 1, 19] = WeightedProgress(summaryItems, item => item.ActualPercent, costWeighted);
                values[index + 1, 20] = (double)values[index + 1, 19]! - (double)values[index + 1, 18]!;
                values[index + 1, 23] = $"{summaryItems.Count:N0} activities | {summaryItems.Count(item => item.FinishVarianceDays > 0.001):N0} delayed | {summaryItems.Count(item => item.UpdateCritical):N0} critical | {summaryItems.Count(item => item.ChangeType == XerActivityChangeType.Added):N0} added | {summaryItems.Count(item => item.ChangeType == XerActivityChangeType.Deleted):N0} deleted";
                continue;
            }
            var item = row.Item!;
            values[index + 1, 0] = item.ChangeType.ToString();
            values[index + 1, 1] = item.WbsPath;
            values[index + 1, 2] = item.ActivityCode;
            values[index + 1, 3] = new string(' ', Math.Max(0, row.Level - 1) * 3) + item.ActivityName;
            values[index + 1, 4] = item.BaselineStatus;
            values[index + 1, 5] = item.UpdateStatus;
            values[index + 1, 6] = item.BaselineStart?.ToOADate();
            values[index + 1, 7] = item.UpdateStart?.ToOADate();
            values[index + 1, 8] = item.StartVarianceDays;
            values[index + 1, 9] = item.BaselineFinish?.ToOADate();
            values[index + 1, 10] = item.UpdateFinish?.ToOADate();
            values[index + 1, 11] = item.FinishVarianceDays;
            values[index + 1, 12] = item.BaselineDurationDays;
            values[index + 1, 13] = item.UpdateDurationDays;
            values[index + 1, 14] = item.DurationVarianceDays;
            values[index + 1, 15] = item.BaselineFloatDays;
            values[index + 1, 16] = item.UpdateFloatDays;
            values[index + 1, 17] = item.FloatVarianceDays;
            values[index + 1, 18] = item.PlannedPercent;
            values[index + 1, 19] = item.ActualPercent;
            values[index + 1, 20] = item.ProgressVariance;
            values[index + 1, 21] = item.BaselineCritical ? "Yes" : string.Empty;
            values[index + 1, 22] = item.UpdateCritical ? "Yes" : string.Empty;
            values[index + 1, 23] = Assess(item);
        }
        ReportTheme.ApplyTitle(sheet.Range["A1"]);
        ReportTheme.ApplySubtitle(sheet.Range["A2"]);
        dynamic target = sheet.Cells[HeaderRow, 1].Resize[rows.Count + 1, headers.Length];
        target.Value2 = values;
        ReportTheme.ApplyHeaderRow(sheet.Range[$"A{HeaderRow}:X{HeaderRow}"]);
        if (rows.Count > 0)
        {
            var lastRow = FirstDataRow + rows.Count - 1;
            ReportTheme.ApplyBodyFont(sheet.Range[$"A{FirstDataRow}:X{lastRow}"]);
            ReportTheme.ApplyGridBorders(sheet.Range[$"A{HeaderRow}:X{lastRow}"]);
            sheet.Range[$"G{FirstDataRow}:H{lastRow}"].NumberFormat = "dd-mmm-yyyy";
            sheet.Range[$"J{FirstDataRow}:K{lastRow}"].NumberFormat = "dd-mmm-yyyy";
            sheet.Range[$"I{FirstDataRow}:R{lastRow}"].NumberFormat = "0.0";
            sheet.Range[$"S{FirstDataRow}:U{lastRow}"].NumberFormat = "0.0%";
        }
        for (var index = 0; index < rows.Count; index++)
        {
            var excelRow = FirstDataRow + index;
            if (rows[index].IsSummary)
            {
                sheet.Range[$"A{excelRow}:X{excelRow}"].Font.Bold = true;
                sheet.Range[$"A{excelRow}:X{excelRow}"].Interior.Color = ReportTheme.WbsLevelFill(rows[index].Level);
                continue;
            }
            var item = rows[index].Item!;
            var color = item.ChangeType switch
            {
                XerActivityChangeType.Added => ReportTheme.PositiveFill,
                XerActivityChangeType.Deleted => ReportTheme.NeutralFill,
                _ when item.FinishVarianceDays > 0.001 => ReportTheme.NegativeFill,
                _ => 0xFFFFFF
            };
            sheet.Range[$"A{excelRow}:X{excelRow}"].Interior.Color = color;
            if (item.UpdateCritical) sheet.Range[$"A{excelRow}:X{excelRow}"].Font.Color = ReportTheme.CriticalFont;
        }
        sheet.Outline.SummaryRow = 0;
        foreach (var group in layout.Groups.Where(group => group.Level <= 7).OrderByDescending(group => group.Level))
        {
            var first = FirstDataRow + group.StartIndex;
            var last = FirstDataRow + group.EndIndex;
            if (last >= first) sheet.Rows[$"{first}:{last}"].Group();
        }
        sheet.Outline.ShowLevels(RowLevels: 2);
        sheet.Rows[HeaderRow].AutoFilter();
        sheet.Columns[1].ColumnWidth = 12;
        sheet.Columns[2].ColumnWidth = 25;
        sheet.Columns[3].ColumnWidth = 16;
        sheet.Columns[4].ColumnWidth = 42;
        sheet.Columns[5].Resize[1, 20].ColumnWidth = 14;
        sheet.Columns[24].ColumnWidth = 36;
        sheet.Application.ActiveWindow.SplitRow = HeaderRow;
        sheet.Application.ActiveWindow.FreezePanes = true;
        ReportTheme.ApplyPrintSetup(sheet, (string)sheet.Name);
    }

    private static ComparisonLayout BuildComparisonLayout(XerScheduleComparison comparison, IReadOnlyList<XerActivityComparison> activities)
    {
        var nodes = new Dictionary<string, ComparisonWbsNode>(StringComparer.OrdinalIgnoreCase);
        AddScheduleWbs(comparison.Baseline, nodes);
        AddScheduleWbs(comparison.Update, nodes);
        foreach (var node in nodes.Values)
            if (node.ParentPath is not null && nodes.TryGetValue(node.ParentPath, out var parent) && !ReferenceEquals(node, parent) && !parent.Children.Contains(node))
                parent.Children.Add(node);

        var direct = activities.GroupBy(item => item.WbsPath, StringComparer.OrdinalIgnoreCase).ToDictionary(group => group.Key, group => group.OrderBy(item => item.ActivityCode, StringComparer.OrdinalIgnoreCase).ToList(), StringComparer.OrdinalIgnoreCase);
        var rows = new List<ComparisonDisplayRow>();
        var groups = new List<ComparisonGroupRange>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddNode(ComparisonWbsNode node, int level)
        {
            if (!visited.Add(node.Path)) return;
            var descendants = CollectComparisonItems(node, direct, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            if (descendants.Count == 0) return;
            var summaryIndex = rows.Count;
            rows.Add(new ComparisonDisplayRow(node.Path, node.Name, level, null, descendants));
            if (direct.TryGetValue(node.Path, out var nodeItems))
                rows.AddRange(nodeItems.Select(item => new ComparisonDisplayRow(node.Path, node.Name, level + 1, item, null)));
            foreach (var child in node.Children.OrderBy(child => child.Sequence).ThenBy(child => child.Path, StringComparer.OrdinalIgnoreCase))
                AddNode(child, level + 1);
            if (rows.Count - 1 > summaryIndex) groups.Add(new ComparisonGroupRange(summaryIndex + 1, rows.Count - 1, level));
        }

        foreach (var root in nodes.Values.Where(node => node.ParentPath is null || !nodes.ContainsKey(node.ParentPath)).OrderBy(node => node.Sequence).ThenBy(node => node.Path, StringComparer.OrdinalIgnoreCase))
            AddNode(root, 1);
        foreach (var orphan in nodes.Values.Where(node => !visited.Contains(node.Path)).OrderBy(node => node.Sequence)) AddNode(orphan, 1);
        foreach (var unrepresented in direct.Where(pair => !nodes.ContainsKey(pair.Key)).OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase))
        {
            var summaryIndex = rows.Count;
            rows.Add(new ComparisonDisplayRow(unrepresented.Key, unrepresented.Key, 1, null, unrepresented.Value));
            rows.AddRange(unrepresented.Value.Select(item => new ComparisonDisplayRow(unrepresented.Key, unrepresented.Key, 2, item, null)));
            groups.Add(new ComparisonGroupRange(summaryIndex + 1, rows.Count - 1, 1));
        }
        return new ComparisonLayout(rows, groups);
    }

    private static void AddScheduleWbs(XerSchedule schedule, IDictionary<string, ComparisonWbsNode> nodes)
    {
        var byId = schedule.WbsItems.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
        var paths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string Resolve(XerWbsItem item, ISet<string> visiting)
        {
            if (paths.TryGetValue(item.Id, out var existing)) return existing;
            if (!visiting.Add(item.Id)) return item.Code;
            var path = item.ParentId is not null && byId.TryGetValue(item.ParentId, out var parent) ? $"{Resolve(parent, visiting)}.{item.Code}" : item.Code;
            paths[item.Id] = path;
            return path;
        }
        foreach (var item in schedule.WbsItems)
        {
            var path = Resolve(item, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            string? parentPath = item.ParentId is not null && byId.TryGetValue(item.ParentId, out var parent) ? Resolve(parent, new HashSet<string>(StringComparer.OrdinalIgnoreCase)) : null;
            if (!nodes.ContainsKey(path)) nodes[path] = new ComparisonWbsNode(path, item.Name, parentPath, item.Sequence);
        }
    }

    private static List<XerActivityComparison> CollectComparisonItems(ComparisonWbsNode node, IReadOnlyDictionary<string, List<XerActivityComparison>> direct, ISet<string> visiting)
    {
        if (!visiting.Add(node.Path)) return [];
        var result = direct.TryGetValue(node.Path, out var items) ? new List<XerActivityComparison>(items) : [];
        foreach (var child in node.Children) result.AddRange(CollectComparisonItems(child, direct, visiting));
        return result;
    }

    private static double WeightedProgress(IReadOnlyList<XerActivityComparison> items, Func<XerActivityComparison, double> selector, bool costWeighted) =>
        XerProgressWeighting.Calculate(items, selector, costWeighted);

    private sealed class ComparisonWbsNode(string path, string name, string? parentPath, int sequence)
    {
        public string Path { get; } = path;
        public string Name { get; } = name;
        public string? ParentPath { get; } = parentPath;
        public int Sequence { get; } = sequence;
        public List<ComparisonWbsNode> Children { get; } = [];
    }

    private sealed record ComparisonDisplayRow(string WbsPath, string WbsName, int Level, XerActivityComparison? Item, IReadOnlyList<XerActivityComparison>? SummaryItems)
    {
        public bool IsSummary => SummaryItems is not null;
    }
    private sealed record ComparisonGroupRange(int StartIndex, int EndIndex, int Level);
    private sealed record ComparisonLayout(List<ComparisonDisplayRow> Rows, List<ComparisonGroupRange> Groups);

    private static string Assess(XerActivityComparison item)
    {
        if (item.ChangeType == XerActivityChangeType.Added) return "Added in update";
        if (item.ChangeType == XerActivityChangeType.Deleted) return "Deleted from update";
        if (!item.BaselineCritical && item.UpdateCritical) return "Newly critical";
        if (item.FinishVarianceDays > 0.001) return $"Delayed {item.FinishVarianceDays:N1} day(s)";
        if (item.FinishVarianceDays < -0.001) return $"Earlier {Math.Abs(item.FinishVarianceDays):N1} day(s)";
        if (item.FloatVarianceDays < -0.001) return "Float reduced";
        return item.ChangeType == XerActivityChangeType.Unchanged ? "No schedule change" : "Changed";
    }
}

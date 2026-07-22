using PlanningToolkit.Core.Xer;

namespace PlanningToolkit.Excel.Excel;

internal static class XerPmsReport
{
    private const int HeaderRow = 4;
    private const int FirstDataRow = 5;

    public static (double Planned, double Actual, double Spi, int LookAhead, int Critical) Create(dynamic application, string baselinePath, string updatePath, int lookAheadWeeks)
    {
        var baselineDocument = XerParser.Parse(baselinePath);
        var updateDocument = XerParser.Parse(updatePath);
        EnsureValid(baselineDocument, "Baseline");
        EnsureValid(updateDocument, "Update");
        var hours = AppServices.CurrentSettings.WorkingHoursPerDay;
        var baseline = XerScheduleBuilder.Build(baselineDocument, hours, XerScheduleDateMode.Baseline);
        var update = XerScheduleBuilder.Build(updateDocument, hours, XerScheduleDateMode.Current);
        var analysis = XerProgressAnalyzer.Analyze(baseline, update, lookAheadWeeks);

        dynamic workbook = application.Workbooks.Add();
        dynamic dashboard = workbook.Worksheets[1];
        dashboard.Name = "Dashboard";
        dynamic curve = workbook.Worksheets.Add(After: dashboard);
        curve.Name = "S-Curve Data";
        dynamic lookAhead = workbook.Worksheets.Add(After: curve);
        lookAhead.Name = $"{lookAheadWeeks}-Week Look Ahead";
        dynamic critical = workbook.Worksheets.Add(After: lookAhead);
        critical.Name = "Critical Path";
        dynamic wbs = workbook.Worksheets.Add(After: critical);
        wbs.Name = "WBS Progress";

        WriteCurve(curve, analysis);
        WriteActivityList(lookAhead, $"{lookAheadWeeks}-Week Look Ahead", analysis.Comparison, analysis.LookAhead, analysis.CostWeighted);
        WriteActivityList(critical, "Critical Path", analysis.Comparison, analysis.CriticalPath, analysis.CostWeighted);
        WriteWbsProgress(wbs, analysis.Comparison, analysis.CostWeighted);
        WriteDashboard(dashboard, curve, analysis);
        dashboard.Activate();
        return (analysis.PlannedPercent, analysis.ActualPercent, analysis.Spi, analysis.LookAhead.Count, analysis.CriticalPath.Count);
    }

    private static void EnsureValid(XerDocument document, string label)
    {
        var result = XerValidator.Validate(document);
        if (!result.IsValid) throw new InvalidOperationException($"{label} XER validation found {result.ErrorCount:N0} error(s). Validate the file first.");
    }

    private static void WriteDashboard(dynamic sheet, dynamic curveSheet, XerProgressAnalysis analysis)
    {
        sheet.Cells[1, 1] = "Planning Toolkit — Project Controls Dashboard";
        sheet.Range["A1:F1"].Merge();
        ReportTheme.ApplyHeaderRow(sheet.Range["A1:F1"]);
        sheet.Range["A1:F1"].Font.Size = ReportTheme.TitleSize;
        sheet.Cells[2, 1] = "Project";
        sheet.Cells[2, 2] = analysis.Comparison.Update.ProjectName;
        sheet.Cells[3, 1] = "Data Date";
        sheet.Cells[3, 2] = analysis.DataDate.ToOADate();
        sheet.Cells[3, 2].NumberFormat = "dd-mmm-yyyy";
        sheet.Cells[2, 4] = "Weight Basis";
        sheet.Cells[2, 5] = analysis.WeightBasis;
        sheet.Cells[3, 4] = "Baseline BAC";
        sheet.Cells[3, 5] = analysis.BaselineBudgetCost;
        sheet.Cells[3, 5].NumberFormat = "#,##0.00";

        string[] kpiNames = ["Planned %", "Actual %", "Variance", "SPI", "Finish Variance", "Critical Activities", $"{analysis.LookAheadWeeks}-Week Look Ahead", "Completed", "In Progress", "Not Started"];
        object[] kpiValues = [analysis.PlannedPercent, analysis.ActualPercent, analysis.ProgressVariance, analysis.Spi, analysis.Comparison.ProjectFinishVarianceDays, analysis.CriticalPath.Count, analysis.LookAhead.Count, analysis.CompletedCount, analysis.InProgressCount, analysis.NotStartedCount];
        for (var index = 0; index < kpiNames.Length; index++)
        {
            var column = 1 + index % 5;
            var row = 5 + (index / 5) * 3;
            sheet.Cells[row, column] = kpiNames[index];
            sheet.Cells[row + 1, column] = kpiValues[index];
            ReportTheme.ApplyKpiCard(sheet.Cells[row, column], sheet.Cells[row + 1, column]);
        }
        sheet.Range["A6:C6"].NumberFormat = "0.0%";
        sheet.Range["D6"].NumberFormat = "0.00";
        sheet.Range["E6"].NumberFormat = "0.0 \"days\"";
        sheet.Cells[11, 8] = "Status";
        sheet.Cells[11, 9] = "Count";
        sheet.Cells[12, 8] = "Completed";
        sheet.Cells[12, 9] = analysis.CompletedCount;
        sheet.Cells[13, 8] = "In Progress";
        sheet.Cells[13, 9] = analysis.InProgressCount;
        sheet.Cells[14, 8] = "Not Started";
        sheet.Cells[14, 9] = analysis.NotStartedCount;
        ReportTheme.ApplyHeaderRow(sheet.Range["H11:I11"]);
        ReportTheme.ApplyGridBorders(sheet.Range["H11:I14"]);
        sheet.Columns[1].Resize[1, 5].ColumnWidth = 18;
        sheet.Columns[2].ColumnWidth = 28;
        TryCreateCharts(sheet, curveSheet, analysis.Curve.Count);
        ReportTheme.ApplyPrintSetup(sheet, "Project Controls Dashboard");
    }

    private static void TryCreateCharts(dynamic dashboard, dynamic curveSheet, int pointCount)
    {
        try
        {
            dynamic curveObject = dashboard.ChartObjects().Add(20, 230, 820, 330);
            dynamic curveChart = curveObject.Chart;
            curveChart.ChartType = 4;
            curveChart.SetSourceData(curveSheet.Range[$"A4:D{4 + pointCount}"]);
            curveChart.HasTitle = true;
            curveChart.ChartTitle.Text = "S-Curve — Planned vs Actual vs Forecast";
            curveChart.HasLegend = true;
            dynamic statusObject = dashboard.ChartObjects().Add(860, 230, 360, 330);
            dynamic statusChart = statusObject.Chart;
            statusChart.ChartType = -4120;
            statusChart.SetSourceData(dashboard.Range["H11:I14"]);
            statusChart.HasTitle = true;
            statusChart.ChartTitle.Text = "Activity Status";
        }
        catch (Exception exception)
        {
            AppServices.Logger.Error("Dashboard charts could not be created; data sheets remain available.", exception);
        }
    }

    private static void WriteCurve(dynamic sheet, XerProgressAnalysis analysis)
    {
        var values = new object?[analysis.Curve.Count + 1, 4];
        values[0, 0] = "Date";
        values[0, 1] = "Baseline Planned";
        values[0, 2] = "Estimated Actual";
        values[0, 3] = "Current Forecast";
        for (var index = 0; index < analysis.Curve.Count; index++)
        {
            var point = analysis.Curve[index];
            values[index + 1, 0] = point.Date.ToOADate();
            values[index + 1, 1] = point.Planned;
            values[index + 1, 2] = point.Actual;
            values[index + 1, 3] = point.Forecast;
        }
        sheet.Cells[1, 1] = $"S-Curve Data ({analysis.WeightBasis})";
        ReportTheme.ApplyTitle(sheet.Range["A1"]);
        sheet.Cells[2, 1] = "Actual history is estimated from current earned progress; use period data for certified historical curves.";
        sheet.Cells[3, 1] = $"Baseline BAC: {analysis.BaselineBudgetCost:N2}";
        ReportTheme.ApplySubtitle(sheet.Range["A2:A3"]);
        dynamic target = sheet.Cells[4, 1].Resize[analysis.Curve.Count + 1, 4];
        target.Value2 = values;
        ReportTheme.ApplyHeaderRow(sheet.Range["A4:D4"]);
        if (analysis.Curve.Count > 0)
        {
            var lastCurveRow = 4 + analysis.Curve.Count;
            sheet.Range[$"A5:A{lastCurveRow}"].NumberFormat = "dd-mmm-yyyy";
            sheet.Range[$"B5:D{lastCurveRow}"].NumberFormat = "0.0%";
            ReportTheme.ApplyBodyFont(sheet.Range[$"A5:D{lastCurveRow}"]);
            ReportTheme.ApplyGridBorders(sheet.Range[$"A4:D{lastCurveRow}"]);
        }
        sheet.Columns[1].ColumnWidth = 16;
        sheet.Columns[2].Resize[1, 3].ColumnWidth = 20;
        ReportTheme.ApplyPrintSetup(sheet, "S-Curve Data");
    }

    private static void WriteActivityList(dynamic sheet, string title, XerScheduleComparison comparison, IReadOnlyList<XerActivityComparison> activities, bool costWeighted)
    {
        string[] headers = ["Type", "WBS Path", "Activity ID", "Activity / WBS Name", "Status", "Start", "Finish", "Planned %", "Actual %", "Variance", "Total Float", "Critical", "Assessment"];
        var layout = BuildWbsLayout(comparison, activities);
        var values = new object?[layout.Rows.Count + 1, headers.Length];
        for (var column = 0; column < headers.Length; column++) values[0, column] = headers[column];
        for (var index = 0; index < layout.Rows.Count; index++)
        {
            var row = layout.Rows[index];
            if (row.IsSummary)
            {
                var items = row.SummaryItems!;
                var planned = Weighted(items, item => item.PlannedPercent, costWeighted);
                var actual = Weighted(items, item => item.ActualPercent, costWeighted);
                values[index + 1, 0] = "WBS Subtotal";
                values[index + 1, 1] = row.WbsPath;
                values[index + 1, 3] = Indent(row.WbsName, row.Level);
                values[index + 1, 4] = $"{items.Count:N0} activities";
                values[index + 1, 5] = MinimumDate(items.Select(item => item.UpdateStart));
                values[index + 1, 6] = MaximumDate(items.Select(item => item.UpdateFinish));
                values[index + 1, 7] = planned;
                values[index + 1, 8] = actual;
                values[index + 1, 9] = actual - planned;
                values[index + 1, 10] = items.Select(item => item.UpdateFloatDays).DefaultIfEmpty(0d).Min();
                values[index + 1, 11] = items.Count(item => item.UpdateCritical);
                values[index + 1, 12] = $"{items.Count(item => item.FinishVarianceDays > 0.001):N0} delayed | {items.Count(item => item.UpdateCritical):N0} critical";
                continue;
            }

            var item = row.Item!;
            values[index + 1, 0] = "Activity";
            values[index + 1, 1] = item.WbsPath;
            values[index + 1, 2] = item.ActivityCode;
            values[index + 1, 3] = Indent(item.ActivityName, row.Level);
            values[index + 1, 4] = item.UpdateStatus;
            values[index + 1, 5] = item.UpdateStart?.ToOADate();
            values[index + 1, 6] = item.UpdateFinish?.ToOADate();
            values[index + 1, 7] = item.PlannedPercent;
            values[index + 1, 8] = item.ActualPercent;
            values[index + 1, 9] = item.ProgressVariance;
            values[index + 1, 10] = item.UpdateFloatDays;
            values[index + 1, 11] = item.UpdateCritical ? "Yes" : string.Empty;
            values[index + 1, 12] = item.ProgressVariance < -0.000001 ? "Behind plan" : item.UpdateCritical ? "Critical" : "On / ahead of plan";
        }

        sheet.Cells[1, 1] = title;
        ReportTheme.ApplyTitle(sheet.Range["A1"]);
        sheet.Cells[2, 1] = "WBS subtotal rows can be expanded or collapsed using the Excel outline controls.";
        sheet.Cells[3, 1] = costWeighted
            ? $"Weight basis: Baseline Budgeted Cost | BAC: {XerProgressWeighting.TotalBaselineCost(comparison.Activities):N2}"
            : "Weight basis: Duration fallback because no Baseline TASKRSRC/PROJCOST cost was found.";
        ReportTheme.ApplySubtitle(sheet.Range["A2:A3"]);
        dynamic target = sheet.Cells[HeaderRow, 1].Resize[layout.Rows.Count + 1, headers.Length];
        target.Value2 = values;
        FormatHeader(sheet, "M");
        if (layout.Rows.Count > 0)
        {
            var lastRow = FirstDataRow + layout.Rows.Count - 1;
            sheet.Range[$"F{FirstDataRow}:G{lastRow}"].NumberFormat = "dd-mmm-yyyy";
            sheet.Range[$"H{FirstDataRow}:J{lastRow}"].NumberFormat = "0.0%";
            sheet.Range[$"K{FirstDataRow}:K{lastRow}"].NumberFormat = "0.0";
            ReportTheme.ApplyBodyFont(sheet.Range[$"A{FirstDataRow}:M{lastRow}"]);
            ReportTheme.ApplyGridBorders(sheet.Range[$"A{HeaderRow}:M{lastRow}"]);
        }
        FormatBodyRows(sheet, layout, "M");
        ApplyOutline(sheet, layout);
        sheet.Columns[1].ColumnWidth = 15;
        sheet.Columns[2].ColumnWidth = 26;
        sheet.Columns[3].ColumnWidth = 16;
        sheet.Columns[4].ColumnWidth = 44;
        sheet.Columns[5].Resize[1, 9].ColumnWidth = 14;
        sheet.Columns[13].ColumnWidth = 34;
        FreezeAndFilter(sheet, title);
    }

    private static void WriteWbsProgress(dynamic sheet, XerScheduleComparison comparison, bool costWeighted)
    {
        string[] headers = ["Type", "WBS Path", "Activity ID", "WBS / Activity Name", "Activities", "Baseline Cost", "Planned %", "Actual %", "Variance", "Delayed", "Critical", "Finish Variance"];
        var layout = BuildWbsLayout(comparison, comparison.Activities);
        var values = new object?[layout.Rows.Count + 1, headers.Length];
        for (var column = 0; column < headers.Length; column++) values[0, column] = headers[column];
        for (var index = 0; index < layout.Rows.Count; index++)
        {
            var row = layout.Rows[index];
            if (row.IsSummary)
            {
                var items = row.SummaryItems!;
                var planned = Weighted(items, item => item.PlannedPercent, costWeighted);
                var actual = Weighted(items, item => item.ActualPercent, costWeighted);
                values[index + 1, 0] = "WBS Subtotal";
                values[index + 1, 1] = row.WbsPath;
                values[index + 1, 3] = Indent(row.WbsName, row.Level);
                values[index + 1, 4] = items.Count;
                values[index + 1, 5] = items.Sum(item => Math.Max(item.BaselineBudgetCost, 0d));
                values[index + 1, 6] = planned;
                values[index + 1, 7] = actual;
                values[index + 1, 8] = actual - planned;
                values[index + 1, 9] = items.Count(item => item.FinishVarianceDays > 0.001);
                values[index + 1, 10] = items.Count(item => item.UpdateCritical);
                values[index + 1, 11] = items.Select(item => item.FinishVarianceDays).DefaultIfEmpty(0d).Max();
                continue;
            }

            var item = row.Item!;
            values[index + 1, 0] = "Activity";
            values[index + 1, 1] = item.WbsPath;
            values[index + 1, 2] = item.ActivityCode;
            values[index + 1, 3] = Indent(item.ActivityName, row.Level);
            values[index + 1, 4] = 1;
            values[index + 1, 5] = item.BaselineBudgetCost;
            values[index + 1, 6] = item.PlannedPercent;
            values[index + 1, 7] = item.ActualPercent;
            values[index + 1, 8] = item.ProgressVariance;
            values[index + 1, 9] = item.FinishVarianceDays > 0.001 ? "Yes" : string.Empty;
            values[index + 1, 10] = item.UpdateCritical ? "Yes" : string.Empty;
            values[index + 1, 11] = item.FinishVarianceDays;
        }

        sheet.Cells[1, 1] = "WBS Progress — Hierarchical Roll-up";
        ReportTheme.ApplyTitle(sheet.Range["A1"]);
        sheet.Cells[2, 1] = "Every WBS subtotal includes all descendant activities; expand groups for activity detail.";
        sheet.Cells[3, 1] = costWeighted
            ? $"Weight basis: Baseline Budgeted Cost | BAC: {XerProgressWeighting.TotalBaselineCost(comparison.Activities):N2}"
            : "Weight basis: Duration fallback because no Baseline TASKRSRC/PROJCOST cost was found.";
        ReportTheme.ApplySubtitle(sheet.Range["A2:A3"]);
        dynamic target = sheet.Cells[HeaderRow, 1].Resize[layout.Rows.Count + 1, headers.Length];
        target.Value2 = values;
        FormatHeader(sheet, "L");
        if (layout.Rows.Count > 0)
        {
            var lastRow = FirstDataRow + layout.Rows.Count - 1;
            sheet.Range[$"F{FirstDataRow}:F{lastRow}"].NumberFormat = "#,##0.00";
            sheet.Range[$"G{FirstDataRow}:I{lastRow}"].NumberFormat = "0.0%";
            sheet.Range[$"L{FirstDataRow}:L{lastRow}"].NumberFormat = "0.0";
            ReportTheme.ApplyBodyFont(sheet.Range[$"A{FirstDataRow}:L{lastRow}"]);
            ReportTheme.ApplyGridBorders(sheet.Range[$"A{HeaderRow}:L{lastRow}"]);
        }
        FormatBodyRows(sheet, layout, "L");
        ApplyOutline(sheet, layout);
        sheet.Columns[1].ColumnWidth = 15;
        sheet.Columns[2].ColumnWidth = 28;
        sheet.Columns[3].ColumnWidth = 16;
        sheet.Columns[4].ColumnWidth = 44;
        sheet.Columns[5].Resize[1, 8].ColumnWidth = 14;
        FreezeAndFilter(sheet, "WBS Progress");
    }

    private static PmsLayout BuildWbsLayout(XerScheduleComparison comparison, IReadOnlyList<XerActivityComparison> activities)
    {
        var nodes = new Dictionary<string, PmsWbsNode>(StringComparer.OrdinalIgnoreCase);
        AddScheduleWbs(comparison.Baseline, nodes);
        AddScheduleWbs(comparison.Update, nodes);
        foreach (var node in nodes.Values)
        {
            if (node.ParentPath is not null && nodes.TryGetValue(node.ParentPath, out var parent) && !ReferenceEquals(node, parent) && !parent.Children.Contains(node))
                parent.Children.Add(node);
        }

        var direct = activities
            .GroupBy(item => item.WbsPath, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => group.OrderBy(item => item.UpdateStart ?? DateTime.MaxValue).ThenBy(item => item.ActivityCode, StringComparer.OrdinalIgnoreCase).ToList(),
                StringComparer.OrdinalIgnoreCase);
        var rows = new List<PmsDisplayRow>();
        var groups = new List<PmsGroupRange>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        void AddNode(PmsWbsNode node, int level)
        {
            if (!visited.Add(node.Path)) return;
            var descendants = CollectItems(node, direct, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            if (descendants.Count == 0) return;
            var summaryIndex = rows.Count;
            rows.Add(new PmsDisplayRow(node.Path, node.Name, level, null, descendants));
            if (direct.TryGetValue(node.Path, out var nodeItems))
                rows.AddRange(nodeItems.Select(item => new PmsDisplayRow(node.Path, node.Name, level + 1, item, null)));
            foreach (var child in node.Children.OrderBy(child => child.Sequence).ThenBy(child => child.Path, StringComparer.OrdinalIgnoreCase))
                AddNode(child, level + 1);
            if (rows.Count - 1 > summaryIndex)
                groups.Add(new PmsGroupRange(summaryIndex + 1, rows.Count - 1, level));
        }

        foreach (var root in nodes.Values
                     .Where(node => node.ParentPath is null || !nodes.ContainsKey(node.ParentPath))
                     .OrderBy(node => node.Sequence)
                     .ThenBy(node => node.Path, StringComparer.OrdinalIgnoreCase))
            AddNode(root, 1);
        foreach (var orphan in nodes.Values.Where(node => !visited.Contains(node.Path)).OrderBy(node => node.Sequence))
            AddNode(orphan, 1);
        foreach (var unrepresented in direct.Where(pair => !nodes.ContainsKey(pair.Key)).OrderBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase))
        {
            var summaryIndex = rows.Count;
            rows.Add(new PmsDisplayRow(unrepresented.Key, unrepresented.Key, 1, null, unrepresented.Value));
            rows.AddRange(unrepresented.Value.Select(item => new PmsDisplayRow(unrepresented.Key, unrepresented.Key, 2, item, null)));
            groups.Add(new PmsGroupRange(summaryIndex + 1, rows.Count - 1, 1));
        }
        return new PmsLayout(rows, groups);
    }

    private static void AddScheduleWbs(XerSchedule schedule, IDictionary<string, PmsWbsNode> nodes)
    {
        var byId = schedule.WbsItems.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
        var paths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        string Resolve(XerWbsItem item, ISet<string> visiting)
        {
            if (paths.TryGetValue(item.Id, out var existing)) return existing;
            if (!visiting.Add(item.Id)) return item.Code;
            var path = item.ParentId is not null && byId.TryGetValue(item.ParentId, out var parent)
                ? $"{Resolve(parent, visiting)}.{item.Code}"
                : item.Code;
            paths[item.Id] = path;
            return path;
        }

        foreach (var item in schedule.WbsItems)
        {
            var path = Resolve(item, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            string? parentPath = item.ParentId is not null && byId.TryGetValue(item.ParentId, out var parent)
                ? Resolve(parent, new HashSet<string>(StringComparer.OrdinalIgnoreCase))
                : null;
            if (!nodes.ContainsKey(path))
                nodes[path] = new PmsWbsNode(path, item.Name, parentPath, item.Sequence);
        }
    }

    private static List<XerActivityComparison> CollectItems(PmsWbsNode node, IReadOnlyDictionary<string, List<XerActivityComparison>> direct, ISet<string> visiting)
    {
        if (!visiting.Add(node.Path)) return [];
        var result = direct.TryGetValue(node.Path, out var items) ? new List<XerActivityComparison>(items) : [];
        foreach (var child in node.Children)
            result.AddRange(CollectItems(child, direct, visiting));
        return result;
    }

    private static void FormatHeader(dynamic sheet, string lastColumn) =>
        ReportTheme.ApplyHeaderRow(sheet.Range[$"A{HeaderRow}:{lastColumn}{HeaderRow}"]);

    private static void FormatBodyRows(dynamic sheet, PmsLayout layout, string lastColumn)
    {
        for (var index = 0; index < layout.Rows.Count; index++)
        {
            var excelRow = FirstDataRow + index;
            var row = layout.Rows[index];
            if (row.IsSummary)
            {
                sheet.Range[$"A{excelRow}:{lastColumn}{excelRow}"].Font.Bold = true;
                sheet.Range[$"A{excelRow}:{lastColumn}{excelRow}"].Interior.Color = ReportTheme.WbsLevelFill(row.Level);
                continue;
            }

            var item = row.Item!;
            if (item.FinishVarianceDays > 0.001)
                sheet.Range[$"A{excelRow}:{lastColumn}{excelRow}"].Interior.Color = ReportTheme.NegativeFill;
            if (item.UpdateCritical)
                sheet.Range[$"A{excelRow}:{lastColumn}{excelRow}"].Font.Color = ReportTheme.CriticalFont;
        }
    }

    private static void ApplyOutline(dynamic sheet, PmsLayout layout)
    {
        if (layout.Groups.Count == 0) return;
        sheet.Outline.SummaryRow = 0;
        foreach (var group in layout.Groups.Where(group => group.Level <= 7).OrderByDescending(group => group.Level))
        {
            var first = FirstDataRow + group.StartIndex;
            var last = FirstDataRow + group.EndIndex;
            if (last >= first)
                sheet.Rows[$"{first}:{last}"].Group();
        }
        sheet.Outline.ShowLevels(RowLevels: 2);
    }

    private static void FreezeAndFilter(dynamic sheet, string printTitle)
    {
        sheet.Rows[HeaderRow].AutoFilter();
        sheet.Activate();
        sheet.Application.ActiveWindow.SplitRow = HeaderRow;
        sheet.Application.ActiveWindow.FreezePanes = true;
        ReportTheme.ApplyPrintSetup(sheet, printTitle);
    }

    private static string Indent(string value, int level) => new string(' ', Math.Max(0, level - 1) * 3) + value;

    private static double? MinimumDate(IEnumerable<DateTime?> dates)
    {
        var values = dates.Where(date => date.HasValue).Select(date => date!.Value).ToList();
        return values.Count == 0 ? null : values.Min().ToOADate();
    }

    private static double? MaximumDate(IEnumerable<DateTime?> dates)
    {
        var values = dates.Where(date => date.HasValue).Select(date => date!.Value).ToList();
        return values.Count == 0 ? null : values.Max().ToOADate();
    }

    private static double Weighted(IReadOnlyList<XerActivityComparison> items, Func<XerActivityComparison, double> selector, bool costWeighted) =>
        XerProgressWeighting.Calculate(items, selector, costWeighted);

    private sealed class PmsWbsNode(string path, string name, string? parentPath, int sequence)
    {
        public string Path { get; } = path;
        public string Name { get; } = name;
        public string? ParentPath { get; } = parentPath;
        public int Sequence { get; } = sequence;
        public List<PmsWbsNode> Children { get; } = [];
    }

    private sealed record PmsDisplayRow(string WbsPath, string WbsName, int Level, XerActivityComparison? Item, IReadOnlyList<XerActivityComparison>? SummaryItems)
    {
        public bool IsSummary => SummaryItems is not null;
    }

    private sealed record PmsGroupRange(int StartIndex, int EndIndex, int Level);
    private sealed record PmsLayout(List<PmsDisplayRow> Rows, List<PmsGroupRange> Groups);
}

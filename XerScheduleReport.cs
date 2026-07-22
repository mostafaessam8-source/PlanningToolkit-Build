using System.Globalization;
using PlanningToolkit.Core.Xer;

namespace PlanningToolkit.Excel.Excel;

internal static class XerScheduleReport
{
    private const int HeaderRow = 5;
    private const int FirstDataRow = 6;

    public static (int Activities, int WbsLevels, int CriticalActivities) Create(dynamic application, string path)
    {
        var document = XerParser.Parse(path);
        var validation = XerValidator.Validate(document);
        if (!validation.IsValid)
            throw new InvalidOperationException($"XER validation found {validation.ErrorCount:N0} error(s). Use Validate XER for details.");

        var schedule = XerScheduleBuilder.Build(document, AppServices.CurrentSettings.WorkingHoursPerDay);
        if (schedule.Activities.Count == 0)
            throw new InvalidOperationException("The XER TASK table does not contain activities.");

        var layout = BuildLayout(schedule);
        dynamic workbook = application.Workbooks.Add();
        dynamic scheduleSheet = workbook.Worksheets[1];
        scheduleSheet.Name = "WBS Schedule";
        WriteSchedule(scheduleSheet, schedule, layout);

        dynamic ganttSheet = workbook.Worksheets.Add(After: scheduleSheet);
        ganttSheet.Name = "Gantt Chart";
        WriteGantt(ganttSheet, schedule, layout.Rows);
        scheduleSheet.Activate();
        return (schedule.Activities.Count, layout.MaximumLevel, schedule.Activities.Count(activity => activity.IsCritical));
    }

    private static Layout BuildLayout(XerSchedule schedule)
    {
        var nodes = schedule.WbsItems.ToDictionary(item => item.Id, item => new WbsNode(item), StringComparer.OrdinalIgnoreCase);
        foreach (var node in nodes.Values)
            if (node.Item.ParentId is not null && nodes.TryGetValue(node.Item.ParentId, out var parent) && !ReferenceEquals(node, parent))
                parent.Children.Add(node);

        var activitiesByWbs = schedule.Activities
            .Where(activity => activity.WbsId is not null && nodes.ContainsKey(activity.WbsId))
            .GroupBy(activity => activity.WbsId!, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.OrderBy(a => a.Start).ThenBy(a => a.Code).ToList(), StringComparer.OrdinalIgnoreCase);

        var rows = new List<DisplayRow>();
        var groups = new List<GroupRange>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var maximumLevel = 0;

        void AddNode(WbsNode node, int level, string parentPath)
        {
            if (!visited.Add(node.Item.Id))
                return;
            maximumLevel = Math.Max(maximumLevel, level);
            var path = string.IsNullOrEmpty(parentPath) ? node.Item.Code : $"{parentPath}.{node.Item.Code}";
            var descendants = CollectActivities(node, activitiesByWbs, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
            var summary = Summarize(descendants);
            var summaryIndex = rows.Count;
            rows.Add(DisplayRow.ForWbs(node.Item, path, level, summary));

            if (activitiesByWbs.TryGetValue(node.Item.Id, out var directActivities))
                rows.AddRange(directActivities.Select(activity => DisplayRow.ForActivity(activity, path, level + 1)));
            foreach (var child in node.Children.OrderBy(child => child.Item.Sequence).ThenBy(child => child.Item.Code))
                AddNode(child, level + 1, path);

            if (rows.Count - 1 > summaryIndex)
                groups.Add(new GroupRange(summaryIndex + 1, rows.Count - 1, level));
        }

        foreach (var root in nodes.Values.Where(node => node.Item.ParentId is null || !nodes.ContainsKey(node.Item.ParentId)).OrderBy(node => node.Item.Sequence).ThenBy(node => node.Item.Code))
            AddNode(root, 1, string.Empty);
        foreach (var orphan in nodes.Values.Where(node => !visited.Contains(node.Item.Id)).OrderBy(node => node.Item.Sequence))
            AddNode(orphan, 1, string.Empty);

        var unassigned = schedule.Activities.Where(activity => activity.WbsId is null || !nodes.ContainsKey(activity.WbsId)).OrderBy(a => a.Start).ThenBy(a => a.Code).ToList();
        if (unassigned.Count > 0)
        {
            var start = rows.Count;
            rows.Add(DisplayRow.ForWbs(new XerWbsItem("UNASSIGNED", null, "UNASSIGNED", "Unassigned Activities", int.MaxValue), "UNASSIGNED", 1, Summarize(unassigned)));
            rows.AddRange(unassigned.Select(activity => DisplayRow.ForActivity(activity, "UNASSIGNED", 2)));
            groups.Add(new GroupRange(start + 1, rows.Count - 1, 1));
        }

        return new Layout(rows, groups, maximumLevel);
    }

    private static List<XerScheduleActivity> CollectActivities(WbsNode node, IReadOnlyDictionary<string, List<XerScheduleActivity>> byWbs, ISet<string> visiting)
    {
        if (!visiting.Add(node.Item.Id))
            return [];
        var result = byWbs.TryGetValue(node.Item.Id, out var direct) ? new List<XerScheduleActivity>(direct) : [];
        foreach (var child in node.Children)
            result.AddRange(CollectActivities(child, byWbs, visiting));
        return result;
    }

    private static Summary Summarize(IReadOnlyList<XerScheduleActivity> activities)
    {
        var start = activities.Where(a => a.Start.HasValue).Select(a => a.Start!.Value).DefaultIfEmpty().Min();
        var finish = activities.Where(a => a.Finish.HasValue).Select(a => a.Finish!.Value).DefaultIfEmpty().Max();
        var weight = activities.Sum(a => Math.Max(a.OriginalDurationDays, 0d));
        var progress = weight > 0
            ? activities.Sum(a => a.PercentComplete * Math.Max(a.OriginalDurationDays, 0d)) / weight
            : activities.Count == 0 ? 0d : activities.Average(a => a.PercentComplete);
        var minimumFloat = activities.Where(a => !string.Equals(a.Status, "Completed", StringComparison.OrdinalIgnoreCase)).Select(a => a.TotalFloatDays).DefaultIfEmpty(0d).Min();
        return new Summary(start == default ? null : start, finish == default ? null : finish, progress, activities.Sum(a => a.OriginalDurationDays), activities.Sum(a => a.RemainingDurationDays), minimumFloat, activities.Any(a => a.IsCritical));
    }

    private static void WriteSchedule(dynamic sheet, XerSchedule schedule, Layout layout)
    {
        string[] headers = ["Type", "WBS Path", "Activity ID", "Activity / WBS Name", "Status", "Original Duration", "Remaining Duration", "Start", "Finish", "% Complete", "Total Float", "Critical", "Calendar", "Constraint", "Level"];
        var values = new object?[layout.Rows.Count + 1, headers.Length];
        for (var column = 0; column < headers.Length; column++) values[0, column] = headers[column];
        for (var index = 0; index < layout.Rows.Count; index++)
        {
            var row = layout.Rows[index];
            values[index + 1, 0] = row.Type;
            values[index + 1, 1] = row.WbsPath;
            values[index + 1, 2] = row.Code;
            values[index + 1, 3] = new string(' ', Math.Max(0, row.Level - 1) * 3) + row.Name;
            values[index + 1, 4] = row.Status;
            values[index + 1, 5] = row.OriginalDuration;
            values[index + 1, 6] = row.RemainingDuration;
            values[index + 1, 7] = row.Start?.ToOADate();
            values[index + 1, 8] = row.Finish?.ToOADate();
            values[index + 1, 9] = row.PercentComplete;
            values[index + 1, 10] = row.TotalFloat;
            values[index + 1, 11] = row.IsCritical ? "Yes" : string.Empty;
            values[index + 1, 12] = row.Calendar;
            values[index + 1, 13] = row.Constraint;
            values[index + 1, 14] = row.Level;
        }

        sheet.Cells[1, 1] = "Planning Toolkit — WBS Schedule";
        ReportTheme.ApplyTitle(sheet.Range["A1"]);
        sheet.Cells[2, 1] = "Project";
        sheet.Cells[2, 2] = schedule.ProjectName;
        sheet.Cells[3, 1] = "Data Date";
        sheet.Cells[3, 2] = schedule.DataDate?.ToOADate();
        sheet.Cells[3, 2].NumberFormat = "dd-mmm-yyyy";
        ReportTheme.ApplySubtitle(sheet.Range["A2:B3"]);
        dynamic target = sheet.Cells[HeaderRow, 1].Resize[layout.Rows.Count + 1, headers.Length];
        target.Value2 = values;

        var lastRow = FirstDataRow + layout.Rows.Count - 1;
        dynamic headerRange = sheet.Range[$"A{HeaderRow}:O{HeaderRow}"];
        dynamic bodyRange = sheet.Range[$"A{FirstDataRow}:O{lastRow}"];
        ReportTheme.ApplyHeaderRow(headerRange);
        ReportTheme.ApplyBodyFont(bodyRange);
        ReportTheme.ApplyGridBorders(sheet.Range[$"A{HeaderRow}:O{lastRow}"]);
        sheet.Range[$"H{FirstDataRow}:I{lastRow}"].NumberFormat = "dd-mmm-yyyy";
        sheet.Range[$"J{FirstDataRow}:J{lastRow}"].NumberFormat = "0.0%";
        sheet.Range[$"F{FirstDataRow}:G{lastRow}"].NumberFormat = "0.0";
        sheet.Range[$"K{FirstDataRow}:K{lastRow}"].NumberFormat = "0.0";

        for (var index = 0; index < layout.Rows.Count; index++)
        {
            var excelRow = FirstDataRow + index;
            if (layout.Rows[index].Type == "WBS")
            {
                sheet.Range[$"A{excelRow}:O{excelRow}"].Font.Bold = true;
                sheet.Range[$"A{excelRow}:O{excelRow}"].Interior.Color = ReportTheme.WbsLevelFill(layout.Rows[index].Level);
            }
            else if (layout.Rows[index].IsCritical)
            {
                sheet.Range[$"A{excelRow}:O{excelRow}"].Font.Color = ReportTheme.CriticalFont;
                sheet.Cells[excelRow, 11].Interior.Color = ReportTheme.CriticalFill;
            }
        }

        // NOTE: this sheet uses row Outline grouping for WBS collapse/expand, which Excel
        // does not support inside a native Table (ListObject) — so we keep it as a themed
        // range with AutoFilter rather than converting it to a table.
        sheet.Outline.SummaryRow = 0;
        foreach (var group in layout.Groups.Where(group => group.Level <= 7).OrderByDescending(group => group.Level))
        {
            var first = FirstDataRow + group.StartIndex;
            var last = FirstDataRow + group.EndIndex;
            if (last >= first) sheet.Rows[$"{first}:{last}"].Group();
        }
        sheet.Outline.ShowLevels(RowLevels: 2);
        sheet.Rows[HeaderRow].AutoFilter();
        sheet.Columns[1].ColumnWidth = 10;
        sheet.Columns[2].ColumnWidth = 24;
        sheet.Columns[3].ColumnWidth = 16;
        sheet.Columns[4].ColumnWidth = 48;
        sheet.Columns[5].ColumnWidth = 14;
        sheet.Columns[6].Resize[1, 10].ColumnWidth = 14;
        sheet.Range["A1:O3"].Font.Bold = true;
        sheet.Application.ActiveWindow.SplitRow = HeaderRow;
        sheet.Application.ActiveWindow.FreezePanes = true;
        ReportTheme.ApplyPrintSetup(sheet, $"WBS Schedule — {schedule.ProjectName}");
    }

    private static void WriteGantt(dynamic sheet, XerSchedule schedule, IReadOnlyList<DisplayRow> rows)
    {
        var dated = rows.Where(row => row.Start.HasValue && row.Finish.HasValue).ToList();
        if (dated.Count == 0)
            throw new InvalidOperationException("The XER activities do not contain usable Start and Finish dates for a Gantt chart.");
        var minimum = dated.Min(row => row.Start!.Value).Date;
        var maximum = dated.Max(row => row.Finish!.Value).Date;
        var timelineStart = minimum.AddDays(-(7 + (int)minimum.DayOfWeek - (int)DayOfWeek.Monday) % 7);
        var weeks = Math.Min(260, Math.Max(1, (int)Math.Ceiling((maximum - timelineStart).TotalDays / 7d) + 1));
        var fixedColumns = 6;
        var values = new object?[rows.Count + 1, fixedColumns + weeks];
        string[] headers = ["Type", "WBS / Activity", "Name", "Start", "Finish", "% Complete"];
        for (var column = 0; column < fixedColumns; column++) values[0, column] = headers[column];
        for (var week = 0; week < weeks; week++) values[0, fixedColumns + week] = timelineStart.AddDays(week * 7).ToOADate();

        for (var index = 0; index < rows.Count; index++)
        {
            var row = rows[index];
            values[index + 1, 0] = row.Type;
            values[index + 1, 1] = row.Code;
            values[index + 1, 2] = new string(' ', Math.Max(0, row.Level - 1) * 2) + row.Name;
            values[index + 1, 3] = row.Start?.ToOADate();
            values[index + 1, 4] = row.Finish?.ToOADate();
            values[index + 1, 5] = row.PercentComplete;
            if (!row.Start.HasValue || !row.Finish.HasValue) continue;
            for (var week = 0; week < weeks; week++)
            {
                var periodStart = timelineStart.AddDays(week * 7);
                var periodEnd = periodStart.AddDays(6);
                if (row.Start.Value.Date <= periodEnd && row.Finish.Value.Date >= periodStart)
                    values[index + 1, fixedColumns + week] = row.Type == "WBS" ? "━" : "■";
            }
        }

        sheet.Cells[1, 1] = $"Planning Toolkit — Gantt Chart — {schedule.ProjectName}";
        ReportTheme.ApplyTitle(sheet.Range["A1"]);
        dynamic target = sheet.Cells[HeaderRow, 1].Resize[rows.Count + 1, fixedColumns + weeks];
        target.Value2 = values;
        dynamic headerRange = sheet.Range[sheet.Cells[HeaderRow, 1], sheet.Cells[HeaderRow, fixedColumns + weeks]];
        dynamic bodyRange = sheet.Range[sheet.Cells[FirstDataRow, 1], sheet.Cells[HeaderRow + rows.Count, fixedColumns + weeks]];
        ReportTheme.ApplyHeaderRow(headerRange);
        ReportTheme.ApplyBodyFont(bodyRange);
        ReportTheme.ApplyGridBorders(sheet.Range[sheet.Cells[HeaderRow, 1], sheet.Cells[HeaderRow + rows.Count, fixedColumns + weeks]]);
        sheet.Range[$"D{HeaderRow}:E{HeaderRow + rows.Count}"].NumberFormat = "dd-mmm-yyyy";
        sheet.Range[$"F{FirstDataRow}:F{FirstDataRow + rows.Count - 1}"].NumberFormat = "0.0%";
        sheet.Range[sheet.Cells[HeaderRow, fixedColumns + 1], sheet.Cells[HeaderRow, fixedColumns + weeks]].NumberFormat = "dd-mmm";
        sheet.Columns[1].ColumnWidth = 9;
        sheet.Columns[2].ColumnWidth = 16;
        sheet.Columns[3].ColumnWidth = 44;
        sheet.Columns[4].Resize[1, 3].ColumnWidth = 13;
        sheet.Range[sheet.Cells[1, fixedColumns + 1], sheet.Cells[rows.Count + HeaderRow, fixedColumns + weeks]].ColumnWidth = 3;
        for (var index = 0; index < rows.Count; index++)
        {
            var excelRow = FirstDataRow + index;
            if (rows[index].Type == "WBS") sheet.Range[sheet.Cells[excelRow, 1], sheet.Cells[excelRow, fixedColumns + weeks]].Font.Bold = true;
            if (rows[index].IsCritical) sheet.Range[sheet.Cells[excelRow, fixedColumns + 1], sheet.Cells[excelRow, fixedColumns + weeks]].Font.Color = ReportTheme.CriticalFont;
            else sheet.Range[sheet.Cells[excelRow, fixedColumns + 1], sheet.Cells[excelRow, fixedColumns + weeks]].Font.Color = ReportTheme.WarningFont;
        }
        if (schedule.DataDate.HasValue)
        {
            var dataWeek = (int)Math.Floor((schedule.DataDate.Value.Date - timelineStart).TotalDays / 7d);
            if (dataWeek >= 0 && dataWeek < weeks)
                sheet.Columns[fixedColumns + dataWeek + 1].Interior.Color = ReportTheme.DataDateFill;
        }
        sheet.Application.ActiveWindow.SplitColumn = fixedColumns;
        sheet.Application.ActiveWindow.SplitRow = HeaderRow;
        sheet.Application.ActiveWindow.FreezePanes = true;
        ReportTheme.ApplyPrintSetup(sheet, $"Gantt Chart — {schedule.ProjectName}");
    }

    private sealed class WbsNode(XerWbsItem item)
    {
        public XerWbsItem Item { get; } = item;
        public List<WbsNode> Children { get; } = [];
    }

    private sealed record Summary(DateTime? Start, DateTime? Finish, double Progress, double Duration, double Remaining, double Float, bool Critical);
    private sealed record GroupRange(int StartIndex, int EndIndex, int Level);
    private sealed record Layout(List<DisplayRow> Rows, List<GroupRange> Groups, int MaximumLevel);
    private sealed record DisplayRow(string Type, string WbsPath, string Code, string Name, string Status, double OriginalDuration, double RemainingDuration, DateTime? Start, DateTime? Finish, double PercentComplete, double TotalFloat, bool IsCritical, string Calendar, string Constraint, int Level)
    {
        public static DisplayRow ForWbs(XerWbsItem item, string path, int level, Summary summary) => new("WBS", path, item.Code, item.Name, "Summary", summary.Duration, summary.Remaining, summary.Start, summary.Finish, summary.Progress, summary.Float, summary.Critical, string.Empty, string.Empty, level);
        public static DisplayRow ForActivity(XerScheduleActivity activity, string path, int level) => new("Activity", path, activity.Code, activity.Name, activity.Status, activity.OriginalDurationDays, activity.RemainingDurationDays, activity.Start, activity.Finish, activity.PercentComplete, activity.TotalFloatDays, activity.IsCritical, activity.Calendar, activity.Constraint, level);
    }
}

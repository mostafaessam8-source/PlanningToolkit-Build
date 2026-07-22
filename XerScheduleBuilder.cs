using System.Globalization;

namespace PlanningToolkit.Core.Xer;

public sealed record XerWbsItem(string Id, string? ParentId, string Code, string Name, int Sequence);

public enum XerScheduleDateMode
{
    Current,
    Baseline
}

public sealed record XerScheduleActivity(
    string Id,
    string Code,
    string Name,
    string? WbsId,
    string Status,
    double OriginalDurationDays,
    double RemainingDurationDays,
    double BudgetedCost,
    DateTime? Start,
    DateTime? Finish,
    double PercentComplete,
    double TotalFloatDays,
    bool IsCritical,
    string Calendar,
    string Constraint);

public sealed class XerSchedule
{
    public string ProjectName { get; internal set; } = "Primavera Schedule";
    public DateTime? DataDate { get; set; }
    public List<XerWbsItem> WbsItems { get; } = new();
    public List<XerScheduleActivity> Activities { get; } = new();
}

public static class XerScheduleBuilder
{
    public static XerSchedule Build(XerDocument document, double hoursPerDay = 8d, XerScheduleDateMode dateMode = XerScheduleDateMode.Current)
    {
        ArgumentNullException.ThrowIfNull(document);
        if (hoursPerDay <= 0)
            throw new ArgumentOutOfRangeException(nameof(hoursPerDay));

        var schedule = new XerSchedule();
        ReadProject(document.FindTable("PROJECT"), schedule);
        ReadWbs(document.FindTable("PROJWBS"), schedule);
        var budgetedCosts = ReadBudgetedCosts(document);
        ReadActivities(document.FindTable("TASK"), schedule, hoursPerDay, dateMode, budgetedCosts);
        return schedule;
    }

    private static void ReadProject(XerTable? table, XerSchedule schedule)
    {
        if (table is null || table.Rows.Count == 0)
            return;
        var row = new RowReader(table, table.Rows[0]);
        schedule.ProjectName = row.Text("proj_short_name", "proj_name") ?? schedule.ProjectName;
        schedule.DataDate = row.Date("last_recalc_date", "data_date");
    }

    private static void ReadWbs(XerTable? table, XerSchedule schedule)
    {
        if (table is null)
            return;
        foreach (var source in table.Rows)
        {
            var row = new RowReader(table, source);
            var id = row.Text("wbs_id");
            if (string.IsNullOrWhiteSpace(id))
                continue;
            schedule.WbsItems.Add(new XerWbsItem(
                id,
                row.Text("parent_wbs_id"),
                row.Text("wbs_short_name", "wbs_code") ?? id,
                row.Text("wbs_name") ?? id,
                row.Integer("seq_num", "wbs_seq_num")));
        }
    }

    private static void ReadActivities(XerTable? table, XerSchedule schedule, double hoursPerDay, XerScheduleDateMode dateMode, IReadOnlyDictionary<string, double> budgetedCosts)
    {
        if (table is null)
            throw new InvalidOperationException("The XER file does not contain a TASK table.");
        foreach (var source in table.Rows)
        {
            var row = new RowReader(table, source);
            var id = row.Text("task_id");
            if (string.IsNullOrWhiteSpace(id))
                continue;
            var status = NormalizeStatus(row.Text("status_code"));
            var totalFloat = row.Number("total_float_hr_cnt") / hoursPerDay;
            var percent = NormalizePercent(row.Number("phys_complete_pct", "complete_pct"));
            var start = dateMode == XerScheduleDateMode.Baseline
                ? row.Date("target_start_date", "early_start_date", "act_start_date", "restart_date")
                : row.Date("act_start_date", "early_start_date", "restart_date", "target_start_date");
            var finish = dateMode == XerScheduleDateMode.Baseline
                ? row.Date("target_end_date", "early_end_date", "act_end_date", "reend_date")
                : row.Date("act_end_date", "early_end_date", "reend_date", "target_end_date");
            budgetedCosts.TryGetValue(id, out var budgetedCost);
            schedule.Activities.Add(new XerScheduleActivity(
                id,
                row.Text("task_code") ?? id,
                row.Text("task_name") ?? id,
                row.Text("wbs_id"),
                status,
                row.Number("target_drtn_hr_cnt", "orig_drtn_hr_cnt") / hoursPerDay,
                row.Number("remain_drtn_hr_cnt") / hoursPerDay,
                budgetedCost,
                start,
                finish,
                percent,
                totalFloat,
                totalFloat <= 0.0001 && !string.Equals(status, "Completed", StringComparison.OrdinalIgnoreCase),
                row.Text("clndr_id", "calendar_id") ?? string.Empty,
                row.Text("cstr_type", "constraint_type") ?? string.Empty));
        }
    }

    private static Dictionary<string, double> ReadBudgetedCosts(XerDocument document)
    {
        var result = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        AddCosts(document.FindTable("TASKRSRC"), result, ["target_cost", "budgeted_cost", "planned_cost"]);
        AddCosts(document.FindTable("PROJCOST"), result, ["cost_load", "target_cost", "budgeted_cost", "planned_cost"]);
        return result;
    }

    private static void AddCosts(XerTable? table, IDictionary<string, double> result, string[] costFields)
    {
        if (table is null) return;
        foreach (var source in table.Rows)
        {
            var row = new RowReader(table, source);
            var taskId = row.Text("task_id");
            if (string.IsNullOrWhiteSpace(taskId)) continue;
            var cost = row.FirstNumber(costFields);
            if (cost <= 0d && row.TryNumber(out var quantity, "target_qty", "budgeted_qty", "planned_qty") && row.TryNumber(out var rate, "cost_per_qty", "price_per_unit"))
                cost = quantity * rate;
            if (cost <= 0d) continue;
            result[taskId] = result.TryGetValue(taskId, out var existing) ? existing + cost : cost;
        }
    }

    private static double NormalizePercent(double value) => Math.Clamp(value > 1d ? value / 100d : value, 0d, 1d);

    private static string NormalizeStatus(string? value) => value switch
    {
        "TK_Complete" => "Completed",
        "TK_Active" => "In Progress",
        "TK_NotStart" => "Not Started",
        _ => value ?? "Unknown"
    };

    private sealed class RowReader
    {
        private readonly XerRow _row;
        private readonly Dictionary<string, int> _fields;

        public RowReader(XerTable table, XerRow row)
        {
            _row = row;
            _fields = table.Fields.Select((name, index) => (name, index)).ToDictionary(x => x.name, x => x.index, StringComparer.OrdinalIgnoreCase);
        }

        public string? Text(params string[] names)
        {
            foreach (var name in names)
                if (_fields.TryGetValue(name, out var index) && index < _row.Values.Count && !string.IsNullOrWhiteSpace(_row.Values[index]))
                    return _row.Values[index].Trim();
            return null;
        }

        public double Number(params string[] names) => double.TryParse(Text(names), NumberStyles.Any, CultureInfo.InvariantCulture, out var value) ? value : 0d;
        public double FirstNumber(params string[] names) => TryNumber(out var value, names) ? value : 0d;
        public bool TryNumber(out double value, params string[] names)
        {
            foreach (var name in names)
            {
                var text = Text(name);
                if (text is not null && double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                    return true;
            }
            value = 0d;
            return false;
        }
        public int Integer(params string[] names) => int.TryParse(Text(names), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value) ? value : 0;

        public DateTime? Date(params string[] names)
        {
            var text = Text(names);
            if (text is null)
                return null;
            string[] formats = ["yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "dd-MMM-yy", "dd-MMM-yyyy", "M/d/yyyy H:mm", "M/d/yyyy"];
            if (DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var exact))
                return exact;
            return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var parsed) ? parsed : null;
        }
    }
}

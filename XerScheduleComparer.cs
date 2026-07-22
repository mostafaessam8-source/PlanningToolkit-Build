namespace PlanningToolkit.Core.Xer;

public enum XerActivityChangeType
{
    Unchanged,
    Changed,
    Added,
    Deleted
}

public sealed record XerActivityComparison(
    string ActivityCode,
    string ActivityName,
    string WbsPath,
    XerActivityChangeType ChangeType,
    string BaselineStatus,
    string UpdateStatus,
    DateTime? BaselineStart,
    DateTime? UpdateStart,
    double StartVarianceDays,
    DateTime? BaselineFinish,
    DateTime? UpdateFinish,
    double FinishVarianceDays,
    double BaselineDurationDays,
    double UpdateDurationDays,
    double DurationVarianceDays,
    double BaselineFloatDays,
    double UpdateFloatDays,
    double FloatVarianceDays,
    double BaselineBudgetCost,
    double UpdateBudgetCost,
    double PlannedPercent,
    double ActualPercent,
    double ProgressVariance,
    bool BaselineCritical,
    bool UpdateCritical);

public sealed class XerScheduleComparison
{
    public required XerSchedule Baseline { get; init; }
    public required XerSchedule Update { get; init; }
    public List<XerActivityComparison> Activities { get; } = new();
    public int AddedCount => Activities.Count(item => item.ChangeType == XerActivityChangeType.Added);
    public int DeletedCount => Activities.Count(item => item.ChangeType == XerActivityChangeType.Deleted);
    public int DelayedCount => Activities.Count(item => item.FinishVarianceDays > 0.001);
    public int CriticalCount => Activities.Count(item => item.UpdateCritical);
    public DateTime? BaselineProjectFinish => Baseline.Activities.Where(a => a.Finish.HasValue).Select(a => a.Finish).Max();
    public DateTime? UpdateProjectFinish => Update.Activities.Where(a => a.Finish.HasValue).Select(a => a.Finish).Max();
    public double ProjectFinishVarianceDays => Difference(BaselineProjectFinish, UpdateProjectFinish);

    private static double Difference(DateTime? baseline, DateTime? update) => baseline.HasValue && update.HasValue ? (update.Value - baseline.Value).TotalDays : 0d;
}

public static class XerScheduleComparer
{
    public static XerScheduleComparison Compare(XerSchedule baseline, XerSchedule update)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(update);
        var result = new XerScheduleComparison { Baseline = baseline, Update = update };
        var baselineByCode = baseline.Activities.GroupBy(a => a.Code, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        var updateByCode = update.Activities.GroupBy(a => a.Code, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        var baselinePaths = BuildWbsPaths(baseline);
        var updatePaths = BuildWbsPaths(update);
        var progressDate = update.DataDate ?? baseline.DataDate ?? DateTime.Today;
        var codes = baselineByCode.Keys.Union(updateByCode.Keys, StringComparer.OrdinalIgnoreCase).OrderBy(code => code, StringComparer.OrdinalIgnoreCase);

        foreach (var code in codes)
        {
            baselineByCode.TryGetValue(code, out var oldActivity);
            updateByCode.TryGetValue(code, out var newActivity);
            var changeType = oldActivity is null ? XerActivityChangeType.Added : newActivity is null ? XerActivityChangeType.Deleted : IsChanged(oldActivity, newActivity) ? XerActivityChangeType.Changed : XerActivityChangeType.Unchanged;
            var path = GetPath(newActivity, updatePaths) ?? GetPath(oldActivity, baselinePaths) ?? "UNASSIGNED";
            var plannedPercent = PlannedPercent(oldActivity, progressDate);
            var actualPercent = newActivity?.PercentComplete ?? 0d;
            result.Activities.Add(new XerActivityComparison(
                code,
                newActivity?.Name ?? oldActivity?.Name ?? code,
                path,
                changeType,
                oldActivity?.Status ?? string.Empty,
                newActivity?.Status ?? string.Empty,
                oldActivity?.Start,
                newActivity?.Start,
                Difference(oldActivity?.Start, newActivity?.Start),
                oldActivity?.Finish,
                newActivity?.Finish,
                Difference(oldActivity?.Finish, newActivity?.Finish),
                oldActivity?.OriginalDurationDays ?? 0d,
                newActivity?.OriginalDurationDays ?? 0d,
                (newActivity?.OriginalDurationDays ?? 0d) - (oldActivity?.OriginalDurationDays ?? 0d),
                oldActivity?.TotalFloatDays ?? 0d,
                newActivity?.TotalFloatDays ?? 0d,
                (newActivity?.TotalFloatDays ?? 0d) - (oldActivity?.TotalFloatDays ?? 0d),
                oldActivity?.BudgetedCost ?? 0d,
                newActivity?.BudgetedCost ?? 0d,
                plannedPercent,
                actualPercent,
                actualPercent - plannedPercent,
                oldActivity?.IsCritical ?? false,
                newActivity?.IsCritical ?? false));
        }
        return result;
    }

    private static bool IsChanged(XerScheduleActivity baseline, XerScheduleActivity update) =>
        baseline.Start != update.Start || baseline.Finish != update.Finish ||
        Math.Abs(baseline.OriginalDurationDays - update.OriginalDurationDays) > 0.001 ||
        Math.Abs(baseline.TotalFloatDays - update.TotalFloatDays) > 0.001 ||
        !string.Equals(baseline.Status, update.Status, StringComparison.OrdinalIgnoreCase) ||
        baseline.IsCritical != update.IsCritical;

    private static double Difference(DateTime? baseline, DateTime? update) => baseline.HasValue && update.HasValue ? (update.Value - baseline.Value).TotalDays : 0d;

    private static double PlannedPercent(XerScheduleActivity? baseline, DateTime dataDate)
    {
        if (baseline?.Start is null || baseline.Finish is null)
            return 0d;
        if (dataDate <= baseline.Start.Value)
            return 0d;
        if (dataDate >= baseline.Finish.Value)
            return 1d;
        var duration = (baseline.Finish.Value - baseline.Start.Value).TotalDays;
        return duration <= 0d ? 0d : Math.Clamp((dataDate - baseline.Start.Value).TotalDays / duration, 0d, 1d);
    }

    private static string? GetPath(XerScheduleActivity? activity, IReadOnlyDictionary<string, string> paths) =>
        activity?.WbsId is not null && paths.TryGetValue(activity.WbsId, out var path) ? path : null;

    private static Dictionary<string, string> BuildWbsPaths(XerSchedule schedule)
    {
        var byId = schedule.WbsItems.ToDictionary(item => item.Id, StringComparer.OrdinalIgnoreCase);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string Resolve(XerWbsItem item, ISet<string> visiting)
        {
            if (result.TryGetValue(item.Id, out var existing)) return existing;
            if (!visiting.Add(item.Id)) return item.Code;
            var path = item.ParentId is not null && byId.TryGetValue(item.ParentId, out var parent) ? $"{Resolve(parent, visiting)}.{item.Code}" : item.Code;
            result[item.Id] = path;
            return path;
        }
        foreach (var item in schedule.WbsItems) Resolve(item, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        return result;
    }
}

namespace PlanningToolkit.Core.Xer;

public sealed record XerProgressPoint(DateTime Date, double Planned, double Actual, double Forecast);

public sealed class XerProgressAnalysis
{
    public required XerScheduleComparison Comparison { get; init; }
    public DateTime DataDate { get; init; }
    public double PlannedPercent { get; init; }
    public double ActualPercent { get; init; }
    public double ProgressVariance => ActualPercent - PlannedPercent;
    public double Spi => PlannedPercent <= 0.000001 ? 0d : ActualPercent / PlannedPercent;
    public bool CostWeighted { get; init; }
    public double BaselineBudgetCost { get; init; }
    public string WeightBasis => CostWeighted ? "Baseline Budgeted Cost" : "Duration fallback (no baseline cost found)";
    public int LookAheadWeeks { get; init; } = 2;
    public List<XerProgressPoint> Curve { get; } = new();
    public List<XerActivityComparison> LookAhead { get; } = new();
    public IReadOnlyList<XerActivityComparison> TwoWeekLookAhead => LookAhead;
    public List<XerActivityComparison> CriticalPath { get; } = new();
    public int CompletedCount => Comparison.Update.Activities.Count(activity => string.Equals(activity.Status, "Completed", StringComparison.OrdinalIgnoreCase));
    public int InProgressCount => Comparison.Update.Activities.Count(activity => string.Equals(activity.Status, "In Progress", StringComparison.OrdinalIgnoreCase));
    public int NotStartedCount => Comparison.Update.Activities.Count - CompletedCount - InProgressCount;
}

public static class XerProgressAnalyzer
{
    public static XerProgressAnalysis Analyze(XerSchedule baseline, XerSchedule update, int lookAheadWeeks = 2)
    {
        if (lookAheadWeeks is < 1 or > 52)
            throw new ArgumentOutOfRangeException(nameof(lookAheadWeeks), "Look-ahead duration must be between 1 and 52 weeks.");

        var comparison = XerScheduleComparer.Compare(baseline, update);
        var dataDate = update.DataDate ?? baseline.DataDate ?? DateTime.Today;
        var costWeighted = XerProgressWeighting.HasBaselineCost(comparison.Activities);
        var analysis = new XerProgressAnalysis
        {
            Comparison = comparison,
            DataDate = dataDate,
            LookAheadWeeks = lookAheadWeeks,
            CostWeighted = costWeighted,
            BaselineBudgetCost = XerProgressWeighting.TotalBaselineCost(comparison.Activities),
            PlannedPercent = XerProgressWeighting.Calculate(comparison.Activities, item => item.PlannedPercent, costWeighted),
            ActualPercent = XerProgressWeighting.Calculate(comparison.Activities, item => item.ActualPercent, costWeighted)
        };

        BuildCurve(analysis);
        analysis.LookAhead.AddRange(comparison.Activities
            .Where(item => !string.Equals(item.UpdateStatus, "Completed", StringComparison.OrdinalIgnoreCase))
            .Where(item => Intersects(item.UpdateStart, item.UpdateFinish, dataDate, dataDate.AddDays(lookAheadWeeks * 7d)))
            .OrderBy(item => item.UpdateStart).ThenBy(item => item.ActivityCode));
        analysis.CriticalPath.AddRange(comparison.Activities
            .Where(item => item.UpdateCritical && !string.Equals(item.UpdateStatus, "Completed", StringComparison.OrdinalIgnoreCase))
            .OrderBy(item => item.UpdateStart).ThenBy(item => item.ActivityCode));
        return analysis;
    }

    private static void BuildCurve(XerProgressAnalysis analysis)
    {
        var items = analysis.Comparison.Activities;
        var starts = items.SelectMany(item => new[] { item.BaselineStart, item.UpdateStart }).Where(date => date.HasValue).Select(date => date!.Value.Date).ToList();
        var finishes = items.SelectMany(item => new[] { item.BaselineFinish, item.UpdateFinish }).Where(date => date.HasValue).Select(date => date!.Value.Date).ToList();
        if (starts.Count == 0 || finishes.Count == 0)
            return;
        var start = starts.Min();
        start = start.AddDays(-((7 + (int)start.DayOfWeek - (int)DayOfWeek.Monday) % 7));
        var finish = finishes.Max();
        var weeks = Math.Min(520, Math.Max(1, (int)Math.Ceiling((finish - start).TotalDays / 7d) + 1));
        for (var week = 0; week < weeks; week++)
        {
            var date = start.AddDays(week * 7);
            analysis.Curve.Add(new XerProgressPoint(
                date,
                XerProgressWeighting.Calculate(items, item => LinearProgress(item.BaselineStart, item.BaselineFinish, date), analysis.CostWeighted),
                XerProgressWeighting.Calculate(items, item => EstimatedActual(item, date, analysis.DataDate), analysis.CostWeighted),
                XerProgressWeighting.Calculate(items, item => Forecast(item, date, analysis.DataDate), analysis.CostWeighted)));
        }
        if (analysis.Curve.All(point => point.Date.Date != analysis.DataDate.Date))
        {
            analysis.Curve.Add(new XerProgressPoint(analysis.DataDate.Date, analysis.PlannedPercent, analysis.ActualPercent, analysis.ActualPercent));
            analysis.Curve.Sort((left, right) => left.Date.CompareTo(right.Date));
        }
    }

    private static double EstimatedActual(XerActivityComparison item, DateTime date, DateTime dataDate)
    {
        if (date >= dataDate) return item.ActualPercent;
        if (!item.UpdateStart.HasValue || date <= item.UpdateStart.Value) return 0d;
        var actualWindow = (dataDate - item.UpdateStart.Value).TotalDays;
        return actualWindow <= 0d ? item.ActualPercent : Math.Clamp(item.ActualPercent * (date - item.UpdateStart.Value).TotalDays / actualWindow, 0d, item.ActualPercent);
    }

    private static double Forecast(XerActivityComparison item, DateTime date, DateTime dataDate)
    {
        if (date <= dataDate) return EstimatedActual(item, date, dataDate);
        if (item.ActualPercent >= 0.999999) return 1d;
        if (!item.UpdateFinish.HasValue || item.UpdateFinish.Value <= dataDate) return item.ActualPercent;
        var remainingWindow = (item.UpdateFinish.Value - dataDate).TotalDays;
        var elapsed = (date - dataDate).TotalDays;
        return Math.Clamp(item.ActualPercent + (1d - item.ActualPercent) * elapsed / remainingWindow, item.ActualPercent, 1d);
    }

    private static double LinearProgress(DateTime? start, DateTime? finish, DateTime date)
    {
        if (!start.HasValue || !finish.HasValue || date <= start.Value) return 0d;
        if (date >= finish.Value) return 1d;
        var duration = (finish.Value - start.Value).TotalDays;
        return duration <= 0d ? 0d : Math.Clamp((date - start.Value).TotalDays / duration, 0d, 1d);
    }

    private static bool Intersects(DateTime? start, DateTime? finish, DateTime windowStart, DateTime windowFinish)
    {
        if (!start.HasValue && !finish.HasValue) return false;
        var activityStart = start ?? finish!.Value;
        var activityFinish = finish ?? start!.Value;
        return activityStart <= windowFinish && activityFinish >= windowStart;
    }

}

public static class XerProgressWeighting
{
    public static bool HasBaselineCost(IReadOnlyList<XerActivityComparison> items) => TotalBaselineCost(items) > 0.000001d;

    public static double TotalBaselineCost(IReadOnlyList<XerActivityComparison> items) => items
        .Where(IsEligible)
        .Sum(item => Math.Max(item.BaselineBudgetCost, 0d));

    public static double Calculate(IReadOnlyList<XerActivityComparison> items, Func<XerActivityComparison, double> selector, bool costWeighted)
    {
        var eligible = items.Where(IsEligible).ToList();
        if (eligible.Count == 0) return 0d;
        var weights = eligible.Select(item => Weight(item, costWeighted)).ToArray();
        var total = weights.Sum();
        if (total <= 0d)
            return costWeighted ? 0d : Math.Clamp(eligible.Average(selector), 0d, 1d);
        var weightedValue = 0d;
        for (var index = 0; index < eligible.Count; index++)
            weightedValue += selector(eligible[index]) * weights[index];
        return Math.Clamp(weightedValue / total, 0d, 1d);
    }

    public static double Weight(XerActivityComparison item, bool costWeighted) => costWeighted
        ? Math.Max(item.BaselineBudgetCost, 0d)
        : DurationWeight(item);

    private static bool IsEligible(XerActivityComparison item) => item.BaselineStart.HasValue && item.BaselineFinish.HasValue;

    private static double DurationWeight(XerActivityComparison item)
    {
        if (item.BaselineDurationDays > 0d) return item.BaselineDurationDays;
        return item.BaselineStart.HasValue && item.BaselineFinish.HasValue
            ? Math.Max((item.BaselineFinish.Value - item.BaselineStart.Value).TotalDays, 0d)
            : 0d;
    }
}

using System.Globalization;
using PlanningToolkit.Core;
using PlanningToolkit.Core.Text;
using PlanningToolkit.Core.Xer;
using PlanningToolkit.Infrastructure.Logging;
using PlanningToolkit.Infrastructure.Settings;

namespace PlanningToolkit.Tests;

internal static class Program
{
    private static int Main()
    {
        var tests = new (string Name, Action Test)[]
        {
            ("TrimWhitespace", TrimWhitespace),
            ("RemoveNonPrintingCharacters", RemoveNonPrintingCharacters),
            ("CaseConversions", CaseConversions),
            ("SentenceCase", SentenceCase),
            ("KnownDateFormats", KnownDateFormats),
            ("InvalidDate", InvalidDate),
            ("DefaultSettingsAreValid", DefaultSettingsAreValid),
            ("InvalidSettingsAreRejected", InvalidSettingsAreRejected),
            ("SettingsRoundTrip", SettingsRoundTrip),
            ("InvalidJsonFallsBackToDefaults", InvalidJsonFallsBackToDefaults),
            ("ParseXerTables", ParseXerTables),
            ("ValidateGoodXer", ValidateGoodXer),
            ("DetectXerColumnMismatch", DetectXerColumnMismatch),
            ("DetectDuplicateTaskId", DetectDuplicateTaskId),
            ("BuildScheduleModel", BuildScheduleModel),
            ("UseBaselineTargetDates", UseBaselineTargetDates),
            ("DetectCriticalActivity", DetectCriticalActivity),
            ("CompareBaselineAndUpdate", CompareBaselineAndUpdate),
            ("DetectAddedAndDeletedActivities", DetectAddedAndDeletedActivities),
            ("CalculatePlannedProgressAtDataDate", CalculatePlannedProgressAtDataDate),
            ("BuildProgressAnalysis", BuildProgressAnalysis),
            ("CostWeightedPlannedProgress", CostWeightedPlannedProgress),
            ("BuildLookAheadAndCriticalPath", BuildLookAheadAndCriticalPath),
            ("ConfigureLookAheadWeeks", ConfigureLookAheadWeeks),
            ("BaselineCurveReaches100WithAddedActivities", BaselineCurveReaches100WithAddedActivities),
            ("SerializeXerRoundTrip", SerializeXerRoundTrip),
            ("ApplyEditableXerChanges", ApplyEditableXerChanges),
            ("RejectReadOnlyXerIdChange", RejectReadOnlyXerIdChange),
            ("DetectBrokenXerReference", DetectBrokenXerReference),
            ("ExportXerWithBackup", ExportXerWithBackup)
        };

        var failures = 0;
        Console.WriteLine($"Running {tests.Length} Planning Toolkit Phase 6 tests...");

        foreach (var (name, test) in tests)
        {
            try
            {
                test();
                Console.WriteLine($"PASS {name}");
            }
            catch (Exception exception)
            {
                failures++;
                Console.Error.WriteLine($"FAIL {name}: {exception.Message}");
            }
        }

        Console.WriteLine(failures == 0
            ? $"All {tests.Length} tests passed."
            : $"{failures} of {tests.Length} tests failed.");
        return failures == 0 ? 0 : 1;
    }

    private static void TrimWhitespace() => Equal(
        "Alpha Beta",
        TextTransforms.TrimWhitespace("  Alpha\t\tBeta  "));

    private static void RemoveNonPrintingCharacters() => Equal(
        "AB\tC",
        TextTransforms.RemoveNonPrintingCharacters("A\u0001B\tC"));

    private static void CaseConversions()
    {
        var culture = CultureInfo.GetCultureInfo("en-US");
        Equal("SCHEDULE UPDATE", TextTransforms.ToUpper("Schedule Update", culture));
        Equal("schedule update", TextTransforms.ToLower("Schedule Update", culture));
        Equal("Schedule Update", TextTransforms.ToProperCase("sCHEDULE uPDATE", culture));
    }

    private static void SentenceCase() => Equal(
        "First sentence. Second sentence! Third?",
        TextTransforms.ToSentenceCase("FIRST SENTENCE. SECOND SENTENCE! THIRD?", CultureInfo.GetCultureInfo("en-US")));

    private static void KnownDateFormats()
    {
        True(DateParser.TryParse("20-Jul-2026", out var first));
        Equal(new DateTime(2026, 7, 20), first.Date);
        True(DateParser.TryParse("2026-07-20", out var second));
        Equal(new DateTime(2026, 7, 20), second.Date);
    }

    private static void InvalidDate() => False(DateParser.TryParse("not-a-date", out _));

    private static void DefaultSettingsAreValid() => Equal(0, new AppSettings().Validate().Count);

    private static void InvalidSettingsAreRejected()
    {
        var settings = new AppSettings { WorkingDaysPerWeek = 9, CurrencyCode = "RIYAL" };
        True(settings.Validate().Count >= 2);
    }

    private static void SettingsRoundTrip()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var path = Path.Combine(directory, "settings.json");
            var store = new JsonSettingsStore(path);
            store.Save(new AppSettings { CurrencyCode = "USD", WorkingHoursPerDay = 9 });
            var loaded = store.Load();
            Equal("USD", loaded.CurrencyCode);
            Equal(9d, loaded.WorkingHoursPerDay);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void InvalidJsonFallsBackToDefaults()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var path = Path.Combine(directory, "settings.json");
            File.WriteAllText(path, "{ invalid json }");
            var logger = new FileAppLogger(Path.Combine(directory, "logs"));
            var loaded = new JsonSettingsStore(path, logger).Load();
            Equal("SAR", loaded.CurrencyCode);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void ParseXerTables()
    {
        var document = ParseSample();
        Equal(3, document.Tables.Count);
        Equal(1, document.FindTable("PROJECT")!.Rows.Count);
        Equal("Activity A", document.FindTable("TASK")!.Rows[0].Values[2]);
    }

    private static void ValidateGoodXer()
    {
        var result = XerValidator.Validate(ParseSample());
        True(result.IsValid);
        Equal(0, result.ErrorCount);
    }

    private static void DetectXerColumnMismatch()
    {
        var text = "ERMHDR\t8.4\n%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\n%E\n";
        var result = XerValidator.Validate(XerParser.Parse(new StringReader(text)));
        True(result.Issues.Any(issue => issue.Code == "COLUMN_COUNT"));
    }

    private static void DetectDuplicateTaskId()
    {
        var text = "ERMHDR\t8.4\n%T\tTASK\n%F\ttask_id\ttask_name\n%R\t1\tA\n%R\t1\tB\n%E\n";
        var result = XerValidator.Validate(XerParser.Parse(new StringReader(text)));
        True(result.Issues.Any(issue => issue.Code == "ID_DUPLICATE"));
    }

    private static XerDocument ParseSample()
    {
        const string text =
            "ERMHDR\t8.4\n" +
            "%T\tPROJECT\n%F\tproj_id\tproj_short_name\n%R\t10\tDemo\n" +
            "%T\tPROJWBS\n%F\twbs_id\tproj_id\twbs_name\n%R\t20\t10\tRoot\n" +
            "%T\tTASK\n%F\ttask_id\tproj_id\ttask_name\n%R\t30\t10\tActivity A\n%E\n";
        return XerParser.Parse(new StringReader(text));
    }

    private static void BuildScheduleModel()
    {
        var document = XerParser.Parse(new StringReader(ScheduleSample));
        var schedule = XerScheduleBuilder.Build(document);
        Equal("Demo Project", schedule.ProjectName);
        Equal(new DateTime(2026, 7, 20), schedule.DataDate!.Value.Date);
        Equal(2, schedule.WbsItems.Count);
        Equal(2, schedule.Activities.Count);
        Equal("In Progress", schedule.Activities[0].Status);
        Equal(10d, schedule.Activities[0].OriginalDurationDays);
        Equal(1000d, schedule.Activities[0].BudgetedCost);
    }

    private static void UseBaselineTargetDates()
    {
        const string text =
            "ERMHDR\t8.4\n" +
            "%T\tTASK\n%F\ttask_id\ttask_code\ttask_name\tstatus_code\ttarget_drtn_hr_cnt\tact_start_date\tact_end_date\tearly_start_date\tearly_end_date\ttarget_start_date\ttarget_end_date\n" +
            "%R\t1\tA1000\tTest\tTK_Complete\t80\t2026-07-03 08:00\t2026-07-09 17:00\t2026-07-04 08:00\t2026-07-10 17:00\t2026-07-01 08:00\t2026-07-15 17:00\n%E\n";
        var document = XerParser.Parse(new StringReader(text));
        var baseline = XerScheduleBuilder.Build(document, 8d, XerScheduleDateMode.Baseline).Activities.Single();
        var current = XerScheduleBuilder.Build(document, 8d, XerScheduleDateMode.Current).Activities.Single();
        Equal(new DateTime(2026, 7, 1, 8, 0, 0), baseline.Start!.Value);
        Equal(new DateTime(2026, 7, 15, 17, 0, 0), baseline.Finish!.Value);
        Equal(new DateTime(2026, 7, 3, 8, 0, 0), current.Start!.Value);
        Equal(new DateTime(2026, 7, 9, 17, 0, 0), current.Finish!.Value);
    }

    private static void DetectCriticalActivity()
    {
        var schedule = XerScheduleBuilder.Build(XerParser.Parse(new StringReader(ScheduleSample)));
        True(schedule.Activities.Single(activity => activity.Code == "A1000").IsCritical);
        False(schedule.Activities.Single(activity => activity.Code == "A1010").IsCritical);
    }

    private static void CompareBaselineAndUpdate()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 10), false));
        var update = new XerSchedule();
        update.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 15), true));
        var result = XerScheduleComparer.Compare(baseline, update);
        Equal(1, result.Activities.Count);
        Equal(5d, result.Activities[0].FinishVarianceDays);
        Equal("Newly critical", result.Activities[0].UpdateCritical && !result.Activities[0].BaselineCritical ? "Newly critical" : string.Empty);
    }

    private static void DetectAddedAndDeletedActivities()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "DELETED", new DateTime(2026, 7, 10), false));
        var update = new XerSchedule();
        update.Activities.Add(Activity("2", "ADDED", new DateTime(2026, 7, 10), false));
        var result = XerScheduleComparer.Compare(baseline, update);
        Equal(1, result.AddedCount);
        Equal(1, result.DeletedCount);
    }

    private static void CalculatePlannedProgressAtDataDate()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 11), false, 0));
        var update = new XerSchedule { DataDate = new DateTime(2026, 7, 8, 12, 0, 0) };
        update.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 12), false, 0.4));
        var item = XerScheduleComparer.Compare(baseline, update).Activities.Single();
        Equal(0.5d, item.PlannedPercent);
        Equal(0.4d, item.ActualPercent);
        True(Math.Abs(item.ProgressVariance - (-0.1d)) < 0.0001);
    }

    private static void BuildProgressAnalysis()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 11), false, 0));
        var update = new XerSchedule { DataDate = new DateTime(2026, 7, 8, 12, 0, 0) };
        update.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 12), false, 0.4));
        var analysis = XerProgressAnalyzer.Analyze(baseline, update);
        Equal(0.5d, analysis.PlannedPercent);
        Equal(0.4d, analysis.ActualPercent);
        True(Math.Abs(analysis.Spi - 0.8d) < 0.0001);
        True(analysis.Curve.Count > 0);
    }

    private static void CostWeightedPlannedProgress()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "LOW_COST", new DateTime(2026, 7, 5), false, budgetedCost: 100d));
        baseline.Activities.Add(Activity("2", "HIGH_COST", new DateTime(2026, 7, 15), false, budgetedCost: 900d));
        var update = new XerSchedule { DataDate = new DateTime(2026, 7, 10) };
        update.Activities.Add(Activity("1", "LOW_COST", new DateTime(2026, 7, 5), false, 1d));
        update.Activities.Add(Activity("2", "HIGH_COST", new DateTime(2026, 7, 15), false, 0d));

        var analysis = XerProgressAnalyzer.Analyze(baseline, update);
        True(analysis.CostWeighted);
        Equal(1000d, analysis.BaselineBudgetCost);
        Equal(0.1d, analysis.PlannedPercent);
        Equal(0.1d, analysis.ActualPercent);
    }

    private static void BuildLookAheadAndCriticalPath()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 15), false));
        var update = new XerSchedule { DataDate = new DateTime(2026, 7, 10) };
        update.Activities.Add(Activity("1", "A1000", new DateTime(2026, 7, 16), true, 0.2));
        var analysis = XerProgressAnalyzer.Analyze(baseline, update);
        Equal(2, analysis.LookAheadWeeks);
        Equal(1, analysis.LookAhead.Count);
        Equal(1, analysis.CriticalPath.Count);
    }

    private static void ConfigureLookAheadWeeks()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "NEAR", new DateTime(2026, 7, 16), false));
        baseline.Activities.Add(Activity("2", "FAR", new DateTime(2026, 7, 25), false));
        var update = new XerSchedule { DataDate = new DateTime(2026, 7, 10) };
        update.Activities.Add(Activity("1", "NEAR", new DateTime(2026, 7, 16), false));
        update.Activities.Add(Activity("2", "FAR", new DateTime(2026, 7, 25), false));

        var oneWeek = XerProgressAnalyzer.Analyze(baseline, update, 1);
        var threeWeeks = XerProgressAnalyzer.Analyze(baseline, update, 3);
        Equal(1, oneWeek.LookAhead.Count);
        Equal(2, threeWeeks.LookAhead.Count);
        Equal(3, threeWeeks.LookAheadWeeks);
    }

    private static void BaselineCurveReaches100WithAddedActivities()
    {
        var baseline = new XerSchedule();
        baseline.Activities.Add(Activity("1", "BASELINE", new DateTime(2026, 7, 11), false));
        var update = new XerSchedule { DataDate = new DateTime(2026, 7, 8) };
        update.Activities.Add(Activity("1", "BASELINE", new DateTime(2026, 7, 12), false, 0.4));
        update.Activities.Add(Activity("2", "ADDED", new DateTime(2026, 7, 30), false, 0.1));
        var analysis = XerProgressAnalyzer.Analyze(baseline, update);
        Equal(1d, analysis.Curve.Last().Planned);
    }

    private static void SerializeXerRoundTrip()
    {
        var original = ParseSample();
        var writer = new StringWriter(CultureInfo.InvariantCulture);
        XerSerializer.Write(original, writer);
        var parsed = XerParser.Parse(new StringReader(writer.ToString()));
        Equal(original.Header, parsed.Header);
        Equal(original.Tables.Count, parsed.Tables.Count);
        Equal("Activity A", parsed.FindTable("TASK")!.Rows[0].Values[2]);
    }

    private static void ApplyEditableXerChanges()
    {
        var document = ParseSample();
        var snapshots = CreateSnapshots(document);
        var task = snapshots.Single(snapshot => snapshot.Name == "TASK");
        ((string[])task.Rows[0])[2] = "Edited Activity";
        var result = XerEditService.Apply(document, snapshots);
        Equal(1, result.EditedCells);
        Equal("Edited Activity", document.FindTable("TASK")!.Rows[0].Values[2]);
    }

    private static void RejectReadOnlyXerIdChange()
    {
        var document = ParseSample();
        var snapshots = CreateSnapshots(document);
        var task = snapshots.Single(snapshot => snapshot.Name == "TASK");
        ((string[])task.Rows[0])[0] = "999";
        Throws<InvalidOperationException>(() => XerEditService.Apply(document, snapshots));
    }

    private static void DetectBrokenXerReference()
    {
        const string text =
            "ERMHDR\t8.4\n" +
            "%T\tPROJECT\n%F\tproj_id\n%R\t10\n" +
            "%T\tPROJWBS\n%F\twbs_id\tproj_id\n%R\t20\t999\n" +
            "%T\tTASK\n%F\ttask_id\tproj_id\n%R\t30\t10\n%E\n";
        var validation = XerValidator.Validate(XerParser.Parse(new StringReader(text)));
        True(validation.Issues.Any(issue => issue.Code == "REFERENCE_MISSING"));
    }

    private static void ExportXerWithBackup()
    {
        var directory = CreateTemporaryDirectory();
        try
        {
            var source = Path.Combine(directory, "source.xer");
            var output = Path.Combine(directory, "edited.xer");
            File.WriteAllText(source,
                "ERMHDR\t8.4\r\n" +
                "%T\tPROJECT\r\n%F\tproj_id\tproj_short_name\r\n%R\t10\tDemo\r\n" +
                "%T\tPROJWBS\r\n%F\twbs_id\tproj_id\twbs_name\r\n%R\t20\t10\tRoot\r\n" +
                "%T\tTASK\r\n%F\ttask_id\tproj_id\ttask_name\r\n%R\t30\t10\tActivity A\r\n%E\r\n",
                System.Text.Encoding.Latin1);
            var document = XerParser.Parse(source);
            var result = XerRoundTripExporter.Export(document, source, output);
            True(File.Exists(output));
            Equal(1, result.BackupPaths.Count);
            True(File.Exists(result.BackupPaths[0]));
            Equal("Activity A", XerParser.Parse(output).FindTable("TASK")!.Rows[0].Values[2]);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static List<XerTableSnapshot> CreateSnapshots(XerDocument document) => document.Tables
        .Select(table => new XerTableSnapshot(
            table.Name,
            table.Fields.ToArray(),
            table.Rows.Select(row => (IReadOnlyList<string>)row.Values.ToArray()).ToArray()))
        .ToList();

    private static XerScheduleActivity Activity(string id, string code, DateTime finish, bool critical, double percent = 0, double budgetedCost = 0) =>
        new(id, code, code, null, "Not Started", 5, 5, budgetedCost, finish.AddDays(-5), finish, percent, critical ? 0 : 5, critical, string.Empty, string.Empty);

    private const string ScheduleSample =
        "ERMHDR\t8.4\n" +
        "%T\tPROJECT\n%F\tproj_id\tproj_short_name\tlast_recalc_date\n%R\t10\tDemo Project\t2026-07-20 08:00\n" +
        "%T\tPROJWBS\n%F\twbs_id\tparent_wbs_id\twbs_short_name\twbs_name\tseq_num\n%R\t20\t\t1\tProject\t1\n%R\t21\t20\t1.1\tConstruction\t2\n" +
        "%T\tTASK\n%F\ttask_id\ttask_code\ttask_name\twbs_id\tstatus_code\ttarget_drtn_hr_cnt\tremain_drtn_hr_cnt\tphys_complete_pct\tearly_start_date\tearly_end_date\ttotal_float_hr_cnt\n" +
        "%R\t30\tA1000\tCritical Work\t21\tTK_Active\t80\t40\t50\t2026-07-20 08:00\t2026-07-30 17:00\t0\n" +
        "%R\t31\tA1010\tNoncritical Work\t21\tTK_NotStart\t40\t40\t0\t2026-07-31 08:00\t2026-08-05 17:00\t24\n" +
        "%T\tTASKRSRC\n%F\ttaskrsrc_id\ttask_id\ttarget_cost\n" +
        "%R\t100\t30\t1000\n" +
        "%R\t101\t31\t3000\n%E\n";

    private static string CreateTemporaryDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "PlanningToolkitTests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static void True(bool value)
    {
        if (!value)
            throw new InvalidOperationException("Expected true but received false.");
    }

    private static void False(bool value) => True(!value);

    private static void Equal<T>(T expected, T actual)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new InvalidOperationException($"Expected '{expected}' but received '{actual}'.");
    }

    private static void Throws<TException>(Action action) where TException : Exception
    {
        try
        {
            action();
        }
        catch (TException)
        {
            return;
        }
        throw new InvalidOperationException($"Expected exception '{typeof(TException).Name}'.");
    }
}

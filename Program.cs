using System.Globalization;
using PlanningToolkit.Core;
using PlanningToolkit.Core.Text;
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
            ("InvalidJsonFallsBackToDefaults", InvalidJsonFallsBackToDefaults)
        };

        var failures = 0;
        Console.WriteLine($"Running {tests.Length} Planning Toolkit Phase 1 tests...");

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
}

using System.ComponentModel;
using System.Globalization;

namespace PlanningToolkit.Core;

public sealed class AppSettings
{
    public const int CurrentSchemaVersion = 1;

    [Browsable(false)]
    public int SchemaVersion { get; set; } = CurrentSchemaVersion;

    [Category("General")]
    [Description("Excel number format used after converting text to dates.")]
    public string DateFormat { get; set; } = "dd-MMM-yyyy";

    [Category("General")]
    public string CurrencyCode { get; set; } = "SAR";

    [Category("Calendar")]
    public double WorkingHoursPerDay { get; set; } = 8;

    [Category("Calendar")]
    public int WorkingDaysPerWeek { get; set; } = 5;

    [Category("Schedule Thresholds")]
    public double CriticalFloatThresholdDays { get; set; }

    [Category("Schedule Thresholds")]
    public double LongDurationThresholdDays { get; set; } = 20;

    [Category("Schedule Thresholds")]
    public double LagThresholdDays { get; set; } = 10;

    [Category("Future Gantt Colors")]
    public string GanttBaselineColor { get; set; } = "#9E9E9E";

    [Category("Future Gantt Colors")]
    public string GanttCurrentColor { get; set; } = "#2196F3";

    [Category("Future Gantt Colors")]
    public string GanttCriticalColor { get; set; } = "#E53935";

    [Category("Future WBS Colors")]
    public string WbsLevel1Color { get; set; } = "#1F4E78";

    [Category("Future WBS Colors")]
    public string WbsLevel2Color { get; set; } = "#5B9BD5";

    [Category("General")]
    public string DefaultOutputFolder { get; set; } = GetDefaultOutputFolder();

    [Category("Excel")]
    [Description("RestorePrevious, KeepAutomatic or KeepManual.")]
    public string CalculationBehavior { get; set; } = "RestorePrevious";

    [Category("Diagnostics")]
    [Description("Error, Warning, Information or Debug.")]
    public string LoggingLevel { get; set; } = "Information";

    public IReadOnlyList<string> Validate()
    {
        var errors = new List<string>();

        if (string.IsNullOrWhiteSpace(DateFormat))
            errors.Add("Date format is required.");
        else
        {
            try
            {
                _ = DateTime.Today.ToString(DateFormat, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                errors.Add("Date format is invalid.");
            }
        }

        if (string.IsNullOrWhiteSpace(CurrencyCode) || CurrencyCode.Trim().Length != 3)
            errors.Add("Currency code must contain three letters.");

        if (WorkingHoursPerDay <= 0 || WorkingHoursPerDay > 24)
            errors.Add("Working hours per day must be between 0 and 24.");

        if (WorkingDaysPerWeek is < 1 or > 7)
            errors.Add("Working days per week must be between 1 and 7.");

        if (CriticalFloatThresholdDays < 0)
            errors.Add("Critical float threshold cannot be negative.");

        if (LongDurationThresholdDays <= 0)
            errors.Add("Long duration threshold must be greater than zero.");

        if (LagThresholdDays < 0)
            errors.Add("Lag threshold cannot be negative.");

        ValidateHexColor(GanttBaselineColor, nameof(GanttBaselineColor), errors);
        ValidateHexColor(GanttCurrentColor, nameof(GanttCurrentColor), errors);
        ValidateHexColor(GanttCriticalColor, nameof(GanttCriticalColor), errors);
        ValidateHexColor(WbsLevel1Color, nameof(WbsLevel1Color), errors);
        ValidateHexColor(WbsLevel2Color, nameof(WbsLevel2Color), errors);

        if (string.IsNullOrWhiteSpace(DefaultOutputFolder))
            errors.Add("Default output folder is required.");

        if (CalculationBehavior is not ("RestorePrevious" or "KeepAutomatic" or "KeepManual"))
            errors.Add("Calculation behavior is invalid.");

        if (LoggingLevel is not ("Error" or "Warning" or "Information" or "Debug"))
            errors.Add("Logging level is invalid.");

        return errors;
    }

    private static void ValidateHexColor(string value, string propertyName, ICollection<string> errors)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length != 7 || value[0] != '#' || !int.TryParse(value[1..], NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
            errors.Add($"{propertyName} must be a color in #RRGGBB format.");
    }

    private static string GetDefaultOutputFolder()
    {
        var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        return string.IsNullOrWhiteSpace(documents) ? Environment.CurrentDirectory : documents;
    }
}

using System.Globalization;

namespace PlanningToolkit.Core.Text;

public static class DateParser
{
    private static readonly string[] KnownFormats =
    {
        "dd-MMM-yyyy",
        "dd-MMM-yy",
        "dd/MM/yyyy",
        "d/M/yyyy",
        "MM/dd/yyyy",
        "M/d/yyyy",
        "yyyy-MM-dd",
        "yyyy/MM/dd",
        "dd.MM.yyyy"
    };

    public static bool TryParse(string value, out DateTime date)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            date = default;
            return false;
        }

        var trimmed = value.Trim();
        var cultures = new[]
        {
            CultureInfo.CurrentCulture,
            CultureInfo.InvariantCulture,
            CultureInfo.GetCultureInfo("en-US"),
            CultureInfo.GetCultureInfo("en-GB"),
            CultureInfo.GetCultureInfo("ar-SA")
        };

        foreach (var culture in cultures.DistinctBy(item => item.Name))
        {
            if (DateTime.TryParseExact(trimmed, KnownFormats, culture, DateTimeStyles.AllowWhiteSpaces, out date))
                return true;

            if (DateTime.TryParse(trimmed, culture, DateTimeStyles.AllowWhiteSpaces, out date))
                return true;
        }

        date = default;
        return false;
    }
}

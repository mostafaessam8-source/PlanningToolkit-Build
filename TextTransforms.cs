using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace PlanningToolkit.Core.Text;

public static partial class TextTransforms
{
    public static string TrimWhitespace(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return RepeatedWhitespace().Replace(value.Trim(), " ");
    }

    public static string RemoveNonPrintingCharacters(string value)
    {
        ArgumentNullException.ThrowIfNull(value);

        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
        {
            if (!char.IsControl(character) || character is '\t' or '\r' or '\n')
                builder.Append(character);
        }

        return builder.ToString();
    }

    public static string ToUpper(string value, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(value);
        return value.ToUpper(culture ?? CultureInfo.CurrentCulture);
    }

    public static string ToLower(string value, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(value);
        return value.ToLower(culture ?? CultureInfo.CurrentCulture);
    }

    public static string ToProperCase(string value, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(value);
        var selectedCulture = culture ?? CultureInfo.CurrentCulture;
        return selectedCulture.TextInfo.ToTitleCase(value.ToLower(selectedCulture));
    }

    public static string ToSentenceCase(string value, CultureInfo? culture = null)
    {
        ArgumentNullException.ThrowIfNull(value);
        if (string.IsNullOrWhiteSpace(value))
            return value;

        var selectedCulture = culture ?? CultureInfo.CurrentCulture;
        var lowered = value.ToLower(selectedCulture);
        var characters = lowered.ToCharArray();
        var capitalizeNext = true;

        for (var index = 0; index < characters.Length; index++)
        {
            if (capitalizeNext && char.IsLetter(characters[index]))
            {
                characters[index] = char.ToUpper(characters[index], selectedCulture);
                capitalizeNext = false;
            }

            if (characters[index] is '.' or '!' or '?' or '\r' or '\n')
                capitalizeNext = true;
        }

        return new string(characters);
    }

    [GeneratedRegex(@"[\p{Zs}\t]+", RegexOptions.CultureInvariant)]
    private static partial Regex RepeatedWhitespace();
}


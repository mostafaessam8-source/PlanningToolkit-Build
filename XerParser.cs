using System.Text;

namespace PlanningToolkit.Core.Xer;

public static class XerParser
{
    public static XerDocument Parse(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("An XER file path is required.", nameof(path));
        if (!File.Exists(path))
            throw new FileNotFoundException("The selected XER file was not found.", path);

        using var stream = File.OpenRead(path);
        using var reader = new StreamReader(stream, DetectEncoding(stream), detectEncodingFromByteOrderMarks: true);
        return Parse(reader);
    }

    public static XerDocument Parse(TextReader reader)
    {
        ArgumentNullException.ThrowIfNull(reader);
        var document = new XerDocument();
        XerTable? currentTable = null;
        string? line;
        var lineNumber = 0;

        while ((line = reader.ReadLine()) is not null)
        {
            lineNumber++;
            if (line.Length == 0)
                continue;

            var parts = line.Split('\t');
            switch (parts[0])
            {
                case "ERMHDR":
                    document.Header = line;
                    break;
                case "%T":
                    if (parts.Length < 2 || string.IsNullOrWhiteSpace(parts[1]))
                        throw new FormatException($"Missing table name at line {lineNumber}.");
                    currentTable = new XerTable(parts[1].Trim(), lineNumber);
                    document.Tables.Add(currentTable);
                    break;
                case "%F":
                    EnsureTable(currentTable, lineNumber, "%F");
                    currentTable!.Fields = parts.Skip(1).Select(field => field.Trim()).ToArray();
                    break;
                case "%R":
                    EnsureTable(currentTable, lineNumber, "%R");
                    currentTable!.Rows.Add(new XerRow(lineNumber, parts.Skip(1).ToArray()));
                    break;
                case "%E":
                    return document;
            }
        }

        return document;
    }

    private static void EnsureTable(XerTable? table, int lineNumber, string recordType)
    {
        if (table is null)
            throw new FormatException($"{recordType} record appears before a %T record at line {lineNumber}.");
    }

    private static Encoding DetectEncoding(Stream stream)
    {
        Span<byte> prefix = stackalloc byte[3];
        var bytesRead = stream.Read(prefix);
        stream.Position = 0;
        if (bytesRead >= 3 && prefix[0] == 0xEF && prefix[1] == 0xBB && prefix[2] == 0xBF)
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: true, throwOnInvalidBytes: true);
        return Encoding.Latin1;
    }
}

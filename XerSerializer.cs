using System.Text;

namespace PlanningToolkit.Core.Xer;

public static class XerSerializer
{
    public static void Write(XerDocument document, string path, Encoding? encoding = null)
    {
        ArgumentNullException.ThrowIfNull(document);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("An output XER file path is required.", nameof(path));

        var selectedEncoding = encoding ?? Encoding.Latin1;
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, selectedEncoding) { NewLine = "\r\n" };
        Write(document, writer);
    }

    public static void Write(XerDocument document, TextWriter writer)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(writer);
        if (string.IsNullOrWhiteSpace(document.Header))
            throw new InvalidOperationException("The XER document does not contain an ERMHDR header.");

        writer.WriteLine(document.Header);
        foreach (var table in document.Tables)
        {
            writer.WriteLine($"%T\t{table.Name}");
            writer.WriteLine("%F\t" + string.Join("\t", table.Fields));
            foreach (var row in table.Rows)
                writer.WriteLine("%R\t" + string.Join("\t", row.Values));
        }
        writer.WriteLine("%E");
    }

    public static Encoding DetectSourceEncoding(string sourcePath)
    {
        using var stream = File.OpenRead(sourcePath);
        Span<byte> prefix = stackalloc byte[3];
        var count = stream.Read(prefix);
        return count >= 3 && prefix[0] == 0xEF && prefix[1] == 0xBB && prefix[2] == 0xBF
            ? new UTF8Encoding(encoderShouldEmitUTF8Identifier: true, throwOnInvalidBytes: true)
            : Encoding.Latin1;
    }
}

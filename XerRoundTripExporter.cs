namespace PlanningToolkit.Core.Xer;

public sealed record XerExportFileResult(
    string OutputPath,
    IReadOnlyList<string> BackupPaths,
    int Tables,
    int Rows);

public static class XerRoundTripExporter
{
    public static XerExportFileResult Export(XerDocument document, string sourcePath, string outputPath)
    {
        ArgumentNullException.ThrowIfNull(document);
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("The original XER source file was not found.", sourcePath);
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("An output XER path is required.", nameof(outputPath));
        if (!string.Equals(Path.GetExtension(outputPath), ".xer", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("The exported file must use the .xer extension.");

        EnsureValid(document, "The edited workbook cannot be exported");

        var fullSourcePath = Path.GetFullPath(sourcePath);
        var fullOutputPath = Path.GetFullPath(outputPath);
        var outputDirectory = Path.GetDirectoryName(fullOutputPath)
            ?? throw new InvalidOperationException("The selected output folder is invalid.");
        Directory.CreateDirectory(outputDirectory);

        var backupPaths = new List<string>
        {
            CreateBackup(fullSourcePath, outputDirectory, "source-backup")
        };
        if (File.Exists(fullOutputPath) && !string.Equals(fullOutputPath, fullSourcePath, StringComparison.OrdinalIgnoreCase))
            backupPaths.Add(CreateBackup(fullOutputPath, outputDirectory, "replaced-backup"));

        var temporaryPath = Path.Combine(
            outputDirectory,
            $".{Path.GetFileName(fullOutputPath)}.{Guid.NewGuid():N}.tmp");
        try
        {
            XerSerializer.Write(document, temporaryPath, XerSerializer.DetectSourceEncoding(fullSourcePath));
            var roundTrip = XerParser.Parse(temporaryPath);
            EnsureValid(roundTrip, "The generated XER failed round-trip validation");
            VerifyExactRoundTrip(document, roundTrip);
            File.Move(temporaryPath, fullOutputPath, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);
        }

        var finalDocument = XerParser.Parse(fullOutputPath);
        EnsureValid(finalDocument, "The final XER failed verification");
        VerifyExactRoundTrip(document, finalDocument);

        return new XerExportFileResult(
            fullOutputPath,
            backupPaths,
            finalDocument.Tables.Count,
            finalDocument.Tables.Sum(table => table.Rows.Count));
    }

    private static string CreateBackup(string path, string outputDirectory, string label)
    {
        var timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        var baseName = Path.GetFileNameWithoutExtension(path);
        var candidate = Path.Combine(outputDirectory, $"{baseName}.{label}-{timestamp}.xer");
        var suffix = 2;
        while (File.Exists(candidate))
            candidate = Path.Combine(outputDirectory, $"{baseName}.{label}-{timestamp}-{suffix++}.xer");
        File.Copy(path, candidate, overwrite: false);
        return candidate;
    }

    private static void EnsureValid(XerDocument document, string prefix)
    {
        var validation = XerValidator.Validate(document);
        if (validation.IsValid)
            return;
        var details = validation.Issues
            .Where(issue => issue.Severity == XerIssueSeverity.Error)
            .Take(5)
            .Select(issue => $"{issue.Code}: {issue.Message}");
        throw new InvalidOperationException(
            $"{prefix}. Found {validation.ErrorCount:N0} critical error(s):\n" + string.Join("\n", details));
    }

    private static void VerifyExactRoundTrip(XerDocument expected, XerDocument actual)
    {
        if (!string.Equals(expected.Header, actual.Header, StringComparison.Ordinal) || expected.Tables.Count != actual.Tables.Count)
            throw new InvalidOperationException("Round-trip verification changed the XER header or table count.");

        for (var tableIndex = 0; tableIndex < expected.Tables.Count; tableIndex++)
        {
            var left = expected.Tables[tableIndex];
            var right = actual.Tables[tableIndex];
            if (!string.Equals(left.Name, right.Name, StringComparison.Ordinal) ||
                !left.Fields.SequenceEqual(right.Fields, StringComparer.Ordinal) ||
                left.Rows.Count != right.Rows.Count)
                throw new InvalidOperationException($"Round-trip verification changed table '{left.Name}'.");

            for (var rowIndex = 0; rowIndex < left.Rows.Count; rowIndex++)
            {
                if (!left.Rows[rowIndex].Values.SequenceEqual(right.Rows[rowIndex].Values, StringComparer.Ordinal))
                    throw new InvalidOperationException(
                        $"Round-trip verification changed table '{left.Name}' row {rowIndex + 1}.");
            }
        }
    }
}

namespace PlanningToolkit.Excel.Excel;

/// <summary>
/// Single source of truth for every color, font, border, table and print setting used across
/// Planning Toolkit report sheets. Report generators (Schedule, Comparison, PMS, Gantt) call
/// into this class instead of hard-coding colors or formatting so the whole workbook stays
/// visually consistent, and the brand can be changed from one place.
/// </summary>
internal static class ReportTheme
{
    // ---- Brand palette (Excel COM colors are 0xBBGGRR, not 0xRRGGBB) --------------------
    public const int HeaderFill = 0x73461F;      // primary brand band (dark teal-brown)
    public const int HeaderFont = 0xFFFFFF;      // white text on header band
    public const int TitleFont = 0x3F2410;       // dark brand tone for report titles
    public const int SubtitleFont = 0x595959;    // muted grey for secondary labels

    public const int Level1Fill = 0xC6E0B4;      // WBS level 1 summary rows (green)
    public const int Level2Fill = 0xDDEBF7;      // WBS level 2 summary rows (blue)
    public const int Level3Fill = 0xE2F0D9;      // WBS level 3+ summary rows (light green)

    public const int CriticalFont = 0x0000C0;    // critical-path activities
    public const int CriticalFill = 0xCEC7FF;    // critical float highlight
    public const int WarningFont = 0xC07000;     // near-critical / non-critical bars
    public const int DataDateFill = 0xD9EAD3;    // data-date column marker

    public const int PositiveFill = 0xD9EAD3;    // favorable variance
    public const int NegativeFill = 0xCEC7FF;    // unfavorable variance
    public const int NeutralFill = 0xF2F2F2;     // no-change rows

    public const int BorderColor = 0xBFBFBF;     // light grey grid lines
    public const int KpiCardFill = 0xDDEBF7;     // KPI tile background

    // ---- Typography ------------------------------------------------------------------------
    public const string FontName = "Segoe UI";
    public const double TitleSize = 16;
    public const double SubtitleSize = 10;
    public const double HeaderSize = 10;
    public const double BodySize = 10;
    public const double KpiValueSize = 18;
    public const double KpiLabelSize = 9;

    // ---- Excel COM constants (kept local so we don't need the typed Interop assembly) ------
    private const int XlContinuous = 1;
    private const int XlThin = 2;
    private const int XlCenter = -4108;
    private const int XlEdgeLeft = 7, XlEdgeTop = 8, XlEdgeBottom = 9, XlEdgeRight = 10;
    private const int XlInsideVertical = 11, XlInsideHorizontal = 12;
    private const int XlLandscape = 2, XlPortrait = 1;
    private const int XlTypePdf = 0;
    private const int XlSrcRange = 1;
    private const int XlYes = 1;

    /// <summary>Report title row (row 1 style) — bold brand-colored text, no fill.</summary>
    public static void ApplyTitle(dynamic range)
    {
        range.Font.Name = FontName;
        range.Font.Bold = true;
        range.Font.Size = TitleSize;
        range.Font.Color = TitleFont;
    }

    /// <summary>Secondary info line under the title (project name, data date, etc).</summary>
    public static void ApplySubtitle(dynamic range)
    {
        range.Font.Name = FontName;
        range.Font.Size = SubtitleSize;
        range.Font.Color = SubtitleFont;
    }

    /// <summary>Standard column-header band used on every report table.</summary>
    public static void ApplyHeaderRow(dynamic range)
    {
        range.Font.Name = FontName;
        range.Font.Bold = true;
        range.Font.Size = HeaderSize;
        range.Interior.Color = HeaderFill;
        range.Font.Color = HeaderFont;
        range.VerticalAlignment = XlCenter;
        range.RowHeight = 20;
    }

    /// <summary>Applies the base body font to a data range without touching fill/color.</summary>
    public static void ApplyBodyFont(dynamic range)
    {
        range.Font.Name = FontName;
        range.Font.Size = BodySize;
    }

    /// <summary>Thin grey grid around and inside a data range so tables never look "bare".</summary>
    public static void ApplyGridBorders(dynamic range)
    {
        foreach (var edge in new[] { XlEdgeLeft, XlEdgeTop, XlEdgeBottom, XlEdgeRight, XlInsideVertical, XlInsideHorizontal })
        {
            dynamic border = range.Borders[edge];
            border.LineStyle = XlContinuous;
            border.Weight = XlThin;
            border.Color = BorderColor;
        }
    }

    /// <summary>
    /// Converts a plain range into a native Excel Table (ListObject) so the user gets AutoFilter,
    /// banded rows and structured references for free, and applies the brand table style.
    /// Falls back silently (returns null) if the range already contains a table.
    /// </summary>
    public static dynamic? ConvertToTable(dynamic sheet, dynamic headerAndDataRange, string tableName)
    {
        try
        {
            dynamic table = sheet.ListObjects.Add(XlSrcRange, headerAndDataRange, Type.Missing, XlYes, Type.Missing);
            table.Name = tableName;
            table.TableStyle = "TableStyleMedium9";
            return table;
        }
        catch (Exception)
        {
            // A range that is already part of a table, or a sheet that does not support
            // ListObjects (rare, e.g. during automated tests), should not break report output.
            return null;
        }
    }

    /// <summary>Consistent, print-ready page setup: A3/A4 landscape, fit-to-width, branded header/footer.</summary>
    public static void ApplyPrintSetup(dynamic sheet, string title, bool landscape = true)
    {
        dynamic pageSetup = sheet.PageSetup;
        pageSetup.Orientation = landscape ? XlLandscape : XlPortrait;
        pageSetup.Zoom = false;
        pageSetup.FitToPagesWide = 1;
        pageSetup.FitToPagesTall = false;
        pageSetup.LeftMargin = sheet.Application.CentimetersToPoints(1.0);
        pageSetup.RightMargin = sheet.Application.CentimetersToPoints(1.0);
        pageSetup.TopMargin = sheet.Application.CentimetersToPoints(1.8);
        pageSetup.BottomMargin = sheet.Application.CentimetersToPoints(1.4);
        pageSetup.HeaderMargin = sheet.Application.CentimetersToPoints(0.8);
        pageSetup.FooterMargin = sheet.Application.CentimetersToPoints(0.8);
        pageSetup.CenterHeader = $"&\"Segoe UI,Bold\"&12{title}";
        pageSetup.RightHeader = "&\"Segoe UI\"&9&D";
        pageSetup.LeftHeader = "&\"Segoe UI\"&9Planning Toolkit";
        pageSetup.CenterFooter = "&\"Segoe UI\"&8Page &P of &N";
        pageSetup.PrintGridlines = false;
        pageSetup.PrintQuality = 600;
    }

    /// <summary>Exports the given worksheet(s) selection to a single PDF file.</summary>
    public static void ExportSheetToPdf(dynamic sheetOrRange, string outputPath)
        => sheetOrRange.ExportAsFixedFormat(
            XlTypePdf, outputPath, Type.Missing, true,
            false, Type.Missing, Type.Missing, false, Type.Missing);

    /// <summary>Draws a small KPI "card": bold label on top, large value underneath, tinted fill.</summary>
    public static void ApplyKpiCard(dynamic labelCell, dynamic valueCell, int? fillOverride = null)
    {
        labelCell.Font.Name = FontName;
        labelCell.Font.Size = KpiLabelSize;
        labelCell.Font.Bold = true;
        labelCell.Interior.Color = fillOverride ?? KpiCardFill;

        valueCell.Font.Name = FontName;
        valueCell.Font.Size = KpiValueSize;
        valueCell.Font.Bold = true;
        valueCell.Interior.Color = fillOverride ?? KpiCardFill;
    }

    /// <summary>Fill color for a WBS summary row given its outline level (1 = top level).</summary>
    public static int WbsLevelFill(int level) => level switch
    {
        1 => Level1Fill,
        2 => Level2Fill,
        _ => Level3Fill,
    };
}

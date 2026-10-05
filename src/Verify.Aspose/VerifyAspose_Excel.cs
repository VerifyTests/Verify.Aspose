using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace VerifyTests;

public static partial class VerifyAspose
{
    class ExcelFontWarningCallback : IWarningCallback
    {
        public void Warning(WarningInfo info)
        {
            if (info.Type != ExceptionType.FontSubstitution)
            {
                return;
            }

            throw new(
                $"""
                 Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
                 Details: {info.Description}
                 """);
        }
    }

    static void RenderEmptySheet()
    {
        //Aspose has an intermitant bug where it will null ref on render.
        //This is an attempt to mitigate that by forcing Aspose to initialize in a non threaded way
        using var book = new Workbook();
        var sheet = book.Worksheets[0];
        var render = new SheetRender(sheet, options);
        using var stream = new MemoryStream();
        render.ToImage(0, stream);
    }

    static ImageOrPrintOptions options = new()
    {
        ImageType = ImageType.Png,
        OnePagePerSheet = true,
        GridlineType = GridlineType.Hair,
        OnlyArea = true,
        PrintingPage = PrintingPageType.IgnoreBlank,
        WarningCallback = new ExcelFontWarningCallback()
    };

    static ConversionResult ConvertExcel(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var book = new Workbook(stream);
        return ConvertExcel(book, settings);
    }

    static ConversionResult ConvertExcel(Workbook book, IReadOnlyDictionary<string, object> settings)
    {
        // force dates in csv export to be consistent
        book.Settings.Region = CountryCode.USA;
        book.Settings.CultureInfo = CultureInfo.InvariantCulture;
        foreach (var sheet in book.Worksheets)
        {
            ScrubCells(sheet);
        }

        // A sheet is a page. Names the image of each, and says which sheets and which outputs the
        // verification wants, so what is left out is neither drawn nor exported.
        var conversion = new PagedConversion(settings)
        {
            Info = GetInfo(book)
        };

        // Building the deterministic xlsx is expensive, so skip it when the xlsx target is excluded.
        if (!settings.IsTargetExcluded("xlsx"))
        {
            conversion.Source(BuildXlsxTarget(book));
        }

        var sheets = book.Worksheets;
        foreach (var number in conversion.Pages(sheets.Count))
        {
            AddSheet(conversion, settings, number, sheets[number - 1]);
        }

        return conversion.Build();
    }

    // The xlsx snapshot is always the full workbook, regardless of PagesToInclude. It is the source
    // of the conversion, so it is never converted again, whichever of xls and xlsx was passed in.
    static Target BuildXlsxTarget(Workbook book)
    {
        using var source = new MemoryStream();
        book.Save(source, SaveFormat.Xlsx);
        var resultStream = DeterministicPackage.Convert(source);

        return new("xlsx", resultStream);
    }

    static object GetInfo(Workbook book) =>
        new
        {
            HasMacro = book.HasMacro.ToString(),
            HasRevisions = book.HasRevisions.ToString(),
            IsDigitallySigned = book.IsDigitallySigned.ToString(),
            Properties = GetProperties(book),
            CustomProperties = GetCustomProperties(book),
            Fonts = GetFonts(book),
            HiddenSheets = HiddenSheets(book.Worksheets)
        };

    // The names of the sheets that are hidden, whether Excel can unhide them or only code can.
    // Null, so left out, when there are none. A hidden sheet is a page as any other, so this is
    // what says it is hidden.
    static List<string>? HiddenSheets(IEnumerable<Worksheet> sheets)
    {
        var hidden = sheets
            .Where(_ => _.VisibilityType != VisibilityType.Visible)
            .Select(_ => _.Name)
            .ToList();
        if (hidden.Count == 0)
        {
            return null;
        }

        return hidden;
    }

    static List<string> GetFonts(Workbook book)
    {
        var fonts = new HashSet<string>();

        foreach (var sheet in book.Worksheets)
        {
            var cells = sheet.Cells;
            var maxRow = cells.MaxDataRow;
            var maxCol = cells.MaxDataColumn;

            for (var row = 0; row <= maxRow; row++)
            {
                for (var col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];
                    if (cell.HasValue())
                    {
                        var style = cell.GetStyle();
                        if (style?.Font?.Name != null)
                        {
                            fonts.Add(style.Font.Name);
                        }
                    }
                }
            }
        }

        // Excel doesn't typically embed fonts in the same way as Word/PDF
        // Fonts are referenced from the system
        return fonts.OrderBy(_ => _).ToList();
    }

    static Dictionary<string, object> GetProperties(Workbook book) =>
        book.BuiltInDocumentProperties
            .Where(_ => _.Name != "LastSavedBy" &&
                        _.Name != "LastSavedTime" &&
                        _.Value.HasValue())
            .ToDictionary(_ => _.Name, _ => _.Value);

    static Dictionary<string, object> GetCustomProperties(Workbook book) =>
        book.CustomDocumentProperties
            .Where(_ => _.Value.HasValue())
            .ToDictionary(_ => _.Name, _ => _.Value);

    static ConversionResult ConvertSheet(Worksheet sheet, IReadOnlyDictionary<string, object> settings)
    {
        ScrubCells(sheet);

        // A sheet verified on its own is the one page, whichever sheet of its workbook it is
        var conversion = new PagedConversion(settings)
        {
            PageCount = 1
        };

        // There is nothing to say of the workbook here but that the sheet is hidden, when it is
        if (HiddenSheets([sheet]) is { } hidden)
        {
            conversion.Info = new
            {
                HiddenSheets = hidden
            };
        }

        AddSheet(conversion, settings, 1, sheet);
        return conversion.Build();
    }

    static Sheet GetInfo(Worksheet sheet) =>
        new(
            sheet.Name,
            GetColumns(sheet).ToList(),
            sheet.CustomProperties
                .ToDictionary(_ => _.Name, _ => _.Value),
            sheet.Hyperlinks);

    // A sheet as the page with the 1 based number: what there is to say of it, its image, and a
    // csv of it.
    static void AddSheet(PagedConversion conversion, IReadOnlyDictionary<string, object> settings, int number, Worksheet sheet)
    {
        var info = GetInfo(sheet);

        var setup = sheet.PageSetup;
        setup.PrintGridlines = true;
        setup.LeftMargin = 0;
        setup.TopMargin = 0;
        setup.RightMargin = 0;
        setup.BottomMargin = 0;

        // Not the text of the page, which would put it in the info file: a csv is a file of its own,
        // named by the sheet.
        if (!settings.IsDerivedTargetExcluded("csv"))
        {
            conversion.AddDerived(new("csv", ToCsv(sheet), sheet.Name));
        }

        Stream? image = null;
        if (conversion.IncludeImages)
        {
            image = RenderSheet(sheet);
        }

        conversion.AddPage(number, image, info: info);
    }

    // Aspose draws nothing for a hidden sheet, and every sheet is a page here. So a hidden one is
    // shown for as long as it takes to draw it, and then is as it was.
    static MemoryStream? RenderSheet(Worksheet sheet)
    {
        var visibility = sheet.VisibilityType;
        if (visibility == VisibilityType.Visible)
        {
            return RenderVisibleSheet(sheet);
        }

        sheet.VisibilityType = VisibilityType.Visible;
        try
        {
            return RenderVisibleSheet(sheet);
        }
        finally
        {
            sheet.VisibilityType = visibility;
        }
    }

    static MemoryStream? RenderVisibleSheet(Worksheet sheet)
    {
        var render = new SheetRender(sheet, options);

        // OnePagePerSheet draws a sheet as the one image, and IgnoreBlank draws none for a sheet
        // with nothing in it.
        if (render.PageCount == 0)
        {
            return null;
        }

        var stream = new MemoryStream();
        render.ToImage(0, stream);
        return stream;
    }

    static void ScrubCells(Worksheet sheet)
    {
        var counter = Counter.Current;
        var cells = sheet.Cells;
        var maxRow = cells.MaxDataRow;
        var maxCol = cells.MaxDataColumn;

        for (var row = 0; row <= maxRow; row++)
        {
            for (var col = 0; col <= maxCol; col++)
            {
                var cell = cells[row, col];
                var (value, replaceCellValue) = GetCellValue(cell, counter);
                if (replaceCellValue)
                {
                    cell.Value = value;
                }
            }
        }
    }

    static (string value, bool replaceCellValue) GetCellValue(Cell cell, Counter counter)
    {
        if (!cell.HasValue())
        {
            return (string.Empty, false);
        }

        switch (cell.Type)
        {
            case CellValueType.IsNumeric:
                var value = cell.DoubleValue;
                if (cell.GetStyle().Custom.Contains('%'))
                {
                    // Percentage
                    return (value.ToString("P", CultureInfo.InvariantCulture), false);
                }

                return (value.ToString(CultureInfo.InvariantCulture), false);

            case CellValueType.IsBool:
                return (cell.BoolValue.ToString(), false);

            case CellValueType.IsDateTime:
                var date = cell.DateTimeValue;
                if (counter.TryConvert(date, out var dateResult))
                {
                    return (dateResult, true);
                }

                return (DateFormatter.Convert(date), false);

            case CellValueType.IsError:
                return (cell.Value.ToString()!, false);

            case CellValueType.IsNull:
                return ("", false);

            default:
                var text = cell.StringValue;
                if (counter.TryConvert(text, out var result))
                {
                    return (result, true);
                }

                return (text, false);
        }
    }

    static string ToCsv(Worksheet sheet)
    {
        var utf8 = Encoding.UTF8;
        var saveOptions = new TxtSaveOptions
        {
            Encoding = utf8,
            TrimTailingBlankCells = true,
            FormatStrategy = CellValueFormatStrategy.DisplayString
        };
        var book = sheet.Workbook;
        book.Worksheets.ActiveSheetName = sheet.Name;
        using var stream = new MemoryStream();
        book.Save(stream, saveOptions);
        stream.Position = 0;
        using var reader = new StreamReader(stream, utf8);
        return reader.ReadToEnd();
    }

    static IEnumerable<ColumnInfo> GetColumns(Worksheet sheet)
    {
        var cells = sheet.Cells;
        var lastCell = cells.LastCell;
        if (lastCell == null)
        {
            yield break;
        }

        // Headers live on the first populated row. Reading a hardcoded row 0 would mistake a
        // leading hidden/empty row for the header and yield columns with null names.
        var headerRow = cells.FirstCell.Row;
        for (var column = 0; column <= lastCell.Column; column++)
        {
            var header = cells[headerRow, column];
            yield return new(
                header.Value,
                cells.GetColumnWidth(column));
        }
    }
}

record ColumnInfo(object Name, double Width);

record Sheet(string Name, List<ColumnInfo> Columns, Dictionary<string, string> Properties, HyperlinkCollection Hyperlinks);
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static ImageOrPrintOptions options = new()
    {
        ImageType = ImageType.Png,
        OnePagePerSheet = true,
        GridlineType = GridlineType.Hair,
        OnlyArea = true,
        PrintingPage = PrintingPageType.IgnoreBlank
    };

    static ConversionResult ConvertExcel(string? targetName, Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var book = new Workbook(stream);
        return ConvertExcel(targetName, book);
    }

    static ConversionResult ConvertExcel(string? targetName, Workbook book)
    {
        // force dates in csv export to be consistent
        book.Settings.Region = CountryCode.USA;
        book.Settings.CultureInfo = CultureInfo.InvariantCulture;
        foreach (var sheet in book.Worksheets)
        {
            ScrubCells(sheet);
        }

        var info = GetInfo(book);

        using var sourceStream = new MemoryStream();
        book.Save(sourceStream, SaveFormat.Xlsx);
        var resultStream = DeterministicPackage.Convert(sourceStream);

        List<Target> targets = [new("xlsx", resultStream, performConversion: false)];

        targets.AddRange(
            book.Worksheets
                .SelectMany(_ => GetSheetStreams(targetName, _)));

        return new(info, targets);
    }

    static object GetInfo(Workbook book) =>
        new
        {
            HasMacro = book.HasMacro.ToString(),
            HasRevisions = book.HasRevisions.ToString(),
            IsDigitallySigned = book.IsDigitallySigned.ToString(),
            Sheets = GetSheetData(book).ToList(),
            Properties = GetProperties(book),
            CustomProperties = GetCustomProperties(book)
        };

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

    static ConversionResult ConvertSheet(string? name, Worksheet sheet)
    {
        ScrubCells(sheet);
        var info = GetInfo(sheet);
        return new(info, GetSheetStreams(name, sheet).ToList());
    }

    static Sheet GetInfo(Worksheet sheet) =>
        new(sheet.Name, GetColumns(sheet).ToList(), sheet.CustomProperties
            .ToDictionary(_ => _.Name, _ => _.Value));

    static IEnumerable<Target> GetSheetStreams(string? targetName, Worksheet sheet)
    {
        var setup = sheet.PageSetup;
        setup.PrintGridlines = true;
        setup.LeftMargin = 0;
        setup.TopMargin = 0;
        setup.RightMargin = 0;
        setup.BottomMargin = 0;

        string targetAndSheet;
        if (targetName == null)
        {
            targetAndSheet = sheet.Name;
        }
        else
        {
            targetAndSheet = $"{targetName}-{sheet.Name}";
        }

        var csv = ToCsv(sheet);
        yield return new("csv", csv, targetAndSheet);
        var render = new SheetRender(sheet, options);

        if (render.PageCount == 1)
        {
            var stream = new MemoryStream();
            render.ToImage(0, stream);
            yield return new("png", stream, targetAndSheet);
            yield break;
        }

        for (var index = 0; index < render.PageCount; index++)
        {
            var stream = new MemoryStream();
            render.ToImage(index, stream);
            yield return new("png", stream, $"{targetAndSheet}_{index}");
        }
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

    static IEnumerable<Sheet> GetSheetData(Workbook book) =>
        book.Worksheets
            .Select(GetInfo);

    static IEnumerable<ColumnInfo> GetColumns(Worksheet sheet)
    {
        var cells = sheet.Cells;
        var lastCell = cells.LastCell;
        if (lastCell == null)
        {
            yield break;
        }

        for (var column = 0; column <= lastCell.Column; column++)
        {
            var header = cells[0, column];
            yield return new(
                header.Value,
                cells.GetColumnWidth(column));
        }
    }
}

record ColumnInfo(object Name, double Width);

record Sheet(string Name, List<ColumnInfo> Columns, Dictionary<string, string> Properties);
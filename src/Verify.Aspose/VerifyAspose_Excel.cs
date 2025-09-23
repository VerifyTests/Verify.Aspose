using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using DeterministicIoPackaging;
using TxtSaveOptions = Aspose.Cells.TxtSaveOptions;

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

    static ConversionResult ConvertExcel(string? name, Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var book = new Workbook(stream);
        return ConvertExcel(name, book);
    }

    static ConversionResult ConvertExcel(string? name, Workbook book)
    {
        //force dates in csv export to be consistent
        book.Settings.Region = CountryCode.USA;
        book.Settings.CultureInfo = CultureInfo.InvariantCulture;
        var info = GetInfo(book);
        return new(info, GetExcelStreams(name, book).ToList());
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
            .Where(_ => _.Value.HasValue())
            .ToDictionary(_ => _.Name, _ => _.Value);

    static Dictionary<string, object> GetCustomProperties(Workbook book) =>
        book.CustomDocumentProperties
            .Where(_ => _.Value.HasValue())
            .ToDictionary(_ => _.Name, _ => _.Value);

    static ConversionResult ConvertSheet(string? name, Worksheet sheet)
    {
        var info = GetInfo(sheet);
        return new(info, GetSheetStreams(name, sheet).ToList());
    }

    static Sheet GetInfo(Worksheet sheet) =>
        new(sheet.Name, GetColumns(sheet).ToList(), sheet.CustomProperties
            .ToDictionary(_ => _.Name, _ => _.Value));

    static IEnumerable<Target> GetExcelStreams(string? name, Workbook book)
    {
        book.BuiltInDocumentProperties.CreatedTime = DeterministicPackage.StableDate;
        book.BuiltInDocumentProperties.LastSavedTime = DeterministicPackage.StableDate;
        book.BuiltInDocumentProperties.Author = null;
        using var sourceStream = new MemoryStream();
        book.Save(sourceStream, SaveFormat.Xlsx);
        sourceStream.Position = 0;
        var resultStream = DeterministicPackage.Convert(sourceStream);

        List<Target> targets = [new("xlsx", resultStream, performConversion: false)];
        targets.AddRange(book.Worksheets.SelectMany(_ => GetSheetStreams(name, _)));
        return targets;
    }

    static IEnumerable<Target> GetSheetStreams(string? name, Worksheet sheet)
    {
        var setup = sheet.PageSetup;
        setup.PrintGridlines = true;
        setup.LeftMargin = 0;
        setup.TopMargin = 0;
        setup.RightMargin = 0;
        setup.BottomMargin = 0;

        var csv = ToCsv(sheet);
        yield return new("csv", csv);
        var render = new SheetRender(sheet, options);

        for (var index = 0; index < render.PageCount; index++)
        {
            var stream = new MemoryStream();
            render.ToImage(index, stream);
            yield return new("png", stream, name);
        }
    }

    static string ToCsv(Worksheet sheet)
    {
        var utf8 = Encoding.UTF8;
        var txtSaveOptions = new TxtSaveOptions
        {
            Encoding = utf8,
            TrimTailingBlankCells = true,
            FormatStrategy = CellValueFormatStrategy.DisplayString
        };
        var book = sheet.Workbook;
        book.Worksheets.ActiveSheetName = sheet.Name;
        using var stream = new MemoryStream();
        book.Save(stream, txtSaveOptions);
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
            var firstRow = cells[1, column];
            yield return new(
                header.Value,
                (uint) cells.GetColumnWidth(column),
                firstRow.Value);
        }
    }
}

record ColumnInfo(object Name, uint Width, object FirstValue);

record Sheet(string Name, List<ColumnInfo> Columns, Dictionary<string, string> Properties);
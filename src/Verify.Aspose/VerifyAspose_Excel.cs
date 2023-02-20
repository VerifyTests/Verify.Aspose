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

    static ConversionResult ConvertExcel(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var book = new Workbook(stream);
        return ConvertExcel(book, settings);
    }

    static ConversionResult ConvertExcel(Workbook book, IReadOnlyDictionary<string, object> settings)
    {
        var info = GetInfo(book);
        return new(info, GetExcelStreams(book).ToList());
    }

    static object GetInfo(Workbook book) =>
        new
        {
            HasMacro = book.HasMacro.ToString(),
            HasRevisions = book.HasRevisions.ToString(),
            IsDigitallySigned = book.IsDigitallySigned.ToString(),
            Sheets = GetSheetData(book).ToList(),
            Properties = GetDocumentProperties(book)
        };

    static Dictionary<string, object> GetDocumentProperties(Workbook book) =>
        book.BuiltInDocumentProperties
            .Where(_ => _.Value.HasValue())
            .ToDictionary(_ => _.Name, _ => _.Value);

    static ConversionResult ConvertSheet(Worksheet sheet, IReadOnlyDictionary<string, object> settings)
    {
        var info = GetInfo(sheet);
        return new(info, GetSheetStreams(sheet).ToList());
    }

    static object GetInfo(Worksheet sheet) =>
        new Sheet(sheet.Name, GetColumns(sheet).ToList());

    static IEnumerable<Target> GetExcelStreams(Workbook book) =>
        book.Worksheets.SelectMany(GetSheetStreams);

    static IEnumerable<Target> GetSheetStreams(Worksheet sheet)
    {
        var setup = sheet.PageSetup;
        setup.PrintGridlines = true;
        setup.LeftMargin = 0;
        setup.TopMargin = 0;
        setup.RightMargin = 0;
        setup.BottomMargin = 0;
        var render = new SheetRender(sheet, options);

        for (var index = 0; index < render.PageCount; index++)
        {
            var stream = new MemoryStream();
            render.ToImage(index, stream);
            yield return new("png", stream);
        }
    }

    static IEnumerable<Sheet> GetSheetData(Workbook book) =>
        book.Worksheets
            .Select(_ => new Sheet(_.Name, GetColumns(_).ToList()));

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

record Sheet(string Name, List<ColumnInfo> Columns);
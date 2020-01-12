using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Verify;

public static partial class VerifyAspose
{
    static ImageOrPrintOptions excelOptions = new ImageOrPrintOptions
    {
        ImageType = ImageType.Png
    };

    static ConversionResult ConvertExcel(Stream stream, VerifySettings settings)
    {
        using var document = new Workbook(stream);
        return ConvertExcel(document, settings);
    }

    static ConversionResult ConvertExcel(Workbook document, VerifySettings settings)
    {
        return new ConversionResult(null, GetExcelStreams(document).ToList());
    }

    static IEnumerable<Stream> GetExcelStreams(Workbook document)
    {
        foreach (var worksheet in document.Worksheets)
        {
            var sheetRender = new SheetRender(worksheet, excelOptions);
            for (var pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
            {
                var outputStream = new MemoryStream();
                sheetRender.ToImage(pageIndex, outputStream);
                yield return outputStream;
            }
        }
    }
}
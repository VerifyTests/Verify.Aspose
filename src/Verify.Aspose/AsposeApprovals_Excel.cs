using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using VerifyXunit;

public static partial class AsposeApprovals
{
    public static async Task VerifyExcel(this VerifyBase verifyBase,string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        using var document = new Workbook(path);
        await VerifyWord(verifyBase, document);
    }

    public static async Task VerifyExcel(this VerifyBase verifyBase,Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        using var document = new Workbook(stream);
        await VerifyWord(verifyBase, document);
    }

    static ImageOrPrintOptions excelOptions = new ImageOrPrintOptions
    {
        ImageType = ImageType.Png
    };

    static Task VerifyWord(this VerifyBase verifyBase, Workbook document)
    {
        return verifyBase.Verify(GetStreams(document));
    }

    static IEnumerable<MemoryStream> GetStreams(Workbook document)
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
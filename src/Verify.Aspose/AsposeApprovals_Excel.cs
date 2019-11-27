using System.IO;
using System.Runtime.ExceptionServices;
using ApprovalTests.Core.Exceptions;
using ApprovalTests.Namers;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using VerifyXunit;

public static partial class AsposeApprovals
{
    public static void VerifyExcel(this VerifyBase verifyBase,string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        using (var document = new Workbook(path))
        {
            await VerifyWord(verifyBase, document);
        }
    }

    public static void VerifyExcel(this VerifyBase verifyBase,Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        using var document = new Workbook(stream);
        await VerifyWord(verifyBase, document);
    }

    static ImageOrPrintOptions excelOptions = new ImageOrPrintOptions
    {
        ImageType = ImageType.Png
    };

    static void VerifyWord(this VerifyBase verifyBase, Workbook document)
    {
        for (var sheetIndex = 0; sheetIndex < document.Worksheets.Count; sheetIndex++)
        {
            var worksheet = document.Worksheets[sheetIndex];
            var sheetRender = new SheetRender(worksheet, excelOptions);
            for (var pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
            {
                var pageNumber = pageIndex + 1;
                var sheetNumber = sheetIndex + 1;

                using var outputStream = new MemoryStream();
                sheetRender.ToImage(pageIndex, outputStream);
                VerifyBinary(outputStream, ref exception);
            }
        }
    }
}
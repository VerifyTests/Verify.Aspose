using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

namespace VerifyTests
{
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
            var info = GetInfo(document);
            return new ConversionResult(info, GetExcelStreams(document).ToList());
        }

        static object GetInfo(Workbook document)
        {
            return new
            {
                HasMacro = document.HasMacro.ToString(),
                HasRevisions = document.HasRevisions.ToString(),
                IsDigitallySigned = document.IsDigitallySigned.ToString(),
                Properties = GetDocumentProperties(document)
            };
        }

        static Dictionary<string, object> GetDocumentProperties(Workbook document)
        {
            return document.BuiltInDocumentProperties
                .Cast<DocumentProperty>()
                .Where(x => x.Value.HasValue())
                .ToDictionary(x => x.Name, x => x.Value);
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
}
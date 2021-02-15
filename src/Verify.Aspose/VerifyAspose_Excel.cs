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
        static ImageOrPrintOptions excelOptions = new()
        {
            ImageType = ImageType.Png
        };

        static ConversionResult ConvertExcel(Stream stream, IReadOnlyDictionary<string, object> settings)
        {
            using Workbook document = new(stream);
            return ConvertExcel(document, settings);
        }

        static ConversionResult ConvertExcel(Workbook document, IReadOnlyDictionary<string, object> settings)
        {
            var info = GetInfo(document);
            return new(info, GetExcelStreams(document).ToList());
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

        static IEnumerable<Target> GetExcelStreams(Workbook document)
        {
            foreach (var worksheet in document.Worksheets)
            {
                SheetRender sheetRender = new(worksheet, excelOptions);
                for (var pageIndex = 0; pageIndex < sheetRender.PageCount; pageIndex++)
                {
                    MemoryStream stream = new();
                    sheetRender.ToImage(pageIndex, stream);
                    yield return new("png", stream);
                }
            }
        }
    }
}
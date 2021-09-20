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

        static object GetInfo(Workbook book)
        {
            return new
            {
                HasMacro = book.HasMacro.ToString(),
                HasRevisions = book.HasRevisions.ToString(),
                IsDigitallySigned = book.IsDigitallySigned.ToString(),
                Properties = GetDocumentProperties(book)
            };
        }

        static Dictionary<string, object> GetDocumentProperties(Workbook book)
        {
            return book.BuiltInDocumentProperties
                .Cast<DocumentProperty>()
                .Where(x => x.Value.HasValue())
                .ToDictionary(x => x.Name, x => x.Value);
        }

        static IEnumerable<Target> GetExcelStreams(Workbook book)
        {
            foreach (var sheet in book.Worksheets)
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
        }
    }
}
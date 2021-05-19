using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace VerifyTests
{
    public static partial class VerifyAspose
    {
        static ConversionResult ConvertWord(Stream stream, IReadOnlyDictionary<string, object> settings)
        {
            Document document = new(stream);
            return ConvertWord(document, settings);
        }

        static ConversionResult ConvertWord(Document document, IReadOnlyDictionary<string, object> settings)
        {
            return new(GetInfo(document), GetWordStreams(document, settings).ToList());
        }

        static object GetInfo(Document document)
        {
            return new
            {
                HasRevisions = document.HasRevisions.ToString(),
                DefaultLocale = (EditingLanguage)document.Styles.DefaultFont.LocaleId,
                Properties = GetDocumentProperties(document)
            };
        }

        static Dictionary<string, object> GetDocumentProperties(Document document)
        {
            return document.BuiltInDocumentProperties
                .Where(x => x.Name != "Bytes" &&
                            x.Value.HasValue())
                .ToDictionary(x => x.Name, x => x.Value);
        }

        static IEnumerable<Target> GetWordStreams(Document document, IReadOnlyDictionary<string, object> settings)
        {
            var pagesToInclude = settings.GetPagesToInclude(document.PageCount);
            for (var pageIndex = 0; pageIndex < pagesToInclude; pageIndex++)
            {
                ImageSaveOptions options = new(SaveFormat.Png)
                {
                    PageSet = new(pageIndex)
                };
                MemoryStream stream = new();
                document.Save(stream, options);
                yield return new("png", stream);
            }
        }
    }
}
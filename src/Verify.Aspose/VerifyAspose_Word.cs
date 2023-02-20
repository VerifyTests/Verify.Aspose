using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using VerifyTestsAspose;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static ConversionResult ConvertWord(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        var document = new Document(stream);
        return ConvertWord(document, settings);
    }

    static ConversionResult ConvertWord(Document document, IReadOnlyDictionary<string, object> settings) => new(GetInfo(document), GetWordStreams(document, settings).ToList());

    static object GetInfo(Document document) =>
        new
        {
            HasRevisions = document.HasRevisions.ToString(),
            DefaultLocale = (EditingLanguage) document.Styles.DefaultFont.LocaleId,
            Properties = GetDocumentProperties(document)
        };

    static Dictionary<string, object> GetDocumentProperties(Document document) =>
        document.BuiltInDocumentProperties
            .Where(_ =>
            {
                if (_.Name == "Bytes")
                {
                    return false;
                }

                if (_.Name == "Version")
                {
                    return false;
                }

                if (_.Name == "TotalEditingTime")
                {
                    return false;
                }

                if (_.Name == "NameOfApplication")
                {
                    return false;
                }

                if (!_.Value.HasValue())
                {
                    return false;
                }

                if (_.Name == "Template" &&
                    (string) _.Value == "Normal.dot")
                {
                    return false;
                }

                if (_.Name == "TitlesOfParts")
                {
                    var strings = (string[]) _.Value;
                    if (strings.Length == 0)
                    {
                        return false;
                    }

                    if (strings.Length == 1)
                    {
                        return strings[0] != "";
                    }
                }

                return true;
            })
            .ToDictionary(_ => _.Name, _ => _.Value);

    static IEnumerable<Target> GetWordStreams(Document document, IReadOnlyDictionary<string, object> settings)
    {
        var pagesToInclude = settings.GetPagesToInclude(document.PageCount);
        for (var pageIndex = 0; pageIndex < pagesToInclude; pageIndex++)
        {
            var saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new(pageIndex)
            };
            var stream = new MemoryStream();
            document.Save(stream, saveOptions);
            yield return new("png", stream);
        }
    }
}
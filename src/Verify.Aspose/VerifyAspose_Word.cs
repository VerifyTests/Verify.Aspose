using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Properties;
using Aspose.Words.Saving;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static ConversionResult ConvertWord(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        //Aspose makes shitty assumptions about streams. like they are writable.
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        memoryStream.Position = 0;
        var document = new Document(memoryStream);
        return ConvertWord(document, settings);
    }

    static ConversionResult ConvertWord(Document document, IReadOnlyDictionary<string, object> settings) =>
        new(GetInfo(document), GetWordStreams(document, settings).ToList());

    static object GetInfo(Document document) =>
        new WordInfo
        {
            HasRevisions = document.HasRevisions.ToString(),
            DefaultLocale = (EditingLanguage) document.Styles.DefaultFont.LocaleId,
            Properties = GetProperties(document),
            CustomProperties = GetCustomProperties(document),
            Text = GetDocumentText(document)
        };

    static Dictionary<string, object> GetProperties(Document document) =>
        document.BuiltInDocumentProperties
            .Where(ShouldIncludeProperty)
            .ToDictionary(_ => _.Name, _ => _.Value);

    static Dictionary<string, object> GetCustomProperties(Document document) =>
        document.CustomDocumentProperties
            .Where(ShouldIncludeProperty)
            .ToDictionary(_ => _.Name, _ => _.Value);

    static bool ShouldIncludeProperty(DocumentProperty property)
    {
        var name = property.Name;

        if (name == "Bytes")
        {
            return false;
        }

        if (name == "Version")
        {
            return false;
        }

        if (name == "TotalEditingTime")
        {
            return false;
        }

        if (name == "NameOfApplication")
        {
            return false;
        }

        if (!property.Value.HasValue())
        {
            return false;
        }

        if (name == "Template" &&
            (string) property.Value == "Normal.dot")
        {
            return false;
        }

        if (name == "TitlesOfParts")
        {
            var strings = (string[]) property.Value;
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
    }

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

    static string GetDocumentText(Document document)
    {
        using var directory = new TempDirectory();
        var path = Path.Combine(directory, "content.md");
        document.Save(
            path,
            new MarkdownSaveOptions());
        return File.ReadAllText(path);
    }
}
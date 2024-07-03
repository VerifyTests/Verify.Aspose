using System.IO.Compression;
using System.Xml.Linq;
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
        new(GetInfo(document), GetWordStreams(document, settings)
            .ToList());

    static object GetInfo(Document document) =>
        new WordInfo
        {
            HasRevisions = document.HasRevisions.ToString(),
            DefaultLocale = (EditingLanguage)document.Styles.DefaultFont.LocaleId,
            Properties = GetProperties(document),
            CustomProperties = GetCustomProperties(document),
            Text = GetDocumentText(document),
        };

    static Dictionary<string, object> GetProperties(Document document) =>
        document
            .BuiltInDocumentProperties
            .Where(ShouldIncludeProperty)
            .ToDictionary(_ => _.Name, _ => _.Value);

    static Dictionary<string, object> GetCustomProperties(Document document) =>
        document
            .CustomDocumentProperties
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
            (string)property.Value == "Normal.dot")
        {
            return false;
        }

        if (name == "TitlesOfParts")
        {
            var strings = (string[])property.Value;
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
        if (settings.GetIncludeWordStyles())
        {
            yield return new("xml", GetStyles(document));
        }

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

    static string GetStyles(Document document)
    {
        using var memoryStream = new MemoryStream();
        document.Save(memoryStream, SaveFormat.Dotx);
        memoryStream.Position = 0;
        using var archive = new ZipArchive(memoryStream);

        var entry = archive.GetEntry("word/styles.xml")!;
        using var entryStream = entry.Open();
        var xmlDocument = XDocument.Load(entryStream);
        CleanupXml(xmlDocument);
        return xmlDocument.ToString();
    }


    static Dictionary<string, string> nodeRenames = new()
    {
        { "b", "bold" },
        { "bCs", "boldComplexScript" },
        { "i", "italic" },
        { "iCs", "italicComplexScript" },
        { "sz", "size" },
        { "szCs", "sizeComplexScript" },
        { "jc", "justification" },
        { "rFonts", "fonts" },
        { "rPr", "run" },
        { "pPr", "paragraph" },
        { "rPrDefault", "runDefault" },
        { "pPrDefault", "paragraphDefault" },
        { "lsdException", "latentStyleException" },
        { "bidi", "complexScriptLanguage" },


    };
    static Dictionary<string, string> attributeRenames = new()
    {
        { "styleId", "Id" },
    };
    static void CleanupXml(XDocument xmlDocument)
    {
        foreach (var node in xmlDocument.Descendants().ToList())
        {
            node.RemoveNamespaceAttributes();
            var name = FixName(node);

            if (name == "name" && node.Parent?.Name == "style")
            {
                node.Remove();
                continue;
            }
            if (name == "lang")
            {
                node.Remove();
                continue;
            }

            if (RemoveComplexSciptIfSameAsNonComplex(node))
            {
                continue;
            }

            var attributes = node.Attributes().ToList();

            foreach (var attribute in attributes)
            {
                if (ShouldRemoveAttribute(attribute))
                {
                    attribute.Remove();
                }
            }
        }
    }

    static string FixName(XElement node)
    {
        var name = node.Name.LocalName;
        if (nodeRenames.TryGetValue(name, out var newName))
        {
            node.Name = newName;
        }
        else
        {
            node.Name = name;
        }

        return name;
    }


    static void RemoveNamespaceAttributes(this XElement node)
    {
        var attributes = node.NamespaceAttributes();

        // Create and add new attributes with local name only for those with a namespace
        foreach (var attribute in attributes
                     .Where(_ => _.Name.Namespace != XNamespace.None))
        {
            node.Add(new XAttribute(attribute.Name.LocalName, attribute.Value));
        }

        // Remove the original attributes
        attributes.Remove();
    }

    static List<XAttribute> NamespaceAttributes(this XElement node) =>
        // Identify xmlns:* attributes and attributes with a namespace
        node
            .Attributes()
            .Where(_ => _.IsNamespaceDeclaration ||
                        _.Name.Namespace != XNamespace.None)
            .ToList();

    static bool RemoveComplexSciptIfSameAsNonComplex(XElement node)
    {
        if (node.Name.LocalName.EndsWith("ComplexScript"))
        {
            var sibling = node.Parent?.Element(node.Name.LocalName[..^13]);
            if (sibling != null && HaveSameAttributes(node, sibling))
            {
                node.Remove();
                return true;
            }
        }

        return false;
    }

    static bool HaveSameAttributes(XElement element1, XElement element2)
    {
        // Check if both elements have the same number of attributes
        if (element1.Attributes().Count() != element2.Attributes().Count())
        {
            return false;
        }

        foreach (var attr1 in element1.Attributes())
        {
            var attr2 = element2.Attribute(attr1.Name);
            // Check if the attribute exists in the second element and has the same value
            if (attr2 == null || attr1.Value != attr2.Value)
            {
                return false;
            }
        }

        return true;
    }

    static bool ShouldRemoveAttribute(XAttribute attribute) =>
        attribute.Name == "unhideWhenUsed" && attribute.Value == "0" ||
        attribute.Name == "semiHidden" && attribute.Value == "0";

}
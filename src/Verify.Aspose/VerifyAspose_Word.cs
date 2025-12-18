using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Properties;
using Aspose.Words.Saving;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static Aspose.Words.Loading.LoadOptions loadOptions = new()
    {
        WarningCallback = new FontWarningCallback()
    };

    class FontWarningCallback : IWarningCallback
    {
        public void Warning(WarningInfo info)
        {
            if (info.WarningType != WarningType.FontSubstitution)
            {
                return;
            }

            throw new(
                $"""
                 Font substitution detected. This can cause inconsitent rendering of documents. Either ensure all dev machines the full set of required conts, or use font embedding.
                 Details: {info.Description}
                 """);
        }
    }

    static ConversionResult ConvertWord(string? name, Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        //Aspose makes shitty assumptions about streams. like they are writable.
        using var memoryStream = new MemoryStream();
        stream.CopyTo(memoryStream);
        memoryStream.Position = 0;
        var document = new Document(memoryStream, loadOptions);
        return ConvertWord(name, document, settings);
    }

    static ConversionResult ConvertWord(string? name, Document document, IReadOnlyDictionary<string, object> settings)
    {
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
        Thread.CurrentThread.CurrentUICulture = CultureInfo.InvariantCulture;
        var info = GetInfo(document);
        List<Target> targets = [BuildDocxTarget(document)];
        targets.AddRange(GetWordStreams(name, document, settings));
        return new(info, targets);
    }

    static Target BuildDocxTarget(Document book)
    {
        using var source = new MemoryStream();
        book.Save(source, SaveFormat.Docx);
        var resultStream = DeterministicPackage.Convert(source);

        return new("docx", resultStream, performConversion: false);
    }

    static WordInfo GetInfo(Document document) =>
        new()
        {
            HasRevisions = document.HasRevisions.ToString(),
            DefaultLocale = (EditingLanguage)document.Styles.DefaultFont.LocaleId,
            Properties = GetProperties(document),
            CustomProperties = GetCustomProperties(document),
            Text = GetDocumentText(document),
            ShadeFormData = document.ShadeFormData
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
    static IEnumerable<Target> GetWordStreams(string? name, Document document, IReadOnlyDictionary<string, object> settings)
    {
        if (settings.GetIncludeWordStyles())
        {
            yield return new("xml", GetStyles(document), name);
        }

        var pagesToInclude = settings.GetPagesToInclude(document.PageCount);
        for (var pageIndex = 0; pageIndex < pagesToInclude; pageIndex++)
        {
            var saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new(pageIndex),
                Resolution = 96,
                Scale = 1.0f,
                UseAntiAliasing = true,
                UseHighQualityRendering = true,
                HorizontalResolution = 96,
                VerticalResolution = 96,
            };
            var stream = new MemoryStream();
            document.Save(stream, saveOptions);
            yield return new("png", stream, name);
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
        { "u", "underline" },
        { "uCs", "underlineComplexScript" },
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
        { "ind", "indent" },
        { "tblPr", "table" },
        { "tblInd", "leadingMarginIndent" },
        { "tblCellMar", "cellMargin" },
        { "outlineLvl", "outlineLevel" },
    };
    static void CleanupXml(XDocument xmlDocument)
    {
        foreach (var node in xmlDocument
                     .Descendants())
        {
            FixName(node);
            CleanupAttributes(node);
        }

        foreach (var node in xmlDocument
                     .Descendants()
                     .ToList())
        {
            var name = node.Name.LocalName;

            if (name == "name" && node.Parent?.Name == "style")
            {
                node.Remove();
                continue;
            }

            if (name is "lang" or "rsid")
            {
                node.Remove();
                continue;
            }

            if (name == "fonts")
            {
                node.Attribute("ascii")?.Remove();
                node.Attribute("eastAsia")?.Remove();
                node.Attribute("hAnsi")?.Remove();
                node.Attribute("asciiTheme")?.Remove();
                node.Attribute("eastAsiaTheme")?.Remove();
                node.Attribute("hAnsiTheme")?.Remove();
            }

            if (RemoveComplexScriptIfSameAsNonComplex(node))
            {
                continue;
            }
        }
        foreach (var node in xmlDocument
                     .Descendants()
                     .ToList())
        {
            RemoveRedundantZeroValAttribute(node);
            HandleValueAttribute(node);
        }
    }


    static bool RemoveComplexScriptIfSameAsNonComplex(XElement node)
    {
        var parent = node.Parent;
        if (parent == null)
        {
            return false;
        }

        var complexScriptNode = parent.Element($"{node.Name.LocalName}ComplexScript");
        if (complexScriptNode == null)
        {
            return false;
        }

        if (HaveSameAttributes(node, complexScriptNode))
        {
            complexScriptNode.Remove();
            return true;
        }

        return false;
    }
    static void HandleValueAttribute(XElement node)
    {
        HandleValueAttribute(node, "val");
        HandleValueAttribute(node, "cs");
    }

    static void HandleValueAttribute(XElement node, string xName)
    {
        var name = node.Name.LocalName;
        var attribute = node.Attribute(xName);
        if (attribute == null)
        {
            return;
        }

        if (!node.HasElements &&
            node.Attributes().Count() == 1)
        {
            var parent = node.Parent;
            if (parent != null)
            {
                parent.Add(new XAttribute(name, attribute.Value));
                node.Remove();
            }
        }
    }

    static void RemoveRedundantZeroValAttribute(XElement node)
    {
        var name = node.Name.LocalName;
        var attribute = node.Attribute("val");
        if (attribute == null)
        {
            return;
        }

        if (name is "bold" or "italic" or "underline")
        {
            if (attribute.Value == "0")
            {
                attribute.Remove();
            }
        }

    }

    static Dictionary<string, string> attributeRenames = new()
    {
        { "styleId", "id" },
        { "customStyle", "custom" },
    };

    static void CleanupAttributes(XElement node)
    {
        var attributes = node.Attributes().ToList();

        foreach (var attribute in attributes)
        {
            if (ShouldRemoveAttribute(attribute))
            {
                attribute.Remove();
                continue;
            }

            var name = attribute.Name.LocalName;
            if (attributeRenames.TryGetValue(name, out var newName))
            {
                name = newName;
            }

            // Create and add new attributes with local name only for those with a namespace
            if (attribute.Name.Namespace != XNamespace.None)
            {
                node.Add(new XAttribute(name, attribute.Value));
                attribute.Remove();
                continue;
            }
        }
    }

    static bool ShouldRemoveAttribute(XAttribute attribute) =>
        attribute.IsNamespaceDeclaration ||
        attribute.Name.LocalName == "aliases" ||
        attribute.Name.LocalName == "unhideWhenUsed" && attribute.Value == "0" ||
        attribute.Name.LocalName == "semiHidden" && attribute.Value == "0";

    static void FixName(XElement node)
    {
        var name = node.Name.LocalName;
        if (nodeRenames.TryGetValue(name, out var newName))
        {
            node.Name = newName;
            return;
        }

        node.Name = name;
    }

    static bool HaveSameAttributes(XElement element1, XElement element2)
    {
        // Check if both elements have the same number of attributes
        if (element1.Attributes().Count() != element2.Attributes().Count())
        {
            return false;
        }

        foreach (var attribute1 in element1.Attributes())
        {
            var attribute2 = element2.Attribute(attribute1.Name);
            // Check if the attribute exists in the second element and has the same value
            if (attribute2 == null ||
                attribute1.Value != attribute2.Value)
            {
                return false;
            }
        }

        return true;
    }

}
using Aspose.Slides;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static ConversionResult ConvertPowerPoint(string? name, Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = new Presentation(stream);
        return ConvertPowerPoint(name, document, settings);
    }

    static ConversionResult ConvertPowerPoint(string? name, Presentation document, IReadOnlyDictionary<string, object> settings)
    {
        var properties = document.DocumentProperties;
        if (properties.NameOfApplication.Contains("Aspose"))
        {
            properties.NameOfApplication = null;
        }

        // Check for font substitutions
        var substitutions = document.FontsManager.GetSubstitutions().ToList();
        if (substitutions.Count > 0)
        {
            var details = string.Join("; ", substitutions.Select(s => $"'{s.OriginalFontName}' -> '{s.SubstitutedFontName}'"));
            throw new(
                $"""
                 Font substitution detected. This can cause inconsitent rendering of documents. Either ensure all dev machines the full set of required conts, or use font embedding.
                 Details: {details}
                 """);
        }

        var (fonts, embeddedFonts) = GetPowerPointFonts(document);
        var info = new
        {
            Properties = properties,
            Fonts = fonts,
            EmbeddedFonts = embeddedFonts
        };

        return new(info, GetPowerPointStreams(name, document, settings).ToList());
    }

    static (List<string> fonts, List<string> embeddedFonts) GetPowerPointFonts(Presentation document)
    {
        var fonts = new HashSet<string>();
        var embeddedFonts = new HashSet<string>();

        foreach (var font in document.FontsManager.GetFonts())
        {
            fonts.Add(font.FontName);
        }

        foreach (var font in document.FontsManager.GetEmbeddedFonts())
        {
            embeddedFonts.Add(font.FontName);
        }

        return (fonts.OrderBy(_ => _).ToList(), embeddedFonts.OrderBy(_ => _).ToList());
    }

    static IEnumerable<Target> GetPowerPointStreams(string? name, Presentation document, IReadOnlyDictionary<string, object> settings)
    {
        var pagesToInclude = settings.GetPagesToInclude(document.Slides.Count);
        for (var index = 0; index < pagesToInclude; index++)
        {
            var slide = document.Slides[index];
            using var bitmap = slide.GetImage(1f, 1f);
            var stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Png);
            yield return new("png", stream, name);
        }
    }
}
using Aspose.Slides;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static ConversionResult ConvertPowerPoint(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = new Presentation(stream);
        return ConvertPowerPoint(document, settings);
    }

    static ConversionResult ConvertPowerPoint(Presentation document, IReadOnlyDictionary<string, object> settings)
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
                 Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
                 Details: {details}
                 """);
        }

        var (fonts, embeddedFonts) = GetPowerPointFonts(document);

        // Names the slides, and says which of them the verification wants and whether it wants
        // them drawn at all, so what is left out is not drawn. A slide is a page.
        var conversion = new PagedConversion(settings)
        {
            Info = new
            {
                Properties = properties,
                Fonts = fonts,
                EmbeddedFonts = embeddedFonts
            }
        };

        var includeImages = conversion.IncludeImages;

        // Aspose.Slides is not safe to render from multiple threads: when two presentations render
        // concurrently a chart can be silently left out of a slide image.
        lock (powerPointRenderLock)
        {
            // Building the deterministic pptx is expensive, so skip it when the pptx target is
            // excluded. A ppt is given back as a pptx, as a doc is as a docx.
            if (!settings.IsTargetExcluded("pptx"))
            {
                conversion.Source(BuildPptxTarget(document));
            }

            var includeText = conversion.IncludeText;
            foreach (var number in conversion.Pages(document.Slides.Count))
            {
                var slide = document.Slides[number - 1];

                Stream? image = null;
                if (includeImages)
                {
                    image = RenderSlide(slide);
                }

                string? text = null;
                if (includeText)
                {
                    text = GetSlideText(slide);
                }

                conversion.AddPage(number, image, text);
            }
        }

        return conversion.Build();
    }

    // The text of a slide: a line for each paragraph that has any, in the order of its shapes.
    static string GetSlideText(ISlide slide)
    {
        var builder = new StringBuilder();
        foreach (var frame in Aspose.Slides.Util.SlideUtil.GetAllTextBoxes(slide))
        {
            foreach (var paragraph in frame.Paragraphs)
            {
                var text = paragraph.Text;
                if (!string.IsNullOrEmpty(text))
                {
                    builder.AppendLine(text);
                }
            }
        }

        return builder.ToString();
    }

    // The pptx snapshot is always the whole presentation, regardless of PagesToInclude, which only
    // limits the slides that are drawn. Mirrors BuildXlsxTarget/BuildDocxTarget.
    static Target BuildPptxTarget(Presentation document)
    {
        using var source = new MemoryStream();
        document.Save(source, Aspose.Slides.Export.SaveFormat.Pptx);
        source.Position = 0;
        var resultStream = DeterministicPackage.Convert(source);

        return new("pptx", resultStream);
    }

    static Lock powerPointRenderLock = new();

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

    static MemoryStream RenderSlide(ISlide slide)
    {
        using var bitmap = slide.GetImage(1f, 1f);
        var stream = new MemoryStream();
        bitmap.Save(stream, ImageFormat.Png);
        return stream;
    }
}
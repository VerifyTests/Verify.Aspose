using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static void OnPdfFontSubstitution(Font oldFont, Font newFont) =>
        throw new(
            $"""
             Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
             Details: '{oldFont.FontName}' -> '{newFont.FontName}'
             """);

    static (List<string> fonts, List<string> embeddedFonts) GetPdfFonts(Document document)
    {
        var fonts = document.FontUtilities.GetAllFonts();
        var all = new List<string>();
        var embedded = new List<string>();

        foreach (var font in fonts)
        {
            all.Add(font.FontName);
            if (font.IsEmbedded)
            {
                embedded.Add(font.FontName);
            }
        }

        return (all.OrderBy(_ => _).ToList(), embedded.OrderBy(_ => _).ToList());
    }

    static void CheckPdfFonts(Document document)
    {
        var fonts = document.FontUtilities.GetAllFonts();
        var missingFonts = new List<string>();

        foreach (var font in fonts)
        {
            // Skip embedded fonts - they don't need substitution
            if (font.IsEmbedded)
            {
                continue;
            }

            // Check if font is available on the system
            try
            {
                var found = FontRepository.FindFont(font.FontName);
                if (found == null)
                {
                    missingFonts.Add(font.FontName);
                }
            }
            catch
            {
                missingFonts.Add(font.FontName);
            }
        }

        if (missingFonts.Count > 0)
        {
            var details = string.Join(", ", missingFonts.Select(f => $"'{f}'"));
            throw new(
                $"""
                 Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
                 Details: Missing fonts: {details}
                 """);
        }
    }

    static ConversionResult ConvertPdf(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = new Document(stream);
        // Subscribe to font substitution events immediately after loading
        document.FontSubstitution += OnPdfFontSubstitution;
        return ConvertPdf(document, settings);
    }

    static ConversionResult ConvertPdf(Document document, IReadOnlyDictionary<string, object> settings)
    {
        // Subscribe to font substitution events (for when Document is passed directly)
        document.FontSubstitution += OnPdfFontSubstitution;

        // Check for fonts that will be substituted
        CheckPdfFonts(document);

        var info = document.Info;
        if (info.Title == "Aspose" ||
            info.Subject == "Aspose" ||
            info.Author == "Aspose")
        {
            throw new("The default value of 'Aspose' for Title, Subject, or Author is not allowed.");
        }

        var (fonts, embeddedFonts) = GetPdfFonts(document);

        // Names the pages, places the text, and says which pages and which outputs the verification
        // wants, so what is left out is neither drawn nor read.
        var conversion = new PagedConversion(settings)
        {
            Info = new
            {
                document.AllowReusePageContent,
                document.CenterWindow,
                document.DisplayDocTitle,
                document.Direction,
                document.Duplex,
                FitWindow = document.FitWindow.ToString(),
                HideMenubar = document.HideMenubar.ToString(),
                HideToolBar = document.HideToolBar.ToString(),
                HideWindowUI = document.HideWindowUI.ToString(),
                IgnoreCorruptedObjects = document.IgnoreCorruptedObjects.ToString(),
                Info = GetInfo(document),
                IsEncrypted = document.IsEncrypted.ToString(),
                IsLinearized = document.IsLinearized.ToString(),
                IsPdfaCompliant = document.IsPdfaCompliant.ToString(),
                IsPdfUaCompliant = document.IsPdfUaCompliant.ToString(),
                IsXrefGapsAllowed = document.IsXrefGapsAllowed.ToString(),
                document.NonFullScreenPageMode,
                OptimizeSize = document.OptimizeSize.ToString(),
                document.PageLabels,
                document.PageLayout,
                document.PageMode,
                document.PdfFormat,
                document.Version,
                Fonts = fonts,
                EmbeddedFonts = embeddedFonts
            },
            // The text is read as markdown
            TextExtension = "md"
        };

        // The text is that of the whole document, so it is given as one text rather than page by
        // page. Read before the pdf is built and the pages are drawn, preserving the original
        // evaluation order: GetDocumentText converts the document to extract its text, which leaves
        // marks in what is saved afterwards.
        if (conversion.IncludeText)
        {
            conversion.Text(GetDocumentText(document));
        }

        // Building the deterministic pdf is expensive, so skip it when the pdf target is excluded.
        if (!settings.IsTargetExcluded("pdf"))
        {
            conversion.Source(BuildPdfTarget(document, settings.NormalizePdf()));
        }

        var includeImages = conversion.IncludeImages;
        // Aspose numbers the pages of a pdf from 1, as Verify does
        foreach (var number in conversion.Pages(document.Pages.Count))
        {
            if (includeImages)
            {
                conversion.AddPage(number, RenderPdfPage(document.Pages[number], settings));
            }
        }

        return conversion.Build();
    }

    // The pdf snapshot is always the full document, regardless of PagesToInclude: PagesToInclude
    // only limits the pages that are drawn. Mirrors BuildXlsxTarget/BuildDocxTarget, which do the
    // same for the OOXML formats via DeterministicPackage.
    static Target BuildPdfTarget(Document document, bool normalize)
    {
        using var source = new MemoryStream();
        document.Save(source);
        source.Position = 0;
        var resultStream = normalize ? PdfNormalizer.Normalize(source) : new(source.ToArray());

        return new("pdf", resultStream);
    }

    static string GetDocumentText(Document document)
    {
        using var stream = new MemoryStream();
        // Aspose.Pdf is not safe to convert to doc from multiple threads: when two documents convert
        // concurrently the save can throw "This is not a structured storage file".
        lock (pdfToDocLock)
        {
            document.Save(stream, new DocSaveOptions());
        }

        stream.Position = 0;
        return GetDocumentText(new Aspose.Words.Document(stream));
    }

    static Lock pdfToDocLock = new();

    static Dictionary<string, string> GetInfo(Document document) =>
        document.Info
            .Where(_ => _.Value.HasValue() &&
                        !_.Key.Contains("Date") &&
                        !_.Value.Contains("Aspose"))
            .ToDictionary(_ => _.Key, _ => _.Value);

    static MemoryStream RenderPdfPage(Page page, IReadOnlyDictionary<string, object> settings)
    {
        var stream = new MemoryStream();
        var pngDevice = settings.GetPdfPngDevice(page);
        pngDevice.Process(page, stream);
        return stream;
    }
}

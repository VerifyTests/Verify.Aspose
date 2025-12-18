using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static void OnPdfFontSubstitution(Font oldFont, Font newFont) =>
        throw new(
            $"""
             Font substitution detected. This can cause inconsitent rendering of documents. Either ensure all dev machines the full set of required conts, or use font embedding.
             Details: '{oldFont.FontName}' -> '{newFont.FontName}'
             """);

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
                 Font substitution detected. This can cause inconsitent rendering of documents. Either ensure all dev machines the full set of required conts, or use font embedding.
                 Details: Missing fonts: {details}
                 """);
        }
    }

    static ConversionResult ConvertPdf(string? name, Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = new Document(stream);
        // Subscribe to font substitution events immediately after loading
        document.FontSubstitution += OnPdfFontSubstitution;
        return ConvertPdf(name, document, settings);
    }

    static ConversionResult ConvertPdf(string? name, Document document, IReadOnlyDictionary<string, object> settings)
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

        return new(
            new
            {
                Pages = document.Pages.Count,
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
                Text = GetDocumentText(document)
            },
            GetPdfStreams(name, document, settings).ToList());
    }

    static string GetDocumentText(Document document)
    {
        using var stream = new MemoryStream();
        document.Save(stream,new DocSaveOptions());
        stream.Position = 0;
        return GetDocumentText(new Aspose.Words.Document(stream));
    }

    static Dictionary<string, string> GetInfo(Document document) =>
        document.Info
            .Where(_ => _.Value.HasValue() &&
                        !_.Key.Contains("Date") &&
                        !_.Value.Contains("Aspose"))
            .ToDictionary(_ => _.Key, _ => _.Value);

    static IEnumerable<Target> GetPdfStreams(string? name, Document document, IReadOnlyDictionary<string, object> settings)
    {
        var pagesToInclude = settings.GetPagesToInclude(document.Pages.Count);
        for (var index = 0; index < pagesToInclude; index++)
        {
            var page = document.Pages[index + 1];
            var stream = new MemoryStream();
            var pngDevice = settings.GetPdfPngDevice(page);
            pngDevice.Process(page, stream);
            yield return new("png", stream, name);
        }
    }
}
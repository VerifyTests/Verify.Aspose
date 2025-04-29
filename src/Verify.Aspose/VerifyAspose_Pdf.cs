using Aspose.Pdf;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static ConversionResult ConvertPdf(string? name, Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = new Document(stream);
        return ConvertPdf(name, document, settings);
    }

    static ConversionResult ConvertPdf(string? name, Document document, IReadOnlyDictionary<string, object> settings)
    {
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
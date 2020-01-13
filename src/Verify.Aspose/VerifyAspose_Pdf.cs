using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using Verify;

public static partial class VerifyAspose
{
    static PngDevice pngDevice = new PngDevice();

    static ConversionResult ConvertPdf(Stream stream, VerifySettings settings)
    {
        using var document = new Document(stream);
        return ConvertPdf(document, settings);
    }

    static ConversionResult ConvertPdf(Document document, VerifySettings settings)
    {
        return new ConversionResult(
            new
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
                document.Version

            },
            GetPdfStreams(document).ToList());
    }

    static Dictionary<string, string> GetInfo(Document document)
    {
        return document.Info
            .Where(x => x.Value.HasValue() &&
                        !x.Key.Contains("Date"))
            .ToDictionary(x => x.Key, x => x.Value);
    }

    static IEnumerable<Stream> GetPdfStreams(Document document)
    {
        foreach (var page in document.Pages)
        {
            var outputStream = new MemoryStream();
            pngDevice.Process(page, outputStream);
            yield return outputStream;
        }
    }
}
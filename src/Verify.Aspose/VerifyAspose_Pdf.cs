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
        return new ConversionResult(null, GetPdfStreams(document).ToList());
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
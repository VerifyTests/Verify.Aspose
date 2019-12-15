using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Pdf;
using Aspose.Pdf.Devices;

public static partial class VerifyAspose
{
    static PngDevice pngDevice = new PngDevice();

    static List<Stream> GetPdfStreams(Stream stream)
    {
        using var document = new Document(stream);
        return GetPdfStreams(document).ToList();
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
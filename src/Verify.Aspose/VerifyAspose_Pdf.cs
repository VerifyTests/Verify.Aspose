using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using VerifyXunit;

public static partial class VerifyAspose
{
    public static async Task VerifyPdf(this VerifyBase verifyBase, string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        using var document = new Document(path);
        await VerifyPdf(verifyBase, document);
    }

    public static async Task VerifyPdf(this VerifyBase verifyBase, Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        using var document = new Document(stream);
        await VerifyPdf(verifyBase, document);
    }

    static PngDevice pngDevice = new PngDevice();

    static Task VerifyPdf(this VerifyBase verifyBase, Document document)
    {
        return verifyBase.VerifyBinary(GetStreams(document), "png");
    }

    static IEnumerable<Stream> GetStreams(Document document)
    {
        foreach (var page in document.Pages)
        {
            var outputStream = new MemoryStream();
            pngDevice.Process(page, outputStream);
            yield return outputStream;
        }
    }
}
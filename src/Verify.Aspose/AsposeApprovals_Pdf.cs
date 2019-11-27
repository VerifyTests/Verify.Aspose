using System.IO;
using System.Runtime.ExceptionServices;
using ApprovalTests.Core.Exceptions;
using ApprovalTests.Namers;
using Aspose.Pdf;
using Aspose.Pdf.Devices;
using VerifyXunit;

public static partial class AsposeApprovals
{
    public static void VerifyPdf(this VerifyBase verifyBase,string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        using var document = new Document(path);
        await VerifyPdf(verifyBase, document);
    }

    public static void VerifyPdf(this VerifyBase verifyBase,Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        using var document = new Document(stream);
        await VerifyPdf(verifyBase, document);
    }

    static PngDevice pngDevice = new PngDevice();

    static void VerifyPdf(this VerifyBase verifyBase,Document document)
    {
        foreach (var page in document.Pages)
        {
            using var outputStream = new MemoryStream();
            pngDevice.Process(page, outputStream);
            VerifyBinary(outputStream, ref exception);
        }
    }
}
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Slides;
using VerifyXunit;

public static partial class VerifyAspose
{
    public static async Task VerifyPowerPoint(this VerifyBase verifyBase, string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        using var document = new Presentation(path);
        await VerifyPowerPoint(verifyBase, document);
    }

    public static async Task VerifyPowerPoint(this VerifyBase verifyBase, Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        using var document = new Presentation(stream);
        await VerifyPowerPoint(verifyBase, document);
    }

    static Task VerifyPowerPoint(this VerifyBase verifyBase, Presentation document)
    {
        return verifyBase.Verify(GetStreams(document), "svg");
    }

    static IEnumerable<Stream> GetStreams(Presentation document)
    {
        foreach (var slide in document.Slides)
        {
            var outputStream = new MemoryStream();
            slide.WriteAsSvg(outputStream);
            yield return outputStream;
        }
    }
}
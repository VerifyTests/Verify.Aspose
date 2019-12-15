using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Slides;

public static partial class VerifyAspose
{
    static List<Stream> GetPowerPointStreams(Stream stream)
    {
        using var document = new Presentation(stream);
        return GetPowerPointStreams(document).ToList();
    }

    static IEnumerable<Stream> GetPowerPointStreams(Presentation document)
    {
        foreach (var slide in document.Slides)
        {
            var outputStream = new MemoryStream();
            slide.WriteAsSvg(outputStream);
            yield return outputStream;
        }
    }
}
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Slides;
using Verify;

public static partial class VerifyAspose
{
    static ConversionResult ConvertPowerPoint(Stream stream, VerifySettings settings)
    {
        using var document = new Presentation(stream);
        return ConvertPowerPoint(document, settings);
    }

    static ConversionResult ConvertPowerPoint(Presentation document, VerifySettings settings)
    {
        return new ConversionResult(document.DocumentProperties, GetPowerPointStreams(document).ToList());
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
using System.Drawing.Imaging;
using Aspose.Slides;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static ConversionResult ConvertPowerPoint(Stream stream, IReadOnlyDictionary<string, object> settings)
    {
        using var document = new Presentation(stream);
        return ConvertPowerPoint(document, settings);
    }

    static ConversionResult ConvertPowerPoint(Presentation document, IReadOnlyDictionary<string, object> settings) =>
        new(document.DocumentProperties, GetPowerPointStreams(document, settings).ToList());

    static IEnumerable<Target> GetPowerPointStreams(Presentation document, IReadOnlyDictionary<string, object> settings)
    {
        var pagesToInclude = settings.GetPagesToInclude(document.Slides.Count);
        for (var index = 0; index < pagesToInclude; index++)
        {
            var slide = document.Slides[index];
            using var bitmap = slide.GetThumbnail(1f, 1f);
            var stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Png);
            yield return new("png", stream, null);
        }
    }
}
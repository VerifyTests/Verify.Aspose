using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Slides;

namespace VerifyTests
{
    public static partial class VerifyAspose
    {
        static ConversionResult ConvertPowerPoint(Stream stream, VerifySettings settings)
        {
            using var document = new Presentation(stream);
            return ConvertPowerPoint(document, settings);
        }

        static ConversionResult ConvertPowerPoint(Presentation document, VerifySettings settings)
        {
            return new ConversionResult(document.DocumentProperties, GetPowerPointStreams(document, settings).ToList());
        }

        static IEnumerable<ConversionStream> GetPowerPointStreams(Presentation document, VerifySettings settings)
        {
            var pagesToInclude = settings.GetPagesToInclude(document.Slides.Count);
            for (var index = 0; index < pagesToInclude; index++)
            {
                var slide = document.Slides[index];
                var stream = new MemoryStream();
                slide.WriteAsSvg(stream);
                yield return new ConversionStream("svg", stream);
            }
        }
    }
}
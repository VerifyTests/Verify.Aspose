using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Verify;

public static partial class VerifyAspose
{
    static ConversionResult ConvertWord(Stream stream, VerifySettings settings)
    {
        var document = new Document(stream);
        return ConvertWord(document, settings);
    }

    static ConversionResult ConvertWord(Document document, VerifySettings settings)
    {
        return new ConversionResult(null, GetWordStreams(document).ToList());
    }

    static IEnumerable<Stream> GetWordStreams(Document document)
    {
        for (var pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
        {
            var options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageIndex = pageIndex
            };
            var outputStream = new MemoryStream();
            document.Save(outputStream, options);
            yield return outputStream;
        }
    }
}
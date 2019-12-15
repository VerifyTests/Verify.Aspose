using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public static partial class VerifyAspose
{
    static List<Stream> GetWordStreams(Stream stream)
    {
        var document = new Document(stream);
        return GetWordStreams(document).ToList();
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
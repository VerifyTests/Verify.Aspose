using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;
using VerifyXunit;

public static partial class VerifyAspose
{
    public static async Task VerifyWord(this VerifyBase verifyBase, string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        var document = new Document(path);
        await VerifyWord(verifyBase,document);
    }

    public static async Task VerifyWord(this VerifyBase verifyBase, Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        var document = new Document(stream);
        await VerifyWord(verifyBase,document);
    }

    static Task VerifyWord(this VerifyBase verifyBase, Document document)
    {
        return verifyBase.Verify(GetStreams(document), "png");
    }

    static IEnumerable<Stream> GetStreams(Document document)
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
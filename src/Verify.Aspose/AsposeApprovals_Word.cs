using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using VerifyXunit;

public static partial class AsposeApprovals
{
    public static void VerifyWord(this VerifyBase verifyBase, string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        var document = new Document(path);
        await VerifyWord(document);
    }

    public static void VerifyWord(this VerifyBase verifyBase, Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        var document = new Document(stream);
        await VerifyWord(document);
    }

    static void VerifyWord(this VerifyBase verifyBase, Document document)
    {
        for (var pageIndex = 0; pageIndex < document.PageCount; pageIndex++)
        {
            var options = new ImageSaveOptions(SaveFormat.Png)
            {
                PageIndex = pageIndex
            };
            using var outputStream = new MemoryStream();
            document.Save(outputStream, options);
            VerifyBinary(outputStream, ref exception);
        }
    }
}
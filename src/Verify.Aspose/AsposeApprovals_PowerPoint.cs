using System.IO;
using ApprovalTests.Core.Exceptions;
using ApprovalTests.Namers;
using Aspose.Slides;
using VerifyXunit;

public static partial class AsposeApprovals
{
    public static void VerifyPowerPoint(this VerifyBase verifyBase,string path)
    {
        Guard.AgainstNullOrEmpty(path, nameof(path));
        using var document = new Presentation(path);
        await VerifyPowerPoint(verifyBase,document);
    }

    public static void VerifyPowerPoint(this VerifyBase verifyBase,Stream stream)
    {
        Guard.AgainstNull(stream, nameof(stream));
        using var document = new Presentation(stream);
        await VerifyPowerPoint(verifyBase,document);
    }

    static void VerifyPowerPoint(this VerifyBase verifyBase,Presentation document)
    {
        for (var pageIndex = 0; pageIndex < document.Slides.Count; pageIndex++)
        {
            var slide = document.Slides[pageIndex];
            using var outputStream = new MemoryStream();
            slide.WriteAsSvg(outputStream);
            VerifyBinary(outputStream, ref exception, ".svg");
        }
    }
}
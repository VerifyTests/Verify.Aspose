using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Pdf.Devices;
using VerifyTests;
using VerifyNUnit;
using NUnit.Framework;

#region TestDefinition
[TestFixture]
public class Samples
{
    static Samples()
    {
        VerifyAspose.Initialize();
    }
    #endregion

    #region VerifyPdf

    [Test]
    public Task VerifyPdf()
    {
        return Verifier.VerifyFile("sample.pdf");
    }

    #endregion

    [Test]
    public Task VerifyPdfResolution()
    {
        return Verifier.VerifyFile("sample.pdf")
            .PdfPngDevice(page =>
            {
                Resolution resolution = new(100);
                var artBox = page.ArtBox;
                var width = Convert.ToInt32(artBox.Width);
                var height = Convert.ToInt32(artBox.Height);
                return new(width, height, resolution);
            });
    }

    #region VerifyPdfStream

    [Test]
    public Task VerifyPdfStream()
    {
        return Verifier.Verify(File.OpenRead("sample.pdf"))
            .UseExtension("pdf");
    }

    #endregion

#if DEBUG

    #region VerifyPowerPoint

    [Test]
    public Task VerifyPowerPoint()
    {
        return Verifier.VerifyFile("sample.pptx");
    }

    #endregion

    #region VerifyPowerPointStream

    [Test]
    public Task VerifyPowerPointStream()
    {
        return Verifier.Verify(File.OpenRead("sample.pptx"))
            .UseExtension("pptx");
    }

    #endregion

#endif

    #region VerifyExcel

    [Test]
    public Task VerifyExcel()
    {
        return Verifier.VerifyFile("sample.xlsx");
    }

    #endregion

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        return Verifier.Verify(File.OpenRead("sample.xlsx"))
            .UseExtension("xlsx");
    }

    #endregion

    #region VerifyWord

    [Test]
    public Task VerifyWord()
    {
        return Verifier.VerifyFile("sample.docx");
    }

    #endregion

    #region VerifyWordStream

    [Test]
    public Task VerifyWordStream()
    {
        return Verifier.Verify(File.OpenRead("sample.docx"))
            .UseExtension("docx");
    }

    #endregion
}
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
        var resolution = new Resolution(100);
        var settings = new VerifySettings();
        settings.PdfPngDevice(page =>
        {
            var artBox = page.ArtBox;
            var width = Convert.ToInt32(artBox.Width);
            var height = Convert.ToInt32(artBox.Height);
            return new PngDevice(width, height, resolution);
        });
        return Verifier.VerifyFile("sample.pdf", settings);
    }

    #region VerifyPdfStream

    [Test]
    public Task VerifyPdfStream()
    {
        var settings = new VerifySettings();
        settings.UseExtension("pdf");
        return Verifier.Verify(File.OpenRead("sample.pdf"), settings);
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
        var settings = new VerifySettings();
        settings.UseExtension("pptx");
        return Verifier.Verify(File.OpenRead("sample.pptx"), settings);
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
        var settings = new VerifySettings();
        settings.UseExtension("xlsx");
        return Verifier.Verify(File.OpenRead("sample.xlsx"), settings);
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
        var settings = new VerifySettings();
        settings.UseExtension("docx");
        return Verifier.Verify(File.OpenRead("sample.docx"), settings);
    }

    #endregion
}
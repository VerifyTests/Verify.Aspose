using System.IO;
using System.Threading.Tasks;
using Verify;
using VerifyXunit;
using Xunit;
using Xunit.Abstractions;

public class Samples :
    VerifyBase
{
    #region VerifyPdf

    [Fact]
    public Task VerifyPdf()
    {
        return VerifyFile("sample.pdf");
    }

    #endregion

    #region VerifyPdfStream

    [Fact]
    public Task VerifyPdfStream()
    {
        var settings = new VerifySettings();
        settings.UseExtension("pdf");
        return Verify(File.OpenRead("sample.pdf"), settings);
    }

    #endregion

#if DEBUG

    #region VerifyPowerPoint

    [Fact]
    public Task VerifyPowerPoint()
    {
        return VerifyFile("sample.pptx");
    }

    #endregion

    #region VerifyPowerPointStream

    [Fact]
    public Task VerifyPowerPointStream()
    {
        var settings = new VerifySettings();
        settings.UseExtension("pptx");
        return Verify(File.OpenRead("sample.pptx"), settings);
    }

    #endregion

#endif

    #region VerifyExcel

    [Fact]
    public Task VerifyExcel()
    {
        return Verify("sample.xlsx");
    }

    #endregion

    #region VerifyExcelStream

    [Fact]
    public Task VerifyExcelStream()
    {
        var settings = new VerifySettings();
        settings.UseExtension("xlsx");
        return Verify(File.OpenRead("sample.xlsx"), settings);
    }

    #endregion

    #region VerifyWord

    [Fact]
    public Task VerifyWord()
    {
        return Verify("sample.docx");
    }

    #endregion

    #region VerifyWordStream

    [Fact]
    public Task VerifyWordStream()
    {
        var settings = new VerifySettings();
        settings.UseExtension("docx");
        return Verify(File.OpenRead("sample.docx"), settings);
    }

    #endregion

    public Samples(ITestOutputHelper output) :
        base(output)
    {
    }

    static Samples()
    {
        VerifyAspose.Initialize();
    }
}
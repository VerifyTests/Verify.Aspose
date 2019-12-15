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
    public async Task VerifyPdf()
    {
        await using var stream = File.OpenRead("sample.pdf");
        var settings = new VerifySettings();
        settings.UseExtension("pdf");
        await Verify(stream, settings);
    }

    #endregion

#if DEBUG

    #region VerifyPowerPoint

    [Fact]
    public async Task VerifyPowerPoint()
    {
        await using var stream = File.OpenRead("sample.pptx");
        var settings = new VerifySettings();
        settings.UseExtension("pptx");
        await Verify(stream, settings);
    }

    #endregion

#endif

    #region VerifyExcel

    [Fact]
    public async Task VerifyExcel()
    {
        await using var stream = File.OpenRead("sample.xlsx");
        var settings = new VerifySettings();
        settings.UseExtension("xlsx");
        await Verify(stream, settings);
    }

    #endregion

    #region VerifyWord

    [Fact]
    public async Task VerifyWord()
    {
        await using var stream = File.OpenRead("sample.docx");
        var settings = new VerifySettings();
        settings.UseExtension("docx");
        await Verify(stream, settings);
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
using System.Threading.Tasks;
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
        await this.VerifyPdf("sample.pdf");
    }
    #endregion

    #region VerifyPowerPoint
    [Fact]
    public async Task VerifyPowerPoint()
    {
        await this.VerifyPowerPoint("sample.pptx");
    }
    #endregion

    #region VerifyExcel
    [Fact]
    public async Task VerifyExcel()
    {
        await this.VerifyExcel("sample.xlsx");
    }
    #endregion

    #region VerifyWord
    [Fact]
    public async Task VerifyWord()
    {
        await this.VerifyWord("sample.docx");
    }
    #endregion

    public Samples(ITestOutputHelper output) :
        base(output)
    {
    }
}
using VerifyXunit;
using Xunit;
using Xunit.Abstractions;

public class Samples :
    VerifyBase
{
    [Fact]
    public void VerifyPdf()
    {
        #region VerifyPdf

        AsposeApprovals.VerifyPdf("sample.pdf");

        #endregion
    }

    [Fact]
    public void VerifyPowerPoint()
    {
        #region VerifyPowerPoint

        AsposeApprovals.VerifyPowerPoint("sample.pptx");

        #endregion
    }

    [Fact]
    [Trait("Category", "Integration")]
    public void VerifyExcel()
    {
        #region VerifyExcel

        AsposeApprovals.VerifyExcel("sample.xlsx");

        #endregion
    }

    [Fact]
    public void VerifyWord()
    {
        #region VerifyWord

        AsposeApprovals.VerifyWord("sample.docx");

        #endregion
    }

    public Samples(ITestOutputHelper output) :
        base(output)
    {
    }
}
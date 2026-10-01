// Initialize is called with AsposeOutputs.Text, so no png or csv targets are produced
public class OutputsTests
{
    [Test]
    public Task Pdf() =>
        VerifyFile("sample.pdf");

    [Test]
    public Task Word() =>
        VerifyFile("sample.docx");

    [Test]
    public Task Excel() =>
        VerifyFile("sample.xlsx");
}

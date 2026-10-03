// Initialize is called with AsposeOutputs.Text, so no png or csv targets are produced
public class OutputsTests
{
    [Test]
    public Task Pdf() =>
        VerifyFile(ProjectFiles.sample_pdf.Path);

    [Test]
    public Task Word() =>
        VerifyFile(ProjectFiles.sample_docx.Path);

    [Test]
    public Task Excel() =>
        VerifyFile(ProjectFiles.sample_xlsx.Path);
}

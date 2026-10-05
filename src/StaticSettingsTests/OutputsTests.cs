// The module initializer excludes the png and csv derived targets, so no page is drawn and no
// sheet is exported
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

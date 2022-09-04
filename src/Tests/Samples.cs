using Aspose.Cells;
using Aspose.Pdf.Devices;
using VerifyTestsAspose;

[TestFixture]
public class Samples
{
    #region VerifyPdf

    [Test]
    public Task VerifyPdf() =>
        VerifyFile("sample.pdf");

    #endregion

    [Test]
    public Task VerifyPdfResolution() =>
        VerifyFile("sample.pdf")
            .PdfPngDevice(page =>
            {
                var resolution = new Resolution(100);
                var artBox = page.ArtBox;
                var width = Convert.ToInt32(artBox.Width);
                var height = Convert.ToInt32(artBox.Height);
                return new(width, height, resolution);
            });

    #region VerifyPdfStream

    [Test]
    public Task VerifyPdfStream() =>
        Verify(File.OpenRead("sample.pdf"))
            .UseExtension("pdf");

    #endregion

#if DEBUG

    #region VerifyPowerPoint

    [Test]
    public Task VerifyPowerPoint() =>
        VerifyFile("sample.pptx");

    #endregion

    #region VerifyPowerPointStream

    [Test]
    public Task VerifyPowerPointStream() =>
        Verify(File.OpenRead("sample.pptx"))
            .UseExtension("pptx");

    #endregion

#endif

    #region VerifyExcel

    [Test]
    public Task VerifyExcel() =>
        VerifyFile("sample.xlsx");

    #endregion

    #region VerifySheet

    [Test]
    public Task VerifySheet()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets.Add("New Sheet");

        var cells = sheet.Cells;

        cells[0, 0].PutValue("Some Text");

        return Verify(sheet);
    }

    #endregion

    #region VerifyWorkbook

    [Test]
    public Task VerifyWorkbook()
    {
        var book = new Workbook();

        var sheet = book.Worksheets.Add("New Sheet");

        var cells = sheet.Cells;

        cells[0, 0].PutValue("Some Text");
        return Verify(book);
    }

    #endregion

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream() =>
        Verify(File.OpenRead("sample.xlsx"))
            .UseExtension("xlsx");

    #endregion

    #region VerifyWord

    [Test]
    public Task VerifyWord() =>
        VerifyFile("sample.docx");

    #endregion

    #region VerifyWordStream

    [Test]
    public Task VerifyWordStream() =>
        Verify(File.OpenRead("sample.docx"))
            .UseExtension("docx");

    #endregion
}
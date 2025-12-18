
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
    public Task VerifyPdfStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.pdf"));
        return Verify(stream, "pdf");
    }

    #endregion

#if DEBUG

    #region VerifyPowerPoint

    [Test]
    public Task VerifyPowerPoint() =>
        VerifyFile("sample.pptx");

    #endregion

    #region VerifyPowerPointStream

    [Test]
    public Task VerifyPowerPointStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.pptx"));
        return Verify(stream, "pptx");
    }

    #endregion

    [Test]
    public Task VerifyPowerPointDoc()
    {
        var presentation = new Presentation();
        presentation.DocumentProperties["Key"] = "Value";
        return Verify(presentation);
    }

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
        sheet.CustomProperties.Add("key", "value");
        var cells = sheet.Cells;

        cells[0, 0].PutValue("Some Text");

        return Verify(sheet);
    }

    #endregion

    [Test]
    public Task VerifySheetWithHyperlinks()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets.Add("New Sheet");
        sheet.CustomProperties.Add("key", "value");
        var cells = sheet.Cells;

        cells[0, 0].PutValue("Some Text");
        sheet.Hyperlinks.Add(0, 0, 1, 1, "theUrl");

        return Verify(sheet);
    }

    #region VerifyWorkbook

    [Test]
    public Task VerifyWorkbook()
    {
        var book = new Workbook
        {
            BuiltInDocumentProperties =
            {
                Comments = "the comments"
            }
        };
        book.CustomDocumentProperties.Add("key", "value");

        var sheet = book.Worksheets.Add("New Sheet");

        var cells = sheet.Cells;

        cells[0, 0].PutValue("Some Text");
        return Verify(book);
    }

    #endregion

    [Test]
    public async Task Cell()
    {
        using var workbook = new Workbook();

        var sheet = workbook.Worksheets[0];
        var cell = sheet.Cells["A1"];
        cell.PutValue("Hello World!");
        await Verify(cell);
    }

    #region VerifyExcelStream

    [Test]
    public Task VerifyExcelStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
        return Verify(stream, "xlsx");
    }

    #endregion

    [Test]
    public void FontSubstitution()
    {
        var exception = Assert.ThrowsAsync<Exception>(() => VerifyFile("fontSubstitution.docx"))!;
        AreEqual(
            """
            Font substitution detected. This can cause inconsitent rendering of documents. Either ensure all dev machines the full set of required conts, or use font embedding.
            Details: Font 'Droid Sans Fallback' has not been found. Using 'Times New Roman' font instead. Reason: default font substitution.
            """,
            exception.Message);
    }

    #region VerifyWord

    [Test]
    public Task VerifyWord() =>
        VerifyFile("sample.docx");

    #endregion

    #region VerifyWordStream

    [Test]
    public Task VerifyWordStream()
    {
        var stream = new MemoryStream(File.ReadAllBytes("sample.docx"));
        return Verify(stream, "docx");
    }

    #endregion

    [Test]
    public Task VerifyWordStyles() =>
        VerifyFile("sample.docx").IncludeWordStyles();

    [Test]
    public Task VerifyWordDocument()
    {
        var document = new Document
        {
            BuiltInDocumentProperties =
            {
                Comments = "the comments"
            }
        };
        document.CustomDocumentProperties.Add("key", "value");
        return Verify(document);
    }

    [Test]
    public Task ShadeFormData()
    {
        var document = new Document
        {
            ShadeFormData = false
        };
        document.CustomDocumentProperties.Add("key", "value");
        return Verify(document);
    }

    [Test]
    public Task AsposeGenerator() =>
        VerifyFile("sample.WithAsposeGenerator.html");
}
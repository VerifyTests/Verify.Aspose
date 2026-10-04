
public class Samples
{
    #region VerifyPdf

    [Test]
    public Task VerifyPdf() =>
        VerifyFile("sample.pdf");

    #endregion

    [Test]
    public Task VerifyPdfResolution() =>
        VerifyFile(ProjectFiles.sample_pdf.Path)
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

    #region ExcludeXlsx

    [Test]
    public Task ExcludeXlsx() =>
        // ExcludeTargets skips the expensive deterministic xlsx build.
        VerifyFile("sample.xlsx")
            .ExcludeTargets("xlsx");

    #endregion

    [Test]
    public Task HiddenRow() =>
        VerifyFile(ProjectFiles.sample_hidden_row_xlsx.Path);

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

    [Test]
    public Task VerifySheetWithHyperlinkNoText()
    {
        using var book = new Workbook();

        var sheet = book.Worksheets.Add("New Sheet");
        // Hyperlink without cell value causes Aspose Hyperlink.TextToDisplay to throw
        sheet.Hyperlinks.Add(0, 0, 1, 1, "https://example.com");

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
    public async Task FontSubstitutionWord()
    {
        var exception = await Assert.That(async () => await VerifyFile(ProjectFiles.fontSubstitution_docx.Path)).ThrowsExactly<Exception>();
        await Assert.That(exception!.Message).IsEqualTo(
            """
            Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
            Details: Font 'Liberation Serif;Times New Roma' has not been found. Using 'Times New Roman' font instead. Reason: default font substitution.
            """);
    }

    [Test]
    public async Task FontSubstitutionPdf()
    {
        var exception = await Assert.That(async () => await VerifyFile("fontSubstitution.pdf")).ThrowsExactly<Exception>();
        await Assert.That(exception!.Message).StartsWith(
            """
            Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
            Details:
            """);
    }

    [Test]
    public async Task FontSubstitutionExcel()
    {
        var exception = await Assert.That(async () => await VerifyFile(ProjectFiles.fontSubstitution_xlsx.Path)).ThrowsExactly<Exception>();
        await Assert.That(exception!.Message).StartsWith(
            """
            Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
            Details:
            """);
    }

#if DEBUG
    [Test]
    public async Task FontSubstitutionPowerPoint()
    {
        var exception = await Assert.That(async () => await VerifyFile(ProjectFiles.fontSubstitution_pptx.Path)).ThrowsExactly<Exception>();
        await Assert.That(exception!.Message).StartsWith(
            """
            Font substitution detected. This can cause inconsistent rendering of documents. Either ensure all dev machines the full set of required fonts, or use font embedding.
            Details:
            """);
    }
#endif

    #region VerifyWord

    [Test]
    public Task VerifyWord() =>
        VerifyFile("sample.docx");

    #endregion

    #region ExcludeDocx

    [Test]
    public Task ExcludeDocx() =>
        // ExcludeTargets skips the expensive deterministic docx build.
        VerifyFile("sample.docx")
            .ExcludeTargets("docx");

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
        VerifyFile(ProjectFiles.sample_docx.Path).IncludeWordStyles();

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
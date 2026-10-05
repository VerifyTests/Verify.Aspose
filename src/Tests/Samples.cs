
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

    // A sheet is a page, so PagesToInclude limits the sheets that are drawn, exported and described.
    // PageCount is still the number of sheets in the workbook.
    [Test]
    public Task PagesToIncludeSheets()
    {
        var book = new Workbook();
        book.Worksheets[0].Cells[0, 0].PutValue("First sheet");
        book.Worksheets.Add("Second").Cells[0, 0].PutValue("Second sheet");

        return Verify(book)
            .PagesToInclude(1)
            .ExcludeTargets("xlsx");
    }

    // A hidden sheet is verified as any other: it is drawn, exported and described. What says it is
    // hidden is HiddenSheets in the info file.
    [Test]
    public Task HiddenSheet()
    {
        var book = new Workbook();
        book.Worksheets[0].Cells[0, 0].PutValue("First sheet");
        var hidden = book.Worksheets.Add("Second");
        hidden.Cells[0, 0].PutValue("Hidden sheet");
        hidden.VisibilityType = VisibilityType.Hidden;

        return Verify(book)
            .ExcludeTargets("xlsx");
    }

    // A hidden sheet is counted as any other, so here it is the first page: it is the one that is
    // drawn and exported, and the sheet that is shown is left out.
    [Test]
    public Task HiddenSheetIsCountedByPagesToInclude()
    {
        var book = new Workbook();
        book.Worksheets[0].Cells[0, 0].PutValue("Hidden sheet");
        book.Worksheets.Add("Second").Cells[0, 0].PutValue("Second sheet");
        // Hidden once there is another sheet to show, since a workbook has to show one
        book.Worksheets[0].VisibilityType = VisibilityType.Hidden;

        return Verify(book)
            .PagesToInclude(1)
            .ExcludeTargets("xlsx");
    }

    // A sheet that only code can unhide is verified as one Excel can
    [Test]
    public Task VeryHiddenSheet()
    {
        var book = new Workbook();
        book.Worksheets[0].Cells[0, 0].PutValue("First sheet");
        var hidden = book.Worksheets.Add("Second");
        hidden.Cells[0, 0].PutValue("Very hidden sheet");
        hidden.VisibilityType = VisibilityType.VeryHidden;

        return Verify(book)
            .ExcludeTargets("xlsx");
    }

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

    // A doc is given back as a docx. That docx is the source of the conversion, so the docx
    // converter does not convert it again: there is one info, not two.
    [Test]
    public Task VerifyDocStream()
    {
        var document = new Document();
        var stream = new MemoryStream();
        document.Save(stream, Aspose.Words.SaveFormat.Doc);
        return Verify(stream, "doc")
            .ExcludeDerivedTargets("png");
    }

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

    #region PageTextPerPage

    [Test]
    public Task PageTextPerPage() =>
        VerifyFile("sample.docx")
            .PageText(PageTextPlacement.PerPage)
            .ExcludeDerivedTargets("png");

    #endregion

    #region TextOnly

    [Test]
    public Task TextOnly() =>
        VerifyFile("sample.pdf")
            .ExcludeDerivedTargets("png");

    #endregion

    // Only what Aspose says of the pdf is left: no text is read, no page is drawn, and the pdf is
    // not built
    [Test]
    public Task NoText() =>
        VerifyFile(ProjectFiles.sample_pdf.Path)
            .PageText(PageTextPlacement.None)
            .ExcludeDerivedTargets("png")
            .ExcludeTargets("pdf");

    [Test]
    public Task AsposeGenerator() =>
        VerifyFile("sample.WithAsposeGenerator.html");
}
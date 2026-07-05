[TestFixture]
public class WordPageCountTests
{
    // WordInfo.PageCount is the full laid-out page count of the source document, independent of
    // PagesToInclude which only trims the rendered page images. A three page document included as
    // a single page still reports PageCount: 3 and emits the full docx, alongside a single png.
    [Test]
    public Task PageCountIsIndependentOfPagesToInclude()
    {
        var document = BuildThreePageDocument();
        return Verify(document)
            .PagesToInclude(1);
    }

    // Without a filter every page is rendered, so PageCount matches the number of png pages.
    [Test]
    public Task PageCountMatchesAllRenderedPages()
    {
        var document = BuildThreePageDocument();
        return Verify(document);
    }

    static Document BuildThreePageDocument()
    {
        var document = new Document();
        var builder = new DocumentBuilder(document);
        builder.Writeln("Page one");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page two");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page three");
        return document;
    }
}

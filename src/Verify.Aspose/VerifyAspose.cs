using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static VerifyAspose() =>
        RenderEmptySheet();

    public static bool Initialized { get; private set; }

    public static void Initialize()
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;

        InnerVerifier.ThrowIfVerifyHasBeenRun();

        var cellConverter = new CellConverter();
        var hyperlinkConverter = new HyperlinkConverter();
        var cellAreaConverter = new CellAreaConverter();

        VerifierSettings.AddExtraSettings(_ =>
        {
            _.Converters.Add(cellConverter);
            _.Converters.Add(hyperlinkConverter);
            _.Converters.Add(cellAreaConverter);
        });

        VerifierSettings.AddScrubber("html", RemoveGeneratorInfo);

        // The name a stream converter is passed is not used. Verify names what a converter returns
        // relative to the target that was converted.
        VerifierSettings.RegisterStreamConverter("xlsx", (_, stream, context) => ConvertExcel(stream, context));
        VerifierSettings.RegisterStreamConverter("xls", (_, stream, context) => ConvertExcel(stream, context));
        VerifierSettings.IgnoreMember<IDocumentProperties>(_ => _.AppVersion);
        VerifierSettings.RegisterFileConverter<Workbook>(ConvertExcel);
        VerifierSettings.RegisterFileConverter<Worksheet>(ConvertSheet);

        VerifierSettings.RegisterStreamConverter("pdf", (_, stream, context) => ConvertPdf(stream, context));
        VerifierSettings.RegisterFileConverter<Aspose.Pdf.Document>(ConvertPdf);

        VerifierSettings.RegisterStreamConverter("pptx", (_, stream, context) => ConvertPowerPoint(stream, context));
        VerifierSettings.RegisterStreamConverter("ppt", (_, stream, context) => ConvertPowerPoint(stream, context));
        VerifierSettings.RegisterFileConverter<Presentation>(ConvertPowerPoint);

        VerifierSettings.RegisterStreamConverter("docx", (_, stream, context) => ConvertWord(stream, context));
        VerifierSettings.RegisterStreamConverter("doc", (_, stream, context) => ConvertWord(stream, context));
        VerifierSettings.RegisterFileConverter<Document>(ConvertWord);
    }

    static void RemoveGeneratorInfo(StringBuilder builder)
    {
        var input = builder.ToString();
        const string startPattern = "<meta name=\"generator\" content=\"Aspose";

        var startIndex = input.IndexOf(startPattern, StringComparison.OrdinalIgnoreCase);
        if (startIndex != -1)
        {
            var endIndex = input.IndexOf('>', startIndex);
            if (endIndex != -1)
            {
                builder.Remove(startIndex, endIndex - startIndex + 1);
            }
        }
    }
}

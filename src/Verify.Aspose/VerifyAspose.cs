using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;

namespace VerifyTests;

public static partial class VerifyAspose
{
    static VerifyAspose() =>
        RenderEmptySheet();

    public static bool Initialized { get; private set; }

    static AsposeOutputs outputs = AsposeOutputs.All;

    /// <param name="outputs">The outputs that documents are split into. Defaults to <see cref="AsposeOutputs.All"/>.</param>
    public static void Initialize(AsposeOutputs outputs = AsposeOutputs.All)
    {
        if (Initialized)
        {
            throw new("Already Initialized");
        }

        Initialized = true;
        VerifyAspose.outputs = outputs;

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
        VerifierSettings.RegisterStreamConverter("xlsx", ConvertExcel);
        VerifierSettings.RegisterStreamConverter("xls", ConvertExcel);
        VerifierSettings.IgnoreMember<IDocumentProperties>(_ => _.AppVersion);
        VerifierSettings.RegisterFileConverter<Workbook>((target, context) => ConvertExcel(null, target, context));
        VerifierSettings.RegisterFileConverter<Worksheet>((target, _) => ConvertSheet(null, target));

        VerifierSettings.RegisterStreamConverter("pdf", ConvertPdf);
        VerifierSettings.RegisterFileConverter<Aspose.Pdf.Document>((target, context) => ConvertPdf(null, target, context));

        VerifierSettings.RegisterStreamConverter("pptx", ConvertPowerPoint);
        VerifierSettings.RegisterStreamConverter("ppt", ConvertPowerPoint);
        VerifierSettings.RegisterFileConverter<Presentation>((target, context) => ConvertPowerPoint(null, target, context));

        VerifierSettings.RegisterStreamConverter("docx", ConvertWord);
        VerifierSettings.RegisterStreamConverter("doc", ConvertWord);
        VerifierSettings.RegisterFileConverter<Document>((target, context) => ConvertWord(null, target, context));
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

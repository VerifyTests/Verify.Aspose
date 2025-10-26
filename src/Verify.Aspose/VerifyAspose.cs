using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;

namespace VerifyTests;

public static partial class VerifyAspose
{
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

        VerifierSettings.AddScrubber(RemoveGeneratorInfo);
        VerifierSettings.RegisterStreamConverter("xlsx", ConvertExcel);
        VerifierSettings.RegisterStreamConverter("xls", ConvertExcel);
        VerifierSettings.IgnoreMember<IDocumentProperties>(_ => _.AppVersion);
        VerifierSettings.RegisterFileConverter<Workbook>((target, _) => ConvertExcel(null, target));
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
        var index = 0;
        const string startPattern = "<meta name=\"generator\" content=\"Aspose";
        var stringValue = builder.ToString();
        while ((index = stringValue.IndexOf(startPattern, index, StringComparison.Ordinal)) != -1)
        {
            var endIndex = stringValue.IndexOf('>', index + startPattern.Length);

            if (endIndex != -1)
            {
                var length = endIndex + 1 - index;
                builder.Remove(index, length);
            }
            else
            {
                index += startPattern.Length;
            }
        }
    }
}
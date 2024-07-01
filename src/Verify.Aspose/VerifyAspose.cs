using Aspose.Cells;
using Aspose.Slides;
using Aspose.Words;
using MarkdownSaveOptions = Aspose.Words.Saving.MarkdownSaveOptions;

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

        VerifierSettings.AddExtraSettings(_ => _.Converters.Add(new CellConverter()));
        VerifierSettings.ScrubLines(_ => _.Contains("<meta name=\"generator\" content=\"Aspose"));
        VerifierSettings.RegisterFileConverter("xlsx", ConvertExcel);
        VerifierSettings.RegisterFileConverter("xls", ConvertExcel);
        VerifierSettings.IgnoreMember<IDocumentProperties>(_ => _.AppVersion);
        VerifierSettings.RegisterFileConverter<Workbook>(ConvertExcel);
        VerifierSettings.RegisterFileConverter<Worksheet>(ConvertSheet);

        VerifierSettings.RegisterFileConverter("pdf", ConvertPdf);
        VerifierSettings.RegisterFileConverter<Aspose.Pdf.Document>(ConvertPdf);

        VerifierSettings.RegisterFileConverter("pptx", ConvertPowerPoint);
        VerifierSettings.RegisterFileConverter("ppt", ConvertPowerPoint);
        VerifierSettings.RegisterFileConverter<Presentation>(ConvertPowerPoint);

        VerifierSettings.RegisterFileConverter("docx", ConvertWord);
        VerifierSettings.RegisterFileConverter("doc", ConvertWord);
        VerifierSettings.RegisterFileConverter<Document>(ConvertWord);
    }
}
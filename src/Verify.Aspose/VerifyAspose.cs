using Aspose.Cells;
using Aspose.Slides;

namespace VerifyTests;

public static partial class VerifyAspose
{
    public static void Initialize()
    {
        VerifierSettings.RegisterFileConverter("xlsx", ConvertExcel);
        VerifierSettings.RegisterFileConverter("xls", ConvertExcel);
        VerifierSettings.RegisterFileConverter<Workbook>(ConvertExcel);

        VerifierSettings.RegisterFileConverter("pdf", ConvertPdf);
        VerifierSettings.RegisterFileConverter<Aspose.Pdf.Document>(ConvertPdf);

        VerifierSettings.RegisterFileConverter("pptx", ConvertPowerPoint);
        VerifierSettings.RegisterFileConverter("ppt", ConvertPowerPoint);
        VerifierSettings.RegisterFileConverter<Presentation>(ConvertPowerPoint);

        VerifierSettings.RegisterFileConverter("docx", ConvertWord);
        VerifierSettings.RegisterFileConverter("doc", ConvertWord);
        VerifierSettings.RegisterFileConverter<Aspose.Words.Document>(ConvertWord);
    }
}
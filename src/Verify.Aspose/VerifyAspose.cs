using Aspose.Cells;
using Aspose.Slides;

namespace VerifyTests
{
    public static partial class VerifyAspose
    {
        public static void Initialize()
        {
            VerifierSettings.RegisterFileConverter("xlsx", "png", ConvertExcel);
            VerifierSettings.RegisterFileConverter("xls", "png", ConvertExcel);
            VerifierSettings.RegisterFileConverter<Workbook>("png", ConvertExcel);

            VerifierSettings.RegisterFileConverter("pdf", "png", ConvertPdf);
            VerifierSettings.RegisterFileConverter<Aspose.Pdf.Document>("png", ConvertPdf);

            VerifierSettings.RegisterFileConverter("pptx", "svg", ConvertPowerPoint);
            VerifierSettings.RegisterFileConverter("ppt", "svg", ConvertPowerPoint);
            VerifierSettings.RegisterFileConverter<Presentation>("svg", ConvertPowerPoint);

            VerifierSettings.RegisterFileConverter("docx", "png", ConvertWord);
            VerifierSettings.RegisterFileConverter("doc", "png", ConvertWord);
            VerifierSettings.RegisterFileConverter<Aspose.Words.Document>("png", ConvertWord);
        }
    }
}
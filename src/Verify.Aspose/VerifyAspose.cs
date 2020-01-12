using Aspose.Cells;
using Aspose.Slides;
using Verify;

public static partial class VerifyAspose
{
    public static void Initialize()
    {
        SharedVerifySettings.RegisterFileConverter("xlsx", "png", ConvertExcel);
        SharedVerifySettings.RegisterFileConverter("xls", "png", ConvertExcel);
        SharedVerifySettings.RegisterFileConverter<Workbook>("png", ConvertExcel);

        SharedVerifySettings.RegisterFileConverter("pdf", "png", ConvertPdf);
        SharedVerifySettings.RegisterFileConverter<Aspose.Pdf.Document>("png", ConvertPdf);

        SharedVerifySettings.RegisterFileConverter("pptx", "svg", ConvertPowerPoint);
        SharedVerifySettings.RegisterFileConverter("ppt", "svg", ConvertPowerPoint);
        SharedVerifySettings.RegisterFileConverter<Presentation>("svg", ConvertPowerPoint);

        SharedVerifySettings.RegisterFileConverter("docx", "png", ConvertWord);
        SharedVerifySettings.RegisterFileConverter("doc", "png", ConvertWord);
        SharedVerifySettings.RegisterFileConverter<Aspose.Words.Document>("png", ConvertWord);
    }
}
using Aspose.Cells;
using Aspose.Slides;
using Verify;

public static partial class VerifyAspose
{
    public static void Initialize()
    {
        SharedVerifySettings.RegisterFileConverter("xlsx", "png", GetExcelStreams);
        SharedVerifySettings.RegisterFileConverter("xls", "png", GetExcelStreams);
        SharedVerifySettings.RegisterFileConverter<Workbook>("png", GetExcelStreams);

        SharedVerifySettings.RegisterFileConverter("pdf", "png", GetPdfStreams);
        SharedVerifySettings.RegisterFileConverter<Aspose.Pdf.Document>("png", GetPdfStreams);

        SharedVerifySettings.RegisterFileConverter("pptx", "svg", GetPowerPointStreams);
        SharedVerifySettings.RegisterFileConverter("ppt", "svg", GetPowerPointStreams);
        SharedVerifySettings.RegisterFileConverter<Presentation>("svg", GetPowerPointStreams);

        SharedVerifySettings.RegisterFileConverter("docx", "png", GetWordStreams);
        SharedVerifySettings.RegisterFileConverter("doc", "png", GetWordStreams);
        SharedVerifySettings.RegisterFileConverter<Aspose.Words.Document>("png", GetWordStreams);
    }
}
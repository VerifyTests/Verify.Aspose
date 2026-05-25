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
        VerifierSettings.RegisterStreamConverter("xlsx", (name, stream, settings) => Locked(() => ConvertExcel(name, stream, settings)));
        VerifierSettings.RegisterStreamConverter("xls", (name, stream, settings) => Locked(() => ConvertExcel(name, stream, settings)));
        VerifierSettings.IgnoreMember<IDocumentProperties>(_ => _.AppVersion);
        VerifierSettings.RegisterFileConverter<Workbook>((target, _) => Locked(() => ConvertExcel(null, target)));
        VerifierSettings.RegisterFileConverter<Worksheet>((target, _) => Locked(() => ConvertSheet(null, target)));

        VerifierSettings.RegisterStreamConverter("pdf", (name, stream, settings) => Locked(() => ConvertPdf(name, stream, settings)));
        VerifierSettings.RegisterFileConverter<Aspose.Pdf.Document>((target, context) => Locked(() => ConvertPdf(null, target, context)));

        VerifierSettings.RegisterStreamConverter("pptx", (name, stream, settings) => Locked(() => ConvertPowerPoint(name, stream, settings)));
        VerifierSettings.RegisterStreamConverter("ppt", (name, stream, settings) => Locked(() => ConvertPowerPoint(name, stream, settings)));
        VerifierSettings.RegisterFileConverter<Presentation>((target, context) => Locked(() => ConvertPowerPoint(null, target, context)));

        VerifierSettings.RegisterStreamConverter("docx", (name, stream, settings) => Locked(() => ConvertWord(name, stream, settings)));
        VerifierSettings.RegisterStreamConverter("doc", (name, stream, settings) => Locked(() => ConvertWord(name, stream, settings)));
        VerifierSettings.RegisterFileConverter<Document>((target, context) => Locked(() => ConvertWord(null, target, context)));
    }

    // Aspose document processing relies on global, non-thread-safe state (font caches,
    // FontSettings.DefaultInstance, rendering engine internals). Aspose.Pdf conversion also
    // invokes Aspose.Words internally, so a single global lock is used across all formats.
    // Without this, rendering can be non-deterministic when the consuming test suite runs in
    // parallel, producing intermittent image-snapshot diffs.
    static readonly Lock asposeLock = new();

    static ConversionResult Locked(Func<ConversionResult> convert)
    {
        lock (asposeLock)
        {
            return convert();
        }
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

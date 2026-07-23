namespace VerifyTestsAspose;

public static class VerifyAsposeSettings
{
    public static void IncludeWordStyles(this VerifySettings settings) =>
        settings.Context["VerifyAsposeIncludeWordStyles"] = true;

    public static SettingsTask IncludeWordStyles(this SettingsTask settings)
    {
        settings.CurrentSettings.IncludeWordStyles();
        return settings;
    }

    internal static bool GetIncludeWordStyles(this IReadOnlyDictionary<string, object> settings)
    {
        if (settings.TryGetValue("VerifyAsposeIncludeWordStyles", out var value))
        {
            return (bool)value;
        }

        return false;
    }

    /// <summary>
    /// Limits the number of rendered page/slide <c>png</c> snapshots to the first
    /// <paramref name="count"/>. Any full-document binary target (for example the <c>docx</c>
    /// emitted for Word) is unaffected and always contains the full source document.
    /// </summary>
    public static void PagesToInclude(this VerifySettings settings, int count) =>
        settings.Context["VerifyAsposePagesToInclude"] = count;

    /// <inheritdoc cref="PagesToInclude(VerifySettings, int)"/>
    public static SettingsTask PagesToInclude(this SettingsTask settings, int count)
    {
        settings.CurrentSettings.PagesToInclude(count);
        return settings;
    }

    internal static int GetPagesToInclude(this IReadOnlyDictionary<string, object> settings, int count)
    {
        if (!settings.TryGetValue("VerifyAsposePagesToInclude", out var value))
        {
            return count;
        }

        return Math.Min(count, (int) value);
    }

    public static void PdfPngDevice(this VerifySettings settings, Func<Aspose.Pdf.Page, PngDevice> func) =>
        settings.Context["VerifyAsposePdfPngDevice"] = func;

    public static SettingsTask PdfPngDevice(this SettingsTask settings, Func<Aspose.Pdf.Page, PngDevice> func)
    {
        settings.CurrentSettings.PdfPngDevice(func);
        return settings;
    }

    static PngDevice defaultDevice = new();

    internal static PngDevice GetPdfPngDevice(this IReadOnlyDictionary<string, object> settings, Aspose.Pdf.Page page)
    {
        if (!settings.TryGetValue("VerifyAsposePdfPngDevice", out var value))
        {
            return defaultDevice;
        }

        var func = (Func<Aspose.Pdf.Page, PngDevice>) value;
        return func(page);
    }

    /// <summary>
    /// Snapshots the pdf bytes exactly as produced, skipping the normalization that neutralizes the
    /// trailer <c>/ID</c>, the <c>/CreationDate</c> and <c>/ModDate</c>, and the XMP dates and
    /// identifiers. Use it when the producer already emits byte-deterministic documents, since
    /// normalizing them again copies the whole buffer, rescans it, and — when the XMP packet is
    /// canonicalized — rebuilds it and repairs the cross-reference table, all to change nothing.
    /// </summary>
    /// <remarks>
    /// Only skip this when the producer is genuinely deterministic. Without it a freshly generated
    /// pdf carries a wall-clock <c>/CreationDate</c> and a fresh <c>/ID</c>, so the snapshot differs
    /// on every run.
    /// <para>
    /// The XMP canonicalization is worth calling out because it is the pass that changes bytes for
    /// an already-deterministic producer: it collapses the packet's whitespace, so enabling or
    /// disabling this setting on an existing suite shifts the stored <c>.verified.pdf</c> even
    /// though nothing about the document changed. Expect to re-accept those snapshots once.
    /// </para>
    /// </remarks>
    public static void SkipPdfNormalization(this VerifySettings settings) =>
        settings.Context["VerifyAsposeSkipPdfNormalization"] = true;

    /// <inheritdoc cref="SkipPdfNormalization(VerifySettings)"/>
    public static SettingsTask SkipPdfNormalization(this SettingsTask settings)
    {
        settings.CurrentSettings.SkipPdfNormalization();
        return settings;
    }

    internal static bool NormalizePdf(this IReadOnlyDictionary<string, object> settings) =>
        !settings.TryGetValue("VerifyAsposeSkipPdfNormalization", out var value) ||
        value is not true;
}

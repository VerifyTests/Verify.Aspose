namespace VerifyTests;

/// <summary>
/// Controls which outputs a document is split into when verified. Passed to <see cref="VerifyAspose.Initialize"/>.
/// The source document (pdf/docx/xlsx/pptx) is not controlled by this; use <c>VerifierSettings.ExcludeTargets</c> for that.
/// </summary>
[Flags]
public enum AsposeOutputs
{
    /// <summary>
    /// Render pages, slides, and sheets to png images.
    /// </summary>
    Png = 1,

    /// <summary>
    /// Extract the document text (markdown) into the <c>Text</c> property of the info for pdf and Word documents.
    /// </summary>
    Text = 2,

    /// <summary>
    /// Export each Excel sheet to csv.
    /// </summary>
    Csv = 4,

    /// <summary>
    /// All outputs. The default.
    /// </summary>
    All = Png | Text | Csv
}

namespace VerifyTests;

/// <summary>
/// What <c>Initialize</c> used to be given to choose the outputs a document is split into. Settings
/// of Verify choose that now, and nothing reads this: it is here so that code still naming it is
/// told what to use in its place.
/// </summary>
[Obsolete(
    "AsposeOutputs and the outputs argument of Initialize are replaced by settings of Verify: VerifierSettings.ExcludeDerivedTargets(\"png\") to leave out the page images, VerifierSettings.PageText(PageTextPlacement.None) to leave out the text and VerifierSettings.ExcludeDerivedTargets(\"csv\") to leave out the csv files. See https://github.com/VerifyTests/Verify.Aspose#migrating-from-5x",
    true)]
[Flags]
public enum AsposeOutputs
{
    /// <summary>
    /// No outputs. Only the source document and info are emitted.
    /// </summary>
    None = 0,

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

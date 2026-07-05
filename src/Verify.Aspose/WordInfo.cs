using Aspose.Words.Loading;

readonly record struct WordInfo
{
    /// <summary>
    /// Total number of pages in the source document. This is the full count regardless of any
    /// <c>PagesToInclude</c> filter, which only trims the rendered page images.
    /// </summary>
    public int PageCount { get; init; }
    public string HasRevisions { get; init; }
    public EditingLanguage DefaultLocale { get; init; }
    public Dictionary<string, object> Properties { get; init; }
    public Dictionary<string, object> CustomProperties { get; init; }
    public bool ShadeFormData { get; init; }
    public List<string> Fonts { get; init; }
    public List<string> EmbeddedFonts { get; init; }
    public string Text { get; init; }
}
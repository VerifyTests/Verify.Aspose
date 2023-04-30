using Aspose.Words.Loading;

readonly record struct WordInfo
{
    public string HasRevisions { get; init; }
    public EditingLanguage DefaultLocale { get; init; }
    public Dictionary<string, object> Properties { get; init; }
    public Dictionary<string, object> CustomProperties { get; init; }
    public string Text { get; init; }
}
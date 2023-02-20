namespace VerifyTests;

class TempDirectory : IDisposable
{
    string directory;

    public TempDirectory()
    {
        directory = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
        Directory.CreateDirectory(directory);
    }

    public void Dispose() =>
        Directory.Delete(directory, true);

    public static implicit operator string(TempDirectory tempDirectory) => tempDirectory.directory;
}
public static class ModuleInitializer
{
    #region enable

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyAspose.Initialize();

    #endregion

    [ModuleInitializer]
    public static void InitializeOther()
    {
        ApplyAsposeLicense();

        VerifyDiffPlex.Initialize();
        VerifyImageMagick.RegisterComparers(.01);
        VerifierSettings.IgnoreMember("Width");
        VerifierSettings.ScrubLinesContaining(
            "Created with an evaluation",
            "Evaluation Only");
    }

    static void ApplyAsposeLicense()
    {
        var stream = new MemoryStream();
        var writer = new StreamWriter(stream);
        writer.Write(Environment.GetEnvironmentVariable("AsposeLicense"));
        writer.Flush();
        stream.Position = 0;
        var license = new Aspose.Cells.License();
        license.SetLicense(stream);
    }
}
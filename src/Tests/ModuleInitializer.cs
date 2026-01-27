[assembly: Culture("en-AU")]

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
    }

    static void ApplyAsposeLicense()
    {
        var licenseText = Environment.GetEnvironmentVariable("AsposeLicense");
        if (licenseText == null)
        {
            throw new("Expected a `AsposeLicense` environment variable");
        }

        var stream = new MemoryStream();
        var writer = new StreamWriter(stream);
        writer.Write(licenseText);
        writer.Flush();

        Apply(stream);
    }

    public static void Apply(Stream stream)
    {
        //Email(stream);
        Pdf(stream);
        Cells(stream);
        Word(stream);
        Slides(stream);
    }

    static void Slides(Stream stream)
    {
        var lic = new Aspose.Slides.License();
        stream.Position = 0;
        lic.SetLicense(stream);
    }

    static void Word(Stream stream)
    {
        var lic = new Aspose.Words.License();
        stream.Position = 0;
        lic.SetLicense(stream);
    }

    // static void Email(Stream stream)
    // {
    //     var lic = new Aspose.Email.License();
    //     stream.Position = 0;
    //     lic.SetLicense(stream);
    // }

    static void Pdf(Stream stream)
    {
        var lic = new Aspose.Pdf.License();
        stream.Position = 0;
        lic.SetLicense(stream);
    }

    static void Cells(Stream stream)
    {
        var lic = new Aspose.Cells.License();
        stream.Position = 0;
        lic.SetLicense(stream);
    }
}

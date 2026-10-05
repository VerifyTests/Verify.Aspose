public static class ModuleInitializer
{
    #region InitializeOutputs

    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyAspose.Initialize();

        // For every test: no page is drawn and no sheet is exported to csv,
        // so only the documents and their text are verified
        VerifierSettings.ExcludeDerivedTargets("png", "csv");
    }

    #endregion

    [ModuleInitializer]
    public static void InitializeOther()
    {
        var culture = new CultureInfo("en-AU");
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;

        VerifierSettings.UseSsimForPng();
        VerifierSettings.IgnoreMember("Width");
    }
}

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
        var culture = new CultureInfo("en-AU");
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;

        VerifierSettings.UseSsimForPng();
        VerifierSettings.IgnoreMember("Width");
        VerifierSettings.InitializePlugins();
    }
}

public static class ModuleInitializer
{
    #region InitializeOutputs

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyAspose.Initialize(AsposeOutputs.Text);

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

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
        VerifyImageMagick.RegisterComparers(.3);
        VerifierSettings.InitializePlugins();
        VerifierSettings.IgnoreMember("Width");
    }
}
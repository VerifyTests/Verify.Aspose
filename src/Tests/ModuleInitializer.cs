public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyDiffPlex.Initialize();
        VerifyAspose.Initialize();
        VerifyImageMagick.RegisterComparers(.22);
        VerifierSettings.IgnoreMember("Width");
    }
}
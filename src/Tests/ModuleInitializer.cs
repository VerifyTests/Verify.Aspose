public static class ModuleInitializer
{
#region enable
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyAspose.Initialize();
        #endregion
        VerifyDiffPlex.Initialize();
        VerifyImageMagick.RegisterComparers(.05);
        VerifierSettings.IgnoreMember("Width");
    }
}
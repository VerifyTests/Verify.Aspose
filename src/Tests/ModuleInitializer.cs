﻿public static class ModuleInitializer
{
    #region enable

    [ModuleInitializer]
    public static void Initialize() =>
        VerifyAspose.Initialize();

    #endregion

    [ModuleInitializer]
    public static void InitializeOther()
    {
        VerifyDiffPlex.Initialize();
        VerifyImageMagick.RegisterComparers(.3);
        VerifierSettings.IgnoreMember("Width");
    }
}
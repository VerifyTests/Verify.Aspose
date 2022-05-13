public static class ModuleInitializer
{
    [ModuleInitializer]
    public static void Initialize()
    {
        VerifyAspose.Initialize();
        VerifyImageMagick.RegisterComparers(.05);
    }
}
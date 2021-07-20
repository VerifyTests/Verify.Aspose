using System;
using System.Collections.Generic;
using Aspose.Pdf.Devices;

namespace VerifyTests
{
    public static class VerifyAsposeSettings
    {
        public static void PagesToInclude(this VerifySettings settings, int count)
        {
            settings.Context["VerifyAsposePagesToInclude"] = count;
        }

        public static SettingsTask PagesToInclude(this SettingsTask settings, int count)
        {
            settings.CurrentSettings.PagesToInclude(count);
            return settings;
        }

        internal static int GetPagesToInclude(this IReadOnlyDictionary<string, object> settings, int count)
        {
            if (!settings.TryGetValue("VerifyAsposePagesToInclude", out var value))
            {
                return count;
            }

            return Math.Min(count, (int) value);
        }

        public static void PdfPngDevice(this VerifySettings settings, Func<Aspose.Pdf.Page, PngDevice> func)
        {
            settings.Context["VerifyAsposePdfPngDevice"] = func;
        }

        public static SettingsTask PdfPngDevice(this SettingsTask settings, Func<Aspose.Pdf.Page, PngDevice> func)
        {
            settings.CurrentSettings.PdfPngDevice(func);
            return settings;
        }

        static PngDevice defaultDevice = new();

        internal static PngDevice GetPdfPngDevice(this IReadOnlyDictionary<string, object> settings, Aspose.Pdf.Page page)
        {
            if (!settings.TryGetValue("VerifyAsposePdfPngDevice", out var value))
            {
                return defaultDevice;
            }

            var func = (Func<Aspose.Pdf.Page, PngDevice>) value;
            return func(page);
        }
    }
}
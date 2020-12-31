using System;
using System.Collections.Generic;
using Aspose.Pdf.Devices;

namespace VerifyTests
{
    public static class VerifyAsposeSettings
    {
        public static void PagesToInclude(this VerifySettings settings, int count)
        {
            Guard.AgainstNull(settings, nameof(settings));
            settings.Context["VerifyAsposePagesToInclude"] = count;
        }

        public static SettingsTask PagesToInclude(this SettingsTask settings, int count)
        {
            Guard.AgainstNull(settings, nameof(settings));
            settings.CurrentSettings.PagesToInclude(count);
            return settings;
        }

        internal static int GetPagesToInclude(this IReadOnlyDictionary<string, object> settings, int count)
        {
            Guard.AgainstNull(settings, nameof(settings));
            if (!settings.TryGetValue("VerifyAsposePagesToInclude", out var value))
            {
                return count;
            }

            return Math.Min(count, (int) value);
        }

        public static void PdfPngDevice(this VerifySettings settings, Func<Aspose.Pdf.Page, PngDevice> func)
        {
            Guard.AgainstNull(settings, nameof(settings));
            Guard.AgainstNull(func, nameof(func));
            settings.Context["VerifyAsposePdfPngDevice"] = func;
        }

        public static SettingsTask PdfPngDevice(this SettingsTask settings, Func<Aspose.Pdf.Page, PngDevice> func)
        {
            Guard.AgainstNull(settings, nameof(settings));
            settings.CurrentSettings.PdfPngDevice(func);
            return settings;
        }

        static PngDevice defaultDevice = new();

        internal static PngDevice GetPdfPngDevice(this IReadOnlyDictionary<string, object> settings, Aspose.Pdf.Page page)
        {
            Guard.AgainstNull(settings, nameof(settings));
            if (!settings.TryGetValue("VerifyAsposePdfPngDevice", out var value))
            {
                return defaultDevice;
            }

            var func = (Func<Aspose.Pdf.Page, PngDevice>) value;
            return func(page);
        }
    }
}
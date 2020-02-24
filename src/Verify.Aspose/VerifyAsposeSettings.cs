using System;

namespace Verify
{
    public static class VerifyAsposeSettings
    {
        public static void PagesToInclude(this VerifySettings settings, int count)
        {
            Guard.AgainstNull(settings, nameof(settings));
            settings.Data["VerifyAsposePagesToInclude"] = count;
        }

        internal static int GetPagesToInclude(this VerifySettings settings, int count)
        {
            Guard.AgainstNull(settings, nameof(settings));
            if (!settings.Data.TryGetValue("VerifyAsposePagesToInclude", out var value))
            {
                return count;
            }
            return Math.Min(count, (int)value);
        }
    }
}
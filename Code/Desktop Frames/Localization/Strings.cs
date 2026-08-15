using System;
using System.Globalization;
using System.Resources;

namespace Desktop_Frames.Localization
{
    /// <summary>
    /// User-facing text, kept out of the code so it can be translated.
    ///
    /// Only text the user actually reads belongs here: window titles, menu
    /// entries, labels, tooltips and message boxes. Log messages stay in
    /// English in the code — they are read by the developer, not by the user,
    /// and a translated log is a log nobody can search.
    ///
    /// The same rule protects the values that are written to frames.json
    /// ("Medium", "Details", "Gray"…): those are keys, not text, and must never
    /// go through here or the saved configuration stops being readable.
    ///
    /// The language follows Windows unless <see cref="OverrideCulture"/> is set.
    /// </summary>
    internal static class Strings
    {
        private static readonly ResourceManager Manager =
            new ResourceManager("Desktop_Frames.Localization.Strings", typeof(Strings).Assembly);

        /// <summary>Forces a language; null (the default) follows Windows.</summary>
        public static CultureInfo? OverrideCulture { get; set; }

        private static CultureInfo Culture => OverrideCulture ?? CultureInfo.CurrentUICulture;

        /// <summary>
        /// Text for a key. Falls back to the key itself if the resource is
        /// missing, so a forgotten entry shows up as a visible label instead of
        /// crashing a window that was about to open.
        /// </summary>
        public static string Get(string key)
        {
            try
            {
                return Manager.GetString(key, Culture) ?? key;
            }
            catch (MissingManifestResourceException)
            {
                return key;
            }
        }

        /// <summary>Text for a key, with {0}, {1}… replaced.</summary>
        public static string Get(string key, params object[] args)
        {
            string format = Get(key);
            try
            {
                return string.Format(Culture, format, args);
            }
            catch (FormatException)
            {
                return format;
            }
        }

        // ── Tray ────────────────────────────────────────────────────────────
        public static string TrayExportDoneTitle => Get("TrayExportDoneTitle");
        public static string TrayExportDoneText => Get("TrayExportDoneText");
        public static string TrayExportFailedTitle => Get("TrayExportFailedTitle");
        public static string TrayExportFailedText => Get("TrayExportFailedText");
        public static string TrayReloadingFrames => Get("TrayReloadingFrames");
    }
}

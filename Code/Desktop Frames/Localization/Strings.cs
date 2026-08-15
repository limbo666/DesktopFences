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

        // ── Buttons, shared by every dialog ─────────────────────────────────
        public static string ButtonYes => Get("ButtonYes");
        public static string ButtonNo => Get("ButtonNo");
        public static string ButtonOk => Get("ButtonOk");

        // ── Confirmation dialogs ────────────────────────────────────────────
        public static string DeleteTabTitle => Get("DeleteTabTitle");
        public static string ConfirmDeleteTitle => Get("ConfirmDeleteTitle");
        public static string DeleteFrameHeading => Get("DeleteFrameHeading");
        public static string DeleteFrameQuestion => Get("DeleteFrameQuestion");

        // ── Tray ────────────────────────────────────────────────────────────
        public static string TrayExportDoneTitle => Get("TrayExportDoneTitle");
        public static string TrayExportDoneText => Get("TrayExportDoneText");
        public static string TrayExportFailedTitle => Get("TrayExportFailedTitle");
        public static string TrayExportFailedText => Get("TrayExportFailedText");
        public static string TrayReloadingFrames => Get("TrayReloadingFrames");

        // ── Context menus and frame tooltips ─────────────────────────
        public static string TooltipGoUp => Get("TooltipGoUp");
        public static string TooltipSetHome => Get("TooltipSetHome");
        public static string TooltipFilterFiles => Get("TooltipFilterFiles");
        public static string TooltipClearFilter => Get("TooltipClearFilter");
        public static string TooltipBlankSpacer => Get("TooltipBlankSpacer");
        public static string MenuAbout => Get("MenuAbout");
        public static string MenuOptions => Get("MenuOptions");
        public static string MenuNewFrame => Get("MenuNewFrame");
        public static string MenuNewPortalFrame => Get("MenuNewPortalFrame");
        public static string MenuNewNoteFrame => Get("MenuNewNoteFrame");
        public static string MenuAddPlugin => Get("MenuAddPlugin");
        public static string MenuNoPluginsAvailable => Get("MenuNoPluginsAvailable");
        public static string MenuEnableTabs => Get("MenuEnableTabs");
        public static string MenuDeleteThisFrame => Get("MenuDeleteThisFrame");
        public static string MenuExportThisFrame => Get("MenuExportThisFrame");
        public static string MenuImportFrame => Get("MenuImportFrame");
        public static string MenuRestoreLastDeleted => Get("MenuRestoreLastDeleted");
        public static string MenuExit => Get("MenuExit");
        public static string MenuEdit => Get("MenuEdit");
        public static string MenuMove => Get("MenuMove");
        public static string MenuRemove => Get("MenuRemove");
        public static string MenuCopyItem => Get("MenuCopyItem");
        public static string MenuRunAsAdmin => Get("MenuRunAsAdmin");
        public static string MenuAlwaysRunAsAdmin => Get("MenuAlwaysRunAsAdmin");
        public static string MenuCopyPath => Get("MenuCopyPath");
        public static string MenuFolderPath => Get("MenuFolderPath");
        public static string MenuFullPath => Get("MenuFullPath");
        public static string MenuOpenTargetFolder => Get("MenuOpenTargetFolder");
        public static string MenuSendToDesktop => Get("MenuSendToDesktop");
        public static string MenuRunAsDifferentUser => Get("MenuRunAsDifferentUser");
        public static string MenuAlwaysRunAsDifferentUser => Get("MenuAlwaysRunAsDifferentUser");
        public static string MenuAutoRoll => Get("MenuAutoRoll");
        public static string MenuAlwaysOnTop => Get("MenuAlwaysOnTop");
        public static string MenuHideFrame => Get("MenuHideFrame");
        public static string MenuPeekBehind => Get("MenuPeekBehind");
        public static string MenuClearDeadShortcuts => Get("MenuClearDeadShortcuts");
        public static string MenuOpenFrameFolder => Get("MenuOpenFrameFolder");
        public static string MenuViewAsDetails => Get("MenuViewAsDetails");
        public static string MenuPasteItem => Get("MenuPasteItem");
        public static string MenuPluginSettings => Get("MenuPluginSettings");
        public static string MenuCustomize => Get("MenuCustomize");
        public static string MenuExportAllIcons => Get("MenuExportAllIcons");
        public static string MenuAddSpacer => Get("MenuAddSpacer");
        public static string MenuSpacerBlank => Get("MenuSpacerBlank");
        public static string MenuSpacerDot => Get("MenuSpacerDot");
        public static string MenuNameAfterPath => Get("MenuNameAfterPath");
    }
}

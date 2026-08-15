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

        // ── Settings windows ────────────────────────────────────
        public static string BtnApply => Get("BtnApply");
        public static string BtnSave => Get("BtnSave");
        public static string BtnCancel => Get("BtnCancel");
        public static string BtnDefault => Get("BtnDefault");
        public static string CustomizeTitle => Get("CustomizeTitle");
        public static string BtnApplyToAll => Get("BtnApplyToAll");
        public static string BtnSaveToAll => Get("BtnSaveToAll");
        public static string TabFrame => Get("TabFrame");
        public static string TabTitle => Get("TabTitle");
        public static string TabIcons => Get("TabIcons");
        public static string LblCustomColor => Get("LblCustomColor");
        public static string LblCustomLaunchEffect => Get("LblCustomLaunchEffect");
        public static string LblFrameBorderColor => Get("LblFrameBorderColor");
        public static string LblFrameBorderThickness => Get("LblFrameBorderThickness");
        public static string LblTitleTextColor => Get("LblTitleTextColor");
        public static string LblTitleTextSize => Get("LblTitleTextSize");
        public static string LblBoldTitleText => Get("LblBoldTitleText");
        public static string LblPortalView => Get("LblPortalView");
        public static string LblIconSize => Get("LblIconSize");
        public static string LblIconSpacing => Get("LblIconSpacing");
        public static string LblTextColor => Get("LblTextColor");
        public static string LblDisableTextShadow => Get("LblDisableTextShadow");
        public static string LblGrayscaleIcons => Get("LblGrayscaleIcons");
        public static string OptionsTitle => Get("OptionsTitle");
        public static string OptionsHeading => Get("OptionsHeading");
        public static string LblDefaultPortalView => Get("LblDefaultPortalView");
        public static string LblNotificationSound => Get("LblNotificationSound");
        public static string TooltipChameleon => Get("TooltipChameleon");
        public static string NoteAutoRoll => Get("NoteAutoRoll");
        public static string LblMenuIcon => Get("LblMenuIcon");
        public static string LblLockIcon => Get("LblLockIcon");
        public static string TooltipAutoReposition => Get("TooltipAutoReposition");
        public static string LblEnableProfileAutomation => Get("LblEnableProfileAutomation");
        public static string NoteHotkeysRestart => Get("NoteHotkeysRestart");
        public static string NoteAutoOrganize => Get("NoteAutoOrganize");
        public static string LblColor => Get("LblColor");
        public static string LblEffect => Get("LblEffect");
        public static string LblMinimumLogLevel => Get("LblMinimumLogLevel");
        public static string LblDonate => Get("LblDonate");
        public static string BtnDonate => Get("BtnDonate");
    }
}

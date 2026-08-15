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


        /// <summary>
        /// Label for a value that is stored in the configuration ("Medium",
        /// "Details", "Gray"…). Only the label is translated: the value itself
        /// travels untouched, so saved settings stay readable in any language.
        /// </summary>
        public static string Item(string storedValue) =>
            string.IsNullOrEmpty(storedValue) ? storedValue : Get("Item" + storedValue.Replace(" ", ""));


        /// <summary>
        /// Label for a hotkey. Letters, digits and function keys read the same
        /// everywhere and stay as they are; only the named keys get translated.
        /// The key code itself is untouched — it lives in the item's Tag.
        /// </summary>
        public static string KeyLabel(string name) => name switch
        {
            "Comma (,)" => Get("KeyComma"),
            "Period (.)" => Get("KeyPeriod"),
            "Tilde (~)" => Get("KeyTilde"),
            "Space" => Get("KeySpace"),
            "Tab" => Get("KeyTab"),
            "Enter" => Get("KeyEnter"),
            "Escape" => Get("KeyEscape"),
            _ => name
        };

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

        // ── Dialogs, panels and secondary windows ───────────────────
        public static string AboutTitle => Get("AboutTitle");
        public static string AboutSupportDevelopment => Get("AboutSupportDevelopment");
        public static string AboutVisitGitHub => Get("AboutVisitGitHub");
        public static string AboutLicense => Get("AboutLicense");
        public static string AboutGreatToHaveYou => Get("AboutGreatToHaveYou");
        public static string AboutDontFeedDragons => Get("AboutDontFeedDragons");
        public static string AboutDonateBig => Get("AboutDonateBig");
        public static string AboutWhat => Get("AboutWhat");
        public static string AboutThinkAgain => Get("AboutThinkAgain");
        public static string AboutPleaseDonate => Get("AboutPleaseDonate");
        public static string AutoOrgTitle => Get("AutoOrgTitle");
        public static string AutoOrgAddRule => Get("AutoOrgAddRule");
        public static string AutoOrgRemove => Get("AutoOrgRemove");
        public static string AutoOrgEnableRule => Get("AutoOrgEnableRule");
        public static string AutoOrgRuleName => Get("AutoOrgRuleName");
        public static string AutoOrgIfExtension => Get("AutoOrgIfExtension");
        public static string AutoOrgCustomExtensions => Get("AutoOrgCustomExtensions");
        public static string AutoOrgFilenameContains => Get("AutoOrgFilenameContains");
        public static string AutoOrgMoveTo => Get("AutoOrgMoveTo");
        public static string AutoOrgGeneratePortal => Get("AutoOrgGeneratePortal");
        public static string AutoOrgIfExists => Get("AutoOrgIfExists");
        public static string AutoOrgSaveRules => Get("AutoOrgSaveRules");
        public static string AutoOrgBrowseFolder => Get("AutoOrgBrowseFolder");
        public static string AutomationTitle => Get("AutomationTitle");
        public static string AutomationManage => Get("AutomationManage");
        public static string AutomationExisting => Get("AutomationExisting");
        public static string AutomationDeleteSelected => Get("AutomationDeleteSelected");
        public static string AutomationDoubleClickEdit => Get("AutomationDoubleClickEdit");
        public static string AutomationRuleDefinition => Get("AutomationRuleDefinition");
        public static string AutomationPickWindow => Get("AutomationPickWindow");
        public static string AutomationPersistedMode => Get("AutomationPersistedMode");
        public static string BackupSelectExportFile => Get("BackupSelectExportFile");
        public static string EditShortcutTitle => Get("EditShortcutTitle");
        public static string LblIcon => Get("LblIcon");
        public static string EditSelectTarget => Get("EditSelectTarget");
        public static string EditSelectIcon => Get("EditSelectIcon");
        public static string FocusSearchActive => Get("FocusSearchActive");
        public static string IconPickerTitle => Get("IconPickerTitle");
        public static string IconPickerHint => Get("IconPickerHint");
        public static string ImportTabTitle => Get("ImportTabTitle");
        public static string ImportTabPrompt => Get("ImportTabPrompt");
        public static string ImportTabNoFrames => Get("ImportTabNoFrames");
        public static string MoveItemTitle => Get("MoveItemTitle");
        public static string MoveItemPrompt => Get("MoveItemPrompt");
        public static string NoteClickFinish => Get("NoteClickFinish");
        public static string NoteClickEdit => Get("NoteClickEdit");
        public static string NoteTextFormat => Get("NoteTextFormat");
        public static string NoteCopyAll => Get("NoteCopyAll");
        public static string NoteClearAll => Get("NoteClearAll");
        public static string NotificationTitle => Get("NotificationTitle");
        public static string NotificationDontShowAgain => Get("NotificationDontShowAgain");
        public static string MenuCutItem => Get("MenuCutItem");
        public static string MenuRenameItem => Get("MenuRenameItem");
        public static string MenuDeleteItem => Get("MenuDeleteItem");
        public static string MenuOpenWith => Get("MenuOpenWith");
        public static string MenuCopyItemPath => Get("MenuCopyItemPath");
        public static string MenuCopyFolderPath => Get("MenuCopyFolderPath");
        public static string MenuResetSorting => Get("MenuResetSorting");
        public static string ProfileManagerTitle => Get("ProfileManagerTitle");
        public static string BtnCreate => Get("BtnCreate");
        public static string UpdateAvailable => Get("UpdateAvailable");
        public static string SearchTitle => Get("SearchTitle");
        public static string SearchPlaceholder => Get("SearchPlaceholder");
        public static string TabAddNew => Get("TabAddNew");
        public static string TabRename => Get("TabRename");
        public static string TabMoveLeft => Get("TabMoveLeft");
        public static string TabMoveRight => Get("TabMoveRight");
        public static string TabAddNewHint => Get("TabAddNewHint");
        public static string TextFormatTitle => Get("TextFormatTitle");
        public static string TextAppearance => Get("TextAppearance");
        public static string TextBehavior => Get("TextBehavior");

        // ── Widget plugins ────────────────────────────────────────
        public static string CalcCopy => Get("CalcCopy");
        public static string CalcPaste => Get("CalcPaste");
        public static string CalcClear => Get("CalcClear");
        public static string CalcClearHistory => Get("CalcClearHistory");
        public static string CalcSettings => Get("CalcSettings");
        public static string CalcShowTape => Get("CalcShowTape");
        public static string CalcShowKeypad => Get("CalcShowKeypad");
        public static string CalcClearOnNewInput => Get("CalcClearOnNewInput");
        public static string CalcShowOperationNames => Get("CalcShowOperationNames");
        public static string CalcDisplayColor => Get("CalcDisplayColor");
        public static string CalcHistoryColor => Get("CalcHistoryColor");
        public static string CalcBadgeFade => Get("CalcBadgeFade");
        public static string CalcClearMemory => Get("CalcClearMemory");
        public static string TermSettings => Get("TermSettings");
        public static string TermConsoleOptions => Get("TermConsoleOptions");
        public static string TermTargetShell => Get("TermTargetShell");
        public static string TermStartupDirectory => Get("TermStartupDirectory");
        public static string TermSelectStartupDir => Get("TermSelectStartupDir");
        public static string TermClearHistory => Get("TermClearHistory");
        public static string TermHistoryCleared => Get("TermHistoryCleared");
        public static string PhotoSettings => Get("PhotoSettings");
        public static string PhotoFitMode => Get("PhotoFitMode");
        public static string PhotoCropToFill => Get("PhotoCropToFill");
        public static string PhotoFitInside => Get("PhotoFitInside");
        public static string PhotoStretch => Get("PhotoStretch");
        public static string PhotoOriginalSize => Get("PhotoOriginalSize");
        public static string PhotoTransition => Get("PhotoTransition");
        public static string PhotoCrossfade => Get("PhotoCrossfade");
        public static string PhotoBlurCrossfade => Get("PhotoBlurCrossfade");
        public static string PhotoVerticalWipe => Get("PhotoVerticalWipe");
        public static string PhotoSubtleTwist => Get("PhotoSubtleTwist");
        public static string PhotoNoTransition => Get("PhotoNoTransition");
        public static string PhotoLiveRescan => Get("PhotoLiveRescan");
        public static string NetNoInterfaces => Get("NetNoInterfaces");
        public static string NetPublicWan => Get("NetPublicWan");
        public static string NetIpLabel => Get("NetIpLabel");
        public static string NetSettings => Get("NetSettings");
        public static string NetExternal => Get("NetExternal");
        public static string NetShowPublicWan => Get("NetShowPublicWan");
        public static string NetLocalAdapters => Get("NetLocalAdapters");
        public static string NetShowDisconnected => Get("NetShowDisconnected");
        public static string NetSelectInterfaces => Get("NetSelectInterfaces");
        public static string PerfSettings => Get("PerfSettings");
        public static string PerfThemeLayout => Get("PerfThemeLayout");
        public static string PerfVisualStyle => Get("PerfVisualStyle");
        public static string PerfDynamicColors => Get("PerfDynamicColors");
        public static string PerfStaticBarColor => Get("PerfStaticBarColor");
        public static string PerfSensors => Get("PerfSensors");
        public static string PerfRefreshRate => Get("PerfRefreshRate");
        public static string QueueTitle => Get("QueueTitle");
        public static string QueueInitializing => Get("QueueInitializing");
        public static string QueueWmiError => Get("QueueWmiError");
        public static string QueueConfig => Get("QueueConfig");
        public static string QueueDiagnosticSettings => Get("QueueDiagnosticSettings");
        public static string QueueShowTitle => Get("QueueShowTitle");
        public static string QueueShowValues => Get("QueueShowValues");
        public static string QueueShowCores => Get("QueueShowCores");
        public static string QueueDynamicColors => Get("QueueDynamicColors");
        public static string QueueRefreshRate => Get("QueueRefreshRate");
        public static string VuSettings => Get("VuSettings");
        public static string VuDisplayOptions => Get("VuDisplayOptions");
        public static string VuLayoutMode => Get("VuLayoutMode");
        public static string VuBallistics => Get("VuBallistics");
        public static string VuSignalGain => Get("VuSignalGain");
        public static string VuGainHint => Get("VuGainHint");
        public static string VuAttack => Get("VuAttack");
        public static string VuAttackHint => Get("VuAttackHint");
        public static string VuDecay => Get("VuDecay");
        public static string VuDecayHint => Get("VuDecayHint");

        // ── Details view columns and shared buttons ─────────────────
        public static string ColName => Get("ColName");
        public static string ColDateModified => Get("ColDateModified");
        public static string ColType => Get("ColType");
        public static string ColSize => Get("ColSize");
        public static string MenuResetColumns => Get("MenuResetColumns");
        public static string FocusFrameTitle => Get("FocusFrameTitle");
        public static string BtnClose => Get("BtnClose");
        public static string BtnBrowse => Get("BtnBrowse");
        public static string AutomationClickTarget => Get("AutomationClickTarget");
        public static string ProfileActive => Get("ProfileActive");
        public static string LblFontSize => Get("LblFontSize");
        public static string CalcErrorDisplay => Get("CalcErrorDisplay");

        // ── Options window: tabs, sections and checkboxes ────────────
        public static string TabGeneral => Get("TabGeneral");
        public static string TabStyleFx => Get("TabStyleFx");
        public static string TabTools => Get("TabTools");
        public static string TabProfiles => Get("TabProfiles");
        public static string TabHotkeys => Get("TabHotkeys");
        public static string TabSmartDesktop => Get("TabSmartDesktop");
        public static string TabLookDeeper => Get("TabLookDeeper");
        public static string SecStartup => Get("SecStartup");
        public static string SecSelections => Get("SecSelections");
        public static string SecAppearance => Get("SecAppearance");
        public static string SecIcons => Get("SecIcons");
        public static string SecUtilities => Get("SecUtilities");
        public static string SecMaintenance => Get("SecMaintenance");
        public static string SecLog => Get("SecLog");
        public static string SecLogCategories => Get("SecLogCategories");
        public static string SecLogConfiguration => Get("SecLogConfiguration");
        public static string SecProfileManagement => Get("SecProfileManagement");
        public static string SecProfileSwitching => Get("SecProfileSwitching");
        public static string SecSmartDesktopAuto => Get("SecSmartDesktopAuto");
        public static string SecDesktopIconVisibility => Get("SecDesktopIconVisibility");
        public static string SecIdleFadeOut => Get("SecIdleFadeOut");
        public static string SecIdleAutoRoll => Get("SecIdleAutoRoll");
        public static string SecAutoHideFrames => Get("SecAutoHideFrames");
        public static string OptStartWithWindows => Get("OptStartWithWindows");
        public static string OptSingleClick => Get("OptSingleClick");
        public static string OptSnapNearFrames => Get("OptSnapNearFrames");
        public static string OptDimensionSnap => Get("OptDimensionSnap");
        public static string OptTrayIcon => Get("OptTrayIcon");
        public static string OptRecycleBin => Get("OptRecycleBin");
        public static string OptNewFrameContextMenu => Get("OptNewFrameContextMenu");
        public static string OptPortalWatermark => Get("OptPortalWatermark");
        public static string OptNoteWatermark => Get("OptNoteWatermark");
        public static string OptDisableScrollbars => Get("OptDisableScrollbars");
        public static string OptEnableSounds => Get("OptEnableSounds");
        public static string OptChameleon => Get("OptChameleon");
        public static string OptFrameTint => Get("OptFrameTint");
        public static string OptIdleFadeOut => Get("OptIdleFadeOut");
        public static string OptAutoHideFrames => Get("OptAutoHideFrames");
        public static string OptHideIconsRunning => Get("OptHideIconsRunning");
        public static string OptHideIconsWhenHidden => Get("OptHideIconsWhenHidden");
        public static string OptAutomaticBackup => Get("OptAutomaticBackup");
        public static string OptEnableLogging => Get("OptEnableLogging");
        public static string OptProfileHotkeys => Get("OptProfileHotkeys");
        public static string OptSpotSearchHotkey => Get("OptSpotSearchHotkey");
        public static string OptFocusFrameHotkey => Get("OptFocusFrameHotkey");
        public static string OptAutoOrganize => Get("OptAutoOrganize");
        public static string OptExecutionToasts => Get("OptExecutionToasts");
        public static string BtnReset => Get("BtnReset");

        // ── Options: buttons, sliders and hotkey rows ─────────────────
        public static string BtnBackup => Get("BtnBackup");
        public static string BtnRestore => Get("BtnRestore");
        public static string BtnOpenBackupsFolder => Get("BtnOpenBackupsFolder");
        public static string BtnScreenBoundFrames => Get("BtnScreenBoundFrames");
        public static string BtnResetStyles => Get("BtnResetStyles");
        public static string BtnClearAllData => Get("BtnClearAllData");
        public static string BtnManageProfiles => Get("BtnManageProfiles");
        public static string BtnManageAutomation => Get("BtnManageAutomation");
        public static string BtnSmartDesktopRules => Get("BtnSmartDesktopRules");
        public static string BtnOrganizeNow => Get("BtnOrganizeNow");
        public static string BtnOpenLog => Get("BtnOpenLog");
        public static string SldFrameTint => Get("SldFrameTint");
        public static string SldMenuTint => Get("SldMenuTint");
        public static string SldAutoHideTime => Get("SldAutoHideTime");
        public static string SldIdleTime => Get("SldIdleTime");
        public static string SldFadeTargetOpacity => Get("SldFadeTargetOpacity");
        public static string HkDirectProfile => Get("HkDirectProfile");
        public static string HkPreviousProfile => Get("HkPreviousProfile");
        public static string HkNextProfile => Get("HkNextProfile");
        public static string HkFocusFrame => Get("HkFocusFrame");
        public static string HkSpotSearch => Get("HkSpotSearch");
        public static string LblTotalRules => Get("LblTotalRules");
        public static string LblEnabledRules => Get("LblEnabledRules");

        // ── Dropdown entries whose value travels by index, not by text ──
        public static string SndDefault => Get("SndDefault");
        public static string SndDoubleDing => Get("SndDoubleDing");
        public static string SndSmoothTickle => Get("SndSmoothTickle");
        public static string SndMessageDing => Get("SndMessageDing");
        public static string SndGentleDing => Get("SndGentleDing");
        public static string SndSoftDing => Get("SndSoftDing");
        public static string FxZoom => Get("FxZoom");
        public static string FxBounce => Get("FxBounce");
        public static string FxFadeOut => Get("FxFadeOut");
        public static string FxSlideUp => Get("FxSlideUp");
        public static string FxRotate => Get("FxRotate");
        public static string FxAgitate => Get("FxAgitate");
        public static string FxGrowAndFly => Get("FxGrowAndFly");
        public static string FxPulse => Get("FxPulse");
        public static string FxElastic => Get("FxElastic");
        public static string FxFlip3D => Get("FxFlip3D");
        public static string FxSpiral => Get("FxSpiral");
        public static string FxShockwave => Get("FxShockwave");
        public static string FxMatrix => Get("FxMatrix");
        public static string FxSupernova => Get("FxSupernova");
        public static string FxTeleport => Get("FxTeleport");

        // ── Dropdown entries shown translated, stored in English ───────
        public static string ViewIcons => Get("ViewIcons");
        public static string ViewDetails => Get("ViewDetails");
        public static string ColorGray => Get("ColorGray");
        public static string ColorBlack => Get("ColorBlack");
        public static string ColorWhite => Get("ColorWhite");
        public static string ColorBeige => Get("ColorBeige");
        public static string ColorGreen => Get("ColorGreen");
        public static string ColorPurple => Get("ColorPurple");
        public static string ColorFuchsia => Get("ColorFuchsia");
        public static string ColorYellow => Get("ColorYellow");
        public static string ColorOrange => Get("ColorOrange");
        public static string ColorRed => Get("ColorRed");
        public static string ColorBlue => Get("ColorBlue");
        public static string ColorBismark => Get("ColorBismark");

        // ── Tray menu (WinForms) ──────────────────────────────────
        public static string TrayProfiles => Get("TrayProfiles");
        public static string TrayReloadAllFrames => Get("TrayReloadAllFrames");
        public static string TrayShowHiddenFrames => Get("TrayShowHiddenFrames");
        public static string TrayFocusFrame => Get("TrayFocusFrame");
        public static string TrayCreateNewProfile => Get("TrayCreateNewProfile");
        public static string TrayManageProfiles => Get("TrayManageProfiles");

        // ── About panel ──────────────────────────────────────────
        public static string AboutVersion => Get("AboutVersion");
        public static string AboutSectionAbout => Get("AboutSectionAbout");
        public static string AboutTagline => Get("AboutTagline");
        public static string AboutBody => Get("AboutBody");
        public static string AboutSectionCredits => Get("AboutSectionCredits");
        public static string AboutCreditsBody => Get("AboutCreditsBody");
        public static string AboutSoundCredits => Get("AboutSoundCredits");
        public static string AboutAnd => Get("AboutAnd");
        public static string BtnMoveUp => Get("BtnMoveUp");
        public static string BtnMoveDown => Get("BtnMoveDown");
        public static string NetStatusUp => Get("NetStatusUp");
        public static string NetStatusDown => Get("NetStatusDown");
        public static string LblProcessName => Get("LblProcessName");
        public static string LblTargetProfile => Get("LblTargetProfile");
        public static string LblActivationDelay => Get("LblActivationDelay");

        // ── Found by sweeping every function that takes a string ───────
        public static string PluginPhotoFrame => Get("PluginPhotoFrame");
        public static string PluginCalculator => Get("PluginCalculator");
        public static string PluginVuMeter => Get("PluginVuMeter");
        public static string PluginIpInfo => Get("PluginIpInfo");
        public static string PluginPerformance => Get("PluginPerformance");
        public static string PluginTerminal => Get("PluginTerminal");
        public static string PluginQueueSaturation => Get("PluginQueueSaturation");
        public static string GaugeCpu => Get("GaugeCpu");
        public static string GaugeRam => Get("GaugeRam");
        public static string GaugeLeftL => Get("GaugeLeftL");
        public static string GaugeRightR => Get("GaugeRightR");
        public static string GaugeMasterVu => Get("GaugeMasterVu");
        public static string GaugeLeftVu => Get("GaugeLeftVu");
        public static string GaugeRightVu => Get("GaugeRightVu");
        public static string DiagCpuUtilization => Get("DiagCpuUtilization");
        public static string DiagRamUtilization => Get("DiagRamUtilization");
        public static string DiagCpuQueue => Get("DiagCpuQueue");
        public static string DiagDiskIo => Get("DiagDiskIo");
        public static string LblFontFamily => Get("LblFontFamily");
        public static string LblWordWrap => Get("LblWordWrap");
        public static string LblSpellCheck => Get("LblSpellCheck");
        public static string LblDisplayName => Get("LblDisplayName");
        public static string LblArguments => Get("LblArguments");
        public static string LblTargetPath => Get("LblTargetPath");
        public static string TermRestarting => Get("TermRestarting");
        public static string TermProcessExited => Get("TermProcessExited");
        public static string NetOffline => Get("NetOffline");
        public static string NetUnreachable => Get("NetUnreachable");
        public static string NetChecking => Get("NetChecking");
        public static string MoveMainItems => Get("MoveMainItems");
        public static string CalcEntryCleared => Get("CalcEntryCleared");
        public static string CalcPercentage => Get("CalcPercentage");
        public static string CalcCleared => Get("CalcCleared");
        public static string CalcClearedAll => Get("CalcClearedAll");
        public static string RandomNameAdjectives => Get("RandomNameAdjectives");
        public static string RandomNamePlaces => Get("RandomNamePlaces");
        public static string RandomNameFormat => Get("RandomNameFormat");
    }
}

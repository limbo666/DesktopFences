using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{
    public enum IconVisibilityEffect
    {
        None, Glow, Shadow, Outline, AngelGlow, ColoredGlow, StrongShadow
    }

    /// <summary>
    /// Manages application settings with Strict "Hard Switch" Master support.
    /// </summary>
    public static class SettingsManager
    {
        // --- Properties ---
        public static bool EnableAutoBackup { get; set; } = true;
        public static DateTime LastAutoBackupDate { get; set; } = DateTime.MinValue;
        public static bool ShowPortalExtensions { get; set; } = false;
        public static bool NoWildcardsOnPortalFilter { get; set; } = false;
        public static bool IsSnapEnabled { get; set; } = true;
        public static bool ShowBackgroundImageOnPortalFences { get; set; } = true;
        public static bool UseRecycleBin { get; set; } = true;
        public static bool ShowInTray { get; set; } = true;
        public static bool EnableSounds { get; set; } = true;
        public static int TintValue { get; set; } = 85;
        public static int MenuTintValue { get; set; } = 30;
        public static int MenuIcon { get; set; } = 0;
        public static int LockIcon { get; set; } = 0;
        public static string SelectedColor { get; set; } = "Gray";
        public static bool IsLogEnabled { get; set; } = false;
        public static int MaxDisplayNameLength { get; set; } = 20;
        public static int PortalBackgroundOpacity { get; set; } = 30;
        public static bool EnableIconGlowEffect { get; set; } = true;
        public static bool DisableSingleInstance { get; set; } = false;
        public static bool DeletePreviousLogOnStart { get; set; } = false;
        public static bool EnableBackgroundValidationLogging { get; set; } = false;
        public static bool SuppressLaunchWarnings { get; set; } = false;
        public static bool DisableFenceScrollbars { get; set; } = false;
        public static bool DisableNoteAutoSave { get; set; } = false;

        public static bool EnableChameleonMode { get; set; } = false;
        public static bool EnableProfileAutomation { get; set; } = false;

        public static bool EnableAutoOrganize { get; set; } = false;

        public static bool EnableAutoOrganizeNotifications { get; set; } = true;

        // --- NEW: Hidden Option for Manual Repositioning ---
        public static bool AllowAutoReposition { get; set; } = true;

        // --- NEW: Context Menu Option ---
        public static bool EnableContextMenu { get; set; } = false;

        // --- NEW: Auto-Hide Fences Options ---
        public static bool AutoHideFences { get; set; } = false;
        public static int AutoHideTime { get; set; } = 60;
        public static bool AutoResetHideTimer { get; set; } = true;
        public static bool HideFlashEffect { get; set; } = true;


        public static IconVisibilityEffect IconVisibilityEffect { get; set; } = IconVisibilityEffect.None;
        public static bool ExportShortcutsOnFenceDeletion { get; set; } = false;
        public static bool DeleteOriginalShortcutsOnDrop { get; set; } = false;
        public static bool EnableSpotSearchHotkey { get; set; } = true;
        public static bool EnableProfileHotkeys { get; set; } = true;
        public static bool EnableFocusFenceHotkey { get; set; } = true;
        public static int SpotSearchKey { get; set; } = 192;

        public static string SpotSearchModifier { get; set; } = "Control";
        public static bool EnableDimensionSnap { get; set; } = false;
        public static bool SingleClickToLaunch { get; set; } = true;
        public static LaunchEffectsManager.LaunchEffect LaunchEffect { get; set; } = LaunchEffectsManager.LaunchEffect.Zoom;
        public static LogManager.LogLevel MinLogLevel { get; set; } = LogManager.LogLevel.Info;
        public static List<LogManager.LogCategory> EnabledLogCategories { get; set; } = new List<LogManager.LogCategory>
        {
            LogManager.LogCategory.General,
            LogManager.LogCategory.Error,
            LogManager.LogCategory.ImportExport,
            LogManager.LogCategory.Settings
        };

        private static string _activeOptionsPath;

        public static void LoadSettings()
        {
            // 1. DETERMINE SOURCE
            string appRoot = AppDomain.CurrentDomain.BaseDirectory;
            string masterPath = Path.Combine(appRoot, "MasterOptions.json");
            string localPath = ProfileManager.GetProfileFilePath("options.json");

            if (File.Exists(masterPath))
            {
                _activeOptionsPath = masterPath;
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings, "MasterOptions.json found. Switched to Global Configuration Mode.");
            }
            else
            {
                _activeOptionsPath = localPath;
            }

            // 2. READ DATA
            try
            {
                if (File.Exists(_activeOptionsPath))
                {
                    string jsonContent = File.ReadAllText(_activeOptionsPath);
                    if (!string.IsNullOrWhiteSpace(jsonContent))
                    {
                        try
                        {
                            var optionsData = JsonConvert.DeserializeObject<dynamic>(jsonContent);
                            if (optionsData != null) ApplyJsonToProperties(optionsData);
                        }
                        catch { /* Corrupt file, defaults will apply */ }
                    }
                }

                // --- FIX: Universal File Hydration ---
                // By unconditionally saving after loading the properties into memory, we guarantee:
                // 1. 0-byte profile files are instantly populated with default JSON.
                // 2. MasterOptions.json / options.json files from older app versions 
                //    automatically get newly added settings injected into them.
                SaveSettings();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading settings: {ex.Message}");
                if (string.IsNullOrEmpty(_activeOptionsPath)) _activeOptionsPath = localPath;
                SaveSettings();
            }
        }

        public static void SaveSettings()
        {
            if (string.IsNullOrEmpty(_activeOptionsPath))
            {
                _activeOptionsPath = ProfileManager.GetProfileFilePath("options.json");
            }

            try
            {
                // 3. WRITE TO THE ACTIVE SOURCE (Includes AllowAutoReposition now)
                var optionsData = GetCurrentPropertiesAsObject();
                string formattedJson = JsonConvert.SerializeObject(optionsData, Formatting.Indented);
                File.WriteAllText(_activeOptionsPath, formattedJson);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"Failed to save settings to {_activeOptionsPath}: {ex.Message}");
            }
        }

        // --- Helpers ---

        private static object GetCurrentPropertiesAsObject()
        {
            return new
            {
                IsSnapEnabled,
                ShowBackgroundImageOnPortalFences,
                ShowInTray,
                EnableSounds,
                UseRecycleBin,
                TintValue,
                MenuTintValue,
                MenuIcon,
                LockIcon,
                SelectedColor,
                IsLogEnabled,
                SingleClickToLaunch,
                EnableDimensionSnap,
                PortalBackgroundOpacity,
                MaxDisplayNameLength,
                IconVisibilityEffect = IconVisibilityEffect.ToString(),
                LaunchEffect = LaunchEffect.ToString(),
                MinLogLevel = MinLogLevel.ToString(),
                EnabledLogCategories = EnabledLogCategories.Select(c => c.ToString()).ToList(),
                DeletePreviousLogOnStart,
                SuppressLaunchWarnings,
                EnableBackgroundValidationLogging,
                DisableSingleInstance,
                DisableFenceScrollbars,
                DisableNoteAutoSave,
                ExportShortcutsOnFenceDeletion,
                DeleteOriginalShortcutsOnDrop,
                EnableSpotSearchHotkey,
                SpotSearchKey,
                SpotSearchModifier,
                NoWildcardsOnPortalFilter,
                ShowPortalExtensions,
                EnableAutoBackup,
                LastAutoBackupDate,

                AllowAutoReposition,
                EnableProfileAutomation,
                EnableChameleonMode, // ADD THIS HERE
                EnableAutoOrganize,
                EnableAutoOrganizeNotifications,
                // NEW
                EnableContextMenu,

                // Auto-Hide
                AutoHideFences,
                AutoHideTime,
                AutoResetHideTimer,
                HideFlashEffect,
                // Global Hotkeys
                EnableProfileHotkeys,
                EnableFocusFenceHotkey,
                ProfileSwitchModifier,
                ProfileSwitchKeys,
                ProfilePrevModifier,
                ProfilePrevKey,
                ProfileNextModifier,
                ProfileNextKey,
                FocusFenceModifier,
                FocusFenceKey
            };
        }

        private static void ApplyJsonToProperties(dynamic data)
        {
            try { EnableAutoBackup = data.EnableAutoBackup ?? false; } catch { EnableAutoBackup = false; }
            try { LastAutoBackupDate = data.LastAutoBackupDate ?? DateTime.MinValue; } catch { LastAutoBackupDate = DateTime.MinValue; }
            try { IsSnapEnabled = data.IsSnapEnabled ?? true; } catch { IsSnapEnabled = true; }
            try { ShowBackgroundImageOnPortalFences = data.ShowBackgroundImageOnPortalFences ?? true; } catch { ShowBackgroundImageOnPortalFences = true; }
            try { ShowInTray = data.ShowInTray ?? true; } catch { ShowInTray = true; }
            try { EnableSounds = data.EnableSounds ?? true; } catch { EnableSounds = true; }
            try { UseRecycleBin = data.UseRecycleBin ?? true; } catch { UseRecycleBin = true; }
            try { TintValue = data.TintValue ?? 85; } catch { TintValue = 85; }
            try { MenuTintValue = data.MenuTintValue ?? 30; } catch { MenuTintValue = 30; }
            try { MenuIcon = data.MenuIcon ?? 0; } catch { MenuIcon = 0; }
            try { LockIcon = data.LockIcon ?? 0; } catch { LockIcon = 0; }
            try { SelectedColor = data.SelectedColor ?? "Gray"; } catch { SelectedColor = "Gray"; }
            try { IsLogEnabled = data.IsLogEnabled ?? false; } catch { IsLogEnabled = false; }
            try { SingleClickToLaunch = data.SingleClickToLaunch ?? true; } catch { SingleClickToLaunch = true; }
            try { EnableDimensionSnap = data.EnableDimensionSnap ?? false; } catch { EnableDimensionSnap = false; }
            try { PortalBackgroundOpacity = data.PortalBackgroundOpacity ?? 30; } catch { PortalBackgroundOpacity = 30; }
            try { DisableFenceScrollbars = data.DisableFenceScrollbars ?? false; } catch { DisableFenceScrollbars = false; }
            try { DisableNoteAutoSave = data.DisableNoteAutoSave ?? false; } catch { DisableNoteAutoSave = false; }
            try { ExportShortcutsOnFenceDeletion = data.ExportShortcutsOnFenceDeletion ?? false; } catch { ExportShortcutsOnFenceDeletion = false; }
            try { DeleteOriginalShortcutsOnDrop = data.DeleteOriginalShortcutsOnDrop ?? false; } catch { DeleteOriginalShortcutsOnDrop = false; }
            try { EnableSpotSearchHotkey = data.EnableSpotSearchHotkey ?? true; } catch { EnableSpotSearchHotkey = true; }
            try { SpotSearchModifier = data.SpotSearchModifier?.ToString() ?? "Control"; } catch { SpotSearchModifier = "Control"; }
            try { ShowPortalExtensions = data.ShowPortalExtensions ?? false; } catch { ShowPortalExtensions = false; }
            try { NoWildcardsOnPortalFilter = data.NoWildcardsOnPortalFilter ?? false; } catch { NoWildcardsOnPortalFilter = false; }

            try { AllowAutoReposition = data.AllowAutoReposition ?? true; } catch { AllowAutoReposition = true; }
            try { EnableProfileAutomation = data.EnableProfileAutomation ?? false; } catch { EnableProfileAutomation = false; }
            try { EnableChameleonMode = data.EnableChameleonMode ?? false; } catch { EnableChameleonMode = false; } // 
            try { EnableAutoOrganize = data.EnableAutoOrganize ?? false; } catch { EnableAutoOrganize = false; }
            try { EnableAutoOrganizeNotifications = data.EnableAutoOrganizeNotifications ?? true; } catch { EnableAutoOrganizeNotifications = true; }

            // NEW
            try { EnableContextMenu = data.EnableContextMenu ?? false; } catch { EnableContextMenu = false; }

            // Auto-Hide
            try { AutoHideFences = data.AutoHideFences ?? false; } catch { AutoHideFences = false; }
            try { AutoHideTime = data.AutoHideTime ?? 60; } catch { AutoHideTime = 60; }
            try { AutoResetHideTimer = data.AutoResetHideTimer ?? true; } catch { AutoResetHideTimer = true; }
            try { HideFlashEffect = data.HideFlashEffect ?? true; } catch { HideFlashEffect = true; }

            try { SpotSearchKey = ParseKey(data.SpotSearchKey); } catch { SpotSearchKey = 192; }

            try
            {
                int value = data.MaxDisplayNameLength ?? 20;
                MaxDisplayNameLength = Math.Max(5, Math.Min(50, value));
            }
            catch { MaxDisplayNameLength = 20; }

            try { DisableSingleInstance = data.DisableSingleInstance ?? false; } catch { DisableSingleInstance = false; }

            try
            {
                string effectName = data.IconVisibilityEffect?.ToString() ?? "None";
                if (Enum.TryParse<IconVisibilityEffect>(effectName, true, out IconVisibilityEffect parsedEffect))
                    IconVisibilityEffect = parsedEffect;
                else
                    IconVisibilityEffect = IconVisibilityEffect.None;
            }
            catch { IconVisibilityEffect = IconVisibilityEffect.None; }

            try
            {
                LaunchEffect = data.LaunchEffect != null
                    ? Enum.Parse(typeof(LaunchEffectsManager.LaunchEffect), data.LaunchEffect.ToString())
                    : LaunchEffectsManager.LaunchEffect.Zoom;
            }
            catch { LaunchEffect = LaunchEffectsManager.LaunchEffect.Zoom; }

            try { DeletePreviousLogOnStart = data.DeletePreviousLogOnStart ?? false; } catch { DeletePreviousLogOnStart = false; }
            try { SuppressLaunchWarnings = data.SuppressLaunchWarnings ?? false; } catch { SuppressLaunchWarnings = false; }
            try { EnableBackgroundValidationLogging = data.EnableBackgroundValidationLogging ?? false; } catch { EnableBackgroundValidationLogging = false; }

            try
            {
                MinLogLevel = data.MinLogLevel != null
                    ? Enum.Parse(typeof(LogManager.LogLevel), data.MinLogLevel.ToString())
                    : LogManager.LogLevel.Info;
            }
            catch { MinLogLevel = LogManager.LogLevel.Info; }

            try
            {
                EnabledLogCategories = data.EnabledLogCategories != null
                    ? ((JArray)data.EnabledLogCategories)
                        .Select(c => Enum.Parse(typeof(LogManager.LogCategory), c.ToString()))
                        .Cast<LogManager.LogCategory>()
                        .ToList()
                    : new List<LogManager.LogCategory> { LogManager.LogCategory.General, LogManager.LogCategory.Error, LogManager.LogCategory.ImportExport, LogManager.LogCategory.Settings };
            }
            catch
            {
                EnabledLogCategories = new List<LogManager.LogCategory> { LogManager.LogCategory.General, LogManager.LogCategory.Error, LogManager.LogCategory.ImportExport, LogManager.LogCategory.Settings };
            }

        
            // Global Hotkeys
            try { EnableProfileHotkeys = data.EnableProfileHotkeys ?? true; } catch { EnableProfileHotkeys = true; }
            try { EnableFocusFenceHotkey = data.EnableFocusFenceHotkey ?? true; } catch { EnableFocusFenceHotkey = true; }
            try { if (data.ProfileSwitchModifier != null) ProfileSwitchModifier = data.ProfileSwitchModifier.ToString(); } catch { }
            try { if (data.ProfileSwitchKeys != null) ProfileSwitchKeys = ((JArray)data.ProfileSwitchKeys).Select(x => (int)x).ToArray(); } catch { }
            try { if (data.ProfilePrevModifier != null) ProfilePrevModifier = data.ProfilePrevModifier.ToString(); } catch { }
            try { if (data.ProfilePrevKey != null) ProfilePrevKey = (int)data.ProfilePrevKey; } catch { }
            try { if (data.ProfileNextModifier != null) ProfileNextModifier = data.ProfileNextModifier.ToString(); } catch { }
            try { if (data.ProfileNextKey != null) ProfileNextKey = (int)data.ProfileNextKey; } catch { }
            try { if (data.FocusFenceModifier != null) FocusFenceModifier = data.FocusFenceModifier.ToString(); } catch { }
            try { if (data.FocusFenceKey != null) FocusFenceKey = (int)data.FocusFenceKey; } catch { }

            SanitizeHotkeys(); // Ensure nulls or invalid manual edits are safely overwritten
        }

        private static int ParseKey(dynamic value)
        {
            if (value == null) return 192;
            if (int.TryParse(value.ToString(), out int code)) return code;
            string keyName = value.ToString().ToLower().Trim();
            return keyName switch { "~" => 192, "tilde" => 192, "space" => 32, "q" => 81, "f1" => 112, _ => 192 };
        }

        public static void SetMinLogLevel(LogManager.LogLevel level)
        {
            MinLogLevel = level;
            SaveSettings();
        }

        public static void SetEnabledLogCategories(List<LogManager.LogCategory> categories)
        {
            EnabledLogCategories = categories;
            SaveSettings();
        }

        #region Global Hotkey Configurations
        public static string ProfileSwitchModifier { get; set; } = "Control, Alt";
        public static int[] ProfileSwitchKeys { get; set; } = new int[] { 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39 }; // Defaults to standard 0-9

        public static string ProfilePrevModifier { get; set; } = "Control, Alt";
        public static int ProfilePrevKey { get; set; } = 0xBC; // Default: VK_OEM_COMMA

        public static string ProfileNextModifier { get; set; } = "Control, Alt";
        public static int ProfileNextKey { get; set; } = 0xBE; // Default: VK_OEM_PERIOD

        public static string FocusFenceModifier { get; set; } = "Control, Alt";
        public static int FocusFenceKey { get; set; } = 0x5A; // Default: VK_Z

        /// <summary>
        /// Validates and repairs hotkey configuration to prevent hook crashes from manual JSON edits.
        /// </summary>
        public static void SanitizeHotkeys()
        {
            // Ensure modifiers aren't completely blank or null
            if (string.IsNullOrWhiteSpace(ProfileSwitchModifier)) ProfileSwitchModifier = "Control, Alt";
            if (string.IsNullOrWhiteSpace(ProfilePrevModifier)) ProfilePrevModifier = "Control, Alt";
            if (string.IsNullOrWhiteSpace(ProfileNextModifier)) ProfileNextModifier = "Control, Alt";
            if (string.IsNullOrWhiteSpace(FocusFenceModifier)) FocusFenceModifier = "Control, Alt";

            // Fallback for missing or broken Profile Switch Array
            if (ProfileSwitchKeys == null || ProfileSwitchKeys.Length < 10)
                ProfileSwitchKeys = new int[] { 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39 };

            // Fallback for totally invalid Virtual Key codes (must be within 0x01 and 0xFE)
            if (ProfilePrevKey <= 0 || ProfilePrevKey > 254) ProfilePrevKey = 0xBC;
            if (ProfileNextKey <= 0 || ProfileNextKey > 254) ProfileNextKey = 0xBE;
            if (FocusFenceKey <= 0 || FocusFenceKey > 254) FocusFenceKey = 0x5A;
        }

        /// <summary>
        /// Injects the current hotkey settings into every available profile's options.json.
        /// This ensures uniform hotkey behavior regardless of the active profile, while respecting MasterOptions override.
        /// </summary>
        public static void BroadcastHotkeysToAllProfiles()
        {
            try
            {
                string appRoot = AppDomain.CurrentDomain.BaseDirectory;
                string profilesDir = Path.Combine(appRoot, "Profiles");

                if (!Directory.Exists(profilesDir)) return;

                foreach (string dir in Directory.GetDirectories(profilesDir))
                {
                    string optionsFile = Path.Combine(dir, "options.json");
                    if (File.Exists(optionsFile))
                    {
                        try
                        {
                            string jsonContent = File.ReadAllText(optionsFile);
                            JObject data = JsonConvert.DeserializeObject<JObject>(jsonContent);
                            if (data == null) continue;

                            // Overwrite strictly the hotkey values
                            data["ProfileSwitchModifier"] = ProfileSwitchModifier;
                            data["ProfileSwitchKeys"] = JArray.FromObject(ProfileSwitchKeys ?? new int[] { 0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39 });
                            data["ProfilePrevModifier"] = ProfilePrevModifier;
                            data["ProfilePrevKey"] = ProfilePrevKey;
                            data["ProfileNextModifier"] = ProfileNextModifier;
                            data["ProfileNextKey"] = ProfileNextKey;
                            data["FocusFenceModifier"] = FocusFenceModifier;
                            data["FocusFenceKey"] = FocusFenceKey;
                            data["SpotSearchModifier"] = SpotSearchModifier;
                            data["SpotSearchKey"] = SpotSearchKey;
                            data["EnableProfileHotkeys"] = EnableProfileHotkeys;
                            data["EnableFocusFenceHotkey"] = EnableFocusFenceHotkey;
                            data["EnableSpotSearchHotkey"] = EnableSpotSearchHotkey;

                            File.WriteAllText(optionsFile, JsonConvert.SerializeObject(data, Formatting.Indented));
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"Failed to broadcast hotkeys to {optionsFile}: {ex.Message}");
                        }
                    }
                }
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings, "Global hotkeys successfully broadcasted to all individual profiles.");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"Critical failure broadcasting hotkeys: {ex.Message}");
            }
        }
        #endregion

    }


}
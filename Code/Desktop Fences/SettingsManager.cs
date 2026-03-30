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


        public static bool EnableProfileAutomation { get; set; } = false;
        // --- NEW: Hidden Option for Manual Repositioning ---
        public static bool AllowAutoReposition { get; set; } = true;

        // --- NEW: Context Menu Option ---
        public static bool EnableContextMenu { get; set; } = false;


        public static IconVisibilityEffect IconVisibilityEffect { get; set; } = IconVisibilityEffect.None;
        public static bool ExportShortcutsOnFenceDeletion { get; set; } = false;
        public static bool DeleteOriginalShortcutsOnDrop { get; set; } = false;
        public static bool EnableSpotSearchHotkey { get; set; } = true;
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
                else
                {
                    SaveSettings();
                }
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
                // NEW
                EnableContextMenu
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

            // NEW
            try { EnableContextMenu = data.EnableContextMenu ?? false; } catch { EnableContextMenu = false; }

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
    }
}
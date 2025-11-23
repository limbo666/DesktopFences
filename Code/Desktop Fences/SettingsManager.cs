using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms.VisualStyles;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{


    public enum IconVisibilityEffect
    {
        None, Glow, Shadow, Outline, AngelGlow, ColoredGlow, StrongShadow
    }


    /// <summary>
    /// Manages the application's settings, including snapping, tint, color, and logging preferences.
    /// Loads settings from and saves them to a JSON file (options.json) in the application directory.
    /// </summary>
    public static class SettingsManager
    {
        /// <summary>
        /// Gets or sets whether snapping is enabled for fence alignment.
        /// </summary>
        public static bool IsSnapEnabled { get; set; } = true;

        /// <summary>
        /// Gets or sets whether the background image should be displayed on portal fences.
        /// </summary>
        public static bool ShowBackgroundImageOnPortalFences { get; set; } = true;
        /// <summary>
        /// Gets or sets whether the portal fence deletion method will use the recycle bin.
        /// </summary>
        public static bool UseRecycleBin { get; set; } = true;

        /// <summary>
        /// Gets or sets whether the tray icon will be shown.
        /// </summary>
        public static bool ShowInTray { get; set; } = true;


        /// <summary>
        /// Gets or sets whether the sounds will be enabled.
        /// </summary>
        public static bool EnableSounds { get; set; } = true;


        /// <summary>
        /// Gets or sets the tint value (0-100) that controls fence transparency.
        /// </summary>
        public static int TintValue { get; set; } = 60;


        /// <summary>
        /// Gets or sets the tint value (0-100) that controls menu icons transparency.
        /// </summary>
        public static int MenuTintValue { get; set; } = 30;


        /// <summary>
        /// Gets or sets the tint value (0-3) that controls the menu icon.
        /// </summary>
        public static int MenuIcon { get; set; } = 0;



        /// <summary>
        /// Gets or sets the tint value (0-3) that controls the menu icon.
        /// </summary>
        public static int LockIcon { get; set; } = 0;


        /// <summary>
        /// Gets or sets the selected color name for fence backgrounds.
        /// </summary>
        public static string SelectedColor { get; set; } = "Gray";

        /// <summary>
        /// Gets or sets whether logging is enabled for snapping events.
        /// </summary>
        public static bool IsLogEnabled { get; set; } = false;


        /// <summary>
        /// Gets or sets the maximum number of characters to display before truncating with "...".
        /// Valid range: 5-50 characters.
        /// </summary>
        public static int MaxDisplayNameLength { get; set; } = 20; // Default to current behavior

        /// <summary>
        /// Gets or sets the opacity of the portal fence background image (0-100).
        /// </summary>
        public static int PortalBackgroundOpacity { get; set; } = 0; // Default to existing opacity

        /// <summary>
        /// Gets or sets whether icon glow effect is enabled for better visibility.
        /// </summary>
        public static bool EnableIconGlowEffect { get; set; } = true; // Default disabled

        /// <summary>
        /// Gets or sets whether to disable single instance enforcement, allowing multiple instances to run simultaneously.
        /// Hidden option - not exposed in GUI, manually configurable in options.json
        /// When true, multiple instances can run without triggering effects or forced exits
        /// Default: false (single instance enforced)
        /// </summary>
        public static bool DisableSingleInstance { get; set; } = false; // Default disabled (single instance enforced)


        /// <summary>
        /// Gets or sets whether to delete the previous log file on application startup.
        /// </summary>
        public static bool DeletePreviousLogOnStart { get; set; } = false; // Default disabled

        /// <summary>
        /// Gets or sets whether to enable background validation logging (file existence checks, icon updates).
        /// When disabled, reduces log volume by suppressing repetitive validation operations.
        /// </summary>
        public static bool EnableBackgroundValidationLogging { get; set; } = false; // Default disabled to reduce log noise


        /// <summary>
        /// Gets or sets whether to suppress launch warning messages for applications that return unusual exit codes
        /// Useful for media players like PotPlayer that may trigger false error messages during successful launches
        /// Hidden option - not exposed in GUI, manually configurable in options.json
        /// </summary>
        public static bool SuppressLaunchWarnings { get; set; } = false; // Default disabled



        /// <summary>
        /// Gets or sets whether to disable automatic scrollbars on fence content areas
        /// When true, removes scrollbars from Data and Portal fences for a cleaner appearance
        /// Hidden option - not exposed in GUI, manually configurable in options.json
        /// </summary>
        public static bool DisableFenceScrollbars { get; set; } = false; // Default disabled


        /// <summary>
        /// Gets or sets whether to disable auto-save for note fences, requiring manual save via Done button.
        /// </summary>
        public static bool DisableNoteAutoSave { get; set; } = false;

        /// <summary>
        /// Gets or sets the icon visibility effect type.
        /// Valid values: None, Glow, Shadow, Outline, AngelGlow, ColoredGlow, StrongShadow
        /// </summary>

        public static IconVisibilityEffect IconVisibilityEffect { get; set; } = IconVisibilityEffect.None;

        /// <summary>
        /// Hidden Tweak: Automatically export icons to desktop when deleting a fence.
        /// </summary>
        public static bool ExportShortcutsOnFenceDeletion { get; set; } = false;

        /// <summary>
        /// Hidden Tweak: Delete the original .lnk file from Desktop after dropping it into a fence.
        /// </summary>
        public static bool DeleteOriginalShortcutsOnDrop { get; set; } = false;



        public static LogManager.LogLevel MinLogLevel { get; set; } = LogManager.LogLevel.Info;
        public static List<LogManager.LogCategory> EnabledLogCategories { get; set; } = new List<LogManager.LogCategory>
{
    LogManager.LogCategory.General,
    LogManager.LogCategory.Error,
    LogManager.LogCategory.ImportExport,
    LogManager.LogCategory.Settings
};

        // Default categories for production
        public static bool EnableDimensionSnap { get; set; } = false;

        public static bool SingleClickToLaunch { get; set; } = true;

        // Add LaunchEffect property
        public static LaunchEffectsManager.LaunchEffect LaunchEffect { get; set; } = LaunchEffectsManager.LaunchEffect.Zoom; // Default to Zoom





        public static void LoadSettings()
        {
            // Determine the path to options.json based on the application directory
            string optionsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "options.json");

            try
            {
                if (System.IO.File.Exists(optionsFilePath))
                {
                    // Read and validate JSON content
                    string jsonContent = File.ReadAllText(optionsFilePath);

                    // Validate JSON structure before parsing
                    if (string.IsNullOrWhiteSpace(jsonContent))
                    {
                        SaveSettings();
                        return;
                    }

                    dynamic optionsData;
                    try
                    {
                        optionsData = JsonConvert.DeserializeObject(jsonContent);
                        if (optionsData == null)
                        {
                            SaveSettings();
                            return;
                        }
                    }
                    catch (JsonException)
                    {
                        // JSON is corrupted, recreate with defaults
                        SaveSettings();
                        return;
                    }

                    // Load settings with individual property protection
                    try { IsSnapEnabled = optionsData.IsSnapEnabled ?? true; } catch { IsSnapEnabled = true; }
                    try { ShowBackgroundImageOnPortalFences = optionsData.ShowBackgroundImageOnPortalFences ?? true; } catch { ShowBackgroundImageOnPortalFences = true; }
                    try { ShowInTray = optionsData.ShowInTray ?? true; } catch { ShowInTray = true; }
                    try { EnableSounds = optionsData.EnableSounds ?? true; } catch { EnableSounds = true; }

                    try { UseRecycleBin = optionsData.UseRecycleBin ?? true; } catch { UseRecycleBin = true; }
                    try { TintValue = optionsData.TintValue ?? 60; } catch { TintValue = 60; }
                    try { MenuTintValue = optionsData.MenuTintValue ?? 30; } catch { MenuTintValue = 30; }
                    try { MenuIcon = optionsData.MenuIcon ?? 0; } catch { MenuIcon = 0; }
                    try { LockIcon = optionsData.LockIcon ?? 0; } catch { LockIcon = 0; }
                    try { SelectedColor = optionsData.SelectedColor ?? "Gray"; } catch { SelectedColor = "Gray"; }
                    try { IsLogEnabled = optionsData.IsLogEnabled ?? false; } catch { IsLogEnabled = false; }
                    try { SingleClickToLaunch = optionsData.SingleClickToLaunch ?? true; } catch { SingleClickToLaunch = true; }
                    try { EnableDimensionSnap = optionsData.EnableDimensionSnap ?? false; } catch { EnableDimensionSnap = false; }
                    try { PortalBackgroundOpacity = optionsData.PortalBackgroundOpacity ?? 20; } catch { PortalBackgroundOpacity = 20; }
                    try { DisableFenceScrollbars = optionsData.DisableFenceScrollbars ?? false; } catch { DisableFenceScrollbars = false; }
                    // Load DisableNoteAutoSave with protection
                    try { DisableNoteAutoSave = optionsData.DisableNoteAutoSave ?? false; } catch { DisableNoteAutoSave = false; }
               
                    try { ExportShortcutsOnFenceDeletion = optionsData.ExportShortcutsOnFenceDeletion ?? false; } catch { ExportShortcutsOnFenceDeletion = false; }
                    try { DeleteOriginalShortcutsOnDrop = optionsData.DeleteOriginalShortcutsOnDrop ?? false; } catch { DeleteOriginalShortcutsOnDrop = false; }
                    //   try { MaxDisplayNameLength = optionsData.MaxDisplayNameLength ?? 20; } catch { MaxDisplayNameLength = 20; }
                    try
                    {
                        int value = optionsData.MaxDisplayNameLength ?? 20;
                        MaxDisplayNameLength = Math.Max(5, Math.Min(50, value)); // Clamp between 5-50
                    }
                    catch
                    {
                        MaxDisplayNameLength = 20;
                    }

                    // Load DisableSingleInstance with protection  
                    try { DisableSingleInstance = optionsData.DisableSingleInstance ?? false; }
                    catch { DisableSingleInstance = false; }


                    try
                    {
                        string effectName = optionsData.IconVisibilityEffect?.ToString() ?? "None";

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, "effectName is: " + effectName);
                        if (Enum.TryParse<IconVisibilityEffect>(effectName, true, out IconVisibilityEffect parsedEffect))
                        {
                            IconVisibilityEffect = parsedEffect;
                        }
                        else
                        {
                            IconVisibilityEffect = IconVisibilityEffect.None;
                        }

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, "IconVisibilityEffect is: " + IconVisibilityEffect);
                    }
                    catch
                    {
                        IconVisibilityEffect = IconVisibilityEffect.None;
                    }


                    // Load LaunchEffect with protection
                    try
                    {
                        LaunchEffect = optionsData.LaunchEffect != null
                            ? Enum.Parse(typeof(LaunchEffectsManager.LaunchEffect), optionsData.LaunchEffect.ToString())
                            : LaunchEffectsManager.LaunchEffect.Zoom;
                    }
                    catch
                    {
                        LaunchEffect = LaunchEffectsManager.LaunchEffect.Zoom;
                    }

                    // Load DeletePreviousLogOnStart with protection
                    try { DeletePreviousLogOnStart = optionsData.DeletePreviousLogOnStart ?? false; }
                    catch { DeletePreviousLogOnStart = false; }

                    // Load SuppressLaunchWarnings with protection  
                    try { SuppressLaunchWarnings = optionsData.SuppressLaunchWarnings ?? false; }
                    catch { SuppressLaunchWarnings = false; }

                    // Load EnableBackgroundValidationLogging with protection  
                    try { EnableBackgroundValidationLogging = optionsData.EnableBackgroundValidationLogging ?? false; }
                    catch { EnableBackgroundValidationLogging = false; }

                    // Load MinLogLevel with protection
                    try
                    {
                        MinLogLevel = optionsData.MinLogLevel != null
                            ? Enum.Parse(typeof(LogManager.LogLevel), optionsData.MinLogLevel.ToString())
                            : LogManager.LogLevel.Info;
                    }
                    catch
                    {
                        MinLogLevel = LogManager.LogLevel.Info;
                    }

                    // Load EnabledLogCategories with protection
                    try
                    {
                        EnabledLogCategories = optionsData.EnabledLogCategories != null
                            ? ((JArray)optionsData.EnabledLogCategories)
                                .Select(c => Enum.Parse(typeof(LogManager.LogCategory), c.ToString()))
                                .Cast<LogManager.LogCategory>()
                                .ToList()
                            : new List<LogManager.LogCategory>
                            {
                        LogManager.LogCategory.General,
                        LogManager.LogCategory.Error,
                        LogManager.LogCategory.ImportExport,
                        LogManager.LogCategory.Settings
                            };
                    }
                    catch
                    {
                        EnabledLogCategories = new List<LogManager.LogCategory>
                {
                    LogManager.LogCategory.General,
                    LogManager.LogCategory.Error,
                    LogManager.LogCategory.ImportExport,
                    LogManager.LogCategory.Settings
                };
                    }
                }
                else
                {
                    // If the file doesn't exist, save the default settings
                    SaveSettings();
                }
            }
            catch (Exception ex)
            {
                // Handle potential errors (e.g., file corruption, access issues) by reverting to defaults
                Console.WriteLine($"Error loading settings: {ex.Message}");
                SaveSettings(); // Ensure a valid options.json exists
            }
        }



        public static void SaveSettings()
        {
            // Determine the path to options.json
            string optionsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "options.json");

            try
            {
                // Create an anonymous object with the current settings
                var optionsData = new
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
                    LaunchEffect = LaunchEffect.ToString(), // Save as string for JSON compatibility
                    MinLogLevel = MinLogLevel.ToString(), // Save as string for JSON compatibility
                    EnabledLogCategories = EnabledLogCategories.Select(c => c.ToString()).ToList(),
                    DeletePreviousLogOnStart,
                    SuppressLaunchWarnings,
                    EnableBackgroundValidationLogging, // 
                    DisableSingleInstance, // Multiple instances option  
                    DisableFenceScrollbars, // Scrollbar control option
                    DisableNoteAutoSave, // Note auto-save control option
                    ExportShortcutsOnFenceDeletion,
                    DeleteOriginalShortcutsOnDrop
                };

                // Serialize to JSON with indentation for readability
                string formattedJson = JsonConvert.SerializeObject(optionsData, Formatting.Indented);

                // Write the JSON to the file
                File.WriteAllText(optionsFilePath, formattedJson);
            }
            catch (Exception ex)
            {

            }
        }



        /// Sets the minimum log level and saves the settings.

        public static void SetMinLogLevel(LogManager.LogLevel level)
        {
            MinLogLevel = level;
            SaveSettings();
        }




        /// Sets the enabled log categories and saves the settings.

        public static void SetEnabledLogCategories(List<LogManager.LogCategory> categories)
        {
            EnabledLogCategories = categories;
            SaveSettings();
        }

  



    }
}
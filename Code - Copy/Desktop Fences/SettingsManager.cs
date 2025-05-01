using Newtonsoft.Json;
using System;
using System.IO;

namespace Desktop_Fences
{
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
        /// Gets or sets the tint value (0-100) that controls fence transparency.
        /// </summary>
        public static int TintValue { get; set; } = 60;

        /// <summary>
        /// Gets or sets the selected color name for fence backgrounds.
        /// </summary>
        public static string SelectedColor { get; set; } = "Gray";

        /// <summary>
        /// Gets or sets whether logging is enabled for snapping events.
        /// </summary>
        public static bool IsLogEnabled { get; set; } = false;

        /// <summary>
        /// Loads settings from options.json if it exists, otherwise saves and uses default settings.
        /// </summary>
        /// 

        public static bool SingleClickToLaunch { get; set; } = true;

        // Add LaunchEffect property
        public static FenceManager.LaunchEffect LaunchEffect { get; set; } = FenceManager.LaunchEffect.Zoom; // Default to Zoom

        public static void LoadSettings()
        {
            // Determine the path to options.json based on the application directory
            string optionsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "options.json");

            try
            {
                if (System.IO.File.Exists(optionsFilePath))
                {
                    // Read and deserialize the JSON content
                    string jsonContent = File.ReadAllText(optionsFilePath);
                    dynamic optionsData = JsonConvert.DeserializeObject(jsonContent);

                    // Update settings with values from the file, using defaults if values are missing
                    IsSnapEnabled = optionsData.IsSnapEnabled ?? true;
                    TintValue = optionsData.TintValue ?? 60;
                    SelectedColor = optionsData.SelectedColor ?? "Gray";
                    IsLogEnabled = optionsData.IsLogEnabled ?? false;
                    SingleClickToLaunch = optionsData.SingleClickToLaunch ?? true;
                    // Load LaunchEffect, parsing from string if present
                    LaunchEffect = optionsData.LaunchEffect != null
                        ? Enum.Parse(typeof(FenceManager.LaunchEffect), optionsData.LaunchEffect.ToString())
                        : FenceManager.LaunchEffect.Zoom;
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

        /// <summary>
        /// Saves the current settings to options.json in a formatted JSON structure.
        /// </summary>
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
                    TintValue,
                    SelectedColor,
                    IsLogEnabled,
                    SingleClickToLaunch,
                    LaunchEffect = LaunchEffect.ToString() // Save as string for JSON compatibility
                };

                // Serialize to JSON with indentation for readability
                string formattedJson = JsonConvert.SerializeObject(optionsData, Formatting.Indented);

                // Write the JSON to the file
                File.WriteAllText(optionsFilePath, formattedJson);
            }
            catch (Exception ex)
            {
                // Log any errors during save (console for simplicity; could be expanded to a file log)
                Console.WriteLine($"Error saving settings: {ex.Message}");
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{
    /// <summary>
    /// Manages all fence data operations including JSON loading, saving, and basic migrations
    /// Extracted from FenceManager for better code organization and maintainability
    /// Handles core data persistence and simple validation/migration scenarios
    /// </summary>
    public static class FenceDataManager
    {
        #region Private Fields - Data Storage
        // Main fence data collection - moved from FenceManager
        private static List<dynamic> _fenceData;

        // JSON file path - moved from FenceManager  
        private static string _jsonFilePath;
        #endregion

        #region Public Properties - Data Access
        /// <summary>
        /// Provides access to fence data collection
        /// Used by: FenceManager, TrayManager, CustomizeFenceForm, ItemMoveDialog
        /// Category: Data Access
        /// </summary>
        public static List<dynamic> FenceData
        {
            get => _fenceData;
            set => _fenceData = value;
        }

        /// <summary>
        /// Gets the JSON file path
        /// Used by: FenceManager, backup operations, debugging
        /// Category: File Management
        /// </summary>
        public static string JsonFilePath
        {
            get => _jsonFilePath;
            set => _jsonFilePath = value;
        }
        #endregion

        #region Initialization - Used by: FenceManager startup
        /// <summary>
        /// Initializes the data manager with JSON file path
        /// Used by: FenceManager during application startup
        /// Category: Initialization
        /// </summary>
        public static void Initialize()
        {
            try
            {
                // Set JSON file path relative to executable
                _jsonFilePath = Path.Combine(
                    Path.GetDirectoryName(Assembly.GetEntryAssembly().Location),
                    "fences.json");

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                    $"FenceDataManager initialized with path: {_jsonFilePath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error initializing FenceDataManager: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Core Data Operations - Used by: FenceManager, TrayManager
        /// <summary>
        /// Main fence data loading method - moved from FenceManager.LoadFenceData
        /// Used by: FenceManager.LoadFenceData, application startup
        /// Category: Data Loading
        /// </summary>
        public static void LoadFenceData(TargetChecker targetChecker)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    "FenceDataManager: Starting fence data load");

                // Initialize fence data list
                _fenceData = new List<dynamic>();

                // Check if JSON file exists
                if (!File.Exists(_jsonFilePath))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                        "JSON file not found. Starting with empty fence configuration.");
                    return;
                }

                // Load and parse JSON data
                if (!LoadFenceDataFromJson())
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                        "Failed to load JSON data. Starting with empty configuration.");
                    return;
                }

                // Apply simple migrations and validation
                ApplySimpleMigrations();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Successfully loaded {_fenceData?.Count ?? 0} fences from JSON");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Critical error in LoadFenceData: {ex.Message}");
                _fenceData = new List<dynamic>(); // Fallback to empty list
            }
        }

        /// <summary>
        /// JSON file parsing and loading - moved from FenceManager.LoadFenceDataFromJson
        /// Used by: LoadFenceData method
        /// Category: JSON Operations
        /// </summary>
        private static bool LoadFenceDataFromJson()
        {
            try
            {
                string jsonContent = File.ReadAllText(_jsonFilePath);

                // Check if the file is empty or contains only whitespace
                if (string.IsNullOrWhiteSpace(jsonContent))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FenceCreation,
                        "JSON file is empty or contains only whitespace. Using default fence configuration.");
                    return false;
                }

                // First, try to parse as a list of fences
                try
                {
                    _fenceData = JsonConvert.DeserializeObject<List<dynamic>>(jsonContent);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        $"Successfully loaded {_fenceData?.Count ?? 0} fences from JSON array.");
                    return _fenceData != null;
                }
                catch (JsonException)
                {
                    // If that fails, try to parse as a single fence object
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        "Failed to parse as array, trying single object format.");

                    dynamic singleFence = JsonConvert.DeserializeObject(jsonContent);
                    _fenceData = new List<dynamic> { singleFence };
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        "Successfully loaded single fence from JSON object.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error loading fence data from JSON: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Saves fence data to JSON with consistent formatting - moved from FenceManager.SaveFenceData
        /// Used by: FenceManager operations, fence updates, position changes
        /// Category: JSON Operations
        /// </summary>
        public static void SaveFenceData()
        {
            try
            {
                var serializedData = new List<JObject>();

                foreach (dynamic fence in _fenceData)
                {
                    IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ?
                        dict : ((JObject)fence).ToObject<IDictionary<string, object>>();

                    // Apply simple format consistency
                    ApplyFormatConsistency(fenceDict);

                    serializedData.Add(JObject.FromObject(fenceDict));
                }

                string formattedJson = JsonConvert.SerializeObject(serializedData, Formatting.Indented);
                File.WriteAllText(_jsonFilePath, formattedJson);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings,
                    $"Saved fences.json with consistent formatting for {serializedData.Count} fences");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings,
                    $"Error saving fence data: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Fence Creation - Used by: TrayManager, FenceManager
        /// <summary>
        /// Creates new fence with proper defaults - moved from FenceManager.CreateNewFence
        /// Used by: TrayManager.AddFence, portal fence creation
        /// Category: Fence Creation
        /// </summary>
        public static dynamic CreateNewFence(string title, string itemsType, double x = 20, double y = 20,
            string customColor = null, string customLaunchEffect = null)
        {
            try
            {
                // Generate appropriate fence name
                string fenceName = (itemsType != "Portal") ?
                    CoreUtilities.GenerateRandomName() : title;

                // Create new fence object
                dynamic newFence = new System.Dynamic.ExpandoObject();
                newFence.Id = Guid.NewGuid().ToString();
                IDictionary<string, object> newFenceDict = newFence;

                // Set basic properties
                newFenceDict["Title"] = fenceName;
                newFenceDict["X"] = x;
                newFenceDict["Y"] = y;
                newFenceDict["Width"] = 230;
                newFenceDict["Height"] = 130;
                newFenceDict["ItemsType"] = itemsType;

                // Set items based on fence type
                newFenceDict["Items"] = itemsType == "Portal" ? "" : new JArray();

                // Apply default properties with simple validation
                ApplyFenceDefaults(newFenceDict, customColor, customLaunchEffect);

                // Add to data collection and save
                _fenceData.Add(newFence);
                SaveFenceData();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Created new {itemsType} fence '{fenceName}' with ID {newFence.Id}");

                return newFence;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error creating new fence: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Simple Migration and Validation - Internal Use
        /// <summary>
        /// Applies simple migrations and property defaults
        /// Used by: LoadFenceData during startup
        /// Category: Data Migration (Simple)
        /// </summary>
        private static void ApplySimpleMigrations()
        {
            bool jsonModified = false;

            foreach (dynamic fence in _fenceData.ToList())
            {
                try
                {
                    IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ?
                        dict : ((JObject)fence).ToObject<IDictionary<string, object>>();

                    // Add missing basic properties
                    if (AddMissingBasicProperties(fenceDict))
                        jsonModified = true;

                    // Validate and fix data types
                    if (ValidateDataTypes(fenceDict))
                        jsonModified = true;
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                        $"Error applying simple migration to fence: {ex.Message}");
                }
            }

            // Save if any modifications were made
            if (jsonModified)
            {
                SaveFenceData();
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    "Applied simple migrations and saved updated fence data");
            }
        }

        /// <summary>
        /// Adds missing basic properties with defaults
        /// Used by: ApplySimpleMigrations
        /// Category: Property Validation
        /// </summary>
        private static bool AddMissingBasicProperties(IDictionary<string, object> fenceDict)
        {
            bool modified = false;

            // Add missing ID
            if (!fenceDict.ContainsKey("Id") || string.IsNullOrEmpty(fenceDict["Id"]?.ToString()))
            {
                fenceDict["Id"] = Guid.NewGuid().ToString();
                modified = true;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                    $"Added missing ID to fence '{(fenceDict.ContainsKey("Title") ? fenceDict["Title"] : "Unknown")}'");
            }

            // Add basic tab properties if missing
            if (!fenceDict.ContainsKey("TabsEnabled"))
            {
                fenceDict["TabsEnabled"] = "false";
                modified = true;
            }

            if (!fenceDict.ContainsKey("CurrentTab"))
            {
                fenceDict["CurrentTab"] = 0;
                modified = true;
            }

            if (!fenceDict.ContainsKey("Tabs"))
            {
                fenceDict["Tabs"] = new JArray();
                modified = true;
            }

            // Add basic state properties
            if (!fenceDict.ContainsKey("IsHidden"))
            {
                fenceDict["IsHidden"] = "false";
                modified = true;
            }

            if (!fenceDict.ContainsKey("IsRolled"))
            {
                fenceDict["IsRolled"] = "false";
                modified = true;
            }

            return modified;
        }

        /// <summary>
        /// Validates and fixes data types for consistency
        /// Used by: ApplySimpleMigrations
        /// Category: Data Validation
        /// </summary>
        private static bool ValidateDataTypes(IDictionary<string, object> fenceDict)
        {
            bool modified = false;

            // Ensure UnrolledHeight is a valid number
            if (fenceDict.ContainsKey("UnrolledHeight"))
            {
                if (!double.TryParse(fenceDict["UnrolledHeight"]?.ToString(), out double unrolledHeight) || unrolledHeight <= 0)
                {
                    double defaultHeight = fenceDict.ContainsKey("Height") ?
                        Convert.ToDouble(fenceDict["Height"]) : 130;
                    fenceDict["UnrolledHeight"] = defaultHeight.ToString();
                    modified = true;
                }
            }

            // Ensure Width and Height are valid
            if (!double.TryParse(fenceDict["Width"]?.ToString(), out double width) || width <= 0)
            {
                fenceDict["Width"] = 230;
                modified = true;
            }

            if (!double.TryParse(fenceDict["Height"]?.ToString(), out double height) || height <= 0)
            {
                fenceDict["Height"] = 130;
                modified = true;
            }

            return modified;
        }

        /// <summary>
        /// Applies format consistency for JSON serialization
        /// Used by: SaveFenceData
        /// Category: Format Consistency
        /// </summary>
        private static void ApplyFormatConsistency(IDictionary<string, object> fenceDict)
        {
            // Convert IsHidden to string format
            if (fenceDict.ContainsKey("IsHidden"))
            {
                bool isHidden = false;
                if (fenceDict["IsHidden"] is bool boolValue)
                    isHidden = boolValue;
                else if (fenceDict["IsHidden"] is string stringValue)
                    isHidden = stringValue.ToLower() == "true";

                fenceDict["IsHidden"] = isHidden.ToString().ToLower();
            }

            // Convert IsRolled to string format
            if (fenceDict.ContainsKey("IsRolled"))
            {
                bool isRolled = false;
                if (fenceDict["IsRolled"] is bool boolValue)
                    isRolled = boolValue;
                else if (fenceDict["IsRolled"] is string stringValue)
                    isRolled = stringValue.ToLower() == "true";

                fenceDict["IsRolled"] = isRolled.ToString().ToLower();
            }
        }

        /// <summary>
        /// Applies default properties to new fence
        /// Used by: CreateNewFence
        /// Category: Fence Defaults
        /// </summary>
        private static void ApplyFenceDefaults(IDictionary<string, object> fenceDict,
            string customColor, string customLaunchEffect)
        {
            // Basic state defaults
            fenceDict["IsHidden"] = "false";
            fenceDict["IsRolled"] = "false";
            fenceDict["UnrolledHeight"] = fenceDict["Height"].ToString();

            // Tab defaults
            fenceDict["TabsEnabled"] = "false";
            fenceDict["CurrentTab"] = 0;
            fenceDict["Tabs"] = new JArray();

            // Custom properties
            if (!string.IsNullOrEmpty(customColor))
                fenceDict["CustomColor"] = customColor;

            if (!string.IsNullOrEmpty(customLaunchEffect))
                fenceDict["CustomLaunchEffect"] = customLaunchEffect;
        }
        #endregion

        #region Data Access Helpers - Used by: Various managers
        /// <summary>
        /// Finds fence by ID with null safety
        /// Used by: FenceManager, CustomizeFenceForm, various operations
        /// Category: Data Access
        /// </summary>
        public static dynamic FindFenceById(string fenceId)
        {
            if (string.IsNullOrEmpty(fenceId) || _fenceData == null)
                return null;

            return _fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
        }

        /// <summary>
        /// Removes fence from data collection
        /// Used by: FenceManager.DeleteFence, cleanup operations
        /// Category: Data Modification
        /// </summary>
        public static bool RemoveFence(dynamic fence)
        {
            if (fence == null || _fenceData == null)
                return false;

            bool removed = _fenceData.Remove(fence);
            if (removed)
            {
                SaveFenceData();
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Removed fence '{fence.Title}' from data collection");
            }
            return removed;
        }
        #endregion
    }
}
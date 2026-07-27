using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Desktop_Frames
{
    /// <summary>
    /// Manages all frame data operations including JSON loading, saving, and basic migrations
    /// Extracted from Framemanager for better code organization and maintainability
    /// Handles core data persistence and simple validation/migration scenarios
    /// </summary>
    public static class FrameDataManager
    {
        #region Private Fields - Data Storage
        // Main frame data collection - moved from Framemanager
        private static List<dynamic> _frameData;

        // JSON file path - moved from Framemanager  
        private static string _jsonFilePath;

        // Key: ChildId, Value: List of ParentIds (Runtime 1-to-Many Accordion Snap Graph)
        public static Dictionary<string, List<string>> DockingMap = new Dictionary<string, List<string>>();
        #endregion

        #region Public Properties - Data Access
        /// <summary>
        /// Provides access to frame data collection
        /// Used by: Framemanager, TrayManager, CustomizeFrameForm, ItemMoveDialog
        /// Category: Data Access
        /// </summary>
        public static List<dynamic> FrameData
        {
            get => _frameData;
            set => _frameData = value;




        }

        /// <summary>
        /// Gets the JSON file path
        /// Used by: Framemanager, backup operations, debugging
        /// Category: File Management
        /// </summary>
        public static string JsonFilePath
        {
            get => _jsonFilePath;
            set => _jsonFilePath = value;
        }
        #endregion

        #region Initialization - Used by: Framemanager startup
        /// <summary>
        /// Initializes the data manager with JSON file path
        /// Used by: Framemanager during application startup
        /// Category: Initialization
        /// </summary>
        public static void Initialize()
        {
            try
            {
                // ====================================================================
                // [LEGACY "FENCES" MIGRATION - DO NOT REMOVE]
                // Self-Healing Logic: Checks for legacy fences.json and instantly 
                // renames it to frames.json upon profile load.
                // ====================================================================
                string legacyPath = ProfileManager.GetProfileFilePath("fences.json");
                string newPath = ProfileManager.GetProfileFilePath("frames.json");

                if (!File.Exists(newPath) && File.Exists(legacyPath))
                {
                    try
                    {
                        // KISS: Read raw text, replace strings, write new file, delete old file.
                        string rawJson = File.ReadAllText(legacyPath);

                        rawJson = rawJson.Replace("\"FenceBorderColor\"", "\"FrameBorderColor\"");
                        rawJson = rawJson.Replace("\"frameBorderColor\"", "\"FrameBorderColor\"");
                        rawJson = rawJson.Replace("\"FenceBorderThickness\"", "\"FrameBorderThickness\"");
                        rawJson = rawJson.Replace("\"frameBorderThickness\"", "\"FrameBorderThickness\"");

                        File.WriteAllText(newPath, rawJson);
                        File.Delete(legacyPath);

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Auto-Migrated and scrubbed fences.json to frames.json");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Migration scrub failed: {ex.Message}");
                        newPath = legacyPath; // Fallback to old file if blocked
                    }
                }

                _jsonFilePath = newPath;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                    $"FrameDataManager initialized with path: {_jsonFilePath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Error initializing FrameDataManager: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Core Data Operations - Used by: Framemanager, TrayManager
        /// <summary>
        /// Main frame data loading method - moved from Framemanager.LoadFrameData
        /// Used by: Framemanager.LoadFrameData, application startup
        /// Category: Data Loading
        /// </summary>
        public static void LoadFrameData(TargetChecker targetChecker)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                    "FrameDataManager: Starting frame data load");

                // Initialize frame data list
                _frameData = new List<dynamic>();

                // Check if JSON file exists
                if (!File.Exists(_jsonFilePath))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                        "JSON file not found. Starting with empty frame configuration.");
                    return;
                }

                // Load and parse JSON data
                if (!LoadFrameDataFromJson())
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FrameCreation,
                        "Failed to load JSON data. Starting with empty configuration.");
                    return;
                }

                // Apply simple migrations and validation
                ApplySimpleMigrations();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                    $"Successfully loaded {_frameData?.Count ?? 0} frames from JSON");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Critical error in LoadFrameData: {ex.Message}");
                _frameData = new List<dynamic>(); // Fallback to empty list
            }
        }

        /// <summary>
        /// JSON file parsing and loading - moved from Framemanager.LoadFrameDataFromJson
        /// Used by: LoadFrameData method
        /// Category: JSON Operations
        /// </summary>
        private static bool LoadFrameDataFromJson()
        {
            try
            {
                string jsonContent = File.ReadAllText(_jsonFilePath);

                // Check if the file is empty or contains only whitespace
                if (string.IsNullOrWhiteSpace(jsonContent))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FrameCreation,
                        "JSON file is empty or contains only whitespace. Using default frame configuration.");
                    return false;
                }

                // First, try to parse as a list of frames
                try
                {
                    _frameData = JsonConvert.DeserializeObject<List<dynamic>>(jsonContent);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                        $"Successfully loaded {_frameData?.Count ?? 0} frames from JSON array.");
                    return _frameData != null;
                }
                catch (JsonException)
                {
                    // If that fails, try to parse as a single frames object
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                        "Failed to parse as array, trying single object format.");

                    dynamic singleFrame = JsonConvert.DeserializeObject(jsonContent);
                    _frameData = new List<dynamic> { singleFrame };
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                        "Successfully loaded single frame from JSON object.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Error loading frame data from JSON: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Saves frame data to JSON with consistent formatting - moved from Framemanager.SaveFrameData
        /// Used by: Framemanager operations, frame updates, position changes
        /// Category: JSON Operations
        /// </summary>
        public static void SaveFrameData()
        {
            try
            {
                var serializedData = new List<JObject>();

                foreach (dynamic frame in _frameData)
                {
                    IDictionary<string, object> frameDict = frame is IDictionary<string, object> dict ?
                        dict : ((JObject)frame).ToObject<IDictionary<string, object>>();

                    // ====================================================================
                    // [LEGACY "FENCES" MIGRATION - DO NOT REMOVE]
                    // The Ultimate Interceptor: Forces old or accidentally generated 
                    // "Fence" keys into official "Frame" keys RIGHT BEFORE saving to disk.
                    // ====================================================================
                    void ConsolidateKey(string officialKey, string[] legacyKeys)
                    {
                        object rescuedValue = null;

                        // Extract valid data from old keys and delete them permanently
                        foreach (string oldKey in legacyKeys)
                        {
                            if (frameDict.ContainsKey(oldKey))
                            {
                                // BUG FIX: Removed '!= "0"' check. 0 is a valid integer (e.g., Border Thickness, X, Y).
                                if (frameDict[oldKey] != null && frameDict[oldKey].ToString() != "")
                                {
                                    rescuedValue = frameDict[oldKey];
                                }
                                frameDict.Remove(oldKey); // Vacuum it out
                            }
                        }

                        // Apply rescued data to the official key if it doesn't already have a valid setting
                        // BUG FIX: Removed '!= "0"' check here to prevent wiping valid zero values.
                        bool hasValidOfficial = frameDict.ContainsKey(officialKey) && frameDict[officialKey] != null && frameDict[officialKey].ToString() != "";
                        if (!hasValidOfficial && rescuedValue != null)
                        {
                            frameDict[officialKey] = rescuedValue;
                        }
                    }

                    ConsolidateKey("FrameBorderColor", new[] { "FrameBorderColor", "FrameBorderColor" });
                    ConsolidateKey("FrameBorderThickness", new[] { "FrameBorderThickness", "FrameBorderThickness" });

                    // Apply simple format consistency
                    ApplyFormatConsistency(frameDict);

                    serializedData.Add(JObject.FromObject(frameDict));
                }

                string formattedJson = JsonConvert.SerializeObject(serializedData, Formatting.Indented);
                File.WriteAllText(_jsonFilePath, formattedJson);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings,
                    $"Saved frames.json with consistent formatting for {serializedData.Count} frames");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings,
                    $"Error saving frame data: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Frame Creation - Used by: TrayManager, Framemanager
        /// <summary>
        /// Creates new frame with proper defaults - moved from Framemanager.CreateNewFrame
        /// Category: Frame Creation
        /// </summary>
        public static dynamic CreateNewFrame(string title, string itemsType, double x = 20, double y = 20,
            string customColor = null, string customLaunchEffect = null)
        {
            try
            {
                // Generate appropriate frame name
                string frameName = (itemsType != "Portal") ?
                    CoreUtilities.GenerateRandomName() : title;

                // Create new frame object
                dynamic newFrame = new System.Dynamic.ExpandoObject();
                newFrame.Id = Guid.NewGuid().ToString();
                IDictionary<string, object> newframeDict = newFrame;

                // Set basic properties
                newframeDict["Title"] = frameName;
                newframeDict["X"] = x;
                newframeDict["Y"] = y;
                newframeDict["Width"] = 230;
                newframeDict["Height"] = 130;
                newframeDict["ItemsType"] = itemsType;

                // Set items based on frame type
                newframeDict["Items"] = itemsType == "Portal" ? "" : new JArray();

                // Apply default properties with simple validation
                ApplyFrameDefaults(newframeDict, customColor, customLaunchEffect);

                // Add to data collection and save
                _frameData.Add(newFrame);
                SaveFrameData();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                    $"Created new {itemsType} frame '{frameName}' with ID {newFrame.Id}");

                return newFrame;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Error creating new frame: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Simple Migration and Validation - Internal Use



        /// <summary>
        /// Applies simple migrations and property defaults
        /// Used by: LoadFrameData during startup
        /// Category: Data Migration (Simple)
        /// </summary>
        private static void ApplySimpleMigrations()
        {
            bool jsonModified = false;

            for (int i = 0; i < _frameData.Count; i++)
            {
                try
                {
                    if (_frameData[i] is JObject frameObj)
                    {
                

                        // --- Run existing property/type validations ---
                        IDictionary<string, object> frameDict = frameObj.ToObject<IDictionary<string, object>>();
                        bool dictModified = false;

                        if (AddMissingBasicProperties(frameDict)) dictModified = true;
                        if (ValidateDataTypes(frameDict)) dictModified = true;

                        if (dictModified)
                        {
                            _frameData[i] = JObject.FromObject(frameDict);
                            jsonModified = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                        $"Error applying simple migration to frame: {ex.Message}");
                }
            }

            if (jsonModified)
            {
                SaveFrameData();
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                    "Applied simple migrations and saved updated frame data");
            }
        }

        /// <summary>
        /// Adds missing basic properties with defaults
        /// Used by: ApplySimpleMigrations
        /// Category: Property Validation
        /// </summary>
        private static bool AddMissingBasicProperties(IDictionary<string, object> frameDict)
        {
            bool modified = false;

            if (!frameDict.ContainsKey("Id") || string.IsNullOrEmpty(frameDict["Id"]?.ToString()))
            {
                frameDict["Id"] = Guid.NewGuid().ToString();
                modified = true;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                    $"Added missing ID to frame '{(frameDict.ContainsKey("Title") ? frameDict["Title"] : "Unknown")}'");
            }

            if (!frameDict.ContainsKey("TabsEnabled")) { frameDict["TabsEnabled"] = "false"; modified = true; }
            if (!frameDict.ContainsKey("CurrentTab")) { frameDict["CurrentTab"] = 0; modified = true; }
            if (!frameDict.ContainsKey("Tabs")) { frameDict["Tabs"] = new JArray(); modified = true; }
            if (!frameDict.ContainsKey("IsHidden")) { frameDict["IsHidden"] = "false"; modified = true; }
            if (!frameDict.ContainsKey("IsRolled")) { frameDict["IsRolled"] = "false"; modified = true; }

            return modified;
        }

        /// <summary>
        /// Validates and fixes data types for consistency
        /// Used by: ApplySimpleMigrations
        /// Category: Data Validation
        /// </summary>
        private static bool ValidateDataTypes(IDictionary<string, object> frameDict)
        {
            bool modified = false;

            if (frameDict.ContainsKey("UnrolledHeight"))
            {
                if (!double.TryParse(frameDict["UnrolledHeight"]?.ToString(), out double unrolledHeight) || unrolledHeight <= 0)
                {
                    double defaultHeight = frameDict.ContainsKey("Height") ?
                        Convert.ToDouble(frameDict["Height"]) : 130;
                    frameDict["UnrolledHeight"] = defaultHeight.ToString();
                    modified = true;
                }
            }

            if (!double.TryParse(frameDict["Width"]?.ToString(), out double width) || width <= 0)
            {
                frameDict["Width"] = 230;
                modified = true;
            }

            if (!double.TryParse(frameDict["Height"]?.ToString(), out double height) || height <= 0)
            {
                frameDict["Height"] = 130;
                modified = true;
            }

            return modified;
        }

        /// <summary>
        /// Applies format consistency for JSON serialization
        /// Used by: SaveFrameData
        /// Category: Format Consistency
        /// </summary>
        private static void ApplyFormatConsistency(IDictionary<string, object> frameDict)
        {
            if (frameDict.ContainsKey("IsHidden"))
            {
                bool isHidden = false;
                if (frameDict["IsHidden"] is bool boolValue) isHidden = boolValue;
                else if (frameDict["IsHidden"] is string stringValue) isHidden = stringValue.ToLower() == "true";
                frameDict["IsHidden"] = isHidden.ToString().ToLower();
            }

            if (frameDict.ContainsKey("IsRolled"))
            {
                bool isRolled = false;
                if (frameDict["IsRolled"] is bool boolValue) isRolled = boolValue;
                else if (frameDict["IsRolled"] is string stringValue) isRolled = stringValue.ToLower() == "true";
                frameDict["IsRolled"] = isRolled.ToString().ToLower();
            }
        }

        /// <summary>
        /// Applies default properties to new frame
        /// Used by: CreateNewFrame
        /// Category: Frame Defaults
        /// </summary>
        private static void ApplyFrameDefaults(IDictionary<string, object> frameDict,
             string customColor, string customLaunchEffect)
        {
            frameDict["IsHidden"] = "false";
            frameDict["IsRolled"] = "false";
            frameDict["UnrolledHeight"] = frameDict["Height"].ToString();
            frameDict["TabsEnabled"] = "false";
            frameDict["CurrentTab"] = 0;
            frameDict["Tabs"] = new JArray();

            string itemsType = frameDict["ItemsType"]?.ToString();
            if (itemsType == "Note") NoteFramemanager.ApplyNoteDefaults(frameDict);

            if (!string.IsNullOrEmpty(customColor)) frameDict["CustomColor"] = customColor;
            if (!string.IsNullOrEmpty(customLaunchEffect)) frameDict["CustomLaunchEffect"] = customLaunchEffect;
        }
        #endregion

        #region Data Access Helpers - Used by: Various managers
        /// <summary>
        /// Finds frame by ID with null safety
        /// Used by: Framemanager, CustomizeFrameForm, various operations
        /// Category: Data Access
        /// </summary>
        public static dynamic FindFrameById(string frameId)
        {
            if (string.IsNullOrEmpty(frameId) || _frameData == null)
                return null;

            return _frameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
        }



        /// <summary>
        /// Mutates the master JSON frame array directly to guarantee persistence across restarts
        /// </summary>
        public static void UpdateDockedRelationships(string childId, List<string> parentIds)
        {
            try
            {
                if (string.IsNullOrEmpty(childId) || _frameData == null) return;

                if (parentIds != null && parentIds.Count > 0) DockingMap[childId] = parentIds;
                else DockingMap.Remove(childId);

                int index = _frameData.FindIndex(f => f.Id?.ToString() == childId);
                if (index >= 0)
                {
                    dynamic actualFrame = _frameData[index];
                    IDictionary<string, object> dict = actualFrame is IDictionary<string, object> d ?
                        d : ((JObject)actualFrame).ToObject<IDictionary<string, object>>();

                    if (parentIds != null && parentIds.Count > 0) dict["DockedParentId"] = string.Join(",", parentIds);
                    else if (dict.ContainsKey("DockedParentId")) dict.Remove("DockedParentId");

                    _frameData[index] = JObject.FromObject(dict);
                    SaveFrameData();
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Saved persistent multi-snap: '{childId}' -> '{string.Join(",", parentIds ?? new List<string>())}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error saving docked relationships: {ex.Message}");
            }
        }

        /// <summary>
        /// Scans loaded JSON frame dictionaries, populates multi-parent DockingMap, and self-heals overlapping child coordinates
        /// </summary>
        public static void RebuildDockingMap()
        {
            try
            {
                DockingMap.Clear();
                if (_frameData == null) return;

                // 1. Rebuild runtime snap graph from persistent JSON attributes
                foreach (dynamic frame in _frameData)
                {
                    IDictionary<string, object> dict = frame is IDictionary<string, object> d ?
                        d : ((JObject)frame).ToObject<IDictionary<string, object>>();

                    if (dict.ContainsKey("Id") && dict.ContainsKey("DockedParentId") && dict["DockedParentId"] != null)
                    {
                        string cId = dict["Id"].ToString();
                        string pIdStr = dict["DockedParentId"].ToString();
                        if (!string.IsNullOrEmpty(cId) && !string.IsNullOrEmpty(pIdStr))
                        {
                            var pIds = pIdStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                            if (pIds.Count > 0) DockingMap[cId] = pIds;
                        }
                    }
                }

                // 2. SELF-HEALING STARTUP ALIGNMENT FOR MULTI-PARENTS
                bool layoutHealed = false;
                foreach (var kvp in DockingMap)
                {
                    string childId = kvp.Key;
                    List<string> parentIds = kvp.Value;

                    int childIdx = _frameData.FindIndex(f => f.Id?.ToString() == childId);
                    if (childIdx >= 0 && parentIds.Count > 0)
                    {
                        dynamic childFrame = _frameData[childIdx];
                        double maxExpectedMinY = 0;

                        foreach (string pId in parentIds)
                        {
                            int parentIdx = _frameData.FindIndex(f => f.Id?.ToString() == pId);
                            if (parentIdx >= 0)
                            {
                                dynamic parentFrame = _frameData[parentIdx];
                                double parentY = Convert.ToDouble(parentFrame.Y?.ToString() ?? "0");

                                // --- FIX: State-Aware Height Calculation ---
                                bool isParentRolled = false;
                                try { isParentRolled = parentFrame.IsRolled?.ToString().ToLower() == "true"; } catch { }

                                double activeParentHeight = isParentRolled ? 28.0 : Convert.ToDouble(parentFrame.UnrolledHeight?.ToString() ?? "130");

                                // Anchor exactly 10px below the active bottom edge (matching runtime CascadeStack)
                                double expectedMinY = parentY + activeParentHeight + 10.0;

                                if (expectedMinY > maxExpectedMinY) maxExpectedMinY = expectedMinY;
                            }
                        }

                        double childY = Convert.ToDouble(childFrame.Y?.ToString() ?? "0");

                        // --- FIX: Strict Exact Alignment ---
                        // Math.Abs > 0.5 allows for tiny floating-point margins without endless saving,
                        // while ensuring frames are pulled UP to close gaps and pushed DOWN to clear overlaps.
                        if (maxExpectedMinY > 0 && Math.Abs(childY - maxExpectedMinY) > 0.5)
                        {
                            childFrame.Y = maxExpectedMinY;
                            layoutHealed = true;
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Self-Healed coordinate for multi-docked child '{childFrame.Title}' from Y={childY:F1} to Y={maxExpectedMinY:F1}");
                        }
                    }
                }

                if (layoutHealed)
                {
                    SaveFrameData();
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Restored {DockingMap.Count} persistent multi-snap links from JSON.");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error in RebuildDockingMap: {ex.Message}");
            }
        }

        #endregion
    }
}
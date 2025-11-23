using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using IWshRuntimeLibrary;

namespace Desktop_Fences
{
    /// <summary>
    /// Comprehensive utilities for Desktop Fences+ application
    /// Consolidates all utility methods from FenceManager, IconManager, PortalFenceManager, and FenceUtilities
    /// Organized by functional categories with clear regions for maintainability
    /// </summary>
    public static class CoreUtilities
    {
        #region Win32 API - System Integration
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        /// <summary>
        /// Gets accurate cursor position using Win32 API
        /// Used by: IconDragDropManager.CreateDragPreview, drag operations
        /// Category: Input Handling
        /// </summary>
        public static Point GetCursorPosition()
        {
            POINT point;
            GetCursorPos(out point);
            return new Point(point.X, point.Y);
        }
        #endregion

        #region File Validation - Existence and Type Checking
        /// <summary>
        /// Safe file existence check with Unicode support
        /// Used by: Icon validation, file operations
        /// Category: File Validation
        /// Moved from: FilePathUtilities.SafeFileExists
        /// </summary>
        public static bool SafeFileExists(string filePath)
        {
            try
            {
                return System.IO.File.Exists(filePath);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"File existence check failed for {filePath}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Safe directory existence check with Unicode support
        /// Used by: Folder validation, portal fence operations
        /// Category: Directory Validation
        /// Moved from: FilePathUtilities.SafeDirectoryExists
        /// </summary>
        public static bool SafeDirectoryExists(string directoryPath)
        {
            try
            {
                return Directory.Exists(directoryPath);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"Directory existence check failed for {directoryPath}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Detects if a file is temporary and should be filtered out
        /// Used by: PortalFenceManager for file watching, general file operations
        /// Category: File Filtering
        /// Moved from: PortalFenceManager.IsTemporaryFile
        /// </summary>
        public static bool IsTemporaryFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return true;

            string fileName = Path.GetFileName(filePath);

            // Excel temporary files (this is the main fix for Excel temp file issues)
            if (fileName.StartsWith("~$"))
                return true;

            // Word temporary files
            if (fileName.StartsWith("~WRL") || fileName.StartsWith("~WRD"))
                return true;

            // PowerPoint temporary files  
            if (fileName.StartsWith("~PPT"))
                return true;

            // General temporary patterns
            if (fileName.StartsWith(".tmp") || fileName.EndsWith(".tmp") ||
                fileName.StartsWith("tmp") || fileName.EndsWith(".temp"))
                return true;

            // System files
            if (fileName == "Thumbs.db" || fileName == "desktop.ini" || fileName == ".DS_Store")
                return true;

            // Try to check file attributes if file exists
            try
            {
                if (System.IO.File.Exists(filePath) || Directory.Exists(filePath))
                {
                    FileAttributes attributes = System.IO.File.GetAttributes(filePath);
                    if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden ||
                        (attributes & FileAttributes.System) == FileAttributes.System ||
                        (attributes & FileAttributes.Temporary) == FileAttributes.Temporary)
                    {
                        return true;
                    }
                }
            }
            catch
            {
                // If we can't check attributes, continue with other checks
            }

            return false;
        }
        #endregion

        #region File Type Detection - Extension and Format Validation
        /// <summary>
        /// Checks if file has a specific extension (case-insensitive)
        /// Used by: Various file processing operations
        /// Category: File Type Validation
        /// </summary>
        public static bool HasExtension(string filePath, string extension)
        {
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(extension))
                return false;

            return Path.GetExtension(filePath).Equals(extension, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Checks if file has any of the specified extensions (case-insensitive)
        /// Used by: File filtering operations
        /// Category: File Type Validation
        /// </summary>
        public static bool HasAnyExtension(string filePath, params string[] extensions)
        {
            if (string.IsNullOrEmpty(filePath) || extensions == null || extensions.Length == 0)
                return false;

            string fileExtension = Path.GetExtension(filePath);
            return extensions.Any(ext => fileExtension.Equals(ext, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Checks if file is a shortcut file (.lnk or .url)
        /// Used by: Shortcut processing operations
        /// Category: File Type Validation
        /// </summary>
        public static bool IsShortcutFile(string filePath)
        {
            return HasAnyExtension(filePath, ".lnk", ".url");
        }

        /// <summary>
        /// Checks if file is an executable (.exe, .bat, .cmd, .msi)
        /// Used by: Executable detection and processing
        /// Category: File Type Validation
        /// </summary>
        public static bool IsExecutableFile(string filePath)
        {
            return HasAnyExtension(filePath, ".exe", ".bat", ".cmd", ".msi", ".com", ".scr");
        }
        #endregion

        #region Web Link Validation - URL and Web Shortcut Processing
        /// <summary>
        /// Detects if a file is a web link shortcut (.url or .lnk pointing to web)
        /// Used by: FenceManager for processing dropped web links
        /// Category: Link Validation
        /// Moved from: FenceManager.IsWebLinkShortcut
        /// </summary>
        public static bool IsWebLinkShortcut(string filePath)
        {
            try
            {
                string extension = System.IO.Path.GetExtension(filePath).ToLower();

                if (extension == ".url")
                {
                    // Check if .url file contains web URL
                    if (System.IO.File.Exists(filePath))
                    {
                        string content = System.IO.File.ReadAllText(filePath);
                        return content.Contains("URL=http://") || content.Contains("URL=https://");
                    }
                }
                else if (extension == ".lnk")
                {
                    // Check if .lnk file targets web URL
                    if (System.IO.File.Exists(filePath))
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                        string target = shortcut.TargetPath ?? "";
                        return target.StartsWith("http://") || target.StartsWith("https://");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                    $"Error checking if {filePath} is web link: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Extracts web URL from .url or .lnk files
        /// Used by: FenceManager for web link processing
        /// Category: Link Processing
        /// Moved from: FenceManager.ExtractWebUrlFromFile
        /// </summary>
        public static string ExtractWebUrlFromFile(string filePath)
        {
            try
            {
                string extension = System.IO.Path.GetExtension(filePath).ToLower();

                if (extension == ".url")
                {
                    // Extract URL from .url file
                    if (System.IO.File.Exists(filePath))
                    {
                        string[] lines = System.IO.File.ReadAllLines(filePath);
                        foreach (string line in lines)
                        {
                            if (line.StartsWith("URL="))
                            {
                                return line.Substring(4); // Remove "URL=" prefix
                            }
                        }
                    }
                }
                else if (extension == ".lnk")
                {
                    // Extract URL from .lnk file
                    if (System.IO.File.Exists(filePath))
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                        return shortcut.TargetPath ?? "";
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling,
                    $"Error extracting web URL from {filePath}: {ex.Message}");
            }
            return string.Empty;
        }

        /// <summary>
        /// Extracts URL from .url file format
        /// Used by: FenceManager.AddIcon for web links
        /// Category: File Processing
        /// Moved from: FenceUtilities.ExtractUrlFromFile
        /// </summary>
        public static string ExtractUrlFromFile(string urlFilePath)
        {
            try
            {
                if (!System.IO.File.Exists(urlFilePath))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                        $"URL file does not exist: {urlFilePath}");
                    return null;
                }

                string[] lines = System.IO.File.ReadAllLines(urlFilePath);
                foreach (string line in lines)
                {
                    if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                    {
                        string url = line.Substring(4).Trim(); // Remove "URL=" prefix
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                            $"Extracted URL from {urlFilePath}: {url}");
                        return url;
                    }
                }

                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                    $"No URL= line found in .url file: {urlFilePath}");
                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"Error reading .url file {urlFilePath}: {ex.Message}");
                return null;
            }
        }
        #endregion

        #region String Validation and Processing - Input and Name Validation
        /// <summary>
        /// Validates if string is not null, empty, or whitespace only
        /// Used by: Input validation, data processing
        /// Category: Input Validation
        /// </summary>
        public static bool IsValidString(string input)
        {
            return !string.IsNullOrWhiteSpace(input);
        }

        /// <summary>
        /// Validates string and ensures it meets minimum length requirement
        /// Used by: Name validation, input validation
        /// Category: Input Validation
        /// </summary>
        public static bool IsValidString(string input, int minLength)
        {
            return IsValidString(input) && input.Trim().Length >= minLength;
        }

        /// <summary>
        /// Validates that a name is suitable for file/folder operations
        /// Used by: Fence naming, file operations
        /// Category: Name Validation
        /// </summary>
        public static bool IsValidName(string name)
        {
            if (!IsValidString(name))
                return false;

            // Check for invalid file system characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            return !name.Any(c => invalidChars.Contains(c));
        }

        /// <summary>
        /// Checks if a path contains Unicode characters that may need special handling
        /// Used by: FenceManager, shortcut operations
        /// Category: Path Analysis
        /// Moved from: FilePathUtilities.ContainsUnicodeCharacters
        /// </summary>
        public static bool ContainsUnicodeCharacters(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            return path.Any(c => c > 127);
        }

        /// <summary>
        /// Truncates display name based on settings with ellipsis
        /// Used by: Icon display, UI text processing
        /// Category: Text Processing
        /// Extracted from: IconManager display name logic
        /// </summary>
        public static string TruncateDisplayName(string displayName, int maxLength)
        {
            if (string.IsNullOrEmpty(displayName) || displayName.Length <= maxLength)
                return displayName;

            return displayName.Substring(0, maxLength) + "...";
        }

        /// <summary>
        /// Cleans and sanitizes file paths for safe operations
        /// Used by: File operations, path processing
        /// Category: Path Processing
        /// </summary>
        public static string CleanPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            return path.Trim().Replace("\\\\", "\\");
        }
        #endregion

        #region UI Size and Layout Utilities - Icon Sizing and Spacing
        /// <summary>
        /// Converts icon size name to pixel dimensions
        /// Used by: ApplyIconSize, icon rendering
        /// Category: Icon Sizing
        /// Moved from: IconManager.GetIconSizePixels
        /// </summary>
        public static int GetIconSizePixels(string iconSize)
        {
            return iconSize switch
            {
                "Tiny" => 16,
                "Small" => 24,
                "Medium" => 32,
                "Large" => 48,
                "Huge" => 64,
                _ => 32 // Default medium size
            };
        }

        /// <summary>
        /// Gets default icon spacing value with validation
        /// Used by: Icon layout, fence spacing calculations
        /// Category: Layout Utilities
        /// </summary>
        public static int GetDefaultIconSpacing()
        {
            return 5; // Default spacing in pixels
        }

        /// <summary>
        /// Validates and normalizes icon spacing values
        /// Used by: Icon layout calculations
        /// Category: Layout Utilities
        /// </summary>
        public static int ValidateIconSpacing(object spacingValue)
        {
            try
            {
                if (spacingValue == null)
                    return GetDefaultIconSpacing();

                if (int.TryParse(spacingValue.ToString(), out int spacing))
                {
                    // Ensure spacing is within reasonable bounds
                    return Math.Max(0, Math.Min(spacing, 50));
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"Error validating icon spacing: {ex.Message}");
            }

            return GetDefaultIconSpacing();
        }
        #endregion

        #region Random Name Generation - Fence Naming Utilities
        // Adjectives for random fence name generation
        private static readonly string[] adjectives = {
            "High", "Low", "Tiny", "Vast", "Wide", "Slim", "Flat", "Bold", "Cold", "Warm",
            "Soft", "Hard", "Dark", "Pale", "Fast", "Slow", "Deep", "Tall", "Short", "Bent",
            "Thin", "Bright", "Light", "Sharp", "Dull", "Loud", "Mute", "Grim", "Kind", "Neat",
            "Rough", "Smooth", "Brave", "Fierce", "Plain", "Worn", "Dry", "Damp", "Strong", "Weak"
        };

        // Places for random fence name generation
        private static readonly string[] places = {
            "Bay", "Hill", "Lake", "Cove", "Peak", "Reef", "Dune", "Glen", "Moor", "Vale",
            "Rock", "Shore", "Bank", "Ford", "Cape", "Crag", "Marsh", "Pond", "Cliff", "Wood",
            "Dell", "Pass", "Cave", "Ridge", "Falls", "Grove", "Creek", "Bluff", "Trail", "Point"
        };

        /// <summary>
        /// Generates a random fence name using adjective + place pattern
        /// Used by: FenceManager.CreateNewFence for default naming
        /// Category: Name Generation
        /// Moved from: FenceUtilities.GenerateRandomFenceName
        /// </summary>
        public static string GenerateRandomFenceName()
        {
            Random random = new Random();
            string adjective = adjectives[random.Next(adjectives.Length)];
            string place = places[random.Next(places.Length)];
            return $"{adjective} {place}";
        }

        /// <summary>
        /// Generates a random name using adjective + place pattern (legacy compatibility)
        /// Used by: FenceManager, FenceDataManager (existing calls)
        /// Category: Name Generation
        /// Moved from: FenceUtilities.GenerateRandomName
        /// </summary>
        public static string GenerateRandomName()
        {
            Random random = new Random();
            string adjective = adjectives[random.Next(adjectives.Length)];
            string place = places[random.Next(places.Length)];
            return $"{adjective} {place}";
        }

        /// <summary>
        /// Generates a unique fence name by checking against existing names
        /// Used by: Fence creation to avoid duplicates
        /// Category: Name Generation
        /// </summary>
        public static string GenerateUniqueFenceName()
        {
            string baseName = GenerateRandomFenceName();
            string uniqueName = baseName;
            int counter = 1;

            // Check against existing fence names
            try
            {
                var existingNames = FenceDataManager.FenceData
                    .Select(f => f.Title?.ToString())
                    .Where(name => !string.IsNullOrEmpty(name))
                    .ToList();

                while (existingNames.Any(name => string.Equals(name, uniqueName, StringComparison.OrdinalIgnoreCase)))
                {
                    uniqueName = $"{baseName} {counter}";
                    counter++;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"Error checking unique fence name: {ex.Message}");
            }

            return uniqueName;
        }
        #endregion

        #region UI Creation Helpers - Text and Tooltip Generation
        /// <summary>
        /// Creates standardized text block for icon labels
        /// Used by: Icon creation, UI consistency
        /// Category: UI Creation
        /// Extracted from: IconManager.CreateDisplayNameTextBlock logic
        /// </summary>
        public static TextBlock CreateIconTextBlock(string displayName, int maxLength = 15)
        {
            string truncatedName = TruncateDisplayName(displayName, maxLength);

            return new TextBlock
            {
                Text = truncatedName,
                Foreground = System.Windows.Media.Brushes.White,
                TextAlignment = TextAlignment.Center,
                FontSize = 10,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 60,
                Effect = new DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 315,
                    ShadowDepth = 1,
                    BlurRadius = 2
                }
            };
        }

        /// <summary>
        /// Creates tooltip text with file information
        /// Used by: Icon tooltip creation
        /// Category: UI Creation
        /// Extracted from: IconManager.CreateIconTooltip logic
        /// </summary>
        public static string CreateTooltipText(string filePath, string targetPath = null, string arguments = null)
        {
            string toolTipText = $"File: {Path.GetFileName(filePath)}";

            if (!string.IsNullOrEmpty(targetPath) && targetPath != filePath)
            {
                toolTipText += $"\nTarget: {targetPath}";
            }

            if (!string.IsNullOrEmpty(arguments))
            {
                toolTipText += $"\nParameters: {arguments}";
            }

            return toolTipText;
        }
        #endregion

        #region Configuration Validation - Settings and Default Values
        /// <summary>
        /// Validates fence color values with fallback defaults
        /// Used by: Fence creation, settings validation
        /// Category: Configuration Validation
        /// </summary>
        public static string ValidateFenceColor(string colorValue, string defaultColor = "#FF1E1E1E")
        {
            if (string.IsNullOrEmpty(colorValue))
                return defaultColor;

            // Basic hex color validation
            if (colorValue.StartsWith("#") && (colorValue.Length == 7 || colorValue.Length == 9))
            {
                return colorValue;
            }

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                $"Invalid color value '{colorValue}', using default: {defaultColor}");
            return defaultColor;
        }

        /// <summary>
        /// Validates boolean configuration values with safe defaults
        /// Used by: Settings processing, configuration loading
        /// Category: Configuration Validation
        /// </summary>
        public static bool ValidateBooleanSetting(object value, bool defaultValue = false)
        {
            if (value == null)
                return defaultValue;

            if (value is bool boolValue)
                return boolValue;

            if (bool.TryParse(value.ToString(), out bool parsedValue))
                return parsedValue;

            return defaultValue;
        }

        /// <summary>
        /// Validates integer configuration values with range checking
        /// Used by: Settings processing, numeric configuration
        /// Category: Configuration Validation
        /// </summary>
        public static int ValidateIntegerSetting(object value, int defaultValue, int minValue = int.MinValue, int maxValue = int.MaxValue)
        {
            if (value == null)
                return defaultValue;

            if (int.TryParse(value.ToString(), out int intValue))
            {
                return Math.Max(minValue, Math.Min(maxValue, intValue));
            }

            return defaultValue;
        }
        #endregion

        #region Path and Unicode Utilities - Advanced Path Processing
        /// <summary>
        /// Checks if path contains only ASCII characters
        /// Used by: UpdateIcon for Unicode handling
        /// Category: Unicode Support
        /// Moved from: IconManager.IsAsciiPath
        /// </summary>
        public static bool IsAsciiPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return true;

            return path.All(c => c <= 127);
        }



        #endregion
    }
}
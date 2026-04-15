using IWshRuntimeLibrary;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Desktop_Fences
{
    
    
    
    /// <summary>
    /// Manages all icon operations including extraction, caching, rendering, and updates
    /// Extracted from FenceManager for better code organization and maintainability
    /// Handles icon lifecycle from extraction to display with advanced caching and Unicode support
    /// </summary>
    public static class IconManager
    {
        #region DLL Imports - Windows Icon API
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

        [DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr hIcon);
        #endregion

        #region Private Fields - Icon Cache
        // Icon cache for performance optimization - moved from FenceManager
        // FIX: Made case-insensitive to survive FileSystemWatcher string mutations
        private static readonly Dictionary<string, ImageSource> iconCache = new Dictionary<string, ImageSource>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Public Properties - Cache Access
        /// <summary>
        /// Provides access to icon cache for performance monitoring and cleanup
        /// Used by: FenceManager (for compatibility), debugging operations
        /// Category: Cache Management
        /// </summary>
        public static Dictionary<string, ImageSource> IconCache => iconCache;
        #endregion

        #region Async Thread-Safe Helpers
        /// <summary>
        /// Safely creates a frozen BitmapImage that can be passed across background threads without crashing WPF.
        /// </summary>
        public static BitmapImage CreateFrozenBitmap(string uriString)
        {
            var bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.CacheOption = BitmapCacheOption.OnLoad; // CRITICAL: Forces synchronous load
            bmp.UriSource = new Uri(uriString);
            bmp.EndInit();
            bmp.Freeze(); // CRITICAL: Makes cross-thread safe
            return bmp;
        }

        private static ImageSource FreezeIcon(ImageSource source)
        {
            if (source != null && source.CanFreeze && !source.IsFrozen)
            {
                source.Freeze();
            }
            return source;
        }
        #endregion

        #region Main Icon Operations - Used by: FenceManager, PortalFenceManager, IconDragDropManager
        /// <summary>
        /// Main icon addition method with comprehensive icon handling and caching
        /// Used by: FenceManager.CreateFence, FenceManager.RefreshFenceContent, PortalFenceManager.AddIcon, IconDragDropManager.RefreshUI
        /// Category: Icon Rendering
        /// Moved from: FenceManager.AddIcon
        /// </summary>
        public static void AddIcon(dynamic icon, WrapPanel wpcont)
        {
            try
            {
                // Extract icon properties
                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ?
                    dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
                bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];
                bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
                bool isShortcut = Path.GetExtension(filePath).ToLower() == ".lnk";

                // Enhanced target path resolution with Unicode support and folder detection
                string targetPath = filePath;
                if (isShortcut)
                {
                    targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);

                    // Re-check if the target is actually a folder for Unicode shortcuts
                    if (!string.IsNullOrEmpty(targetPath) && System.IO.Directory.Exists(targetPath))
                    {
                        isFolder = true; // Update folder flag for shortcuts targeting folders
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Corrected isFolder to true for Unicode shortcut {filePath} targeting folder {targetPath}");
                    }
                }
                string arguments = iconDict.ContainsKey("Arguments") ? (string)iconDict["Arguments"] : null;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"AddIcon: {filePath} | IsFolder:{isFolder} | IsLink:{isLink} | IsShortcut:{isShortcut}");

                // Get icon spacing from fence data
                int iconSpacing = GetIconSpacingForFile(filePath);

                // Create main StackPanel container
                StackPanel sp = new StackPanel
                {
                    Margin = new Thickness(iconSpacing),
                    Width = 60
                };

                // Create and add icon image
                System.Windows.Controls.Image ico = new System.Windows.Controls.Image();

                // Apply icon size settings FIRST so placeholders scale correctly
                ApplyIconSize(ico, filePath);

                ImageSource cachedIcon = null;
                lock (iconCache)
                {
                    if (iconCache.TryGetValue(filePath, out var cached)) cachedIcon = cached;
                }

                if (cachedIcon != null)
                {
                    ico.Source = cachedIcon;
                    ApplyPostCreationIconOverride(ico, filePath);
                }
                else
                {
                    ico.Source = CreateFrozenBitmap("pack://application:,,,/Resources/file-WhiteX.png");

                    LazyIconLoader.RequestIcon(new IconLoadRequest
                    {
                        FilePath = filePath,
                        TargetPath = targetPath,
                        IsFolder = isFolder,
                        IsLink = isLink,
                        IsShortcut = isShortcut,
                        IconDict = iconDict,
                        TargetImage = ico,
                        OnLoaded = () => ApplyPostCreationIconOverride(ico, filePath)
                    });
                }

                sp.Children.Add(ico);

                // Create and add text label
                TextBlock lbl = CreateIconLabel(iconDict, filePath);
                sp.Children.Add(lbl);

                // Set tag for event handling (matches ClickEventAdder expectation)
                sp.Tag = new { FilePath = filePath, IsFolder = isFolder, Arguments = arguments };

                // Create tooltip
                CreateIconTooltip(sp, filePath, targetPath, arguments);

                // Add to container
                wpcont.Children.Add(sp);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Successfully added icon for {filePath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error in AddIcon for {icon}: {ex.Message}");
            }
        }

      

        public static ImageSource GetIconForFile(string targetPath, string filePath, bool isFolder = false,
                            bool isLink = false, bool isShortcut = false, IDictionary<string, object> iconDict = null)
        {
            try
            {
                // --- ULTIMATE SAFETY NET: Filename Fallback ---
                // If the file is locked and we can't read the target, but the file is literally named "Spotify", force the icon.
                string fileLower = Path.GetFileName(filePath).ToLower();
                if (fileLower.Contains("spotify"))
                {
                    var spotIcon = CreateFrozenBitmap("pack://application:,,,/Resources/spotify-White.png");
                    lock (iconCache) { iconCache[filePath] = spotIcon; }
                    return spotIcon;
                }
                if (fileLower.Contains("steam") && isShortcut)
                {
                    var steamIcon = CreateFrozenBitmap("pack://application:,,,/Resources/steam-White.png");
                    lock (iconCache) { iconCache[filePath] = steamIcon; }
                    return steamIcon;
                }

                // Check cache first (Thread-safe)
                lock (iconCache)
                {
                    if (iconCache.ContainsKey(filePath)) return iconCache[filePath];
                }

                ImageSource extractedIcon = null;

                // ==========================================
                // 1. RESOLVE THE TRUE TARGET OF .URL FILES
                // ==========================================
                string actualTarget = targetPath ?? "";

                if (Path.GetExtension(filePath)?.ToLower() == ".url" && System.IO.File.Exists(filePath))
                {
                    // For .url files, 'targetPath' is just the file path. We MUST extract the true destination.
                    string extractedUrl = CoreUtilities.ExtractUrlFromFile(filePath);
                    if (!string.IsNullOrEmpty(extractedUrl))
                    {
                        actualTarget = extractedUrl;
                    }
                    else
                    {
                        // Failsafe: Manually scrape the URL= line if the utility fails
                        try
                        {
                            string[] lines = System.IO.File.ReadAllLines(filePath);
                            foreach (string line in lines)
                            {
                                if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                                {
                                    actualTarget = line.Substring(4).Trim();
                                    break;
                                }
                            }
                        }
                        catch { }
                    }
                }

                string targetLower = actualTarget.ToLower();

                // ==========================================
                // 2. SCAN THE TRUE TARGET FOR CUSTOM PROTOCOLS
                // ==========================================
                if (targetLower.Contains("spotify:") || targetLower.Contains("spotify.com"))
                {
                    extractedIcon = CreateFrozenBitmap("pack://application:,,,/Resources/spotify-White.png");
                }
                else if (targetLower.Contains("steam://"))
                {
                    extractedIcon = CreateFrozenBitmap("pack://application:,,,/Resources/steam-White.png");
                }

                // ==========================================
                // 3. EXTRACT CUSTOM ICONS (Epic Games, etc.)
                // ==========================================
                if (extractedIcon == null && Path.GetExtension(filePath)?.ToLower() == ".url")
                {
                    extractedIcon = ExtractCustomIconFromUrl(filePath);
                    if (extractedIcon == null) extractedIcon = FreezeIcon(Utility.GetShellIcon(filePath, false));
                }

                // ==========================================
                // 4. STANDARD FALLBACKS
                // ==========================================
                if (extractedIcon == null)
                {
                    if (isLink || Path.GetExtension(filePath)?.ToLower() == ".url" || targetLower.StartsWith("http"))
                    {
                        extractedIcon = CreateFrozenBitmap("pack://application:,,,/Resources/link-White.png"); // check
                    }
                    else if (isShortcut)
                    {
                        extractedIcon = ExtractShortcutIcon(filePath, targetPath);
                    }
                    else if (isFolder)
                    {
                        extractedIcon = GetFolderIcon(targetPath);
                    }
                    else
                    {
                        extractedIcon = GetFileIcon(targetPath);
                    }
                }

                // Cache and return (Thread-safe)
                if (extractedIcon != null)
                {
                    lock (iconCache)
                    {
                        iconCache[filePath] = extractedIcon;
                    }
                }
                return extractedIcon;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Error extracting icon for {filePath}: {ex.Message}");
                var fallbackIcon = CreateFrozenBitmap("pack://application:,,,/Resources/file-WhiteX.png");
                lock (iconCache) { iconCache[filePath] = fallbackIcon; }
                return fallbackIcon;
            }
        }

        #endregion

        #region Icon Update Operations - Used by: FenceManager.UpdateIcon, TargetChecker
        /// <summary>
        /// Updates existing icon with new state (exists/missing)
        /// Used by: FenceManager.UpdateIcon, TargetChecker validation
        /// Category: Icon Updates
        /// Moved from: FenceManager.UpdateIcon
        /// </summary>
        public static void UpdateIcon(StackPanel sp, string filePath, bool isFolder)
        {
            try
            {
                var ico = sp.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                if (ico == null) return;

                if (!System.IO.File.Exists(filePath) && !Directory.Exists(filePath))
                {
                    // File is missing/deleted
                    var missingIcon = CreateFrozenBitmap("pack://application:,,,/Resources/file-WhiteX.png");
                    if (ico.Source != missingIcon)
                    {
                        ico.Source = missingIcon;
                        lock (iconCache) { iconCache[filePath] = missingIcon; }
                    }
                    return;
                }

                // File exists. Priority 1: Check the cache!
                ImageSource cachedIcon = null;
                lock (iconCache)
                {
                    if (iconCache.TryGetValue(filePath, out var cached))
                    {
                        cachedIcon = cached;
                    }
                }

                if (cachedIcon != null)
                {
                    // The icon is safely in the cache. No need to ruin it with re-extraction.
                    if (ico.Source != cachedIcon) ico.Source = cachedIcon;
                }
                else
                {
                    // It's missing from cache (TargetChecker fired before AddIcon finished).
                    bool isShortcut = Path.GetExtension(filePath).ToLower() == ".lnk";
                    bool isUrl = Path.GetExtension(filePath).ToLower() == ".url";

                    // Extract the true target before passing to the queue
                    string trueTarget = filePath;
                    try
                    {
                        if (isUrl)
                        {
                            string content = System.IO.File.ReadAllText(filePath);
                            var match = System.Text.RegularExpressions.Regex.Match(content, @"URL=([^\r\n]+)");
                            if (match.Success) trueTarget = match.Groups[1].Value.Trim();
                        }
                        else if (isShortcut)
                        {
                            trueTarget = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                            if (string.IsNullOrEmpty(trueTarget)) trueTarget = filePath;
                        }
                    }
                    catch { }

                    LazyIconLoader.RequestIcon(new IconLoadRequest
                    {
                        FilePath = filePath,
                        TargetPath = trueTarget,
                        IsFolder = isFolder,
                        IsLink = isUrl,
                        IsShortcut = isShortcut,
                        TargetImage = ico,
                        OnLoaded = () =>
                        {
                            // Absolute Safety Net: If it's Spotify, lock it in instantly
                            string lowerTarget = trueTarget.ToLower();
                            if (lowerTarget.Contains("spotify:") || lowerTarget.Contains("spotify.com"))
                            {
                                ico.Source = CreateFrozenBitmap("pack://application:,,,/Resources/spotify-White.png");
                            }

                            ApplyPostCreationIconOverride(ico, filePath);
                            lock (iconCache) { iconCache[filePath] = ico.Source; }
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error updating icon for {filePath}: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates fence JSON data after icon refresh operations
        /// Used by: Icon refresh operations, shortcut editing
        /// Category: Data Synchronization
        /// Moved from: FenceManager.UpdateFenceDataForIcon
        /// </summary>
        public static void UpdateFenceDataForIcon(string shortcutPath, string newDisplayName, NonActivatingWindow parentWindow)
        {
            try
            {
                if (string.IsNullOrEmpty(newDisplayName)) return;

                string fenceId = parentWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                var fence = FenceDataManager.FindFenceById(fenceId);
                if (fence?.ItemsType?.ToString() != "Data") return;

                var items = fence.Items as JArray;
                if (items == null) return;

                foreach (var item in items)
                {
                    string itemFilename = item["Filename"]?.ToString();
                    if (!string.IsNullOrEmpty(itemFilename) &&
                        string.Equals(Path.GetFullPath(itemFilename), Path.GetFullPath(shortcutPath), StringComparison.OrdinalIgnoreCase))
                    {
                        item["DisplayName"] = newDisplayName;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"Updated JSON data for: {itemFilename}");
                        break;
                    }
                }

                FenceDataManager.SaveFenceData();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error updating fence JSON: {ex.Message}");
            }
        }
        #endregion



        #region Icon Extraction Helpers - Internal Methods

        /// <summary>
        /// POST-CREATION OVERRIDE: 
        /// Runs immediately after the UI element is created. Forcefully swaps the icon
        /// if the target contains specific app protocols, bypassing all Windows extraction logic.
        /// </summary>
        private static void ApplyPostCreationIconOverride(System.Windows.Controls.Image ico, string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath)) return;

                string targetToScan = "";

                // 1. Extract the raw target string from the physical file
                if (filePath.EndsWith(".url", StringComparison.OrdinalIgnoreCase))
                {
                    targetToScan = CoreUtilities.ExtractUrlFromFile(filePath);
                    // Brute-force fallback if the utility parser fails
                    if (string.IsNullOrEmpty(targetToScan)) targetToScan = System.IO.File.ReadAllText(filePath);
                }
                else if (filePath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase))
                {
                    targetToScan = Utility.GetShortcutTarget(filePath);
                }

                if (string.IsNullOrEmpty(targetToScan)) return;

                // 2. Forcefully override the UI Image Source if the string matches
                if (targetToScan.IndexOf("spotify:", StringComparison.OrdinalIgnoreCase) >= 0 || targetToScan.IndexOf("spotify.com", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ico.Source = CreateFrozenBitmap("pack://application:,,,/Resources/spotify-White.png");
                    lock (iconCache) { iconCache[filePath] = ico.Source; } // Poison the cache safely
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling, $"Post-Creation Override: Forced Spotify icon for {filePath}");
                }
                else if (targetToScan.IndexOf("steam://", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    ico.Source = CreateFrozenBitmap("pack://application:,,,/Resources/steam-White.png");
                    lock (iconCache) { iconCache[filePath] = ico.Source; } // Poison the cache safely
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling, $"Post-Creation Override: Forced Steam icon for {filePath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Post-Creation override failed for {filePath}: {ex.Message}");
            }
        }



        /// <summary>
        /// Extracts icon from shortcut files with custom icon handling
        /// Used by: GetIconForFile
        /// Category: Shortcut Processing
        /// </summary>
        private static ImageSource ExtractShortcutIcon(string filePath, string targetPath)
        {
            try
            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);

                // Handle custom IconLocation with index - but prioritize missing folder icons
                if (!string.IsNullOrEmpty(shortcut.IconLocation))
                {
                    string[] iconParts = shortcut.IconLocation.Split(',');
                    string iconPath = iconParts[0];
                    int iconIndex = 0;

                    if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
                    {
                        iconIndex = parsedIndex;
                    }

                    // Check if this is a folder shortcut with missing target
                    bool isTargetMissing = string.IsNullOrEmpty(targetPath) ||
                                         (!System.IO.File.Exists(targetPath) && !Directory.Exists(targetPath));
                    bool isFolderShortcut = (!string.IsNullOrEmpty(targetPath) && Directory.Exists(targetPath)) ||
                                          (shortcut.TargetPath?.ToLower().Contains("explorer.exe") == true);

                    // If it's a missing folder shortcut with system folder icon, use our custom missing icon
                    if (isTargetMissing && isFolderShortcut && iconPath.ToLower().Contains("shell32.dll"))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                            $"Using folder-WhiteX.png for missing Unicode folder shortcut {filePath} instead of system icon");
                        return new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                    }

                    if (System.IO.File.Exists(iconPath))
                    {
                        var customIcon = ExtractIconFromFile(iconPath, iconIndex);
                        if (customIcon != null)
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                                $"Extracted custom icon at index {iconIndex} from {iconPath} for {filePath}");
                            return customIcon;
                        }
                    }
                }

                // Fallback to target icon
                if (!string.IsNullOrEmpty(targetPath))
                {
                    if (System.IO.File.Exists(targetPath))
                    {
                        return System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                    }
                    else if (Directory.Exists(targetPath))
                    {
                        return new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    }
                }

                // Final fallback
                return new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error extracting shortcut icon for {filePath}: {ex.Message}");
                return new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
            }
        }

        /// <summary>
        /// Extracts icon from file using Windows API
        /// Used by: ExtractShortcutIcon
        /// Category: Windows API
        /// </summary>
        public static ImageSource ExtractIconFromFile(string iconPath, int iconIndex)
        {
            try
            {
                IntPtr[] hIcon = new IntPtr[1];
                uint result = ExtractIconEx(iconPath, iconIndex, hIcon, null, 1);

                if (result > 0 && hIcon[0] != IntPtr.Zero)
                {
                    try
                    {
                        return Imaging.CreateBitmapSourceFromHIcon(
                            hIcon[0],
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions()
                        );
                    }
                    finally
                    {
                        DestroyIcon(hIcon[0]);
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error extracting icon from {iconPath} at index {iconIndex}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Gets appropriate folder icon based on existence
        /// Used by: GetIconForFile
        /// Category: Folder Icons
        /// </summary>
        private static ImageSource GetFolderIcon(string folderPath)
        {
            return Directory.Exists(folderPath) ?
                CreateFrozenBitmap("pack://application:,,,/Resources/folder-White.png") :
                CreateFrozenBitmap("pack://application:,,,/Resources/folder-WhiteX.png");
        }

        /// <summary>
        /// Gets file icon with fallback handling
        /// Used by: GetIconForFile
        /// Category: File Icons
        /// </summary>
        //private static ImageSource GetFileIcon(string filePath)
        //{
        //    try
        //    {
        //        if (System.IO.File.Exists(filePath))
        //        {
        //            return System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource();
        //        }
        //        else
        //        {
        //            return new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
        //            $"Error extracting file icon for {filePath}: {ex.Message}");
        //        return new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
        //    }
        //}


        private static ImageSource GetFileIcon(string filePath)
        {
            try
            {
                if (System.IO.File.Exists(filePath))
                {
                    return FreezeIcon(System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource());
                }
                else
                {
                    return CreateFrozenBitmap("pack://application:,,,/Resources/file-WhiteX.png");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error extracting file icon for {filePath}: {ex.Message}");

                // FIX: Office/AppX Icon Crash (.pptx, .docx). Fallback to Shell API
                ImageSource shellIcon = Utility.GetShellIcon(filePath, false);
                if (shellIcon != null) return FreezeIcon(shellIcon);

                return CreateFrozenBitmap("pack://application:,,,/Resources/file-WhiteX.png");
            }
        }

        #endregion

        /// <summary>
        /// Rips open a .url file and manually extracts the exact, game-specific icon 
        /// (Solves Epic Games, Steam, and Spotify custom shortcut icons)
        /// </summary>
        private static ImageSource ExtractCustomIconFromUrl(string urlFilePath)
        {
            try
            {
                if (!System.IO.File.Exists(urlFilePath)) return null;

                string[] lines = System.IO.File.ReadAllLines(urlFilePath);
                string iconFile = null;
                int iconIndex = 0;

                foreach (string line in lines)
                {
                    if (line.StartsWith("IconFile=", StringComparison.OrdinalIgnoreCase))
                    {
                        iconFile = line.Substring(9).Trim();
                    }
                    else if (line.StartsWith("IconIndex=", StringComparison.OrdinalIgnoreCase))
                    {
                        int.TryParse(line.Substring(10).Trim(), out iconIndex);
                    }
                }

                // If we found a specific custom icon path inside the file
                if (!string.IsNullOrEmpty(iconFile) && System.IO.File.Exists(iconFile))
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                        $"Found exact game/app IconFile in URL: {iconFile}");

                    // Extract it directly from the source .ico or .exe
                    var customIcon = ExtractIconFromFile(iconFile, iconIndex);
                    if (customIcon != null) return customIcon;

                    // Fallback to standard extraction if ExtractIconFromFile fails
                    return System.Drawing.Icon.ExtractAssociatedIcon(iconFile).ToImageSource();
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Custom URL icon scraper failed: {ex.Message}");
            }
            return null;
        }

        #region UI Creation Helpers - Internal Methods
        /// <summary>
        /// Creates text label for icon with truncation and styling
        /// Used by: AddIcon
        /// Category: UI Creation
        /// </summary>
        private static TextBlock CreateIconLabel(IDictionary<string, object> iconDict, string filePath)
        {
            string displayName = iconDict.ContainsKey("DisplayName") && !string.IsNullOrEmpty((string)iconDict["DisplayName"]) ?
                (string)iconDict["DisplayName"] : Path.GetFileNameWithoutExtension(filePath);

            // Apply truncation based on settings
            if (displayName.Length > SettingsManager.MaxDisplayNameLength)
            {
                displayName = displayName.Substring(0, SettingsManager.MaxDisplayNameLength) + "...";
            }

            return new TextBlock
            {
                Text = displayName,
                Foreground = System.Windows.Media.Brushes.White,
                TextAlignment = TextAlignment.Center,
                FontSize = 10,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 60,
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    Color = Colors.Black,
                    Direction = 315,
                    ShadowDepth = 1,
                    BlurRadius = 2
                }
            };
        }

        /// <summary>
        /// Creates tooltip for icon with file information
        /// Used by: AddIcon
        /// Category: UI Creation
        /// </summary>
        private static void CreateIconTooltip(StackPanel sp, string filePath, string targetPath, string arguments)
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

            sp.ToolTip = new ToolTip { Content = toolTipText };
        }

        /// <summary>
        /// Applies icon size settings based on fence configuration
        /// Used by: AddIcon
        /// Category: Icon Sizing
        /// </summary>
        private static void ApplyIconSize(System.Windows.Controls.Image ico, string filePath)
        {
            // Default size
            ico.Width = 40;
            ico.Height = 40;

            try
            {
                // Get icon size from fence data if available
                foreach (var fenceData in FenceDataManager.FenceData)
                {
                    if (fenceData.ItemsType?.ToString() == "Data")
                    {
                        var fenceItems = fenceData.Items as JArray;
                        if (fenceItems?.Any(i => i["Filename"]?.ToString() == filePath) == true)
                        {
                            string iconSize = fenceData.IconSize?.ToString() ?? "Medium";
                            ico.Width = ico.Height = CoreUtilities.GetIconSizePixels(iconSize);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Error applying icon size for {filePath}: {ex.Message}");
            }
        }
        #endregion

        #region Utility Methods - Helper Functions

        /// <summary>
        /// Applies safe hardware acceleration to existing panels without breaking drag/drop logic.
        /// Category: UI Optimization
        /// </summary>
        public static void OptimizeFencePanel(WrapPanel panel, ScrollViewer scrollViewer)
        {
            if (panel == null) return;

            panel.CacheMode = new BitmapCache { EnableClearType = false, RenderAtScale = 1.0, SnapsToDevicePixels = true };
            panel.UseLayoutRounding = true;

            if (scrollViewer != null)
            {
                VirtualizingPanel.SetScrollUnit(scrollViewer, ScrollUnit.Pixel);
                RenderOptions.SetBitmapScalingMode(scrollViewer, BitmapScalingMode.LowQuality);
            }
        }


        /// <summary>
        /// Enhanced AddIcon with fence context for proper sizing and spacing
        /// Used by: RefreshFenceContentSimple for tabbed fences
        /// Category: Icon Rendering
        /// </summary>
        public static void AddIconWithFenceContext(dynamic icon, WrapPanel wpcont, dynamic fence)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
            "=== AddIconWithFenceContext called ===");

                // Extract icon properties
                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ?
                    dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                string filePath = iconDict.ContainsKey("Filename") ?
                    (string)iconDict["Filename"] : "Unknown";
                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
                bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];
                bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
                string targetPath = isShortcut ? Utility.GetShortcutTarget(filePath) : filePath;
                string arguments = iconDict.ContainsKey("Arguments") ?
                    (string)iconDict["Arguments"] : null;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"AddIconWithFenceContext: {filePath} | IsFolder:{isFolder} | IsLink:{isLink} | IsShortcut:{isShortcut}");

                // Get icon spacing from fence context
                int iconSpacing = GetIconSpacingFromFence(fence);

                // Create main StackPanel container
                StackPanel sp = new StackPanel
                {
                    Margin = new Thickness(iconSpacing),
                    Width = 60
                };

                // Create and add icon image
                System.Windows.Controls.Image ico = new System.Windows.Controls.Image();

                // Apply icon size settings FIRST
                ApplyIconSizeFromFence(ico, fence);

                ImageSource cachedIcon = null;
                lock (iconCache)
                {
                    if (iconCache.TryGetValue(filePath, out var cached)) cachedIcon = cached;
                }

                if (cachedIcon != null)
                {
                    ico.Source = cachedIcon;
                    ApplyPostCreationIconOverride(ico, filePath);
                }
                else
                {
                    ico.Source = CreateFrozenBitmap("pack://application:,,,/Resources/file-WhiteX.png");

                    LazyIconLoader.RequestIcon(new IconLoadRequest
                    {
                        FilePath = filePath,
                        TargetPath = targetPath,
                        IsFolder = isFolder,
                        IsLink = isLink,
                        IsShortcut = isShortcut,
                        IconDict = iconDict,
                        TargetImage = ico,
                        OnLoaded = () => ApplyPostCreationIconOverride(ico, filePath)
                    });
                }

                sp.Children.Add(ico);



                // Create and add text label
                TextBlock lbl = CreateIconLabel(iconDict, filePath);
                sp.Children.Add(lbl);

                // Set tag for event handling (matches ClickEventAdder expectation)
                sp.Tag = new { FilePath = filePath, IsFolder = isFolder, Arguments = arguments };

                // Create tooltip
                CreateIconTooltip(sp, filePath, targetPath, arguments);

                // Add to container
                wpcont.Children.Add(sp);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Successfully added icon with fence context: {filePath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error in AddIconWithFenceContext: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets icon spacing directly from fence data
        /// Used by: AddIconWithFenceContext
        /// Category: Layout Utilities
        /// </summary>
        private static int GetIconSpacingFromFence(dynamic fence)
        {
            try
            {
                if (fence == null) return 5;

                int spacing = Convert.ToInt32(fence.IconSpacing?.ToString() ?? "5");
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Got icon spacing from fence: {spacing}");
                return spacing;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Error getting icon spacing from fence: {ex.Message}");
                return 5; // Default spacing
            }
        }

        /// <summary>
        /// Applies icon size directly from fence data
        /// Used by: AddIconWithFenceContext
        /// Category: Icon Sizing
        /// </summary>
        private static void ApplyIconSizeFromFence(System.Windows.Controls.Image ico, dynamic fence)
        {
            // Default size
            ico.Width = 40;
            ico.Height = 40;

            try
            {
                if (fence == null) return;

                string iconSize = fence.IconSize?.ToString() ?? "Medium";
                ico.Width = ico.Height = CoreUtilities.GetIconSizePixels(iconSize);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Applied icon size from fence: {iconSize} -> {ico.Width}x{ico.Height}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Error applying icon size from fence: {ex.Message}");
            }
        }


        /// <summary>
        /// Gets icon spacing from fence data for specific file
        /// Used by: AddIcon
        /// Category: Layout Utilities
        /// </summary>
        private static int GetIconSpacingForFile(string filePath)
        {
            try
            {
                foreach (var fenceData in FenceDataManager.FenceData)
                {
                    if (fenceData.ItemsType?.ToString() == "Data")
                    {
                        var fenceItems = fenceData.Items as JArray;
                        if (fenceItems?.Any(i => i["Filename"]?.ToString() == filePath) == true)
                        {
                            return Convert.ToInt32(fenceData.IconSpacing?.ToString() ?? "5");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"Error getting icon spacing for {filePath}: {ex.Message}");
            }

            return 5; // Default spacing
        }

        ///// <summary>
        ///// Converts icon size name to pixel dimensions
        ///// Used by: ApplyIconSize
        ///// Category: Icon Sizing
        ///// </summary>
        //private static int GetIconSizePixels(string iconSize)
        //{
        //    return iconSize switch
        //    {
        //        "Tiny" => 16,
        //        "Small" => 24,
        //        "Medium" => 32,
        //        "Large" => 48,
        //        "Huge" => 64,
        //        _ => 32
        //    };
        //}

        ///// <summary>
        ///// Checks if path contains only ASCII characters
        ///// Used by: UpdateIcon for Unicode handling
        ///// Category: Unicode Support
        ///// </summary>
        //private static bool IsAsciiPath(string path)
        //{
        //    return path.All(c => c <= 127);
        //}


        /// <summary>
        /// Handles Unicode shortcut icon updates
        /// Used by: UpdateIcon
        /// Category: Unicode Support
        /// </summary>
        private static ImageSource UpdateUnicodeShortcutIcon(string filePath, bool isFolder)
        {
            // --- ULTIMATE SAFETY NET: Protect Custom Icons from Unicode overwrite ---
            string fileLower = Path.GetFileName(filePath).ToLower();
            if (fileLower.Contains("spotify")) return CreateFrozenBitmap("pack://application:,,,/Resources/spotify-White.png");
            if (fileLower.Contains("steam")) return CreateFrozenBitmap("pack://application:,,,/Resources/steam-White.png");

            try
            {
                string iconTargetPath = GetShortcutTargetUnicodeSafe(filePath);
                if (!string.IsNullOrEmpty(iconTargetPath) && System.IO.File.Exists(iconTargetPath))
                {
                    return System.Drawing.Icon.ExtractAssociatedIcon(iconTargetPath).ToImageSource();
                }
                else
                {
                    return new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error updating Unicode shortcut icon for {filePath}: {ex.Message}");
                return new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
            }
        }

        /// <summary>
        /// Unicode-safe shortcut target resolution
        /// Used by: UpdateUnicodeShortcutIcon
        /// Category: Unicode Support
        /// </summary>
        private static string GetShortcutTargetUnicodeSafe(string shortcutPath)
        {
            try
            {
                // Use existing Utility method as fallback
                return Utility.GetShortcutTarget(shortcutPath);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Unicode shortcut resolution failed for {shortcutPath}: {ex.Message}");
                return string.Empty;
            }
        }
        #endregion

        #region Cache Management - Used by: System cleanup, debugging
        /// <summary>
        /// Clears icon cache for memory management
        /// Used by: System cleanup, debugging operations
        /// Category: Cache Management
        /// </summary>
        public static void ClearIconCache()
        {
            iconCache.Clear();
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                "Icon cache cleared");
        }

        /// <summary>
        /// Gets cache statistics for monitoring
        /// Used by: Debugging, performance monitoring
        /// Category: Cache Management
        /// </summary>
        public static int GetCacheSize()
        {
            return iconCache.Count;
        }
        #endregion
    }
}
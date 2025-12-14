using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using IWshRuntimeLibrary;
using Newtonsoft.Json.Linq;
using Microsoft.VisualBasic;

namespace Desktop_Fences
{
    public class PortalFenceManager
    {
        // New field for the active filter
        private string _currentFilter = null;


        private readonly dynamic _fence;
        private readonly WrapPanel _wpcont;
        private readonly FileSystemWatcher _watcher;
        private string _targetFolderPath;
        private readonly Dispatcher _dispatcher;
        private readonly Dictionary<string, (FileSystemEventArgs Args, string OldPath)> _pendingEvents = new Dictionary<string, (FileSystemEventArgs, string)>();
        private readonly DispatcherTimer _debounceTimer;



        // --- FILTERING ENGINE START ---

        /// <summary>
        /// Updates the current filter and refreshes the visibility of all items.
        /// Publicly called by FenceManager when the user types in the filter bar.
        /// </summary>
        public void ApplyFilter(string filterText)
        {
            _currentFilter = filterText;
            _dispatcher.Invoke(() =>
            {
                foreach (StackPanel sp in _wpcont.Children.OfType<StackPanel>())
                {
                    if (sp.Tag != null)
                    {
                        // Safely retrieve path from anonymous type or object
                        string path = sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)?.ToString();
                        if (!string.IsNullOrEmpty(path))
                        {
                            sp.Visibility = ShouldShowItem(path) ? Visibility.Visible : Visibility.Collapsed;
                        }
                    }
                }
            });
        }



        /// <summary>
        /// Determines if a file should be visible based on the current filter.
        /// Supports "Smart Match" if NoWildcardsOnPortalFilter is enabled.
        /// </summary>
        private bool ShouldShowItem(string filePath)
        {
            if (string.IsNullOrWhiteSpace(_currentFilter)) return true;

            string fileName = System.IO.Path.GetFileName(filePath);
            var terms = _currentFilter.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(t => t.Trim())
                                      .ToList();

            bool hasIncludeRules = terms.Any(t => !t.StartsWith(">"));
            bool matchesInclude = !hasIncludeRules;
            bool matchesExclude = false;

            foreach (var term in terms)
            {
                if (string.IsNullOrEmpty(term)) continue;

                string pattern = term;
                bool isExclude = false;

                // 1. Identify Exclusion
                if (pattern.StartsWith(">"))
                {
                    isExclude = true;
                    pattern = pattern.Substring(1); // Remove '>' prefix
                }

                // 2. Apply Smart Wildcards (Hidden Option)
                // Logic: If user wants "No Wildcards", we treat text as "Contains".
                // We only auto-wrap if the user hasn't typed wildcards themselves.
                if (SettingsManager.NoWildcardsOnPortalFilter)
                {
                    if (!pattern.Contains("*") && !pattern.Contains("?"))
                    {
                        pattern = "*" + pattern + "*";
                    }
                }

                // 3. Match
                if (isExclude)
                {
                    if (IsMatch(fileName, pattern))
                    {
                        matchesExclude = true;
                        break; // Hard fail
                    }
                }
                else
                {
                    if (IsMatch(fileName, pattern))
                    {
                        matchesInclude = true;
                    }
                }
            }

            return !matchesExclude && matchesInclude;
        }




        /// <summary>
        /// Simple glob matching (* and ?)
        /// </summary>
        private bool IsMatch(string text, string pattern)
        {
            // Use VB's Like operator or simple Regex. 
            // For a dependency-free C# solution, we convert glob to regex.
            try
            {
                string regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
                                      .Replace(@"\*", ".*")
                                      .Replace(@"\?", ".") + "$";
                return System.Text.RegularExpressions.Regex.IsMatch(text, regexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            }
            catch { return false; }
        }
        // --- FILTERING ENGINE END ---






        public PortalFenceManager(dynamic fence, WrapPanel wpcont)
        {
            _fence = fence;
            _wpcont = wpcont;
            _dispatcher = _wpcont.Dispatcher;

            // Initialize debounce timer with longer interval for Excel temp files
            _debounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500) // Increased for better stability
            };
            _debounceTimer.Tick += ProcessPendingEvents;

            // Extract folder path
            IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            _targetFolderPath = fenceDict.ContainsKey("Path") ? fenceDict["Path"]?.ToString() : null;

            // FIX: Load saved filter immediately on startup
            if (fenceDict.ContainsKey("FilterString"))
            {
                _currentFilter = fenceDict["FilterString"]?.ToString();
            }

            if (string.IsNullOrEmpty(_targetFolderPath))
            {
                throw new Exception("No folder path defined for Portal Fence. Please recreate the fence.");
            }

            if (!Directory.Exists(_targetFolderPath))
            {
                throw new Exception($"The folder '{_targetFolderPath}' does not exist. Please update the Portal Fence settings.");
            }

            _watcher = new FileSystemWatcher(_targetFolderPath)
            {
                EnableRaisingEvents = true,
                IncludeSubdirectories = false
            };
            _watcher.Created += (s, e) => QueueEvent(e);
            _watcher.Deleted += (s, e) => QueueEvent(e);
            _watcher.Renamed += (s, e) => QueueEvent(e, e.OldFullPath);

            InitializeFenceContents();

          //  // --- TEST CODE START ---
          //  // Hardcode a filter to prove the engine works.
         //   // This simulates a user typing "*.txt" into the filter bar.
         //   ApplyFilter("*.txt");
          //  // --- TEST CODE END ---
        }

        private void QueueEvent(FileSystemEventArgs e, string oldPath = null)
        {
            // Filter out Excel and other temporary files immediately
            if (CoreUtilities.IsTemporaryFile(e.FullPath))
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Ignoring temporary file: {e.FullPath}");
                return;
            }

            lock (_pendingEvents)
            {
                _pendingEvents[e.FullPath] = (e, oldPath);
                _debounceTimer.Stop();
                _debounceTimer.Start();
            }
        }

     
        private void ProcessPendingEvents(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            Dictionary<string, (FileSystemEventArgs Args, string OldPath)> events;
            lock (_pendingEvents)
            {
                events = new Dictionary<string, (FileSystemEventArgs Args, string OldPath)>(_pendingEvents);
                _pendingEvents.Clear();
            }

            _dispatcher.Invoke(() =>
            {
                foreach (var evt in events.Values)
                {
                    switch (evt.Args.ChangeType)
                    {
                        case WatcherChangeTypes.Created:
                            // Double-check file still exists and isn't temporary before adding
                            if ((System.IO.File.Exists(evt.Args.FullPath) || Directory.Exists(evt.Args.FullPath)) &&
                                    !CoreUtilities.IsTemporaryFile(evt.Args.FullPath))
                            {
                                AddIcon(evt.Args.FullPath);
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Processed queued Created event for {evt.Args.FullPath}");
                            }
                            break;
                        case WatcherChangeTypes.Deleted:
                            RemoveIcon(evt.Args.FullPath);
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Processed queued Deleted event for {evt.Args.FullPath}");
                            break;
                        case WatcherChangeTypes.Renamed:
                            RemoveIcon(evt.OldPath);
                            if ((System.IO.File.Exists(evt.Args.FullPath) || Directory.Exists(evt.Args.FullPath)) &&
                                !CoreUtilities.IsTemporaryFile(evt.Args.FullPath))
                            {
                                AddIcon(evt.Args.FullPath);
                            }
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Processed queued Renamed event: {evt.OldPath} -> {evt.Args.FullPath}");
                            break;
                    }
                }
                VerifyFenceContents();
            });
        }

        private void AddIcon(string path)
        {
            // Enhanced filter during add to prevent duplicates
            try
            {
                // Double-check file exists and isn't temporary
                if (!System.IO.File.Exists(path) && !Directory.Exists(path))
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"File doesn't exist, skipping: {path}");
                    return;
                }

                if (CoreUtilities.IsTemporaryFile(path))
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Skipping temporary file: {path}");
                    return;
                }

                // Check if icon already exists in UI
                var existingPanel = _wpcont.Children.OfType<StackPanel>()
                    .FirstOrDefault(sp => sp.Tag != null &&
                                    sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)?.ToString() == path);

                if (existingPanel != null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Icon already exists, skipping: {path}");
                    return;
                }

                FileAttributes attributes = System.IO.File.GetAttributes(path);

                // Skip hidden files and folders
                if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Skipping hidden item: {path}");
                    return;
                }

                // Optional: Also skip system files if desired
                if ((attributes & FileAttributes.System) == FileAttributes.System)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Skipping system item: {path}");
                    return;
                }
            }
            catch (Exception ex)
            {
                // If we can't get attributes, log and continue (file might be inaccessible)
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Cannot check attributes for {path}: {ex.Message}");
                return;
            }

            dynamic icon = new System.Dynamic.ExpandoObject();
            IDictionary<string, object> iconDict = icon;
            iconDict["Filename"] = path;

            bool isFolder = false;
            try { isFolder = Directory.Exists(path); } catch { }
            iconDict["IsFolder"] = isFolder;
            // --- RESTORED: Network Path Detection ---
            iconDict["IsNetwork"] = FenceManager.IsNetworkPath(path);


            string displayName;

            try
            {
                // FIX: Handle Extensions based on Global Setting
                if (SettingsManager.ShowPortalExtensions && !isFolder)
                {
                    // Force display name WITH extension
                    displayName = Path.GetFileName(path);
                }
                else
                {
                    if (isFolder)
                    {
                        // Folders → keep full name even if they contain dots
                        displayName = Path.GetFileName(path);
                    }
                    else
                    {
                        // Files → strip extension (default behavior)
                        displayName = Path.GetFileNameWithoutExtension(path);
                    }
                }
            }
            catch
            {
                // Fallback: act like it's a file without extension
                displayName = Path.GetFileNameWithoutExtension(path);
            }

            iconDict["DisplayName"] = displayName;

            // --- FIX: ONE CALL ONLY ---
            // We use the new signature that passes '_fence' context.
            // This applies the custom settings (Size, Color, etc.) immediately.
            FenceManager.AddIcon(icon, _wpcont, _fence);

            // Now we grab the StackPanel that was just added to attach logic
            StackPanel sp = _wpcont.Children[_wpcont.Children.Count - 1] as StackPanel;
            if (sp != null)
            {
                // FIX: Apply filter immediately upon creation
                sp.Visibility = ShouldShowItem(path) ? Visibility.Visible : Visibility.Collapsed;

                FenceManager.ClickEventAdder(sp, path, Directory.Exists(path));

          
                // Create and attach context menu
                ContextMenu contextMenu = new ContextMenu();

                // 1. Copy Item (File Object)
                MenuItem copyFileItem = new MenuItem { Header = "Copy Item" };
                copyFileItem.Click += (s, e) =>
                {
                    try
                    {
                        // Add file to clipboard as a FileDropList (Standard Windows Copy)
                        var paths = new System.Collections.Specialized.StringCollection();
                        paths.Add(path);
                        Clipboard.SetFileDropList(paths);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Copied item to clipboard: {path}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error copying item: {ex.Message}");
                    }
                };
                contextMenu.Items.Add(copyFileItem);

                // 2. Cut Item (File Object with Move Effect)
                MenuItem cutFileItem = new MenuItem { Header = "Cut Item" };
                cutFileItem.Click += (s, e) =>
                {
                    try
                    {
                        var paths = new System.Collections.Specialized.StringCollection();
                        paths.Add(path);

                        // Create a DataObject to hold both the file list and the "Move" flag
                        DataObject data = new DataObject();
                        data.SetFileDropList(paths);

                        // Set "Preferred DropEffect" to Move (Byte value 2)
                        // This tells Windows Explorer to perform a MOVe operation on Paste
                        byte[] moveEffect = new byte[] { 2, 0, 0, 0 };
                        System.IO.MemoryStream stream = new System.IO.MemoryStream(moveEffect);
                        data.SetData("Preferred DropEffect", stream);

                        Clipboard.SetDataObject(data, true);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Cut item to clipboard: {path}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error cutting item: {ex.Message}");
                    }
                };
                contextMenu.Items.Add(cutFileItem);

                // 3. Rename item (Existing)
                MenuItem renameItem = new MenuItem { Header = "Rename item" };
                renameItem.Click += (s, e) => RenameItem(path, sp);
                contextMenu.Items.Add(renameItem);

                // 4. Delete item (Existing)
                MenuItem deleteItem = new MenuItem { Header = "Delete item" };
                deleteItem.Click += (s, e) => DeleteItem(path, sp);
                contextMenu.Items.Add(deleteItem);

                // 5. Separator
                contextMenu.Items.Add(new Separator());

                // 6. Copy path (Existing - Moved to bottom)
                MenuItem copyPathItem = new MenuItem { Header = "Copy path" };
                copyPathItem.Click += (s, e) => CopyPathOrTarget(path);
                contextMenu.Items.Add(copyPathItem);

                sp.ContextMenu = contextMenu;



            }
        }

        private void RenameItem(string currentPath, StackPanel sp)
        {
            try
            {
                string currentName = Path.GetFileNameWithoutExtension(currentPath);
                string extension = Path.GetExtension(currentPath);

                // Simple input dialog (you can replace with a proper dialog if you have one)
                string newName = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter new name:",
                    "Rename Item",
                    currentName);

                if (string.IsNullOrEmpty(newName) || newName == currentName)
                    return;

                string newPath = Path.Combine(Path.GetDirectoryName(currentPath), newName + extension);

                // Check if target name already exists
                if (System.IO.File.Exists(newPath) || Directory.Exists(newPath))
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("A file or folder with that name already exists.", "Rename Error");
                    return;
                }

                // Perform the rename
                if (Directory.Exists(currentPath))
                {
                    Directory.Move(currentPath, newPath);
                }
                else if (System.IO.File.Exists(currentPath))
                {
                    System.IO.File.Move(currentPath, newPath);
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Renamed {currentPath} to {newPath}");

                // The FileSystemWatcher will automatically handle UI updates
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to rename {currentPath}: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to rename item: {ex.Message}", "Rename Error");
            }
        }

        private void InitializeFenceContents()
        {
            _wpcont.Children.Clear();
            if (Directory.Exists(_targetFolderPath))
            {
                foreach (string path in Directory.GetFileSystemEntries(_targetFolderPath))
                {
                    // Filter out temporary files during initialization
                    if (CoreUtilities.IsTemporaryFile(path))
                        continue;

                    // Filter out hidden files during initialization too
                    try
                    {
                        FileAttributes attributes = System.IO.File.GetAttributes(path);

                        // Skip hidden and system files
                        if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden ||
                            (attributes & FileAttributes.System) == FileAttributes.System)
                        {
                            continue;
                        }
                    }
                    catch
                    {
                        // Skip files we can't read attributes for
                        continue;
                    }

                    AddIcon(path);
                }
            }
            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Initialized fence contents for {_targetFolderPath} with {_wpcont.Children.Count} items (hidden and temporary files excluded)");
        }

        private void VerifyFenceContents()
        {
            if (!Directory.Exists(_targetFolderPath))
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Target folder {_targetFolderPath} does not exist, skipping verification");
                return;
            }

            // Get only non-hidden, non-temporary files
            var currentFiles = Directory.GetFileSystemEntries(_targetFolderPath)
                .Where(path =>
                {
                    if (CoreUtilities.IsTemporaryFile(path))
                        return false;

                    try
                    {
                        FileAttributes attributes = System.IO.File.GetAttributes(path);
                        return (attributes & FileAttributes.Hidden) != FileAttributes.Hidden &&
                               (attributes & FileAttributes.System) != FileAttributes.System;
                    }
                    catch
                    {
                        return false; // Skip files we can't read
                    }
                })
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var panelsToRemove = _wpcont.Children.OfType<StackPanel>()
                .Where(sp => sp.Tag != null &&
                             sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag) is string filePath &&
                             !currentFiles.Contains(filePath))
                .ToList();

            foreach (var sp in panelsToRemove)
            {
                _wpcont.Children.Remove(sp);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Removed stale icon for {sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)} during verification");
            }

            var existingPaths = _wpcont.Children.OfType<StackPanel>()
                .Select(sp => sp.Tag?.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)?.ToString())
                .Where(p => p != null)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var path in currentFiles)
            {
                if (!existingPaths.Contains(path))
                {
                    AddIcon(path);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Added missing icon for {path} during verification");
                }
            }
        }

        private void CopyPathOrTarget(string path)
        {
            try
            {
                string pathToCopy;
                if (Path.GetExtension(path).ToLower() == ".lnk")
                {
                    // If it's a shortcut, get the target path
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(path);
                    pathToCopy = shortcut.TargetPath;
                }
                else
                {
                    // Otherwise, copy the folder path (portal fence path)
                    pathToCopy = Path.GetDirectoryName(path); // Gets the parent directory
                }

                // Copy to clipboard
                Clipboard.SetText(pathToCopy);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Copied path to clipboard: {pathToCopy}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Failed to copy path for {path}: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Unable to copy path.", "Error");
            }
        }

        private void DeleteItem(string path, StackPanel sp)
        {
            bool UseRecycleBin = SettingsManager.UseRecycleBin;
            if (UseRecycleBin == true)
            {
                try
                {
                    // First, check if the item exists
                    if (!Directory.Exists(path) && !System.IO.File.Exists(path))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Item not found for deletion: {path}");
                        return;
                    }

                    // Use SHFileOperation to move to recycle bin
                    SHFILEOPSTRUCT shf = new SHFILEOPSTRUCT();
                    shf.wFunc = FO_DELETE;
                    shf.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION;
                    shf.pFrom = path + '\0' + '\0'; // Double null-terminated string

                    int result = SHFileOperation(ref shf);

                    if (result != 0)
                    {
                        throw new Exception($"Failed to move to recycle bin (error code: {result})");
                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Moved to recycle bin: {path}");

                    // Remove the icon from the UI
                    _wpcont.Children.Remove(sp);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, ($"Removed icon for {path} from UI"));
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to move item {path} to recycle bin: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Unable to move item to recycle bin.", "Error");
                }
            }
            else
            {
                try
                {
                    if (Directory.Exists(path))
                    {
                        // Delete folder
                        Directory.Delete(path, true); // true for recursive deletion
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Deleted folder: {path}");
                    }
                    else if (System.IO.File.Exists(path))
                    {
                        // Delete file
                        System.IO.File.Delete(path);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Deleted file: {path}");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Item not found for deletion: {path}");
                        return;
                    }

                    // Remove the icon from the UI
                    _wpcont.Children.Remove(sp);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Removed icon for {path} from UI");
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Failed to delete item {path}: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Unable to delete item.", "Error");
                }
            }
        }

        // Corrected Win32 API declarations
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            public uint wFunc;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pFrom;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pTo;
            public ushort fFlags;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHFileOperation([In] ref SHFILEOPSTRUCT lpFileOp);

        const uint FO_DELETE = 0x0003;
        const ushort FOF_ALLOWUNDO = 0x0040;
        const ushort FOF_NOCONFIRMATION = 0x0010;

        private void RemoveIcon(string path)
        {
            var sp = _wpcont.Children.OfType<StackPanel>().FirstOrDefault(s =>
                s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == path);
            if (sp != null)
            {
                _wpcont.Children.Remove(sp);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Successfully removed icon for {path}");
            }
            else
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Failed to find StackPanel for {path} in RemoveIcon");
            }
        }

        // TEST: Filter for only text files (REMOVE AFTER TEST)
        // ApplyFilter("*.txt");



        /// <summary>
        /// Safely switches the monitored folder without destroying the fence window.
        /// Used for the "Dive In" navigation feature.
        /// </summary>
        public void NavigateTo(string newPath)
        {
            try
            {
                if (string.IsNullOrEmpty(newPath) || !Directory.Exists(newPath))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Cannot navigate to invalid path: {newPath}");
                    return;
                }

                // 1. Suspend Watcher to prevent event spam during switch
                bool wasEnable = _watcher.EnableRaisingEvents;
                _watcher.EnableRaisingEvents = false;

                // 2. Clear UI
                _dispatcher.Invoke(() => _wpcont.Children.Clear());

                // 3. Switch Target
                _targetFolderPath = newPath;
                _watcher.Path = newPath; // FileSystemWatcher supports dynamic path changing

                // 4. Reload Content
                InitializeFenceContents();

                // 5. Resume Watcher
                _watcher.EnableRaisingEvents = wasEnable;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Portal Fence navigated to: {newPath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Navigation failed: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Could not navigate to folder.\n{ex.Message}", "Navigation Error");
            }
        }

        public void Dispose()
        {
            _watcher?.Dispose();
            _debounceTimer?.Stop();
            _debounceTimer.Tick -= ProcessPendingEvents;
        }
    }
}
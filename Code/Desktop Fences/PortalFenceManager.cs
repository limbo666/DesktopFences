using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Controls;
using Newtonsoft.Json.Linq;
using System.Windows.Threading;
using System.Linq;
using System.Reflection;
using IWshRuntimeLibrary;
using System.Diagnostics;
using System.Windows;
using System.Runtime.InteropServices;

namespace Desktop_Fences
{
    public class PortalFenceManager
    {
        private readonly dynamic _fence;
        private readonly WrapPanel _wpcont;
        private readonly FileSystemWatcher _watcher;
        private string _targetFolderPath;
        private readonly Dispatcher _dispatcher;
        private readonly Dictionary<string, (FileSystemEventArgs Args, string OldPath)> _pendingEvents = new();
        private readonly DispatcherTimer _debounceTimer;
    
       
        public PortalFenceManager(dynamic fence, WrapPanel wpcont)
        {
            _fence = fence;
            _wpcont = wpcont;
            _dispatcher = _wpcont.Dispatcher;

            // Initialize debounce timer
            _debounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _debounceTimer.Tick += ProcessPendingEvents;

            // Extract folder path
            IDictionary<string, object> fenceDict = fence is IDictionary<string, object> dict ? dict : ((JObject)fence).ToObject<IDictionary<string, object>>();
            _targetFolderPath = fenceDict.ContainsKey("Path") ? fenceDict["Path"]?.ToString() : null;

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
        }

        private void QueueEvent(FileSystemEventArgs e, string oldPath = null)
        {
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
                            AddIcon(evt.Args.FullPath);
                            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Processed queued Created event for {evt.Args.FullPath}");
                          //  FenceManager.Log(FenceManager.LogLevel.Info, FenceManager.LogCategory.General, $"Processed queued Created event for {evt.Args.FullPath}");
                            break;
                        case WatcherChangeTypes.Deleted:
                            RemoveIcon(evt.Args.FullPath);
                            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Processed queued Deleted event for {evt.Args.FullPath}");
                            break;
                        case WatcherChangeTypes.Renamed:
                            RemoveIcon(evt.OldPath);
                            AddIcon(evt.Args.FullPath);
                            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Processed queued Renamed event: {evt.OldPath} -> {evt.Args.FullPath}");
                            break;
                    }
                }
                VerifyFenceContents();
            });
        }

        private void InitializeFenceContents()
        {
            _wpcont.Children.Clear();
            if (Directory.Exists(_targetFolderPath))
            {
                foreach (string path in Directory.GetFileSystemEntries(_targetFolderPath))
                {
                    AddIcon(path);
                }
            }
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Initialized fence contents for {_targetFolderPath} with {_wpcont.Children.Count} items");
        }

        private void VerifyFenceContents()
        {
            if (!Directory.Exists(_targetFolderPath))
            {
                FenceManager.Log(FenceManager.LogLevel.Warn, FenceManager.LogCategory.General, $"Target folder {_targetFolderPath} does not exist, skipping verification");
                return;
            }

            var currentFiles = Directory.GetFileSystemEntries(_targetFolderPath).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var panelsToRemove = _wpcont.Children.OfType<StackPanel>()
                .Where(sp => sp.Tag != null &&
                             sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag) is string filePath &&
                             !currentFiles.Contains(filePath))
                .ToList();
            foreach (var sp in panelsToRemove)
            {
                _wpcont.Children.Remove(sp);
                FenceManager.Log(FenceManager.LogLevel.Info, FenceManager.LogCategory.General, $"Removed stale icon for {sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)} during verification");
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
                    FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Added missing icon for {path} during verification");
                }
            }
        }

        private void AddIcon(string path)
        {
            dynamic icon = new System.Dynamic.ExpandoObject();
            IDictionary<string, object> iconDict = icon;
            iconDict["Filename"] = path;
            iconDict["IsFolder"] = Directory.Exists(path);
            iconDict["DisplayName"] = Path.GetFileNameWithoutExtension(path);





            //FenceManager.AddIcon(icon, _wpcont);
            //StackPanel sp = _wpcont.Children[_wpcont.Children.Count - 1] as StackPanel;
            //if (sp != null)
            //{
            //    FenceManager.ClickEventAdder(sp, path, Directory.Exists(path));
            //    Log($"Added icon for {path}");
            //}
            //else
            //{
            //    Log($"Failed to add StackPanel for {path}");
            //}

            FenceManager.AddIcon(icon, _wpcont);
            StackPanel sp = _wpcont.Children[_wpcont.Children.Count - 1] as StackPanel;
            if (sp != null)
            {
                FenceManager.ClickEventAdder(sp, path, Directory.Exists(path));

                // Create and attach context menu
                ContextMenu contextMenu = new ContextMenu();

                // "Copy path (or target)" menu item
                MenuItem copyPathItem = new MenuItem { Header = "Copy path" };
                copyPathItem.Click += (s, e) => CopyPathOrTarget(path);
                contextMenu.Items.Add(copyPathItem);

                // "Delete item" menu item
                MenuItem deleteItem = new MenuItem { Header = "Delete item" };
                deleteItem.Click += (s, e) => DeleteItem(path, sp);
                contextMenu.Items.Add(deleteItem);

                sp.ContextMenu = contextMenu;
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
                FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.UI, $"Copied path to clipboard: {pathToCopy}");
            }
            catch (Exception ex)
            {
                FenceManager.Log(FenceManager.LogLevel.Error, FenceManager.LogCategory.UI, $"Failed to copy path for {path}: {ex.Message}");
                //  MessageBox.Show("Unable to copy path.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Unable to copy path.", "Error");
            }
        }
        // delete  item handler
        private void DeleteItem(string path, StackPanel sp)
        {
           bool  UseRecycleBin= SettingsManager.UseRecycleBin; //  UseRecycleBin is a property in SettingsManager
            if (UseRecycleBin== true) { 
            
            try
            {
                // First, check if the item exists
                if (!Directory.Exists(path) && !System.IO.File.Exists(path))
                {
                        FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Item not found for deletion: {path}");
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

                    FenceManager.Log(FenceManager.LogLevel.Info, FenceManager.LogCategory.UI, $"Moved to recycle bin: {path}");

                // Remove the icon from the UI
                _wpcont.Children.Remove(sp);
                    FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.UI, ($"Removed icon for {path} from UI"));
            }
            catch (Exception ex)
            {
                    FenceManager.Log(FenceManager.LogLevel.Error, FenceManager.LogCategory.General, $"Failed to move item {path} to recycle bin: {ex.Message}");
                    //   MessageBox.Show("Unable to move item to recycle bin.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Unable to move item to recycle bin.", "Error");
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
                        FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Deleted folder: {path}");
                    }
                    else if (System.IO.File.Exists(path))
                    {
                        // Delete file
                        System.IO.File.Delete(path);
                        FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Deleted file: {path}");
                    }
                    else
                    {
                        FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Item not found for deletion: {path}");
                        return;
                    }

                    // Remove the icon from the UI
                    _wpcont.Children.Remove(sp);
                    FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Removed icon for {path} from UI");
                }
                catch (Exception ex)
                {
                    FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Failed to delete item {path}: {ex.Message}");
                    // MessageBox.Show("Unable to delete item.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Unable to delete item.", "Error");
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
                FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Successfully removed icon for {path}");
            }
            else
            {
                FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Failed to find StackPanel for {path} in RemoveIcon");
            }
        }
        public void Dispose()
        {
            _watcher?.Dispose();
            _debounceTimer?.Stop();
            _debounceTimer.Tick -= ProcessPendingEvents;
        }
        //private void Log(string message)
        //{
        //    bool isLogEnabled = true;
        //    if (isLogEnabled)
        //    {
        //        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
        //        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
        //    }
        //}
    }
}
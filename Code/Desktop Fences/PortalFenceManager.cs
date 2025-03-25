using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Controls;
using Newtonsoft.Json.Linq;
using System.Windows.Threading;
using System.Linq;
using System.Reflection;

namespace Desktop_Fences
{
    public class PortalFenceManager
    {
        private readonly dynamic _fence;
        private readonly WrapPanel _wpcont;
        private readonly FileSystemWatcher _watcher;
        private string _targetFolderPath;
        private readonly Dispatcher _dispatcher;

        public PortalFenceManager(dynamic fence, WrapPanel wpcont)
        {
            _fence = fence;
            _wpcont = wpcont;
            _dispatcher = _wpcont.Dispatcher; // Λαμβάνουμε το Dispatcher του UI thread

            // Παίρνουμε το Path από το fence αντί για Items
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
            _watcher.Created += OnFileCreated;
            _watcher.Deleted += OnFileDeleted;
            _watcher.Renamed += OnFileRenamed;

            InitializeFenceContents();
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
        }

        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            _dispatcher.Invoke(() => AddIcon(e.FullPath));
            Log($"Added new item to portal fence {_fence.Title}: {e.FullPath}");
        }

        private void OnFileDeleted(object sender, FileSystemEventArgs e)
        {
            _dispatcher.Invoke(() => RemoveIcon(e.FullPath));
            Log($"Removed item from portal fence {_fence.Title}: {e.FullPath}");
        }

        private void OnFileRenamed(object sender, RenamedEventArgs e)
        {
            _dispatcher.Invoke(() =>
            {
                RemoveIcon(e.OldFullPath);
                AddIcon(e.FullPath);
            });
            Log($"Renamed item in portal fence {_fence.Title}: {e.OldFullPath} -> {e.FullPath}");
        }

        private void AddIcon(string path)
        {
            dynamic icon = new System.Dynamic.ExpandoObject();
            IDictionary<string, object> iconDict = icon;
            iconDict["Filename"] = path;
            iconDict["IsFolder"] = Directory.Exists(path);
            iconDict["DisplayName"] = Path.GetFileNameWithoutExtension(path);

            FenceManager.AddIcon(icon, _wpcont);
            StackPanel sp = _wpcont.Children[_wpcont.Children.Count - 1] as StackPanel;
            if (sp != null)
            {
                FenceManager.ClickEventAdder(sp, path, Directory.Exists(path));
            }
        }

        private void RemoveIcon(string path)
        {
            var sp = _wpcont.Children.OfType<StackPanel>().FirstOrDefault(s => (s.Tag as string) == path);
            if (sp != null)
            {
                _wpcont.Children.Remove(sp);
            }
        }

        private void Log(string message)
        {
            bool isLogEnabled = true; // Αν δεν έχεις SettingsManager, το αφήνουμε true προσωρινά
            if (isLogEnabled)
            {
                string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
            }
        }
    }
}
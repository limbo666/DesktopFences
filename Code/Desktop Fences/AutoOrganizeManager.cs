using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Newtonsoft.Json;
using Microsoft.VisualBasic.FileIO; // Required for sending to Recycle Bin

namespace Desktop_Fences
{
    public enum RuleConflictAction { Rename, Overwrite, Skip }

    public class OrganizeRule
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; }
        public string Extensions { get; set; } = "";
        public string NameContains { get; set; } = "";
        public string TargetFolderPath { get; set; }
        public RuleConflictAction ConflictAction { get; set; } = RuleConflictAction.Rename;
        public bool AutoCreateFence { get; set; } = false;
        public bool IsEnabled { get; set; } = true;
        public int Priority { get; set; } = 0;
        public DateTime? LastRun { get; set; } // NEW: Tracks last execution
    }

    public static class AutoOrganizeManager
    {
        private static FileSystemWatcher _watcher;
        public static List<OrganizeRule> Rules = new List<OrganizeRule>(); // Expose to the UI
        private static string _rulesFilePath => ProfileManager.GetProfileFilePath("auto_organize.json");
        private static readonly string _desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public static void Initialize()
        {
            LoadRules();
            if (SettingsManager.EnableAutoOrganize) Start();
        }

        public static void LoadRules()
        {
            try
            {
                if (File.Exists(_rulesFilePath))
                {
                    string json = File.ReadAllText(_rulesFilePath);
                    Rules = JsonConvert.DeserializeObject<List<OrganizeRule>>(json) ?? new List<OrganizeRule>();
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error loading organize rules: {ex.Message}");
            }
        }

        public static void SaveRules()
        {
            try
            {
                File.WriteAllText(_rulesFilePath, JsonConvert.SerializeObject(Rules, Formatting.Indented));
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error saving organize rules: {ex.Message}");
            }
        }

        public static void Start()
        {
            if (_watcher != null) return;

            try
            {
                _watcher = new FileSystemWatcher(_desktopPath)
                {
                    NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                    EnableRaisingEvents = true
                };

                // --- THE BROWSER SHIELD ---
                // Delay exactly 3 seconds before processing. This prevents browsers from panicking 
                // and deleting the file when we move it before they finish writing the Mark of the Web!
                _watcher.Created += (s, e) => Task.Run(async () => { await Task.Delay(3000); await ProcessFileAsync(e.FullPath); });
                _watcher.Renamed += (s, e) => Task.Run(async () => { await Task.Delay(3000); await ProcessFileAsync(e.FullPath); });

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Auto-Organize engine started.");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to start Auto-Organize: {ex.Message}");
            }
        }

        public static void Stop()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
                _watcher = null;
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Auto-Organize engine stopped.");
            }
        }
  
        // FIX: Added 'silent' parameter so the engine can sweep seamlessly in the background
        public static void ProcessDesktopNow(bool silent = false)
        {
            if (Rules == null || Rules.Count == 0) return;

            Task.Run(async () =>
            {
                var files = Directory.GetFiles(_desktopPath);
                foreach (var file in files)
                {
                    await ProcessFileAsync(file);
                }

                if (!silent)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm("Desktop organization complete!", "Success");
                    });
                }
            });
        }

        private static async Task ProcessFileAsync(string filePath)
        {
            if (!File.Exists(filePath) || !SettingsManager.EnableAutoOrganize) return;

            string fileName = Path.GetFileName(filePath);
            string ext = Path.GetExtension(filePath).ToLower();

            // --- THE SHORTCUT SHIELD ---
            // Ignore virtual items and active temp downloads
            if (ext == ".lnk" || ext == ".url" || ext == ".crdownload" || ext == ".part" || ext == ".tmp") return;

            // Sort rules by priority (lower number = higher priority)
            var activeRules = Rules.Where(r => r.IsEnabled).OrderBy(r => r.Priority).ToList();

            foreach (var rule in activeRules)
            {
                if (DoesFileMatchRule(fileName, rule))
                {
                    await ExecuteMoveAsync(filePath, rule);
                    break; // File processed, stop checking rules
                }
            }
        }

        private static bool DoesFileMatchRule(string fileName, OrganizeRule rule)
        {
            // 1. Check Extensions
            if (!string.IsNullOrWhiteSpace(rule.Extensions) && !rule.Extensions.Contains("*.*") && rule.Extensions.Trim() != "*")
            {
                string fileExt = Path.GetExtension(fileName).ToLower();
                var allowedExts = rule.Extensions.Split(new[] { ';', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                                 .Select(e => e.Trim().Replace("*", "").ToLower());

                // If the file extension is NOT in the allowed list, reject it
                if (!allowedExts.Contains(fileExt)) return false;
            }

            // 2. Check Name Contains (AND condition)
            if (!string.IsNullOrWhiteSpace(rule.NameContains))
            {
                // If the filename does NOT contain the required text, reject it
                if (fileName.IndexOf(rule.NameContains, StringComparison.OrdinalIgnoreCase) == -1)
                    return false;
            }

            // If it survived both filters, it is a perfect match!
            return true;
        }

        private static async Task ExecuteMoveAsync(string sourcePath, OrganizeRule rule)
        {
            if (!Directory.Exists(rule.TargetFolderPath))
            {
                try { Directory.CreateDirectory(rule.TargetFolderPath); }
                catch { return; }
            }

            // --- THE DOWNLOAD WAITER ---
            // Wait up to 60 seconds for the browser/app to release the file lock
            if (!await WaitForFileUnlockAsync(sourcePath, 60000))
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Auto-Organize skipped {Path.GetFileName(sourcePath)}: File was locked.");
                return;
            }

            string fileName = Path.GetFileName(sourcePath);
            string destPath = Path.Combine(rule.TargetFolderPath, fileName);

            try
            {
                // --- CONFLICT PROTOCOL ---
                if (File.Exists(destPath))
                {
                    if (rule.ConflictAction == RuleConflictAction.Skip) return;

                    if (rule.ConflictAction == RuleConflictAction.Overwrite)
                    {
                        // Safely send the existing file to the Recycle Bin before overwriting
                        FileSystem.DeleteFile(destPath, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin);
                    }
                    else if (rule.ConflictAction == RuleConflictAction.Rename)
                    {
                        string nameOnly = Path.GetFileNameWithoutExtension(fileName);
                        string ext = Path.GetExtension(fileName);
                        int count = 1;

                        while (File.Exists(destPath))
                        {
                            destPath = Path.Combine(rule.TargetFolderPath, $"{nameOnly} ({count}){ext}");
                            count++;
                        }
                    }
                }

                // Execute Physical Move
                File.Move(sourcePath, destPath);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Auto-Organize moved {fileName} to {rule.TargetFolderPath}");

                // --- SUCCESS TRACKING & NOTIFICATIONS ---
                rule.LastRun = DateTime.Now;
                SaveRules(); // Save the new timestamp

                if (SettingsManager.EnableAutoOrganizeNotifications)
                {
                    SmartToast.Show($"Rule: {rule.Name} Executed", $"Moved '{fileName}' to {new DirectoryInfo(rule.TargetFolderPath).Name}");
                }


                // --- AUTO-CREATE FENCE PROTOCOL ---
                if (rule.AutoCreateFence)
                {
                    Application.Current.Dispatcher.Invoke(() => CheckAndCreatePortalFence(rule.TargetFolderPath));
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Auto-Organize failed to move {fileName}: {ex.Message}");
            }
        }

        public static void CheckAndCreatePortalFence(string targetFolder)
        {
            // Prevent duplicate fences: Check if a portal fence already points to this folder
            var existingFences = FenceManager.GetFenceData();
            foreach (var f in existingFences)
            {
                try
                {
                    string fType = null, fPath = null;
                    if (f is Newtonsoft.Json.Linq.JObject jObj)
                    {
                        fType = jObj["ItemsType"]?.ToString();
                        fPath = jObj["Path"]?.ToString(); // FIX: Changed from PortalFolderPath to Path
                    }
                    else
                    {
                        fType = f.GetType().GetProperty("ItemsType")?.GetValue(f)?.ToString() ?? f.ItemsType?.ToString();
                        fPath = f.GetType().GetProperty("Path")?.GetValue(f)?.ToString() ?? f.Path?.ToString(); // FIX
                    }

                    if (fType == "Portal" && string.Equals(fPath, targetFolder, StringComparison.OrdinalIgnoreCase))
                    {
                        return; // Fence already exists!
                    }
                }
                catch { }
            }

            // Command FenceManager to generate a new Portal Fence using the exact required schema
            dynamic newFence = new Newtonsoft.Json.Linq.JObject();
            newFence.Id = Guid.NewGuid().ToString();
            newFence.Title = new DirectoryInfo(targetFolder).Name; // FIX: Changed from Name to Title
            newFence.X = 100.0; // FIX: Changed from Left to X
            newFence.Y = 100.0; // FIX: Changed from Top to Y
            newFence.Width = 350.0;
            newFence.Height = 250.0;
            newFence.ItemsType = "Portal";
            newFence.Path = targetFolder; // FIX: Changed from PortalFolderPath to Path
            newFence.Items = ""; // Portal expects empty string for Items

            // FIX: Initialize ALL fence properties with defaults to match JSON structure exactly
            newFence.IsLocked = "false";
            newFence.IsHidden = "false";
            newFence.CustomColor = null;
            newFence.CustomLaunchEffect = null;
            newFence.IsRolled = "false";
            newFence.UnrolledHeight = 250.0;
            newFence.TextColor = null;
            newFence.BoldTitleText = "false";
            newFence.TitleTextColor = null;
            newFence.DisableTextShadow = "false";
            newFence.IconSize = "Medium";
            newFence.GrayscaleIcons = "false";
            newFence.IconSpacing = 5;
            newFence.TitleTextSize = "Medium";
            newFence.FenceBorderColor = null;
            newFence.FenceBorderThickness = 2;

            // TABS FEATURE requirements
            newFence.TabsEnabled = "false";
            newFence.CurrentTab = 0;
            newFence.Tabs = new Newtonsoft.Json.Linq.JArray();

            // 1. Add to the central data manager
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                FenceDataManager.FenceData.Add(newFence);

                // 2. Pass a TargetChecker instance to the creation engine
                FenceManager.CreateFence(newFence, new TargetChecker(1000));

                // 3. Save using the correct data manager method
                FenceDataManager.SaveFenceData();
            });

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Auto-Organize spawned a new Portal Fence for {targetFolder}");
        }

        private static async Task<bool> WaitForFileUnlockAsync(string filePath, int timeoutMs)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    using (FileStream fs = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    {
                        return true; // Lock acquired successfully, file is ready
                    }
                }
                catch (IOException)
                {
                    await Task.Delay(500); // Wait half a second and try again
                }
            }
            return false; // Timed out
        }


        public static void Pause()
        {
            if (_watcher != null) _watcher.EnableRaisingEvents = false;
        }

        public static void Resume()
        {
            if (!SettingsManager.EnableAutoOrganize) return;

            if (_watcher == null)
            {
                Start();
            }
            else
            {
                _watcher.EnableRaisingEvents = true;
            }

            // Silently sweep the desktop to catch any files that landed while we were paused!
            ProcessDesktopNow(true);
        }
    }
}
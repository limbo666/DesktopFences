using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Newtonsoft.Json;
using Desktop_Frames.Localization;
using Microsoft.VisualBasic.FileIO; // Required for sending to Recycle Bin

namespace Desktop_Frames
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
        public bool AutoCreateFrame { get; set; } = false;
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
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.MsgDesktopOrganized, Strings.DlgSuccess);
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

        /// <summary>
        /// True when the file matches one of the patterns in a rule's extension list.
        ///
        /// The asterisk has to behave like a wildcard here. The presets rely on it:
        /// "Documents" is "*.doc*; *.pdf; *.txt; *.rtf; *.xls*; *.ppt*", where the
        /// trailing asterisk is what is meant to cover .docx and .xlsx alongside the
        /// older .doc and .xls.
        ///
        /// The previous version removed every asterisk and compared the result to the
        /// file extension for equality, so "*.doc*" became ".doc" and matched .doc but
        /// not .docx — the preset quietly ignored every modern Office file.
        ///
        /// That exact comparison is kept as a first check, so any list that worked
        /// before still works, including entries someone typed without a wildcard.
        /// </summary>
        private static bool MatchesAnyExtensionPattern(string fileName, string patterns)
        {
            string name = Path.GetFileName(fileName).ToLowerInvariant();
            string ext = Path.GetExtension(fileName).ToLowerInvariant();

            foreach (string raw in patterns.Split(new[] { ';', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string pattern = raw.Trim().ToLowerInvariant();
                if (pattern.Length == 0) continue;

                // Old behaviour, preserved: "*.pdf" stripped of asterisks is ".pdf".
                if (pattern.Replace("*", "") == ext) return true;

                // New: the asterisk now stands for anything, on the whole file name.
                if (pattern.IndexOf('*') >= 0)
                {
                    string expression = "^" + System.Text.RegularExpressions.Regex.Escape(pattern).Replace("\\*", ".*") + "$";
                    if (System.Text.RegularExpressions.Regex.IsMatch(name, expression)) return true;
                }
            }

            return false;
        }

        private static bool DoesFileMatchRule(string fileName, OrganizeRule rule)
        {
            // 1. Check Extensions
            if (!string.IsNullOrWhiteSpace(rule.Extensions) && !rule.Extensions.Contains("*.*") && rule.Extensions.Trim() != "*")
            {
                // If the file matches none of the patterns, reject it
                if (!MatchesAnyExtensionPattern(fileName, rule.Extensions)) return false;
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


                // --- AUTO-CREATE FRAME PROTOCOL ---
                if (rule.AutoCreateFrame)
                {
                    Application.Current.Dispatcher.Invoke(() => CheckAndCreatePortalFrame(rule.TargetFolderPath));
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Auto-Organize failed to move {fileName}: {ex.Message}");
            }
        }

        public static void CheckAndCreatePortalFrame(string targetFolder)
        {
            // Prevent duplicate frames: Check if a portal frame already points to this folder
            var existingFrames = Framemanager.GetFrameData();
            foreach (var f in existingFrames)
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
                        return; // Frame already exists!
                    }
                }
                catch { }
            }

            // Command Framemanager to generate a new Portal Frame using the exact required schema
            dynamic newFrame = new Newtonsoft.Json.Linq.JObject();
            newFrame.Id = Guid.NewGuid().ToString();
            newFrame.Title = new DirectoryInfo(targetFolder).Name; // FIX: Changed from Name to Title
            newFrame.X = 100.0; // FIX: Changed from Left to X
            newFrame.Y = 100.0; // FIX: Changed from Top to Y
            newFrame.Width = 350.0;
            newFrame.Height = 250.0;
            newFrame.ItemsType = "Portal";
            newFrame.Path = targetFolder; // FIX: Changed from PortalFolderPath to Path
            newFrame.Items = ""; // Portal expects empty string for Items

            // FIX: Initialize ALL frame properties with defaults to match JSON structure exactly
            newFrame.IsLocked = "false";
            newFrame.IsHidden = "false";
            newFrame.CustomColor = null;
            newFrame.CustomLaunchEffect = null;
            newFrame.IsRolled = "false";
            newFrame.UnrolledHeight = 250.0;
            newFrame.TextColor = null;
            newFrame.BoldTitleText = "false";
            newFrame.TitleTextColor = null;
            newFrame.DisableTextShadow = "false";
            newFrame.IconSize = "Medium";
            newFrame.GrayscaleIcons = "false";
            newFrame.IconSpacing = 5;
            newFrame.TitleTextSize = "Medium";
            newFrame.FrameBorderColor = null;
            newFrame.FrameBorderThickness = 2;

            // TABS FEATURE requirements
            newFrame.TabsEnabled = "false";
            newFrame.CurrentTab = 0;
            newFrame.Tabs = new Newtonsoft.Json.Linq.JArray();

            // 1. Add to the central data manager
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                FrameDataManager.FrameData.Add(newFrame);

                // 2. Pass a TargetChecker instance to the creation engine
                Framemanager.CreateFrame(newFrame, new TargetChecker(1000));

                // 3. Save using the correct data manager method
                FrameDataManager.SaveFrameData();
            });

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Auto-Organize spawned a new Portal Frame for {targetFolder}");
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
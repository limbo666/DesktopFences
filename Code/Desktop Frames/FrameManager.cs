using Desktop_Frames.Localization;
using IWshRuntimeLibrary;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics.Arm;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;
using Color = System.Drawing.Color;
using Desktop_Frames.Plugins; // --- NEW: Plugins Support ---

namespace Desktop_Frames
{
    public static class Framemanager
    {

        // --- NEW: Centralized Presets ---
        private static readonly Dictionary<string, string> _standardPresets = new Dictionary<string, string>
        {
            { "Images", "*.jpg; *.jpeg; *.png; *.gif; *.bmp; *.webp" },
            { "Documents", "*.doc*; *.pdf; *.txt; *.rtf; *.xls*; *.ppt*" },
            { "Executables", "*.exe; *.bat; *.msi; *.cmd; *.ps1" },
            { "Archives", "*.zip; *.rar; *.7z; *.tar; *.gz" },
            { "Media", "*.mp3; *.wav; *.mp4; *.mkv; *.avi" },
            { "Hide System", ">*.tmp; >desktop.ini; >~$*" }
        };
        // -------------------------------

        // Temporary override flag for the "Screen Bound" button
        public static bool IsManualRepositioning = false;



        // --- WM_GETMINMAXINFO Implementation ---
        private const int WM_GETMINMAXINFO = 0x0024;

        private const int WM_SYSCOMMAND = 0x0112; // <--- NEW
        private const int SC_MAXIMIZE = 0xF030;   // <--- NEW

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MINMAXINFO
        {
            public POINT ptReserved;
            public POINT ptMaxSize;
            public POINT ptMaxPosition;
            public POINT ptMinTrackSize;
            public POINT ptMaxTrackSize;
        }

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);


        // Tracks temporary navigation paths for Portal frames. Key: frameId, Value: CurrentPath
        private static Dictionary<string, string> _portalNavigationStates = new Dictionary<string, string>();

        // Track icon state to prevent GDI leaks from constant re-extraction
        // Key: FilePath
        // Value: (LastWriteTime of the .lnk/file, IsBroken state)
        private static Dictionary<string, (DateTime LastWrite, bool IsBroken)> _iconStates = new Dictionary<string, (DateTime, bool)>();

        private static dynamic _options;
        private static Dictionary<dynamic, PortalFramemanager> _portalFrames = new Dictionary<dynamic, PortalFramemanager>();

        // --- NEW: Active Plugins Registry ---
        public static Dictionary<string, IFramePlugin> _activePlugins = new Dictionary<string, IFramePlugin>();

        // Stores heart TextBlock references for each frame to enable efficient ContextMenu updates
        private static readonly Dictionary<dynamic, TextBlock> _heartTextBlocks = new Dictionary<dynamic, TextBlock>();

        //resize feedback
        private static Window _sizeFeedbackWindow;
        private static System.Windows.Threading.DispatcherTimer _hideTimer;

        // Add near other static fields
        private static TargetChecker _currentTargetChecker;

        #region Auto-Hide frames Engine
        private static DispatcherTimer _masterAutoHideTimer;
        public static bool _areFramesAutoHidden = false;
        private static int _hideSequenceId = 0;

        public static void InitializeAutoHideTimer()
        {
            if (_masterAutoHideTimer == null)
            {
                _masterAutoHideTimer = new DispatcherTimer();
                _masterAutoHideTimer.Tick += (s, e) => ExecuteAutoHideSequence();
            }
            ResetAutoHideTimer();
        }

        public static void ResetAutoHideTimer()
        {
            if (_masterAutoHideTimer == null) return;
            _masterAutoHideTimer.Stop();

            if (SettingsManager.AutoHideFrames && SettingsManager.AutoHideTime > 0)
            {
                _masterAutoHideTimer.Interval = TimeSpan.FromSeconds(SettingsManager.AutoHideTime);
                _masterAutoHideTimer.Start();
            }
        }

        public static void WakeUpFrames()
        {
            _hideSequenceId++;
            _areFramesAutoHidden = false;
            ResetAutoHideTimer();

            // --- NEW: Restore Native Desktop Icons if synced ---
            // BUG FIX: Only restore them if the user hasn't enforced the permanent "Always Hide" rule!
            if (SettingsManager.HideDesktopElementsOnAllFramesHide && !SettingsManager.HideDesktopElementsOnStart)
            {
                DesktopIconManager.SetDesktopIconsVisible(true);
            }

            var frames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
            foreach (var frame in frames)
            {
                if (frame.Visibility == Visibility.Visible && frame.Opacity >= 0.99) continue;

                // --- BUG FIX: Skip Individually Hidden frames (Safe Extraction) ---
                string frameId = frame.Tag?.ToString();
                if (!string.IsNullOrEmpty(frameId))
                {
                    var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                    if (currentFrame != null)
                    {
                        bool isIndividuallyHidden = false;
                        try
                        {
                            if (currentFrame is Newtonsoft.Json.Linq.JObject jFrame)
                                isIndividuallyHidden = jFrame["IsHidden"]?.ToString().ToLower() == "true";
                            else
                                isIndividuallyHidden = (currentFrame.IsHidden?.ToString() ?? "false").ToLower() == "true";
                        }
                        catch { }

                        if (isIndividuallyHidden) continue; // Leave it hidden in the tray
                    }
                }
                // ------------------------------------------------------------------

                // ------------------------------------------------------------------
                // --- DWM & WPF BUFFER FLASH FIX ---
                // 1. Wipe the slate clean of any old animations from the previous hide
                frame.BeginAnimation(UIElement.OpacityProperty, null);

                // 2. Hard-lock the base opacity to exactly 0.0
                frame.Opacity = 0.0;

                // 3. Make visible
                frame.Visibility = Visibility.Visible;

                // 4. THE MAGIC BULLET: Force WPF to synchronously flush the UI queue and render the invisible 0.0 frame right now.
                // This overwrites the Windows DWM cache *before* the new animation is permitted to begin.
                frame.UpdateLayout();

                // 5. Begin the smooth fade-in
                var fadeIn = new System.Windows.Media.Animation.DoubleAnimation
                {
                    From = 0.0,
                    To = 1.0,
                    Duration = TimeSpan.FromMilliseconds(250)
                };

                frame.BeginAnimation(UIElement.OpacityProperty, fadeIn);

                // --- BUG FIX: Sync localized idle timers with global wake-up ---
                frame.TriggerWakeUpIdleReset();
            }

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Frames faded in and restored.");
        }

        private static async void ExecuteAutoHideSequence()
        {
            _masterAutoHideTimer.Stop();
            if (_areFramesAutoHidden) return;

            int currentSequence = ++_hideSequenceId;

            bool isDialogOpen = System.Windows.Application.Current.Windows
                .OfType<System.Windows.Window>()
                .Any(w => w.Visibility == Visibility.Visible
                       && !(w is NonActivatingWindow)
                       && w.Width > 0
                       && !string.IsNullOrWhiteSpace(w.Title));

            if (isDialogOpen)
            {
                ResetAutoHideTimer();
                return;
            }

            var visibleFrames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>()
                .Where(w => w.Visibility == Visibility.Visible)
                .ToList();

            if (visibleFrames.Count == 0) return;
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Auto-Hide: Executing hide sequence.");

            // --- BUG FIX: Clear Held Animations ---
            // Because the Wake-Up fade-in natively holds at 1.0, we MUST release the 
            // animation lock before manually toggling local Opacity values. Otherwise, 
            // the animation engine fights the flash effect, causing visual residue.
            foreach (var win in visibleFrames)
            {
                win.BeginAnimation(System.Windows.UIElement.OpacityProperty, null);
                win.Opacity = 1.0;
            }

            if (SettingsManager.HideFlashEffect)
            {
                for (int i = 0; i < 5; i++)
                {
                    foreach (var win in visibleFrames) win.Opacity = 0.3;
                    await Task.Delay(30);
                    if (currentSequence != _hideSequenceId) return;

                    foreach (var win in visibleFrames) win.Opacity = 1.0;
                    await Task.Delay(30);
                    if (currentSequence != _hideSequenceId) return;
                }
            }

            var fadeOutAnim = new System.Windows.Media.Animation.DoubleAnimation
            {
                From = 1.0,
                To = 0.0,
                Duration = new System.Windows.Duration(TimeSpan.FromMilliseconds(200))
            };

            foreach (var win in visibleFrames)
            {
                win.BeginAnimation(System.Windows.UIElement.OpacityProperty, fadeOutAnim);
            }

            await Task.Delay(200);

            if (currentSequence != _hideSequenceId) return;

            foreach (var win in visibleFrames)
            {
                // --- WPF 1-FRAME FLASH FIX ---
                // Must hard-set to 0.0 BEFORE removing the animation precedence!
                win.Opacity = 0.0;
                win.BeginAnimation(System.Windows.UIElement.OpacityProperty, null);
                win.Visibility = Visibility.Hidden;
            }

            _areFramesAutoHidden = true;

            // --- NEW: Hide Native Desktop Icons if synced ---
            if (SettingsManager.HideDesktopElementsOnAllFramesHide)
            {
                DesktopIconManager.SetDesktopIconsVisible(false);
            }
        }
        #endregion

        // Track frame currently in rollup/rolldown transition to prevent event conflicts
        private static readonly HashSet<string> _framesInTransition = new HashSet<string>();
        // Emergency cleanup timer to prevent permanently stuck transition states
        private static System.Windows.Threading.DispatcherTimer _transitionCleanupTimer;

        // --- NEW: Auto Roll Tracking Engine ---
        private static readonly HashSet<string> _autoRolledFrames = new HashSet<string>();
        private static readonly Dictionary<string, DispatcherTimer> _autoRollTimers = new Dictionary<string, DispatcherTimer>();
        // --------------------------------------

        public static void RefreshScrollbarSettings()
        {
            var allFrames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
            foreach (var win in allFrames)
            {
                var border = win.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();

                if (scrollViewer != null)
                {
                    scrollViewer.VerticalScrollBarVisibility = SettingsManager.DisableFrameScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto;
                }
            }
        }

        public static void RefreshAllIconsTint()
        {
            // Calculate the target opacity exactly as we do in AddIcon
            double targetOpacity = SettingsManager.ApplyTintToIcons ? Math.Max(0.2, SettingsManager.TintValue / 100.0) : 1.0;

            var allFrames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
            foreach (var win in allFrames)
            {
                // --- FIX: Update Plugin Frame Opacity instantly ---
                string frameId = win.Tag?.ToString();
                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);

                if (currentFrame != null && currentFrame.ItemsType?.ToString() == "Plugin")
                {
                    // CRITICAL: Release the WPF animation lock first, or the opacity change is ignored
                    win.BeginAnimation(System.Windows.UIElement.OpacityProperty, null);
                    win.Opacity = targetOpacity;
                    continue; // Plugins don't use the WrapPanel, move to the next window
                }
                // --------------------------------------------------

                var wrapPanel = FindWrapPanel(win);
                if (wrapPanel != null)
                {
                    foreach (var sp in wrapPanel.Children.OfType<StackPanel>())
                    {
                        // 1. Update the Image (Checking direct child, then checking inside Grid for Network Icons)
                        var img = sp.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                        if (img == null)
                        {
                            var grid = sp.Children.OfType<Grid>().FirstOrDefault();
                            img = grid?.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                        }
                        if (img != null) img.Opacity = targetOpacity;

                        // 2. Update the TextLabel(s)
                        foreach (var txt in sp.Children.OfType<TextBlock>())
                        {
                            txt.Opacity = targetOpacity;
                        }
                    }
                }
            }
        }

        public static void RefreshAutoRollSettings()
        {
            int autoRollDelay = 2000;
            try { if (SettingsManager.AutoRollTime > 0) autoRollDelay = SettingsManager.AutoRollTime * 1000; } catch { }
            foreach (var timer in _autoRollTimers.Values)
            {
                if (timer != null)
                {
                    bool wasRunning = timer.IsEnabled;
                    timer.Stop();
                    timer.Interval = TimeSpan.FromMilliseconds(autoRollDelay);
                    if (wasRunning) timer.Start();
                }
            }
        }

        public static void ReloadFrames()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Reloading all frames...");

                // 1. Life Support
                Window lifeSupport = MessageBoxesManager.CreateWaitWindow("Desktop Frames +", "Refreshing configuration...");
                lifeSupport.Show();
                System.Windows.Forms.Application.DoEvents();

                // 2. Clean up Old State
                if (TrayManager.Instance != null)
                {
                    TrayManager.Instance.ClearHiddenFrames(); // Ensure Tray is clean
                }

                var openFrames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
                foreach (var win in openFrames) win.Close();

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                _heartTextBlocks.Clear();
                foreach (var portal in _portalFrames.Values) try { portal.Dispose(); } catch { }
                _portalFrames.Clear();
                FrameDataManager.FrameData?.Clear();
                _currentTargetChecker?.Stop();

                // Clear Auto Roll tracking to prevent memory leaks
                foreach (var timer in _autoRollTimers.Values) timer.Stop();
                _autoRollTimers.Clear();
                _autoRolledFrames.Clear();

                // 3. Re-Load
                _currentTargetChecker = new TargetChecker(1000);
                LoadAndCreateFrames(_currentTargetChecker);

                // Allow a brief moment for WPF to register the new windows
                System.Windows.Forms.Application.DoEvents();

                // ==================== DEBUG FIX START ====================
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Starting Hidden Frame Registration...");

                if (FrameDataManager.FrameData != null)
                {
                    // Use 'dynamic' to bypass JValue compile errors
                    foreach (dynamic data in FrameDataManager.FrameData)
                    {
                        // A. Safe Title Extraction (Trim spaces)
                        string jsonTitle = (string)data.Title;
                        if (jsonTitle != null) jsonTitle = jsonTitle.Trim();

                        // B. Safe 'IsHidden' Extraction
                        // This handles: true (bool), "true" (string), 1 (int)
                        string rawHidden = (string)data.IsHidden?.ToString();
                        bool isHidden = false;
                        if (!string.IsNullOrEmpty(rawHidden))
                        {
                            bool.TryParse(rawHidden, out isHidden);
                        }

                        // C. Debug Log
                        if (isHidden)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Frame '{jsonTitle}' is marked HIDDEN in JSON. Searching for window...");

                            // D. Find Window
                            var frameWindow = System.Windows.Application.Current.Windows
                                .OfType<NonActivatingWindow>()
                                .FirstOrDefault(w => w.Title != null && w.Title.Trim() == jsonTitle);

                            if (frameWindow != null)
                            {
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"-> MATCH FOUND. Hiding '{jsonTitle}'");
                                TrayManager.AddHiddenFrame(frameWindow);
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"-> WARNING: Could not find open window for '{jsonTitle}'. Window list count: {System.Windows.Application.Current.Windows.Count}");
                            }
                        }
                    }
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, "FrameDataManager.FrameData is NULL after load.");
                }
                // ==================== DEBUG FIX END ====================

                lifeSupport.Close();
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Frames reloaded successfully.");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error reloading frames: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgReloadingFramesFailed", ex.Message), Strings.DlgError);
            }
        }
  





        /// <summary>
        /// Helper to detect UNC Roots (e.g. \\Server or \\192.168.1.10).
        /// These technically aren't "Directories" in .NET, but act as Folders in Windows.
        /// </summary>
        private static bool IsUncRoot(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            // Must start with \\ and NOT contain a drive letter (like C:\)
            if (path.StartsWith(@"\\") && !path.Contains(@":\"))
            {
                string clean = path.Substring(2);
                // If no more slashes (Server) or just one trailing slash (Server\), it's a root.
                int slash = clean.IndexOf('\\');
                return (slash < 0 || slash == clean.Length - 1);
            }
            return false;
        }
        // PORTAL FEATURE: Logic to switch folders
        public static void NavigatePortalFrame(dynamic frame, string newPath)
        {
            try
            {
                if (string.IsNullOrEmpty(newPath) || !System.IO.Directory.Exists(newPath)) return;

                string frameId = frame.Id?.ToString();

                // --- BUG FIX: Pull Live frame for Accurate Base Path ---
                var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId) ?? frame;
                string basePath = liveFrame.Path?.ToString(); // The permanent "Home"

                // --- BUG FIX: Strict Path Normalization ---
                string normNew = newPath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);
                string normBase = basePath?.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) ?? "";

                // Update State
                if (string.Equals(normNew, normBase, StringComparison.OrdinalIgnoreCase))
                {
                    _portalNavigationStates.Remove(frameId); // We returned Home
                }
                else
                {
                    _portalNavigationStates[frameId] = normNew; // We are Deep
                }

                // Tell the Manager to switch
                var portalEntry = _portalFrames.FirstOrDefault(kvp => kvp.Key?.Id?.ToString() == frameId);
                if (portalEntry.Value != null)
                {
                    portalEntry.Value.NavigateTo(normNew);
                }

                // Update UI (Show/Hide Bar)
                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                var win = windows.FirstOrDefault(w => w.Tag?.ToString() == frameId);
                if (win != null)
                {
                    RefreshPortalNavBar(win, liveFrame);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Portal Navigation Error: {ex.Message}");
            }
        }






        // PORTAL FEATURE: The UI Strip [ <  Path...  ⚓ ]
        // v2.5.4 .186: Precision Alignment (Anchor under Filter)
        public static void RefreshPortalNavBar(NonActivatingWindow frameWindow, dynamic frame)
        {
            try
            {
                var border = frameWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                if (frame.ItemsType?.ToString() != "Portal") return;

                string frameId = frame.Id?.ToString();
                var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId) ?? frame;
                string basePath = liveFrame.Path?.ToString();

                string normBase = basePath?.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) ?? "";
                string currentPath = _portalNavigationStates.ContainsKey(frameId) ? _portalNavigationStates[frameId] : normBase;
                string normCurrent = currentPath?.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) ?? "";

                bool isNavigating = !string.Equals(normCurrent, normBase, StringComparison.OrdinalIgnoreCase);

                // 1. Locate existing bar
                var existingBar = dockPanel.Children.OfType<Grid>().FirstOrDefault(g => g.Tag?.ToString() == "PORTAL_NAV_BAR");

                if (!isNavigating)
                {
                    if (existingBar != null) dockPanel.Children.Remove(existingBar);
                    return;
                }

                // Generate Display Path
                string displayPath = new System.IO.DirectoryInfo(normCurrent).Name;
                try
                {
                    if (normCurrent.StartsWith(normBase, StringComparison.OrdinalIgnoreCase))
                    {
                        string rootName = new System.IO.DirectoryInfo(normBase).Name;
                        string relativePart = normCurrent.Substring(normBase.Length).TrimStart(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar);
                        displayPath = string.IsNullOrEmpty(relativePart) ? rootName : System.IO.Path.Combine(rootName, relativePart);
                    }
                }
                catch { }

                // 2. BULLETPROOF UPDATE: Inject Data Directly Instead of Destroying
                // By targeting the grid columns directly, we completely avoid WPF rendering caches and stale closures.
                if (existingBar != null && existingBar.Children.Count >= 3)
                {
                    if (existingBar.Children[0] is Button btnBack) btnBack.Tag = normCurrent;
                    if (existingBar.Children[1] is TextBlock lblPath)
                    {
                        lblPath.Text = displayPath;
                        lblPath.ToolTip = normCurrent;
                    }
                    if (existingBar.Children[2] is Button btnSetBase) btnSetBase.Tag = normCurrent;

                    return; // Done! UI is instantly updated.
                }

                // 3. CREATE NEW BAR (First time only)
                Grid navGrid = new Grid
                {
                    Tag = "PORTAL_NAV_BAR",
                    Height = 24,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(60, 0, 0, 0)),
                    Margin = new Thickness(5, 0, 35, 2)
                };

                navGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(25) });
                navGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                navGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(25) });

                Button newBtnBack = new Button
                {
                    Content = "‹",
                    Tag = normCurrent, // Stores state dynamically
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = System.Windows.Media.Brushes.Transparent,
                    Foreground = System.Windows.Media.Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    ToolTip = Strings.TooltipGoUp
                };
                newBtnBack.Click += (s, e) =>
                {
                    if (s is Button btn && btn.Tag != null)
                    {
                        string parentPath = System.IO.Directory.GetParent(btn.Tag.ToString())?.FullName;
                        if (!string.IsNullOrEmpty(parentPath))
                            NavigatePortalFrame(liveFrame, parentPath);
                    }
                };

                TextBlock newLblPath = new TextBlock
                {
                    Text = displayPath,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Left,
                    Foreground = System.Windows.Media.Brushes.White,
                    FontSize = 11,
                    FontWeight = FontWeights.SemiBold,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    ToolTip = normCurrent,
                    Margin = new Thickness(5, 0, 5, 0)
                };

                Button newBtnSetBase = new Button
                {
                    Content = "⚓",
                    Tag = normCurrent, // Stores state dynamically
                    FontSize = 10,
                    Background = System.Windows.Media.Brushes.Transparent,
                    Foreground = System.Windows.Media.Brushes.Orange,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    ToolTip = Strings.TooltipSetHome
                };
                newBtnSetBase.Click += (s, e) =>
                {
                    if (s is Button btn && btn.Tag != null)
                    {
                        string target = btn.Tag.ToString();
                        if (MessageBoxesManager.ShowCustomYesNoMessageBox(Strings.Get("MsgConfirmSetPortalHome", target), Strings.DlgUpdatePortalFrame))
                        {
                            UpdateFrameProperty(liveFrame, "Path", target, "Updated Portal Path");
                            _portalNavigationStates.Remove(frameId);
                            RefreshPortalNavBar(frameWindow, liveFrame);
                        }
                    }
                };

                navGrid.Children.Add(newBtnBack); Grid.SetColumn(newBtnBack, 0);
                navGrid.Children.Add(newLblPath); Grid.SetColumn(newLblPath, 1);
                navGrid.Children.Add(newBtnSetBase); Grid.SetColumn(newBtnSetBase, 2);

                DockPanel.SetDock(navGrid, Dock.Top);

                int insertIndex = 0;
                bool titleFound = false;
                for (int i = 0; i < dockPanel.Children.Count; i++)
                {
                    if (dockPanel.Children[i] is Grid g && g.Children.OfType<TextBlock>().Any(tb => tb.Name == "FrameLockIcon"))
                    {
                        insertIndex = i + 1;
                        titleFound = true;
                        break;
                    }
                }
                if (!titleFound) insertIndex = 1;
                if (insertIndex > dockPanel.Children.Count) insertIndex = dockPanel.Children.Count;

                dockPanel.Children.Insert(insertIndex, navGrid);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing Nav Bar: {ex.Message}");
            }
        }



















        public static void ShowPortalToast(NonActivatingWindow win, string message)
        {
            try
            {
                var cborder = win.Content as Border;
                if (cborder == null) return;

                Grid mainGrid = cborder.Child as Grid;
                if (mainGrid == null)
                {
                    // Convert DockPanel to Grid on the fly to support overlays
                    var dp = cborder.Child as DockPanel;
                    if (dp != null)
                    {
                        mainGrid = new Grid();
                        cborder.Child = mainGrid;
                        mainGrid.Children.Add(dp);
                    }
                    else return;
                }

                // Find or create toast
                Border toast = mainGrid.Children.OfType<Border>().FirstOrDefault(b => b.Name == "PortalToast");
                TextBlock toastText;

                if (toast == null)
                {
                    toast = new Border
                    {
                        Name = "PortalToast",
                        Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(220, 30, 30, 30)),
                        CornerRadius = new CornerRadius(15),
                        Padding = new Thickness(15, 6, 15, 6),
                        Margin = new Thickness(0, 0, 0, 15),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        IsHitTestVisible = false, // Prevents blocking mouse clicks
                        Opacity = 0
                    };

                    // Add subtle shadow for a premium feel
                    toast.Effect = new DropShadowEffect { Color = Colors.Black, BlurRadius = 10, ShadowDepth = 2, Opacity = 0.5 };

                    toastText = new TextBlock
                    {
                        Foreground = System.Windows.Media.Brushes.White,
                        FontSize = 12,
                        FontWeight = FontWeights.SemiBold,
                        TextAlignment = TextAlignment.Center
                    };
                    toast.Child = toastText;
                    mainGrid.Children.Add(toast);
                }
                else
                {
                    toastText = toast.Child as TextBlock;
                }

                if (toastText != null) toastText.Text = message;

                // Stop any running animation
                toast.BeginAnimation(UIElement.OpacityProperty, null);

                // Create the Fade In -> Hold -> Fade Out animation (2 seconds total)
                var fadeInOut = new DoubleAnimationUsingKeyFrames();
                fadeInOut.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.Zero)));
                fadeInOut.KeyFrames.Add(new EasingDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(200)), new QuadraticEase { EasingMode = EasingMode.EaseOut }));
                fadeInOut.KeyFrames.Add(new LinearDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(1500)))); // Hold for 1.5s
                fadeInOut.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(2000)), new QuadraticEase { EasingMode = EasingMode.EaseIn })); // Fade out

                toast.BeginAnimation(UIElement.OpacityProperty, fadeInOut);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing toast: {ex.Message}");
            }
        }

        /// <summary>
        /// Surgically reloads only specific frames by closing and recreating their windows
        /// Used by: ItemMoveDialog to prevent duplicate windows during move operations
        /// Category: Selective Window Management
        /// </summary>
        public static void ReloadSpecificFrames(dynamic sourceFrame, dynamic targetFrame)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Reloading specific frames: '{sourceFrame?.Title}' and '{targetFrame?.Title}'");

                // Get frame IDs for identification
                string sourceframeId = sourceFrame?.Id?.ToString();
                string targetFrameId = targetFrame?.Id?.ToString();

                // Check if we're about to close all frames (which would terminate the app)
                bool isSourceTargetSame = sourceframeId == targetFrameId;
                int totalFrameCount = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().Count();
                int framesToClose = isSourceTargetSame ? 1 : 2;
                bool needsLifeSupport = (framesToClose >= totalFrameCount);

                Window lifeSupportWindow = null;

                // Show "life support" wait window if we're about to close all frames
                if (needsLifeSupport)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Creating life support window - about to close {framesToClose} of {totalFrameCount} total frames");

                    lifeSupportWindow = MessageBoxesManager.CreateWaitWindow("Desktop Frames +", "Refreshing frames, please wait...");
                    lifeSupportWindow.Show();
                }

                // Find and close only the specific frame windows
                var allWindows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
                var windowsToClose = new List<NonActivatingWindow>();

                foreach (var window in allWindows)
                {
                    string windowframeId = window.Tag?.ToString();
                    if (!string.IsNullOrEmpty(windowframeId) &&
                        (windowframeId == sourceframeId || windowframeId == targetFrameId))
                    {
                        windowsToClose.Add(window);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"Marking frame window for reload: {window.Title} (ID: {windowframeId})");
                    }
                }

                // Close the specific windows
                foreach (var window in windowsToClose)
                {
                    try
                    {
                        window.Close();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"Closed frame window: {window.Title}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                            $"Error closing window {window.Title}: {ex.Message}");
                    }
                }

                // Clean up portal frame managers for the specific frames
                var portalFramesToRemove = _portalFrames.Keys
                    .Where(frame => frame?.Id?.ToString() == sourceframeId || frame?.Id?.ToString() == targetFrameId)
                    .ToList();

                foreach (var frame in portalFramesToRemove)
                {
                    _portalFrames.Remove(frame);
                }

                // Recreate only the specific frames with proper TargetChecker
                var FrameData = FrameDataManager.FrameData;
                var framesToRecreate = FrameData.Where(f =>
                    f.Id?.ToString() == sourceframeId || f.Id?.ToString() == targetFrameId).ToList();

                foreach (var frame in framesToRecreate)
                {
                    try
                    {
                        CreateFrame(frame, _currentTargetChecker ?? new TargetChecker(1000));
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate,
                            $"Recreated frame: {frame.Title}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                            $"Error recreating frame {frame.Title}: {ex.Message}");
                    }
                }

                // Close life support window if we used it
                if (lifeSupportWindow != null)
                {
                    try
                    {
                        lifeSupportWindow.Close();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            "Closed life support window");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                            $"Error closing life support window: {ex.Message}");
                    }
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    "Specific frame reload completed successfully");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in ReloadSpecificFrames: {ex.Message}");
                throw;
            }
        }


        private static double GetDpiScaleFactor(Window window)
        {
            // Get the screen where the window is located based on its position
            var screen = System.Windows.Forms.Screen.FromPoint(
                new System.Drawing.Point((int)window.Left, (int)window.Top));

            // Use Graphics to get the screen's DPI
            using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = graphics.DpiX; // Horizontal DPI
                return dpiX / 96.0; // Standard DPI is 96, so scale factor = dpiX / 96
            }
        }



        private static void AdjustFramePositionToScreen(NonActivatingWindow win)
        {
            // 1. CONTROL CHECK: 
            // If Auto-Reposition is OFF ... AND ... we are NOT manually forcing it -> EXIT.
            if (!SettingsManager.AllowAutoReposition && !IsManualRepositioning)
            {
                return;
            }

            // Get the DPI scale factor for the window
            double dpiScale = GetDpiScaleFactor(win);

            // Determine the screen based on the current position in pixels
            var screen = System.Windows.Forms.Screen.FromPoint(
                new System.Drawing.Point((int)(win.Left * dpiScale), (int)(win.Top * dpiScale)));

            // Get the screen's working area in pixels (excludes taskbars)
            var workingArea = screen.WorkingArea;

            // Convert window position and size from DIUs to pixels
            double winLeftPx = win.Left * dpiScale;
            double winTopPx = win.Top * dpiScale;
            double winWidthPx = win.Width * dpiScale;
            double winHeightPx = win.Height * dpiScale;

            // Calculate new position in pixels
            double newLeftPx = winLeftPx;
            double newTopPx = winTopPx;

            // Ensure the right edge doesn't exceed the working area's right boundary
            if (newLeftPx + winWidthPx > workingArea.Right)
            {
                newLeftPx = workingArea.Right - winWidthPx;
            }
            // Ensure the left edge isn't off-screen to the left
            if (newLeftPx < workingArea.Left)
            {
                newLeftPx = workingArea.Left;
            }

            // Ensure the bottom edge doesn't exceed the working area's bottom boundary
            if (newTopPx + winHeightPx > workingArea.Bottom)
            {
                newTopPx = workingArea.Bottom - winHeightPx;
            }
            // Ensure the top edge isn't off-screen to the top
            if (newTopPx < workingArea.Top)
            {
                newTopPx = workingArea.Top;
            }

            // Convert the adjusted position back to DIUs
            double newLeft = newLeftPx / dpiScale;
            double newTop = newTopPx / dpiScale;

            // Apply the new position if it has changed
            if (newLeft != win.Left || newTop != win.Top)
            {
                win.Left = newLeft;
                win.Top = newTop;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.ImportExport, $"Adjusted frame '{win.Title}' position to ({newLeft}, {newTop}) to fit within screen bounds.");
                FrameDataManager.SaveFrameData();
            }
        }




        public static void ForceRepositionallFrames()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Manual maintenance: Forcing screen bound check.");

                // 1. Alter variable to allow execution
                IsManualRepositioning = true;

                // 2. Run the logic (it will now pass the check)
                foreach (var frame in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>())
                {
                    AdjustFramePositionToScreen(frame);
                }

                FrameDataManager.SaveFrameData();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Error, $"Force reposition failed: {ex.Message}");
            }
            finally
            {
                // 3. Return variable to neutral
                IsManualRepositioning = false;
            }
        }





        private static int _registryMonitorTickCount = 0;


        public static Style GetThemedContextMenuStyle(dynamic frame)
        {
            try
            {
                // Safely extract the ID without dynamic binding crashes
                string targetId = null;
                try
                {
                    if (frame is Newtonsoft.Json.Linq.JObject jObj) targetId = jObj["Id"]?.ToString();
                    else targetId = frame.Id?.ToString();
                }
                catch { }

                // Fetch live data safely, avoiding FirstOrDefault cross-contamination
                dynamic liveFrame = frame;
                if (!string.IsNullOrEmpty(targetId))
                {
                    foreach (dynamic f in FrameDataManager.FrameData)
                    {
                        string fId = null;
                        try { fId = (f is Newtonsoft.Json.Linq.JObject jf) ? jf["Id"]?.ToString() : f.Id?.ToString(); } catch { }

                        if (fId == targetId)
                        {
                            liveFrame = f;
                            break;
                        }
                    }
                }

                // 1. Determine Background Color safely
                string frameColorName = null;
                try
                {
                    if (liveFrame is Newtonsoft.Json.Linq.JObject jf) frameColorName = jf["CustomColor"]?.ToString();
                    else frameColorName = liveFrame.CustomColor?.ToString();
                }
                catch { }

                if (string.IsNullOrEmpty(frameColorName)) frameColorName = SettingsManager.SelectedColor;

                System.Windows.Media.Color bgColor = System.Windows.Media.Colors.Gray;
                try
                {
                    var drawingColor = Utility.GetColorFromName(frameColorName);
                    bgColor = System.Windows.Media.Color.FromArgb(255, drawingColor.R, drawingColor.G, drawingColor.B);
                }
                catch { }

                // 2. Determine Foreground Color (Luminance check)
                double luminance = (0.299 * bgColor.R + 0.587 * bgColor.G + 0.114 * bgColor.B) / 255.0;
                System.Windows.Media.Color fgColor = luminance > 0.5 ? System.Windows.Media.Colors.Black : System.Windows.Media.Colors.White;

                string bgHex = $"#{bgColor.A:X2}{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
                string fgHex = $"#{fgColor.A:X2}{fgColor.R:X2}{fgColor.G:X2}{fgColor.B:X2}";

                // 3. Generate XAML Style using the universal banner
                string xamlStyle = $@"
                <Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' 
                       xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' 
                       TargetType='ContextMenu'>
                    <Style.Resources>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.MenuBrushKey}}' Color='{bgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.ControlBrushKey}}' Color='{bgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.WindowBrushKey}}' Color='{bgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.MenuTextBrushKey}}' Color='{fgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.ControlTextBrushKey}}' Color='{fgHex}'/>
         <Style TargetType='MenuItem'>
                            <Setter Property='Background' Value='{bgHex}'/>
                            <Setter Property='Foreground' Value='{fgHex}'/>
                            <Setter Property='BorderThickness' Value='0'/>
                            <Setter Property='BorderBrush' Value='Transparent'/>
                        </Style>
                    </Style.Resources>
                    <Setter Property='Background' Value='{bgHex}'/>
                    <Setter Property='Foreground' Value='{fgHex}'/>
                    <Setter Property='Template'>
                        <Setter.Value>
                            <ControlTemplate TargetType='ContextMenu'>
                                <Border Background='{{TemplateBinding Background}}' BorderBrush='#555555' BorderThickness='1' Padding='1'>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width='30'/>
                                            <ColumnDefinition Width='*'/>
                                        </Grid.ColumnDefinitions>
                                        
                                    <Border Grid.Column='0' Background='#151515'>
                                            <Image Source='pack://application:,,,/Resources/DesktopFramesVertical.png' Stretch='Uniform' VerticalAlignment='Top' Margin='0,5,0,0'/>
                                        </Border>
                                        
                                        <ScrollViewer Grid.Column='1' Margin='2,0,0,0' VerticalScrollBarVisibility='Hidden'>
                                            <ItemsPresenter KeyboardNavigation.DirectionalNavigation='Cycle'/>
                                        </ScrollViewer>
                                    </Grid>
                                </Border>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>";

                return (Style)System.Windows.Markup.XamlReader.Parse(xamlStyle);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Failed to parse themed ContextMenu style: {ex.Message}");
                return null;
            }
        }








        // Builds the heart ContextMenu for a frame with consistent items and dynamic state
        // v2.5.4 .183: Swapped Tabs/Delete position for better UX safety
        private static ContextMenu BuildHeartContextMenu(dynamic frame, bool showTabsOption = false)
        {
            // DEBUG: Log menu building
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                $"Building heart context menu for frame '{frame.Title}'");

            var menu = new ContextMenu();

            Style themedStyle = GetThemedContextMenuStyle(frame);
            if (themedStyle != null) menu.Style = themedStyle;

            // --- AUTO-CLOSE TIMER ---
            System.Windows.Threading.DispatcherTimer menuTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(4)
            };

            menuTimer.Tick += (s, e) =>
            {
                if (menu.IsOpen && !menu.IsMouseOver)
                {
                    menu.IsOpen = false;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Heart Menu auto-closed by timer");
                }
                menuTimer.Stop();
            };

            menu.Opened += (s, e) => { menuTimer.Start(); };
            menu.Closed += (s, e) => menuTimer.Stop();
            menu.MouseEnter += (s, e) => menuTimer.Stop();
            menu.MouseLeave += (s, e) => { menuTimer.Stop(); menuTimer.Start(); };
            // ------------------------

            // About item
            var aboutItem = new MenuItem { Header = Strings.MenuAbout };
            aboutItem.Click += (s, e) => AboutFormManager.ShowAboutForm();
            menu.Items.Add(aboutItem);

            // Options item
            var optionsItem = new MenuItem { Header = Strings.MenuOptions };
            optionsItem.Click += (s, e) => OptionsFormManager.ShowOptionsForm();
            menu.Items.Add(optionsItem);

            // Separator
            menu.Items.Add(new Separator());

            // New frame items
            var newFrameItem = new MenuItem { Header = Strings.MenuNewFrame };
            newFrameItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                CreateNewFrame("", "Data", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newFrameItem);

            var newPortalFrameItem = new MenuItem { Header = Strings.MenuNewPortalFrame };
            newPortalFrameItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                CreateNewFrame("New Portal Frame", "Portal", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newPortalFrameItem);

            MenuItem newNoteFrameItem = new MenuItem { Header = Strings.MenuNewNoteFrame };
            newNoteFrameItem.Click += (s, e) =>
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                CreateNewFrame("", "Note", mousePosition.X, mousePosition.Y);
            };
            menu.Items.Add(newNoteFrameItem);

            // --- NEW: Tiered Dynamic Plugin Menu ---
            if (SettingsManager.PluginAvailabilityLevel > 0)
            {
                var addPluginMenu = new MenuItem { Header = Strings.MenuAddPlugin };
                var availablePlugins = PluginManager.GetAvailablePlugins();

                if (availablePlugins != null && availablePlugins.Count > 0)
                {
                    foreach (var plugin in availablePlugins)
                    {
                        string pId = plugin.Key;
                        var pluginItem = new MenuItem { Header = plugin.Value };
                        pluginItem.Click += (s, e) =>
                        {
                            var mousePosition = System.Windows.Forms.Cursor.Position;
                            CreateNewFrame($"New {plugin.Value}", "Plugin", mousePosition.X, mousePosition.Y, null, null, pId);
                        };
                        addPluginMenu.Items.Add(pluginItem);
                    }
                }
                else
                {
                    addPluginMenu.Items.Add(new MenuItem { Header = Strings.MenuNoPluginsAvailable, IsEnabled = false });
                }
                menu.Items.Add(addPluginMenu);
            }
            // If PluginAvailabilityLevel == 0, the menu option is completely hidden.
            // --------------------------------

            menu.Items.Add(new Separator());

            // --- REORDERED: Tabs Option First ---
            // TABS FEATURE (Native Checkbox Logic)
            bool isDataFrame = frame.ItemsType?.ToString() == "Data";
            if (isDataFrame)
            {
                bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";

                var enableTabsItem = new MenuItem
                {
                    Header = Strings.MenuEnableTabs,
                    IsCheckable = true,   // Shows checkbox gutter
                    IsChecked = tabsEnabled // Visual checkmark
                };

                enableTabsItem.Click += (s, e) => ToggleFrameTabs(frame);
                menu.Items.Add(enableTabsItem);

                // Separator AFTER tabs to separate it from the Delete option
                menu.Items.Add(new Separator());
            }

            // --- REORDERED: Delete Option Second ---
            // Delete this frames
            var deleteThisFrame = new MenuItem { Header = Strings.MenuDeleteThisFrame };
            deleteThisFrame.Click += (s, e) =>
            {
                bool result = MessageBoxesManager.ShowCustomMessageBoxForm();
                if (result == true)
                {
                    if (SettingsManager.ExportShortcutsOnFrameDeletion && frame.ItemsType?.ToString() == "Data")
                    {
                        ExportAllIconsToDesktop(frame, false);
                    }

                    BackupManager.BackupDeletedFrame(frame);

                    FrameDataManager.FrameData.Remove(frame);
                    _heartTextBlocks.Remove(frame);

                    // --- BUG FIX: Avoid JObject HashCode Mutation ---
                    var targetPortal = _portalFrames.FirstOrDefault(kvp => kvp.Key?.Id?.ToString() == frame.Id?.ToString());
                    if (targetPortal.Value != null)
                    {
                        targetPortal.Value.Dispose();
                        _portalFrames.Remove(targetPortal.Key);
                    }

                    FrameDataManager.SaveFrameData();

                    var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                    var win = windows.FirstOrDefault(w => w.Tag?.ToString() == frame.Id?.ToString());
                    if (win != null) win.Close();

                    UpdateAllHeartContextMenus();
                }
            };
            menu.Items.Add(deleteThisFrame);

            menu.Items.Add(new Separator());

            // Export/Import Group
            var exportItem = new MenuItem { Header = Strings.MenuExportThisFrame };
            exportItem.Click += (s, e) => BackupManager.ExportFrame(frame);
            menu.Items.Add(exportItem);

            var importItem = new MenuItem { Header = Strings.MenuImportFrame };
            importItem.Click += (s, e) => BackupManager.ImportFrame();
            menu.Items.Add(importItem);

            // Restore frame item
            var restoreItem = new MenuItem
            {
                Header = Strings.MenuRestoreLastDeleted,
                Visibility = BackupManager.IsRestoreAvailable ? Visibility.Visible : Visibility.Collapsed
            };
            restoreItem.Click += (s, e) => BackupManager.RestoreLastDeletedFrame();
            menu.Items.Add(restoreItem);

            // Separator
            menu.Items.Add(new Separator());

            // Exit item
            var exitItem = new MenuItem { Header = Strings.MenuExit };
            exitItem.Click += (s, e) => System.Windows.Application.Current.Shutdown();
            menu.Items.Add(exitItem);

            return menu;
        }



      
        /// <summary>
        /// Centralized method to attach the standard Context Menu to an icon.
        /// Layout: Edit/Move/Remove -> Copy -> Admin -> Path
        /// Includes LIVE data lookup to ensure Checkable items stay synced.
        /// </summary>
        public static void AttachIconContextMenu(StackPanel sp, dynamic item, dynamic frame, NonActivatingWindow window)
        {
            try
            {
                // Initial snapshot for static properties (Path, Folder status)
                IDictionary<string, object> iconDict = item is IDictionary<string, object> dict ?
                    dict : ((JObject)item).ToObject<IDictionary<string, object>>();

                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];

                // HELPER: Function to find the LIVE item in memory (Handles Tabs vs Main)
                // We need this because 'item' becomes stale after an update.
                Func<dynamic> GetLiveItem = () =>
                {
                    try
                    {
                        string frameId = frame.Id?.ToString();
                        var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                        if (liveFrame == null) return null;

                        JArray itemsArray = liveFrame.Items as JArray ?? new JArray();

                        // Handle Tabs
                        bool tabsEnabled = liveFrame.TabsEnabled?.ToString().ToLower() == "true";
                        if (tabsEnabled)
                        {
                            var tabs = liveFrame.Tabs as JArray ?? new JArray();
                            int currentTabIndex = Convert.ToInt32(liveFrame.CurrentTab?.ToString() ?? "0");
                            if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                            {
                                var activeTab = tabs[currentTabIndex] as JObject;
                                itemsArray = activeTab?["Items"] as JArray ?? itemsArray;
                            }
                        }

                        // Find specific item by filename
                        return itemsArray.FirstOrDefault(i => string.Equals(
                            System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                            System.IO.Path.GetFullPath(filePath),
                            StringComparison.OrdinalIgnoreCase));
                    }
                    catch { return null; }
                };

                ContextMenu iconContextMenu = new ContextMenu();

                Style themedStyle = GetThemedContextMenuStyle(frame);
                if (themedStyle != null) iconContextMenu.Style = themedStyle;

                // --- GROUP 1: MANIPULATION ---
                MenuItem miEdit = new MenuItem { Header = Strings.MenuEdit };
                MenuItem miMove = new MenuItem { Header = Strings.MenuMove };
                MenuItem miRemove = new MenuItem { Header = Strings.MenuRemove };

                iconContextMenu.Items.Add(miEdit);
                iconContextMenu.Items.Add(miMove);
                iconContextMenu.Items.Add(miRemove);

                bool isSpacer = filePath != null && filePath.StartsWith("INTERNAL_BLANK_");

                // --- GROUP 2: CLIPBOARD ---
                Separator sepClipboard = new Separator();
                MenuItem miCopyItem = new MenuItem { Header = Strings.MenuCopyItem };

                iconContextMenu.Items.Add(sepClipboard);
                iconContextMenu.Items.Add(miCopyItem);

                if (isSpacer)
                {
                    sepClipboard.Visibility = Visibility.Collapsed;
                    miCopyItem.Visibility = Visibility.Collapsed;
                }

                // --- GROUP 3: EXECUTION (Conditional) ---
                bool isEligibleForAdmin = !isFolder && !string.IsNullOrEmpty(filePath) &&
                    (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk" || System.IO.Path.GetExtension(filePath).ToLower() == ".exe");

                MenuItem miAlwaysAdmin = null;
                MenuItem miRunAsDifferentUser = null;
                MenuItem miAlwaysRunAsDifferentUser = null;

                if (isEligibleForAdmin)
                {
                    iconContextMenu.Items.Add(new Separator());

                    MenuItem miRunAsAdmin = new MenuItem { Header = Strings.MenuRunAsAdmin };
                    miAlwaysAdmin = new MenuItem
                    {
                        Header = Strings.MenuAlwaysRunAsAdmin,
                        IsCheckable = true
                    };

                    iconContextMenu.Items.Add(miRunAsAdmin);
                    iconContextMenu.Items.Add(miAlwaysAdmin);

                 
                    // Run as Admin Logic
                    miRunAsAdmin.Click += (s, e) => {
                        string target = Utility.GetShortcutTarget(filePath);
                        string args = Utility.GetShortcutArguments(filePath); // Extract arguments

                        ProcessStartInfo psi = new ProcessStartInfo { FileName = target, UseShellExecute = true, Verb = "runas" };

                        if (!string.IsNullOrEmpty(args))
                        {
                            psi.Arguments = args; // Pass arguments to admin launch
                        }

                        if (System.IO.File.Exists(target))
                        {
                            string wd = System.IO.Path.GetDirectoryName(target);
                            if (!string.IsNullOrEmpty(wd)) psi.WorkingDirectory = wd;
                        }
                        try { Process.Start(psi); }
                        catch (Exception ex)
                        {
                            MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgRunAsAdminFailed", ex.Message), Strings.DlgError);
                        }
                    };

                    // START- HERE

                    // Toggle Always Admin Logic
                    miAlwaysAdmin.Click += (sender, e) => {
                        var liveItem = GetLiveItem();
                        if (liveItem != null)
                        {
                            bool newVal = !Convert.ToBoolean(liveItem["AlwaysRunAsAdmin"] ?? false);
                            liveItem["AlwaysRunAsAdmin"] = newVal;
                            miAlwaysAdmin.IsChecked = newVal;

                            // Mutually exclusive: Uncheck Different User if Admin is checked
                            if (newVal)
                            {
                                liveItem["AlwaysRunAsDifferentUser"] = false;
                                if (miAlwaysRunAsDifferentUser != null) miAlwaysRunAsDifferentUser.IsChecked = false;
                            }

                            FrameDataManager.SaveFrameData();
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Toggled AlwaysRunAsAdmin to {newVal} for {filePath}");
                        }
                    };
                }

                // --- GROUP 4: PATH / TARGET ---
                Separator sepPath = new Separator();
                iconContextMenu.Items.Add(sepPath);

                MenuItem miCopyPathRoot = new MenuItem { Header = Strings.MenuCopyPath };
                MenuItem miCopyFolder = new MenuItem { Header = Strings.MenuFolderPath };
                MenuItem miCopyFullPath = new MenuItem { Header = Strings.MenuFullPath };

                miCopyPathRoot.Items.Add(miCopyFolder);
                miCopyPathRoot.Items.Add(miCopyFullPath);

                MenuItem miFindTarget = new MenuItem { Header = Strings.MenuOpenTargetFolder };

                iconContextMenu.Items.Add(miCopyPathRoot);
                iconContextMenu.Items.Add(miFindTarget);

                if (isSpacer)
                {
                    sepPath.Visibility = Visibility.Collapsed;
                    miCopyPathRoot.Visibility = Visibility.Collapsed;
                    miFindTarget.Visibility = Visibility.Collapsed;
                }

                // --- EVENT HANDLERS ---
                miEdit.Click += (s, e) => EditItem(item, frame, window);
                miMove.Click += (s, e) => ItemMoveDialog.ShowMoveDialog(item, frame, window.Dispatcher);

                miRemove.Click += (s, e) =>
                {
                    try
                    {
                        // Use the LIVE lookup to find the item to remove
                        var liveItem = GetLiveItem(); // Returns the JObject from the array

                        // We need the Array itself to remove the item
                        string frameId = frame.Id?.ToString();
                        var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);

                        if (liveFrame != null && liveItem != null)
                        {
                            // Logic to locate the array containing liveItem
                            JArray targetArray = liveFrame.Items as JArray; // Default to main

                            bool tabsEnabled = liveFrame.TabsEnabled?.ToString().ToLower() == "true";
                            if (tabsEnabled)
                            {
                                var tabs = liveFrame.Tabs as JArray;
                                int tabIdx = Convert.ToInt32(liveFrame.CurrentTab?.ToString() ?? "0");
                                if (tabs != null && tabIdx < tabs.Count)
                                {
                                    targetArray = tabs[tabIdx]["Items"] as JArray;
                                }
                            }

                            if (targetArray != null)
                            {
                                targetArray.Remove(liveItem);
                                FrameDataManager.SaveFrameData();
                                var wp = VisualTreeHelper.GetParent(sp) as WrapPanel;
                                if (wp != null) wp.Children.Remove(sp);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error removing: {ex.Message}");
                    }
                };

                miCopyItem.Click += (s, e) => CopyPasteManager.CopyItem(item, frame);

                miFindTarget.Click += (s, e) => {
                    string target = Utility.GetShortcutTarget(filePath);
                    if (!string.IsNullOrEmpty(target) && (System.IO.File.Exists(target) || System.IO.Directory.Exists(target)))
                        Process.Start("explorer.exe", $"/select,\"{target}\"");
                };

                miCopyFolder.Click += (s, e) => {
                    string target = Utility.GetShortcutTarget(filePath);
                    if (!string.IsNullOrEmpty(target)) Clipboard.SetText(System.IO.Path.GetDirectoryName(target));
                };
                miCopyFullPath.Click += (s, e) => {
                    string target = Utility.GetShortcutTarget(filePath);
                    if (!string.IsNullOrEmpty(target)) Clipboard.SetText(target);
                };

                // --- DYNAMIC UPDATES (On Open) ---
                MenuItem miSendToDesktop = null;


                iconContextMenu.Opened += (s, e) => {
                    // --- LIVE THEME FIX: Refresh style right before opening ---
                    try
                    {
                        Style liveStyle = GetThemedContextMenuStyle(frame);
                        if (liveStyle != null) iconContextMenu.Style = liveStyle;
                    }
                    catch { }

                    // Fetch live item ONCE per open to prevent variable shadowing errors
                    var currentLiveItem = GetLiveItem();
                    bool isAlwaysDiffUser = currentLiveItem != null && Convert.ToBoolean(currentLiveItem["AlwaysRunAsDifferentUser"] ?? false);

                    bool isCtrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;

                    // Clear dynamic items
                    if (miSendToDesktop != null && iconContextMenu.Items.Contains(miSendToDesktop))
                        iconContextMenu.Items.Remove(miSendToDesktop);
                    if (miRunAsDifferentUser != null && iconContextMenu.Items.Contains(miRunAsDifferentUser))
                        iconContextMenu.Items.Remove(miRunAsDifferentUser);
                    if (miAlwaysRunAsDifferentUser != null && iconContextMenu.Items.Contains(miAlwaysRunAsDifferentUser))
                        iconContextMenu.Items.Remove(miAlwaysRunAsDifferentUser);

                    // 1. Send to Desktop (Ctrl only, Skip for spacers)
                    if (isCtrl && !isSpacer)
                    {
                        miSendToDesktop = new MenuItem { Header = Strings.MenuSendToDesktop };
                        miSendToDesktop.Click += (sender, args) => CopyPasteManager.SendToDesktop(item);
                        int idxCopy = iconContextMenu.Items.IndexOf(miCopyItem);
                        if (idxCopy != -1) iconContextMenu.Items.Insert(idxCopy + 1, miSendToDesktop);
                    }

                    // 2. Different User Options (Ctrl OR if setting is already checked)
                    if (isEligibleForAdmin && (isCtrl || isAlwaysDiffUser))
                    {
                        miRunAsDifferentUser = new MenuItem { Header = Strings.MenuRunAsDifferentUser };
                        miAlwaysRunAsDifferentUser = new MenuItem { Header = Strings.MenuAlwaysRunAsDifferentUser, IsCheckable = true };

                        miRunAsDifferentUser.Click += (sender, args) => {
                            string target = Utility.GetShortcutTarget(filePath);
                            string argsStr = Utility.GetShortcutArguments(filePath);
                            ProcessStartInfo psi = new ProcessStartInfo { FileName = target, UseShellExecute = true, Verb = "runasuser" };
                            if (!string.IsNullOrEmpty(argsStr)) psi.Arguments = argsStr;
                            if (System.IO.File.Exists(target))
                            {
                                string wd = System.IO.Path.GetDirectoryName(target);
                                if (!string.IsNullOrEmpty(wd)) psi.WorkingDirectory = wd;
                            }
                            try { Process.Start(psi); }
                            catch (Exception ex) { if (!ex.Message.Contains("canceled")) LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error running as different user: {ex.Message}"); }
                        };

                        miAlwaysRunAsDifferentUser.Click += (sender, args) => {
                            var liveItem = GetLiveItem();
                            if (liveItem != null)
                            {
                                bool newVal = !Convert.ToBoolean(liveItem["AlwaysRunAsDifferentUser"] ?? false);
                                liveItem["AlwaysRunAsDifferentUser"] = newVal;
                                miAlwaysRunAsDifferentUser.IsChecked = newVal;

                                // Mutually exclusive: Uncheck Admin if Different User is checked
                                if (newVal && miAlwaysAdmin != null)
                                {
                                    liveItem["AlwaysRunAsAdmin"] = false;
                                    miAlwaysAdmin.IsChecked = false;
                                }
                                FrameDataManager.SaveFrameData();
                            }
                        };

                        int idxAdmin = iconContextMenu.Items.IndexOf(miAlwaysAdmin);
                        if (idxAdmin != -1)
                        {
                            iconContextMenu.Items.Insert(idxAdmin + 1, miRunAsDifferentUser);
                            iconContextMenu.Items.Insert(idxAdmin + 2, miAlwaysRunAsDifferentUser);
                        }
                    }

                    // 3. REFRESH CHECK STATES
                    if (currentLiveItem != null)
                    {
                        if (miAlwaysAdmin != null)
                            miAlwaysAdmin.IsChecked = Convert.ToBoolean(currentLiveItem["AlwaysRunAsAdmin"] ?? false);

                        if (miAlwaysRunAsDifferentUser != null && iconContextMenu.Items.Contains(miAlwaysRunAsDifferentUser))
                            miAlwaysRunAsDifferentUser.IsChecked = Convert.ToBoolean(currentLiveItem["AlwaysRunAsDifferentUser"] ?? false);
                    }
                };


                sp.ContextMenu = iconContextMenu;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error attaching context menu: {ex.Message}");
            }
        }





        // Updates all heart ContextMenus across all frames using stored TextBlock references
        public static void UpdateAllHeartContextMenus()
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                foreach (var entry in _heartTextBlocks)
                {
                    var frame = entry.Key;
                    var heart = entry.Value;
                    if (heart != null)
                    {
                        // heart.ContextMenu = BuildHeartContextMenu(frame);
                        heart.ContextMenu = BuildHeartContextMenu(frame, false); // Normal menu by default
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Updated heart ContextMenu for frame '{frame.Title}'");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Skipped update for frame '{frame.Title}': heart TextBlock is null");
                    }
                }
            });
        }

        private static void CreateWebLinkShortcut(string targetUrl, string shortcutPath, string displayName)
        {
            try
            {
                // For web links, create a .url file instead of .lnk file
                // Change the extension to .url if it's not already
                string urlFilePath = shortcutPath;
                if (System.IO.Path.GetExtension(shortcutPath).ToLower() == ".lnk")
                {
                    urlFilePath = System.IO.Path.ChangeExtension(shortcutPath, ".url");
                }

                // Create a .url file content
                string urlFileContent = $"[InternetShortcut]\r\nURL={targetUrl}\r\nIconIndex=0\r\n";

                // Write the .url file
                System.IO.File.WriteAllText(urlFilePath, urlFileContent, System.Text.Encoding.ASCII);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Created web link URL file: {urlFilePath} -> {targetUrl}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling, $"Error creating web link shortcut {shortcutPath}: {ex.Message}");
                throw;
            }
        }


        /// <summary>
        /// Exports all icons from a Data frames to the desktop
        /// </summary>
        /// <param name="frame">The frame to export icons from</param>
        /// <param name="showConfirmation">If true, shows a message box on completion. False for silent automation.</param>
        private static void ExportAllIconsToDesktop(dynamic frame, bool showConfirmation = true)
        {
            try
            {
                // Verify this is a Data frame
                if (frame.ItemsType?.ToString() != "Data")
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                        $"Export all icons attempted on non-Data frame: {frame.ItemsType}");
                    return;
                }

                int exportedCount = 0;
                int totalCount = 0;

                // Check if frame has tabs
                bool hasTabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";

                if (hasTabsEnabled && frame.Tabs != null)
                {
                    // Handle tabbed frame - export from all tabs
                    var tabs = frame.Tabs as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();

                    foreach (var tab in tabs)
                    {
                        var tabObj = tab as Newtonsoft.Json.Linq.JObject;
                        var tabItems = tabObj?["Items"] as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();

                        foreach (var item in tabItems)
                        {
                            totalCount++;
                            try
                            {
                                CopyPasteManager.SendToDesktop(item);
                                exportedCount++;
                            }
                            catch (Exception itemEx)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                    $"Error exporting item to desktop: {itemEx.Message}");
                            }
                        }
                    }
                }
                else
                {
                    // Handle regular frame (no tabs)
                    var items = frame.Items as Newtonsoft.Json.Linq.JArray ?? new Newtonsoft.Json.Linq.JArray();

                    foreach (var item in items)
                    {
                        totalCount++;
                        try
                        {
                            CopyPasteManager.SendToDesktop(item);
                            exportedCount++;
                        }
                        catch (Exception itemEx)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                                $"Error exporting item to desktop: {itemEx.Message}");
                        }
                    }
                }

                // Show result message after a brief delay to allow desktop refresh
                string resultMessage = $"Exported {exportedCount} of {totalCount} icons to desktop.";
                if (exportedCount != totalCount)
                {
                    resultMessage += $" {totalCount - exportedCount} items failed to export.";
                }

                // Use DispatcherTimer for UI thread delay
                var delayTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(1700) // 1.7 second delay
                };
                delayTimer.Tick += (s, e) =>
                {
                    delayTimer.Stop();

                    // FIX: Check flag before showing message
                    if (showConfirmation)
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm(resultMessage, "Export Complete");
                    }
                };
                delayTimer.Start();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Export all icons completed for frame '{frame.Title}': {exportedCount}/{totalCount} successful");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in ExportAllIconsToDesktop: {ex.Message}");

                // Always show errors, even in silent mode, so user knows why it failed
                MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgExportIconsFailed", ex.Message), Strings.DlgExportError);
            }
        }

        /// <summary>
        /// Extracts URL from different browser drop data formats
        /// </summary>
        private static string ExtractUrlFromDropData(IDataObject dataObject)
        {
            try
            {
                // Try UniformResourceLocator format first (most browsers)
                if (dataObject.GetDataPresent("UniformResourceLocator"))
                {
                    object urlData = dataObject.GetData("UniformResourceLocator");
                    if (urlData is System.IO.MemoryStream stream)
                    {
                        byte[] bytes = new byte[stream.Length];
                        stream.Read(bytes, 0, (int)stream.Length);
                        string url = System.Text.Encoding.ASCII.GetString(bytes).Trim('\0');
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                            $"Extracted URL from UniformResourceLocator: {url}");
                        return url;
                    }
                }

                // Try Firefox format
                if (dataObject.GetDataPresent("text/x-moz-url"))
                {
                    string mozUrl = dataObject.GetData("text/x-moz-url") as string;
                    if (!string.IsNullOrEmpty(mozUrl))
                    {
                        string[] parts = mozUrl.Split('\n');
                        if (parts.Length > 0 && !string.IsNullOrEmpty(parts[0]))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                                $"Extracted URL from text/x-moz-url: {parts[0]}");
                            return parts[0].Trim();
                        }
                    }
                }

                // Try HTML format  
                if (dataObject.GetDataPresent(DataFormats.Html))
                {
                    string html = dataObject.GetData(DataFormats.Html) as string;
                    if (!string.IsNullOrEmpty(html))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(html, @"href=['""]([^'""]+)['""]");
                        if (match.Success)
                        {
                            string url = match.Groups[1].Value;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                                $"Extracted URL from HTML: {url}");
                            return url;
                        }
                    }
                }

                // Try plain text format (last resort)
                if (dataObject.GetDataPresent(DataFormats.Text))
                {
                    string text = dataObject.GetData(DataFormats.Text) as string;
                    if (!string.IsNullOrEmpty(text) && IsValidWebUrl(text.Trim()))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                            $"Extracted URL from text: {text.Trim()}");
                        return text.Trim();
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Error extracting URL from drop data: {ex.Message}");
                return null;
            }
        }

        ///// <summary>
        ///// Validates if a string is a valid web URL
        ///// </summary>
        //private static bool IsValidWebUrl(string url)
        //{
        //    if (string.IsNullOrWhiteSpace(url)) return false;

        //    return Uri.TryCreate(url, UriKind.Absolute, out Uri validUri) &&
        //           (validUri.Scheme == Uri.UriSchemeHttp || validUri.Scheme == Uri.UriSchemeHttps ||
        //            validUri.Scheme.Equals("steam", StringComparison.OrdinalIgnoreCase));
        //}


        /// <summary>
        /// Validates if a string is a valid web URL
        /// </summary>
        private static bool IsValidWebUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;

            // Accept ANY valid URI scheme except local files when dropping. 
            // This allows spotify:, steam://, etc., to be successfully created as shortcuts.
            return Uri.TryCreate(url, UriKind.Absolute, out Uri validUri) &&
                   validUri.Scheme != Uri.UriSchemeFile;
        }



        /// <summary>
        /// Adds a URL shortcut to the frame from dropped URL
        /// </summary>
        private static void AddUrlShortcutToFrame(string url, dynamic frame, WrapPanel wrapPanel)
        {
            try
            {
                // Create shortcuts directory if it doesn't exist
                if (!System.IO.Directory.Exists("Shortcuts"))
                {
                    System.IO.Directory.CreateDirectory("Shortcuts");
                }

                // Generate shortcut name from URL
                Uri uri = new Uri(url);
                string displayName = uri.Host.Replace("www.", "");
                string baseShortcutName = System.IO.Path.Combine("Shortcuts", $"{displayName}.url");
                string shortcutPath = baseShortcutName;

                // Ensure unique filename
                int counter = 1;
                while (System.IO.File.Exists(shortcutPath))
                {
                    shortcutPath = System.IO.Path.Combine("Shortcuts", $"{displayName} ({counter++}).url");
                }

                // Create the URL shortcut file  
                CreateWebLinkShortcut(url, shortcutPath, displayName);

                // Create item data for the frame
                dynamic newItem = new System.Dynamic.ExpandoObject();
                IDictionary<string, object> newItemDict = newItem;
                newItemDict["Filename"] = shortcutPath;
                newItemDict["IsFolder"] = false;
                newItemDict["IsLink"] = true;
                newItemDict["IsNetwork"] = false;
                newItemDict["DisplayName"] = displayName;
                newItemDict["AlwaysRunAsAdmin"] = false;

                // Find the frame in FrameDataManager.FrameData and update it properly
                string frameId = frame.Id?.ToString();
                int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);

                if (frameIndex >= 0)
                {
                    // Get the actual frame from FrameDataManager.FrameData
                    dynamic actualFrame = FrameDataManager.FrameData[frameIndex];
                    var actualItems = actualFrame.Items as JArray ?? new JArray();

                    // Handle tabs if enabled
                    bool tabsEnabled = actualFrame.TabsEnabled?.ToString().ToLower() == "true";
                    JArray targetList = actualItems;

                    if (tabsEnabled)
                    {
                        var tabs = actualFrame.Tabs as JArray ?? new JArray();
                        int currentTab = Convert.ToInt32(actualFrame.CurrentTab?.ToString() ?? "0");

                        if (currentTab >= 0 && currentTab < tabs.Count)
                        {
                            var activeTab = tabs[currentTab] as JObject;
                            if (activeTab != null)
                            {
                                var tabItems = activeTab["Items"] as JArray ?? new JArray();
                                tabItems.Add(JObject.FromObject(newItem));
                                targetList = tabItems; // Mark as target for ordering
                            }
                        }
                    }
                    else
                    {
                        actualItems.Add(JObject.FromObject(newItem));
                    }

                    // Set DisplayOrder
                    newItemDict["DisplayOrder"] = targetList.Count - 1;

                    // Save frame data BEFORE creating the icon
                    FrameDataManager.SaveFrameData();

                    // Add icon to UI using Framemanager.AddIcon for proper sizing
                    Framemanager.AddIcon(newItem, wrapPanel);
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                        $"Could not find frame with ID {frameId} in FrameDataManager.FrameData");
                }

                // Add event handlers to the new icon
                StackPanel sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
                if (sp != null)
                {
                    ClickEventAdder(sp, shortcutPath, false);

                    // FIX: Use the FULL Context Menu (Edit, Move, Copy, etc.)
                    // We need to find the parent window to pass it to the menu builder
                    NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(wrapPanel);
                    if (parentWindow != null)
                    {
                        AttachIconContextMenu(sp, newItem, frame, parentWindow);
                    }
                    else
                    {
                        // Fallback (rare) - Should not happen if UI is loaded
                        CreateBasicContextMenu(sp, newItem, shortcutPath);
                    }
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                    $"Successfully added URL shortcut: {displayName} -> {url}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Error adding URL shortcut to frame: {ex.Message}");
            }
        }





        /// <summary>
        /// Gets the next display order for new items
        /// </summary>
        private static int GetNextDisplayOrder(dynamic frame)
        {
            try
            {
                var items = frame.Items as JArray ?? new JArray();
                int maxOrder = items.Count > 0 ? items.Max(i => i["DisplayOrder"]?.Value<int>() ?? 0) : -1;
                return maxOrder + 1;
            }
            catch { return 0; }
        }

        /// <summary>
        /// Creates basic context menu for URL shortcuts
        /// </summary>
        private static void CreateBasicContextMenu(StackPanel sp, dynamic item, string filePath)
        {
            try
            {
                ContextMenu iconContextMenu = new ContextMenu();

                // Fetch frame dynamically to apply style
                try
                {
                    NonActivatingWindow parentWin = FindVisualParent<NonActivatingWindow>(sp);
                    if (parentWin != null)
                    {
                        string frameId = parentWin.Tag?.ToString();
                        var frame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == frameId);
                        if (frame != null)
                        {
                            Style themedStyle = GetThemedContextMenuStyle(frame);
                            if (themedStyle != null) iconContextMenu.Style = themedStyle;
                        }
                    }
                }
                catch { }

                MenuItem miRemove = new MenuItem { Header = Strings.MenuRemove };


                miRemove.Click += (sender, e) =>
                {
                    try
                    {
                        NonActivatingWindow parentWin = FindVisualParent<NonActivatingWindow>(sp);
                        if (parentWin != null)
                        {
                            string frameId = parentWin.Tag?.ToString();
                            var frame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == frameId);
                            if (frame != null)
                            {
                                var items = frame.Items as JArray;
                                var itemToRemove = items?.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                                if (itemToRemove != null)
                                {
                                    items.Remove(itemToRemove);
                                    FrameDataManager.SaveFrameData();

                                    WrapPanel wrapPanel = FindVisualParent<WrapPanel>(sp);
                                    wrapPanel?.Children.Remove(sp);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                            $"Error removing URL shortcut: {ex.Message}");
                    }
                };




                iconContextMenu.Items.Add(miRemove);

                bool alwaysAdmin = Convert.ToBoolean(item["AlwaysRunAsAdmin"] ?? false);
                MenuItem miAlwaysAdmin = new MenuItem
                {
                    Header = Strings.MenuAlwaysRunAsAdmin,
                    IsCheckable = true,
                    IsChecked = alwaysAdmin
                };
                miAlwaysAdmin.Click += (sender, e) => {
                    try
                    {
                        NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(sp);
                        if (parentWindow != null)
                        {
                            string frameId = parentWindow.Tag?.ToString();
                            if (!string.IsNullOrEmpty(frameId))
                            {
                                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                                if (currentFrame != null && currentFrame.ItemsType?.ToString() == "Data")
                                {
                                    var items = currentFrame.Items as JArray ?? new JArray();
                                    bool tabsEnabled = currentFrame.TabsEnabled?.ToString().ToLower() == "true";
                                    if (tabsEnabled)
                                    {
                                        var tabs = currentFrame.Tabs as JArray ?? new JArray();
                                        int currentTabIndex = Convert.ToInt32(currentFrame.CurrentTab?.ToString() ?? "0");
                                        if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                                        {
                                            var currentTab = tabs[currentTabIndex] as JObject;
                                            items = currentTab?["Items"] as JArray ?? items;
                                        }
                                    }
                                    var matchingItem = items.FirstOrDefault(i => string.Equals(
                                        System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                                        System.IO.Path.GetFullPath(filePath),
                                        StringComparison.OrdinalIgnoreCase));
                                    if (matchingItem != null)
                                    {
                                        bool currentValue = Convert.ToBoolean(matchingItem["AlwaysRunAsAdmin"] ?? false);
                                        bool newValue = !currentValue;
                                        matchingItem["AlwaysRunAsAdmin"] = newValue;
                                        miAlwaysAdmin.IsChecked = newValue;
                                        FrameDataManager.SaveFrameData();
                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Toggled AlwaysRunAsAdmin to {newValue} for item: {filePath}");
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error toggling AlwaysRunAsAdmin for {filePath}: {ex.Message}");
                    }
                };
                // Refresh IsChecked on menu open
                iconContextMenu.Opened += (s, ev) => {
                    try
                    {
                        NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(sp);

                        // --- LIVE THEME FIX ---
                        if (parentWindow != null)
                        {
                            string frameId = parentWindow.Tag?.ToString();
                            var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == frameId);
                            if (liveFrame != null)
                            {
                                Style liveStyle = GetThemedContextMenuStyle(liveFrame);
                                if (liveStyle != null) iconContextMenu.Style = liveStyle;
                            }
                        }
                        if (parentWindow != null)
                        {
                            string frameId = parentWindow.Tag?.ToString();
                            if (!string.IsNullOrEmpty(frameId))
                            {
                                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                                if (currentFrame != null)
                                {
                                    var items = currentFrame.Items as JArray ?? new JArray();
                                    bool tabsEnabled = currentFrame.TabsEnabled?.ToString().ToLower() == "true";
                                    if (tabsEnabled)
                                    {
                                        var tabs = currentFrame.Tabs as JArray ?? new JArray();
                                        int currentTabIndex = Convert.ToInt32(currentFrame.CurrentTab?.ToString() ?? "0");
                                        if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                                        {
                                            var currentTab = tabs[currentTabIndex] as JObject;
                                            items = currentTab?["Items"] as JArray ?? items;
                                        }
                                    }
                                    var matchingItem = items.FirstOrDefault(i => string.Equals(
                                        System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                                        System.IO.Path.GetFullPath(filePath),
                                        StringComparison.OrdinalIgnoreCase));
                                    if (matchingItem != null)
                                    {
                                        miAlwaysAdmin.IsChecked = Convert.ToBoolean(matchingItem["AlwaysRunAsAdmin"] ?? false);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Error refreshing AlwaysRunAsAdmin IsChecked for {filePath}: {ex.Message}");
                    }
                };
                iconContextMenu.Items.Add(miAlwaysAdmin);


          sp.ContextMenu = iconContextMenu;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Error creating context menu: {ex.Message}");
            }
        }
             
        public static List<dynamic> GetFrameData()
        {
            return FrameDataManager.FrameData;
        }

        /// <summary>
        /// Public access to portal frames dictionary for CustomizeFrameForm
        /// </summary>
        public static Dictionary<dynamic, PortalFramemanager> GetPortalFrames()
        {
            return _portalFrames;
        }





        /// <summary>
        /// Updates the filter history for a frame using LRU (Least Recently Used) logic.
        /// Max 5 items. Duplicates move to top.
        /// </summary>
        private static void UpdateFilterHistory(dynamic frame, string newFilter)
        {
            if (string.IsNullOrWhiteSpace(newFilter)) return;

            // Don't save if it matches a Standard Preset value exactly
            if (_standardPresets.ContainsValue(newFilter)) return;

            try
            {
                // 1. GET FRESH DATA (Fixes the "Replace" bug)
                // We cannot rely on the 'frame' parameter because it might be stale 
                // (referencing the state from when the window opened).
                string frameId = frame.Id?.ToString();
                if (string.IsNullOrEmpty(frameId)) return;

                int index = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                if (index < 0) return;

                // Grab the LIVE object from the master list
                dynamic liveFrame = FrameDataManager.FrameData[index];

                // 2. Convert to Dictionary for modification
                IDictionary<string, object> frameDict = liveFrame is IDictionary<string, object> dict
                    ? dict
                    : ((JObject)liveFrame).ToObject<IDictionary<string, object>>();

                // 3. Get existing history
                JArray history;
                if (frameDict.ContainsKey("FilterHistory") && frameDict["FilterHistory"] is JArray existingArr)
                {
                    history = existingArr;
                }
                else
                {
                    history = new JArray();
                }

                // 4. Update Logic (LRU)
                // Remove existing if present (to bump to top)
                for (int i = history.Count - 1; i >= 0; i--)
                {
                    if (string.Equals(history[i].ToString(), newFilter, StringComparison.OrdinalIgnoreCase))
                    {
                        history.RemoveAt(i);
                    }
                }

                // Insert at top
                history.Insert(0, newFilter);

                // Cap at 5 items
                while (history.Count > 5)
                {
                    history.RemoveAt(history.Count - 1);
                }

                // 5. Save back to Master List
                frameDict["FilterHistory"] = history;
                FrameDataManager.FrameData[index] = JObject.FromObject(frameDict);
                FrameDataManager.SaveFrameData();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings,
                    $"Updated Filter History for {frameId}: {string.Join(", ", history)}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Settings, $"History Error: {ex.Message}");
            }
        }


        /// <summary>
        /// Resets visual customizations for all frames to default values.
        /// Preserves position, size, content, and system state (Locked/Rolled).
        /// </summary>
        public static void ResetAllCustomizations()
        {
            try
            {
                bool modificationsMade = false;

                // FIX: Use .ToList() to create a snapshot copy for iteration.
                // This prevents "Collection was modified" errors when we update the real list inside the loop.
                var framesSnapshot = FrameDataManager.FrameData.ToList();

                foreach (dynamic frame in framesSnapshot)
                {
                    // Convert to JObject/Dictionary for safe property access
                    IDictionary<string, object> frameDict = frame is IDictionary<string, object> dict
                        ? dict
                        : ((JObject)frame).ToObject<IDictionary<string, object>>();

                    // --- RESET VISUALS TO DEFAULTS ---

                    // Frame Appearance
                    frameDict["CustomColor"] = null;
                    frameDict["CustomLaunchEffect"] = null;
                    frameDict["FrameBorderColor"] = null;
                    frameDict["FrameBorderThickness"] = 2; // Default thickness

                    // Title Appearance
                    frameDict["TitleTextColor"] = null;
                    frameDict["TitleTextSize"] = "Medium";
                    frameDict["BoldTitleText"] = "false";

                    // Icon Appearance
                    frameDict["TextColor"] = null;
                    frameDict["DisableTextShadow"] = "false";
                    frameDict["IconSize"] = "Medium";
                    frameDict["IconSpacing"] = 5;
                    frameDict["GrayscaleIcons"] = "false";

                    // Update the object in the REAL list
                    int index = FrameDataManager.FrameData.IndexOf(frame);
                    if (index >= 0)
                    {
                        FrameDataManager.FrameData[index] = JObject.FromObject(frameDict);
                        modificationsMade = true;
                    }
                }

                if (modificationsMade)
                {
                    FrameDataManager.SaveFrameData();

                    // Reload to reflect changes immediately
                    ReloadFrames();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings,
                        "All frame customizations have been reset to defaults.");

                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(
                        Strings.MsgCustomizationsReset,
                        Strings.DlgResetComplete);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Error,
                    $"Error resetting customizations: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgResetCustomizationsFailed", ex.Message), Strings.DlgError);
            }
        }



        // Update frame property, save to JSON, and apply runtime changes
        public static void UpdateFrameProperty(dynamic frame, string propertyName, string value, string logMessage)
        {
            try
            {


                string frameId = frame.Id?.ToString();
                if (string.IsNullOrEmpty(frameId))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Frame '{frame.Title}' has no Id");
                    return;
                }

                // Skip updates if frame is in transition (except for IsRolled and UnrolledHeight which are rollup-specific)
                if (_framesInTransition.Contains(frameId) && propertyName != "IsRolled" && propertyName != "UnrolledHeight")
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameUpdate, $"Skipping {propertyName} update for frame '{frameId}' - in transition");
                    return;
                }

                // Find by GUID instead of title or reference
                int index = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                if (index >= 0)
                {
                    // Get the frame from FrameDataManager.FrameData
                    dynamic actualFrame = FrameDataManager.FrameData[index];





                    // Convert to dictionary safely
                    IDictionary<string, object> frameDict = actualFrame as IDictionary<string, object> ?? ((JObject)actualFrame).ToObject<IDictionary<string, object>>();

                    // Handle IsHidden specifically to store as string to match JSON format
                    if (propertyName == "IsHidden")
                    {
                        // Convert boolean-like string input to string "true" or "false"
                        bool parsedValue = value?.ToLower() == "true";
                        frameDict[propertyName] = parsedValue.ToString().ToLower(); // Store as "true" or "false"
                    }
                    else if (propertyName == "FrameBorderThickness")
                    {
                        // BUG FIX: Enforce integer serialization to prevent JSON type wiping/fallback
                        if (int.TryParse(value, out int parsedThickness))
                        {
                            frameDict[propertyName] = parsedThickness;
                        }
                        else
                        {
                            frameDict[propertyName] = 2; // Failsafe
                        }
                    }
                    else
                    {
                        // Update other properties as provided
                        frameDict[propertyName] = value;
                    }

                    // Update the frame in FrameDataManager.FrameData
                    FrameDataManager.FrameData[index] = JObject.FromObject(frameDict);
                    FrameDataManager.SaveFrameData();
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"{logMessage} for frame '{frame.Title}'");

                    // Find the window to apply runtime changes
                    var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                    //var win = windows.FirstOrDefault(w => w.Title == frame.Title.ToString());
                    var win = windows.FirstOrDefault(w => w.Tag?.ToString() == frameId);
                    if (win != null)
                    {
                        // Apply runtime changes
                        if (propertyName == "CustomColor")
                        {
                            Utility.ApplyTintAndColorToFrame(win, value);
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Applied color '{value ?? "Default"}' to frame '{frame.Title}' at runtime");
                        }
                        else if (propertyName == "IsHidden")
                        {
                            // Update visibility based on IsHidden
                            bool isHidden = value?.ToLower() == "true";
                            win.Visibility = isHidden ? Visibility.Hidden : Visibility.Visible;
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Set visibility to {(isHidden ? "Hidden" : "Visible")} for frame '{frame.Title}'");
                        }
                        else if (propertyName == "AlwaysOnTop")
                        {
                            bool alwaysOnTop = value?.ToLower() == "true";
                            win.Topmost = alwaysOnTop;

                            // --- BULLETPROOF FIX: Delay Native OS Call to beat WPF's internal queue ---
                            win.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                // Brute-force WPF to re-evaluate the handle state
                                if (alwaysOnTop) { win.Topmost = false; win.Topmost = true; }

                                var hwnd = new System.Windows.Interop.WindowInteropHelper(win).Handle;
                                if (hwnd != IntPtr.Zero)
                                {
                                    IntPtr HWND_TOPMOST = new IntPtr(-1);
                                    IntPtr HWND_NOTOPMOST = new IntPtr(-2);
                                    uint SWP_NOMOVE = 0x0002;
                                    uint SWP_NOSIZE = 0x0001;
                                    uint SWP_NOACTIVATE = 0x0010;

                                    SetWindowPos(hwnd, alwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                                }
                            }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);

                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Set Topmost to {alwaysOnTop} for frame '{frame.Title}'");
                        }
                        //step 7
                        else if (propertyName == "IsRolled")
                        {
                            bool isRolled = value?.ToLower() == "true";
                            double targetHeight = isRolled ? 28 : Convert.ToDouble(actualFrame.UnrolledHeight?.ToString() ?? "130"); //rolled height
                            var heightAnimation = new DoubleAnimation(win.Height, targetHeight, TimeSpan.FromSeconds(0.3))
                            {
                                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                            };
                            win.BeginAnimation(Window.HeightProperty, heightAnimation);
                            // Update WrapPanel visibility
                            var border = win.Content as Border;
                            if (border != null)
                            {
                                var dockPanel = border.Child as DockPanel;
                                if (dockPanel != null)
                                {
                                    var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                                    if (scrollViewer != null)
                                    {
                                        var wpcont = scrollViewer.Content as WrapPanel;
                                        if (wpcont != null)
                                        {
                                            wpcont.Visibility = isRolled ? Visibility.Collapsed : Visibility.Visible;
                                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Set WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for frame '{actualFrame.Title}'");

                                            // --- MULTI-MONITOR INVISIBLE ICON FIX ---
                                            if (!isRolled)
                                            {
                                                wpcont.UpdateLayout();
                                                wpcont.InvalidateVisual();
                                                wpcont.Opacity = 0.99;
                                                wpcont.Dispatcher.BeginInvoke(new Action(() => { wpcont.Opacity = 1.0; }), System.Windows.Threading.DispatcherPriority.Render);
                                            }
                                        }
                                    }
                                }
                            }
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Set height to {targetHeight} for frame '{actualFrame.Title}' (IsRolled={isRolled})");
                        }
                        else if (propertyName == "UnrolledHeight")
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Set UnrolledHeight to {value} for frame '{actualFrame.Title}'");
                        }
                        else if (propertyName == "FrameBorderThickness")
                        {
                            if (win.Content is Border frameBorder)
                            {
                                int thickness = 2;
                                int.TryParse(value, out thickness);
                                frameBorder.BorderThickness = new Thickness(thickness);

                                if (thickness > 0 && frameBorder.BorderBrush == null)
                                {
                                    string borderColorName = actualFrame.FrameBorderColor?.ToString();
                                    if (!string.IsNullOrEmpty(borderColorName))
                                        frameBorder.BorderBrush = new SolidColorBrush(Utility.GetColorFromName(borderColorName));
                                    else
                                        frameBorder.BorderBrush = new SolidColorBrush(System.Windows.Media.Colors.Gray);
                                }
                                else if (thickness == 0)
                                {
                                    frameBorder.BorderBrush = null;
                                }
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Set runtime FrameBorderThickness to {thickness} for frame '{actualFrame.Title}'");
                            }
                        }

                        // Update context menu checkmarks
                        if (win.ContextMenu != null)
                        {
                            var customizeMenu = win.ContextMenu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header.ToString() == "Customize");
                            if (customizeMenu != null)
                            {
                                var submenu = propertyName == "CustomColor"
                                    ? customizeMenu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header.ToString() == "Color")
                                    : customizeMenu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header.ToString() == "Launch Effect");
                                if (submenu != null)
                                {
                                    foreach (MenuItem item in submenu.Items)
                                    {
                                        item.IsChecked = item.Tag?.ToString() == value || (value == null && item.Header.ToString() == "Default");
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FrameUpdate, $"Failed to find window for frame '{frame.Title}' to apply {propertyName}");
                    }
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FrameUpdate, $"Failed to find frame '{frame.Title}' in FrameDataManager.FrameData for {propertyName} update");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameUpdate, $"Error updating {propertyName} for frame '{frame.Title}': {ex.Message}");
            }
        }



        // TABS FEATURE: Toggle tabs with strict Data Migration and Effect Cleanup
        private static void ToggleFrameTabs(dynamic frame)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Toggling tabs for frame '{frame.Title}'");

                string frameId = frame.Id?.ToString();
                int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                if (frameIndex < 0) return;

                dynamic liveFrame = FrameDataManager.FrameData[frameIndex];
                IDictionary<string, object> frameDict = liveFrame is IDictionary<string, object> dict
                    ? dict : ((JObject)liveFrame).ToObject<IDictionary<string, object>>();

                bool currentTabsEnabled = frameDict.ContainsKey("TabsEnabled") && frameDict["TabsEnabled"]?.ToString().ToLower() == "true";
                bool newTabsEnabled = !currentTabsEnabled;

                var mainItems = frameDict["Items"] as JArray ?? new JArray();
                var tabs = frameDict["Tabs"] as JArray ?? new JArray();

                // 1. Data Migration Logic
                if (newTabsEnabled)
                {
                    // ENABLING TABS: Move Main -> Tab
                    if (tabs.Count == 0)
                    {
                        // Create first tab with Herb Naming
                        dynamic newTab = new JObject();
                        string herbName = FrameUtilities.GenerateRandomHerbName();
                        newTab.TabName = $"0. {herbName}";
                        newTab.Items = mainItems;
                        tabs.Add(newTab);
                        frameDict["CurrentTab"] = 0;
                    }
                    else
                    {
                        int targetIdx = Convert.ToInt32(frameDict["CurrentTab"]?.ToString() ?? "0");
                        if (targetIdx < 0 || targetIdx >= tabs.Count) targetIdx = 0;

                        var targetTab = tabs[targetIdx] as JObject;
                        var targetItems = targetTab["Items"] as JArray ?? new JArray();

                        foreach (var item in mainItems) targetItems.Add(item);
                        targetTab["Items"] = targetItems;
                        frameDict["CurrentTab"] = targetIdx;
                    }
                    // Clear main items after moving them to tab
                    frameDict["Items"] = new JArray();
                    frameDict["Tabs"] = tabs;
                }
                else
                {
                    // DISABLING TABS: Move Active Tab -> Main
                    int currentTabIdx = Convert.ToInt32(frameDict["CurrentTab"]?.ToString() ?? "0");
                    if (tabs.Count > 0 && currentTabIdx < tabs.Count && currentTabIdx >= 0)
                    {
                        var activeTab = tabs[currentTabIdx] as JObject;
                        var activeItems = activeTab["Items"] as JArray;
                        if (activeItems != null)
                        {
                            // FIX: Clear mainItems first! 
                            // This prevents duplication because mainItems might already contain stale copies 
                            // (especially from Tab 0 syncing).
                            mainItems.Clear();

                            foreach (var item in activeItems) mainItems.Add(item);
                            activeTab["Items"] = new JArray();
                        }
                    }
                    frameDict["Items"] = mainItems;
                }

                // 2. Save Updates
                frameDict["TabsEnabled"] = newTabsEnabled.ToString().ToLower();

                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                var currentWindow = windows.FirstOrDefault(w => w.Tag?.ToString() == frameId);
                if (currentWindow != null)
                {
                    frameDict["X"] = currentWindow.Left;
                    frameDict["Y"] = currentWindow.Top;
                    frameDict["Width"] = currentWindow.Width;
                    frameDict["Height"] = currentWindow.Height;
                }

                FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                FrameDataManager.SaveFrameData();

                // 3. Reload Window & CLEANUP EFFECTS
                if (currentWindow != null)
                {
                    // Cleanup Legendary Effects
                    string currentTitle = frame.Title?.ToString() ?? "";
                    if (currentTitle == "Nikos" || currentTitle == "Nikos Georgousis" || currentTitle.Contains(">:"))
                    {
                        InterCore.ProcessTitleChange(liveFrame, "RESET_EFFECT_TEMP", "");
                    }

                    _heartTextBlocks.Remove(frame);
                    if (_portalFrames.ContainsKey(frame))
                    {
                        _portalFrames[frame].Dispose();
                        _portalFrames.Remove(frame);
                    }
                    currentWindow.Close();
                }

                var updatedFrame = FrameDataManager.FrameData[frameIndex];

                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    CreateFrame(updatedFrame, new TargetChecker(1000));
                }));
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error toggling tabs: {ex.Message}");
            }
        }





        // TABS FEATURE: Simplified content refresh that reuses existing infrastructure

        public static void RefreshFrameContentSimple(NonActivatingWindow frameWindow, dynamic frame, int tabIndex)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"RefreshFrameContentSimple: switching to tab {tabIndex}");

                // 1. Find the WrapPanel
                var border = frameWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();
                var wrapPanel = scrollViewer?.Content as WrapPanel;

                if (wrapPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot find WrapPanel for content refresh");
                    return;
                }

                // 2. Clear existing content
                wrapPanel.Children.Clear();

                // 3. Get items from the specified tab
                var tabs = frame.Tabs as JArray ?? new JArray();
                if (tabIndex >= 0 && tabIndex < tabs.Count)
                {
                    var activeTab = tabs[tabIndex] as JObject;
                    if (activeTab != null)
                    {
                        var items = activeTab["Items"] as JArray ?? new JArray();
                        string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {tabIndex}";

                        // 4. Sort items by DisplayOrder
                        var sortedItems = items
                            .OfType<JObject>()
                            .OrderBy(item => item["DisplayOrder"]?.Type == JTokenType.Integer ? item["DisplayOrder"].Value<int>() : 0)
                            .ToList();

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Loading {sortedItems.Count} items from tab '{tabName}'");

                        // 5. Iterate and Add Icons (Using Unified Logic)
                        foreach (dynamic icon in sortedItems)
                        {
                            // FIX A: Use the main Framemanager.AddIcon method!
                            // This ensures network indicators, clean icons, and custom sizes are applied.
                            AddIcon(icon, wrapPanel, frame);

                            StackPanel sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                // FIX B: Extract Properties safely
                                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict
                                    ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];

                                // FIX C: Extract Arguments
                                string arguments = null;
                                if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
                                {
                                    try
                                    {
                                        WshShell shell = new WshShell();
                                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                        arguments = shortcut.Arguments;
                                    }
                                    catch { }
                                }

                                // FIX D: Attach Events
                                ClickEventAdder(sp, filePath, isFolder, arguments);

                                // FIX E: Attach Centralized Context Menu (Crucial for Right-Click)
                                AttachIconContextMenu(sp, icon, frame, frameWindow);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in RefreshFrameContentSimple: {ex.Message}");
            }
        }






        // TABS FEATURE: Remove item from current tab
        private static void RemoveItemFromTab(dynamic item, dynamic frame, int tabIndex, WrapPanel wrapPanel)
        {
            try
            {
                var tabs = frame.Tabs as JArray ?? new JArray();
                if (tabIndex >= 0 && tabIndex < tabs.Count)
                {
                    var activeTab = tabs[tabIndex] as JObject;
                    if (activeTab != null)
                    {
                        var items = activeTab["Items"] as JArray ?? new JArray();
                        string itemFilename = item.Filename?.ToString();

                        // Find and remove the item
                        for (int i = items.Count - 1; i >= 0; i--)
                        {
                            var currentItem = items[i] as JObject;
                            if (currentItem != null && currentItem["Filename"]?.ToString() == itemFilename)
                            {
                                items.RemoveAt(i);
                                break;
                            }
                        }

                        // Update the tab
                        activeTab["Items"] = items;

                        // Save changes
                        string frameId = frame.Id?.ToString();
                        int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                        if (frameIndex >= 0)
                        {
                            FrameDataManager.FrameData[frameIndex] = frame;
                            FrameDataManager.SaveFrameData();
                            // ENHANCED: Add Tab0 synchronization after removal
                            if (tabIndex == 0)
                            {
                                // string frameId = frame.Id?.ToString();
                                if (!string.IsNullOrEmpty(frameId))
                                {
                                    TabManager.SynchronizeTab0Content(frameId, "tab0", "remove");
                                }
                            }

                        }

                        // Refresh the display
                        var frameWindow = FindVisualParent<NonActivatingWindow>(wrapPanel);
                        if (frameWindow != null)
                        {
                            RefreshFrameContentSimple(frameWindow, frame, tabIndex);
                        }

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Removed item '{itemFilename}' from tab {tabIndex}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error removing item from tab: {ex.Message}");
            }
        }






        // TABS FEATURE: Create custom button template with rounded top corners only
        private static ControlTemplate CreateTabButtonTemplate()
        {
            var template = new ControlTemplate(typeof(Button));

            // Create the border with rounded top corners
            var border = new FrameworkElementFactory(typeof(Border));
            border.Name = "border";
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(4, 4, 0, 0)); // Top corners rounded
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Button.BackgroundProperty));
            border.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Button.BorderBrushProperty));
            border.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Button.BorderThicknessProperty));
            border.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));

            // Create the content presenter
            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(Button.ContentProperty));

            border.AppendChild(contentPresenter);
            template.VisualTree = border;

            // Add hover trigger for better interaction feedback
            var trigger = new Trigger();
            trigger.Property = Button.IsMouseOverProperty;
            trigger.Value = true;
            trigger.Setters.Add(new Setter(Button.OpacityProperty, 0.8));

            template.Triggers.Add(trigger);

            return template;
        }






        // Helper event handlers to avoid lambda capture issues
        private static void TabButton_MouseEnter(object sender, RoutedEventArgs e) { }

        private static void TabButton_MouseLeave(object sender, RoutedEventArgs e) { }

        private static void MigrateLegacyJson()
        {
            try
            {
                bool jsonModified = false;
                var validColors = new HashSet<string> { "Red", "Green", "Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple", "Fuchsia", "Yellow", "Orange" };
                var validEffects = Enum.GetNames(typeof(LaunchEffectsManager.LaunchEffect)).ToHashSet();

                for (int i = 0; i < FrameDataManager.FrameData.Count; i++)
                {
                    dynamic frame = FrameDataManager.FrameData[i];
                    IDictionary<string, object> frameDict = frame is IDictionary<string, object> dict
                        ? dict : ((JObject)frame).ToObject<IDictionary<string, object>>();

                    // --- 1. CORE VALIDATION (VITAL) ---
                    // Add GUID if missing
                    if (!frameDict.ContainsKey("Id"))
                    {
                        frameDict["Id"] = Guid.NewGuid().ToString();
                        jsonModified = true;
                    }

                    // --- 2. PORTAL frame MIGRATION ---
                    if (frame.ItemsType?.ToString() == "Portal")
                    {
                        string portalPath = frame.Items?.ToString();
                        if (!string.IsNullOrEmpty(portalPath) && !System.IO.Directory.Exists(portalPath))
                        {
                            frameDict["IsFolder"] = true;
                            jsonModified = true;
                        }
                    }
                    else
                    {
                        var items = frame.Items as JArray ?? new JArray();
                        bool itemsModified = false;
                        foreach (var item in items)
                        {
                            var itemDict = item as JObject;
                            if (itemDict != null)
                            {
                                if (itemDict["IsFolder"] == null)
                                {
                                    string path = itemDict["Filename"]?.ToString();
                                    itemDict["IsFolder"] = System.IO.Directory.Exists(path);
                                    itemsModified = true;
                                }
                                // Ensure IsLink exists
                                if (!itemDict.ContainsKey("IsLink"))
                                {
                                    itemDict["IsLink"] = false;
                                    itemsModified = true;
                                }
                                // Ensure IsNetwork exists
                                if (!itemDict.ContainsKey("IsNetwork"))
                                {
                                    string fname = itemDict["Filename"]?.ToString() ?? "";
                                    itemDict["IsNetwork"] = IsNetworkPath(fname);
                                    itemsModified = true;
                                }
                                // Ensure AlwaysRunAsAdmin exists
                                if (!itemDict.ContainsKey("AlwaysRunAsAdmin"))
                                {
                                    itemDict["AlwaysRunAsAdmin"] = false;
                                    itemsModified = true;
                                }
                            }
                        }
                        if (itemsModified)
                        {
                            frameDict["Items"] = items;
                            jsonModified = true;
                        }
                    }

                    // --- 3. PROPERTY INITIALIZATION (DEFAULTS) ---
                    // Visuals
                    if (!frameDict.ContainsKey("IsLocked")) { frameDict["IsLocked"] = "false"; jsonModified = true; }
                    if (!frameDict.ContainsKey("IsRolled")) { frameDict["IsRolled"] = "false"; jsonModified = true; }
                    if (!frameDict.ContainsKey("AutoRoll")) { frameDict["AutoRoll"] = "false"; jsonModified = true; } // --- NEW ---
                    if (!frameDict.ContainsKey("AlwaysOnTop")) { frameDict["AlwaysOnTop"] = "false"; jsonModified = true; } // --- NEW ---
                    if (!frameDict.ContainsKey("UnrolledHeight"))
                    {
                        double height = frameDict.ContainsKey("Height") ? Convert.ToDouble(frameDict["Height"]) : 130;
                        frameDict["UnrolledHeight"] = height;
                        jsonModified = true;
                    }

                    if (!frameDict.ContainsKey("IconSize")) { frameDict["IconSize"] = "Medium"; jsonModified = true; }
                    if (!frameDict.ContainsKey("IconSpacing")) { frameDict["IconSpacing"] = 5; jsonModified = true; }
                    if (!frameDict.ContainsKey("CustomColor")) { frameDict["CustomColor"] = null; jsonModified = true; }
                    if (!frameDict.ContainsKey("CustomLaunchEffect")) { frameDict["CustomLaunchEffect"] = null; jsonModified = true; }
                    if (!frameDict.ContainsKey("IsHidden")) { frameDict["IsHidden"] = "false"; jsonModified = true; }

                    // Text
                    if (!frameDict.ContainsKey("TextColor")) { frameDict["TextColor"] = null; jsonModified = true; }
                    if (!frameDict.ContainsKey("TitleTextColor")) { frameDict["TitleTextColor"] = null; jsonModified = true; }
                    if (!frameDict.ContainsKey("TitleTextSize")) { frameDict["TitleTextSize"] = "Medium"; jsonModified = true; }
                    if (!frameDict.ContainsKey("BoldTitleText")) { frameDict["BoldTitleText"] = "false"; jsonModified = true; }
                    if (!frameDict.ContainsKey("DisableTextShadow")) { frameDict["DisableTextShadow"] = "false"; jsonModified = true; }
                    if (!frameDict.ContainsKey("GrayscaleIcons")) { frameDict["GrayscaleIcons"] = "false"; jsonModified = true; }

                    // Border
                    if (!frameDict.ContainsKey("FrameBorderColor")) { frameDict["FrameBorderColor"] = null; jsonModified = true; }
                    if (!frameDict.ContainsKey("FrameBorderThickness")) { frameDict["FrameBorderThickness"] = 2; jsonModified = true; }

                    // --- 4. TABS FEATURE (Structure Only, No Merge) ---
                    if (!frameDict.ContainsKey("TabsEnabled"))
                    {
                        frameDict["TabsEnabled"] = "false";
                        jsonModified = true;
                    }
                    if (!frameDict.ContainsKey("CurrentTab"))
                    {
                        frameDict["CurrentTab"] = 0;
                        jsonModified = true;
                    }
                    if (!frameDict.ContainsKey("Tabs"))
                    {
                        frameDict["Tabs"] = new JArray();
                        jsonModified = true;
                    }

                    // --- 5. TAB VALIDATION (Structure Repair) ---
                    // Ensure existing tabs have valid structure (Name, Items Array)
                    // but DO NOT move items around or merge lists.
                    var tabs = frameDict["Tabs"] as JArray ?? new JArray();
                    if (tabs.Count > 0)
                    {
                        bool tabsModified = false;
                        for (int t = 0; t < tabs.Count; t++)
                        {
                            var tab = tabs[t] as JObject;
                            if (tab == null) continue;

                            if (tab["TabName"] == null)
                            {
                                tab["TabName"] = $"Tab {t + 1}";
                                tabsModified = true;
                            }
                            if (tab["Items"] == null || !(tab["Items"] is JArray))
                            {
                                tab["Items"] = new JArray();
                                tabsModified = true;
                            }
                        }
                        if (tabsModified)
                        {
                            frameDict["Tabs"] = tabs;
                            jsonModified = true;
                        }
                    }

                    // Update Master List
                    FrameDataManager.FrameData[i] = JObject.FromObject(frameDict);
                }

                if (jsonModified)
                {
                    FrameDataManager.SaveFrameData();
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Migrated fences.json with updated fields (Validations Only)");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error migrating fences.json: {ex.Message}");
            }
        }


        public static void LoadAndCreateFrames(TargetChecker targetChecker)
        {
            // Get current program version from assembly
            string currentVersion = Assembly.GetEntryAssembly()?.GetName().Version?.ToString() ?? "0.0.0";

            // --- EXECUTABLE MIGRATION ---
            RegistryHelper.PerformRenameMigration();

            RegistryHelper.SetProgramManagementValues(currentVersion);

            // CRITICAL FIX: Load settings EARLY so the single instance check knows the user's preference
            FrameDataManager.Initialize();
            SettingsManager.LoadSettings();


            // --- NEW: Register internal plugins before loading frames ---
            PluginManager.Initialize();

            // --- ONE-TIME ALTGR KEYBOARD CONFLICT WARNING FOR EXISTING USERS ---
            if (!SettingsManager.AltGrWarningShown && SettingsManager.EnableProfileHotkeys)
            {
                // Check if they are actually using the problematic Ctrl+Alt combination
                string mods = SettingsManager.ProfileSwitchModifier?.ToLower() ?? "";
                if (mods.Contains("control") && mods.Contains("alt"))
                {
                    bool disableHotkeys = MessageBoxesManager.ShowCustomYesNoMessageBox(
                                  Strings.MsgConfirmDisableCtrlAltHotkeys,
                                  Strings.DlgKeyboardCheck,
                                  NotificationSound.NadaAlert); // Override the user's sound preference for this critical alert

                    if (disableHotkeys)
                    {
                        SettingsManager.EnableProfileHotkeys = false;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.Settings, "User opted to disable Ctrl+Alt profile hotkeys due to AltGr conflict.");
                    }
                }

                // Mark as shown so we never bother them again, regardless of their choice
                SettingsManager.AltGrWarningShown = true;
                SettingsManager.SaveSettings();
                SettingsManager.BroadcastHotkeysToAllProfiles(); // Sync this decision across profiles
            }
            // -------------------------------------------------------------------

      
            // === PRE-FLIGHT CHECK FOR SINGLE INSTANCE SETTING ===
            // Directly read JSON to ensure we get the active profile's preference before UI fully boots
            bool disableSingleInstance = SettingsManager.DisableSingleInstance;
            try
            {
                string baseDir = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                string optionsPath = System.IO.Path.Combine(baseDir, "options.json"); // Legacy fallback
                string masterOptionsPath = System.IO.Path.Combine(baseDir, "MasterOptions.json");

                if (System.IO.File.Exists(masterOptionsPath))
                {
                    optionsPath = masterOptionsPath;
                }
                else
                {
                    string activeProfilePath = System.IO.Path.Combine(baseDir, "ActiveProfile.txt");
                    if (System.IO.File.Exists(activeProfilePath))
                    {
                        string activeProfile = System.IO.File.ReadAllText(activeProfilePath).Trim();
                        string profileOptionsPath = System.IO.Path.Combine(baseDir, "Profiles", activeProfile, "options.json");
                        if (System.IO.File.Exists(profileOptionsPath))
                        {
                            optionsPath = profileOptionsPath;
                        }
                    }
                }

                if (System.IO.File.Exists(optionsPath))
                {
                    string jsonContent = System.IO.File.ReadAllText(optionsPath);
                    var jObj = JObject.Parse(jsonContent);
                    if (jObj["DisableSingleInstance"] != null)
                    {
                        disableSingleInstance = jObj["DisableSingleInstance"].Value<bool>();
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Pre-flight check failed: {ex.Message}");
            }

            // === SINGLE INSTANCE CHECK (WITH DISABLE OPTION) ===
            try
            {
                // Check if single instance enforcement is disabled
                if (disableSingleInstance)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "SingleInstance: Single instance enforcement disabled. Multiple instances allowed.");
                }
                else
                {
                    // Small delay to ensure process visibility
                    System.Threading.Thread.Sleep(100);

                    Process currentProcess = Process.GetCurrentProcess();
                    string processName = Path.GetFileNameWithoutExtension(currentProcess.ProcessName);
                    Process[] allInstances = Process.GetProcessesByName(processName);

                    // Check if this is a duplicate instance
                    if (allInstances.Length > 1)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                            $"SingleInstance: Duplicate instance detected. Found {allInstances.Length} instances. Writing trigger and exiting.");

                        // Write registry trigger for the original instance
                        bool registryWritten = RegistryHelper.WriteTrigger();

                        if (registryWritten)
                        {
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                "SingleInstance: Registry trigger written successfully. Exiting duplicate instance.");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                                "SingleInstance: Failed to write registry trigger. Still exiting duplicate instance.");
                        }

                        Environment.Exit(0);
                        return;
                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "SingleInstance: This is the first instance. Continuing with normal startup.");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"SingleInstance: Error in single instance check: {ex.Message}. Continuing startup.");
            }





            // REMOVED: Legacy path override. 
            // FrameDataManager.Initialize() now sets the correct Profile path internally.

            // Below added for reload function
            _currentTargetChecker = targetChecker;

            // Delete previous log file if setting is enabled
            if (SettingsManager.DeletePreviousLogOnStart)
            {
                try
                {
                    string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Frames.log");
                    if (System.IO.File.Exists(logPath))
                    {
                        System.IO.File.Delete(logPath);
                        // Note: Can't log this deletion since we just deleted the log file
                    }
                }
                catch
                {
                    // Silently ignore deletion errors to avoid startup issues
                }
            }


            BackupManager.CleanLastDeletedFolder();
            InterCore.Initialize();

            _options = new
            {
                IsSnapEnabled = SettingsManager.IsSnapEnabled,
                ShowBackgroundImageOnPortalFences = SettingsManager.ShowBackgroundImageOnPortalFrames,
                Showintray = SettingsManager.ShowInTray,
                EnableSounds = SettingsManager.EnableSounds,
                TintValue = SettingsManager.TintValue,
                MenuTintValue = SettingsManager.MenuTintValue,
                MenuIcon = SettingsManager.MenuIcon,
                LockIcon = SettingsManager.LockIcon,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect,
                CheckNetworkPaths = false 
            };

            bool jsonLoadSuccessful = false;

            if (System.IO.File.Exists(FrameDataManager.JsonFilePath))
            {
                jsonLoadSuccessful = LoadFrameDataFromJson();
            }

            // If JSON loading failed or file doesn't exist, initialize with defaults
            if (!jsonLoadSuccessful || FrameDataManager.FrameData == null || FrameDataManager.FrameData.Count == 0)
            {
                InitializeDefaultFrame();
            }
            else
            {
                // Only migrate if we successfully loaded existing data
                MigrateLegacyJson();
            }

            // Sanitize Portal Frames with missing target folders
            var invalidFrames = new List<dynamic>();
            foreach (dynamic frame in FrameDataManager.FrameData.ToList()) // Use ToList to avoid collection modification issues
            {
                if (frame.ItemsType?.ToString() == "Portal")
                {
                    string targetPath = frame.Path?.ToString();
                    if (string.IsNullOrEmpty(targetPath) || !System.IO.Directory.Exists(targetPath))
                    {
                        invalidFrames.Add(frame);
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FrameCreation, $"Marked Portal Frame '{frame.Title}' for removal due to missing target folder: {targetPath ?? "null"}");
                    }
                }
            }

            // Remove invalid frames and save
            if (invalidFrames.Any())
            {
                foreach (var frame in invalidFrames)
                {
                    FrameDataManager.FrameData.Remove(frame);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Removed Portal Frame '{frame.Title}' from FrameDataManager.FrameData");
                }
                FrameDataManager.SaveFrameData();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Saved updated frames.json after removing {invalidFrames.Count} invalid Portal Frames");
            }

            // Clear any stuck transition states from previous session
            ClearAllTransitionStates();

            // Start emergency cleanup timer if not already running
            if (_transitionCleanupTimer == null)
            {
                _transitionCleanupTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(10) // Check every 10 seconds
                };
                _transitionCleanupTimer.Tick += (s, e) =>
                {
                    if (_framesInTransition.Count > 0)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.FrameUpdate, $"Emergency cleanup: Found {_framesInTransition.Count} frames stuck in transition state");
                        ClearAllTransitionStates();
                    }
                };
                _transitionCleanupTimer.Start();
            }

            // --- RESTORE PERSISTENT ACCORDION STACK LINKS FROM JSON ---
            FrameDataManager.RebuildDockingMap();

            foreach (dynamic frame in FrameDataManager.FrameData.ToList())
            {
                CreateFrame(frame, targetChecker);
            }

            // --- NEW: Initialize Auto-Hide Engine ---
            InitializeAutoHideTimer();

            // --- NEW: Initialize Desktop Icon Manager ---
            DesktopIconManager.Initialize();

            // --- BUG FIX: Enforce Profile Switching State ---
            // Strictly apply the incoming profile's setting. If false, we explicitly restore visibility!
            DesktopIconManager.SetDesktopIconsVisible(!SettingsManager.HideDesktopElementsOnStart);

            // Initialize Auto-Backup Timer
            BackupManager.InitializeAutoBackup();

        }


        // --- NEW: Draw Mode Integration ---

        public static void StartDrawMode()
        {
            // Called by App.xaml.cs (Context Menu) or InterCore (IPC)
            var overlay = new DrawFrameOverlay();
            overlay.Show();
        }

        public static void CreateFrameFromDraw(Rect r)
        {
            // 1. Generate Random Unique Name (Adjective + Place pattern)
            string defaultName = CoreUtilities.GenerateUniqueFrameName();

            // 2. Ask for Name (Pre-filled with the random name)
            string name = Microsoft.VisualBasic.Interaction.InputBox("Enter name for new Frame:", "Create Frame", defaultName);
            if (string.IsNullOrWhiteSpace(name)) return;

            // 2. Create Data
            // FIX: Changed "Standard" to "Data" so it functions correctly
            var frame = FrameDataManager.CreateNewFrame(name, "Data", r.X, r.Y);

            // 3. Update Dimensions & Title
            // CreateNewFrame uses defaults (230x130) and generates a random name for Data Frames.
            // We must override these with the User's Input and the Drawn Size.
            var frameDict = frame as System.Collections.Generic.IDictionary<string, object>;
            if (frameDict != null)
            {
                frameDict["Width"] = r.Width;
                frameDict["Height"] = r.Height;

                // FIX: Force set the title to what the user actually typed
                frameDict["Title"] = name;

                FrameDataManager.SaveFrameData();
            }

            // 4. Refresh UI
            // Your ReloadFrames() correctly rebuilds the entire UI state
            ReloadFrames();
        }




        private static bool LoadFrameDataFromJson()
        {
            try
            {
                string jsonContent = System.IO.File.ReadAllText(FrameDataManager.JsonFilePath);

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
                    FrameDataManager.FrameData = JsonConvert.DeserializeObject<List<dynamic>>(jsonContent);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                        $"Successfully loaded {FrameDataManager.FrameData?.Count ?? 0} frame from JSON array.");
                    return FrameDataManager.FrameData != null;
                }
                catch (JsonSerializationException)
                {
                    // If that fails, try to parse as a single frame object
                    try
                    {
                        var singleFrame = JsonConvert.DeserializeObject<dynamic>(jsonContent);
                        if (singleFrame != null)
                        {
                            FrameDataManager.FrameData = new List<dynamic> { singleFrame };
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                                "Successfully loaded single frame from JSON object.");
                            return true;
                        }
                    }
                    catch (JsonSerializationException innerEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                            $"Failed to parse JSON as single frame object: {innerEx.Message}");
                    }
                }
            }
            catch (System.IO.IOException ioEx)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"IO error reading frames.json: {ioEx.Message}");
            }
            catch (UnauthorizedAccessException accessEx)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Access denied reading frames.json: {accessEx.Message}");
            }
            catch (JsonReaderException jsonEx)
            {
                // Handle malformed JSON (syntax errors, invalid characters, etc.)
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Malformed JSON detected in frames.json: {jsonEx.Message}");

                // Optionally, create a backup of the corrupted file
                CreateCorruptedFileBackup();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Unexpected error loading frames.json: {ex.Message}");
            }

            return false;
        }

        private static void CreateCorruptedFileBackup()
        {
            try
            {
                string backupPath = FrameDataManager.JsonFilePath + ".corrupted." + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                System.IO.File.Copy(FrameDataManager.JsonFilePath, backupPath, true);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                    $"Created backup of corrupted JSON file: {backupPath}");
            }
            catch (Exception backupEx)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Failed to create backup of corrupted JSON file: {backupEx.Message}");
            }
        }




        private static void InitializeDefaultFrame()
        {
            try
            {
                // 1. Create the Standard Data Frames (For user icons)
                var dataFrame = new
                {
                    Id = Guid.NewGuid().ToString(),
                    Title = "New Frame - Drop your shortcuts here",
                    X = 20.0,
                    Y = 20.0,
                    Width = 360.0,
                    Height = 180.0,
                    ItemsType = "Data",
                    Items = new JArray(),

                    // Defaults
                    IsLocked = "false",
                    IsHidden = "false",
                    IsRolled = "false",
                    AutoRoll = "false", // --- NEW: Auto Roll ---
                    AlwaysOnTop = "false", // --- NEW ---
                    UnrolledHeight = 130.0,
                    TabsEnabled = "false",
                    CurrentTab = 0,
                    Tabs = new JArray(),

                    // Visual Defaults
                    CustomColor = (string)null,
                    FrameBorderThickness = 2
                };


                // 2. Create the "Startup Tips" Note Frame
                var noteFrame = new
                {
                    Id = Guid.NewGuid().ToString(), // Unique ID
                    Title = "Desktop Frames + Startup Tips", // Explicit Name
                    X = 20.0,   // Positioned below the data frame
                    Y = 200.0,  // Data frame ends then this frame begins
                    Width = 555.0,
                    Height = 318.0,
                    ItemsType = "Note",
                    Items = new JArray(),

                    // Visuals from your spec
                    CustomColor = (string)null, // Default
                    TextColor = "Teal",         // Teal text

                    // Note Settings
                    NoteContent = "WELCOME TO DESKTOP FRAMES +\r\n" +
                                  "---------------------------\r\n" +
                                  "• Roll Up/Down: Double-click the frame title bar.\r\n" +
                                  "• Rename: Ctrl + Click the title bar (Enter to save).\r\n" +
                                  "• Search (SpotSearch): Press Ctrl + ` (Tilde) to find any icon instantly.\r\n" +
                                  "• Options: Click the '♥' menu icon (top-left).\r\n" +
                                  "• Reorder Icons on a frame: Ctrl + Drag icon to new position.\r\n" +
                                  "• Context Menu: Right-click icons or Frames for more options.\r\n" +
                                  " \r\n" +
                                  "TIP: Ctrl + Click or Ctrl + Right-click, gives even more options.\r\n\r\n" +
                                  "Try customizing this frame! Right-click the title bar -> Customize...",

                    NoteFontSize = "Medium",
                    NoteFontFamily = "Segoe UI",
                    WordWrap = "true",
                    SpellCheck = "false",

                    // Standard Properties
                    IsHidden = "false",
                    IsLocked = "false",
                    IsRolled = "false",
                    AutoRoll = "false", // --- NEW: Auto Roll ---
                    AlwaysOnTop = "false", // --- NEW ---
                    UnrolledHeight = 318.0,
                    BoldTitleText = "false",
                    DisableTextShadow = "false",
                    IconSize = "Medium",
                    IconSpacing = 5,
                    FrameBorderThickness = 2
                };

                // 3. Combine and Save
                var frames = new List<object> { dataFrame, noteFrame };
                string defaultJson = JsonConvert.SerializeObject(frames, Formatting.Indented);

                System.IO.File.WriteAllText(FrameDataManager.JsonFilePath, defaultJson);

                // Load into memory
                FrameDataManager.FrameData = JsonConvert.DeserializeObject<List<dynamic>>(defaultJson);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                    "Initialized default configuration with Data Frame and Startup Tips");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                    $"Error initializing default frames: {ex.Message}");

                // Fallback to empty list if something explodes
                FrameDataManager.FrameData = new List<dynamic>();
            }
        }


        public static void CreateFrame(dynamic frame, TargetChecker targetChecker)
        {

            // --- FIX: Declare Title TextBox EARLY ---
            TextBox titletb = new TextBox
            {
                HorizontalContentAlignment = HorizontalAlignment.Center,
                Visibility = Visibility.Collapsed,
                Background = System.Windows.Media.Brushes.White,
                Foreground = System.Windows.Media.Brushes.Black,
                Padding = new Thickness(2)
            };

            // --- NEW: Declare Commit Action for robust saving ---
            Action CommitRename = null;
            // ---------------------------------------------------

            // Check for valid Portal Frame target folder
            if (frame.ItemsType?.ToString() == "Portal")
            {
                string targetPath = frame.Path?.ToString();
                if (string.IsNullOrEmpty(targetPath) || !System.IO.Directory.Exists(targetPath))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation, $"Skipping creation of Portal Frame '{frame.Title}' due to missing target folder: {targetPath ?? "null"}");
                    FrameDataManager.FrameData.Remove(frame);
                    FrameDataManager.SaveFrameData();
                    return;
                }
            }
            DockPanel dp = new DockPanel();
            // Get frame border color and thickness from frame data
            System.Windows.Media.Brush borderBrush = null;
            double borderThickness = 0;
            try
            {
                string borderColorName = frame.FrameBorderColor?.ToString();
                int customThickness = Convert.ToInt32(frame.FrameBorderThickness?.ToString() ?? "2");
                if (!string.IsNullOrEmpty(borderColorName))
                {
                    var borderColor = Utility.GetColorFromName(borderColorName);
                    borderBrush = new SolidColorBrush(borderColor);
                    borderThickness = customThickness; // Use custom thickness
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Applied border color '{borderColorName}' with thickness {borderThickness} to frame '{frame.Title}'");
                }
                else if (customThickness > 0)
                {
                    // Even without color, apply thickness with default color
                    borderBrush = new SolidColorBrush(System.Windows.Media.Colors.Gray);
                    borderThickness = customThickness;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Applied default border with thickness {borderThickness} to frame '{frame.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Error applying border: {ex.Message}");
                borderBrush = null;
                borderThickness = 0;
            }
            Border cborder = new Border
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 0, 0, 0)),
                // --- APPLY HIDDEN SETTING: Sharp corners if enabled, otherwise default 6px round ---
                CornerRadius = SettingsManager.FramesWithNoRoundCorners ? new CornerRadius(0) : new CornerRadius(6),
                BorderBrush = borderBrush, // Apply border color
                BorderThickness = new Thickness(borderThickness), // Apply border thickness
                Child = dp
            };
            // Add Double Click Handler
            cborder.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ClickCount == 2)
                {
                    SearchFormManager.ToggleSearch();
                    e.Handled = true;
                }
            };


            //  Add heart symbol in top-left corner
            string MenuSymbol = "♥";

            if (SettingsManager.MenuIcon == 0)
            {
                MenuSymbol = "♥";
              
            }
            else if (SettingsManager.MenuIcon == 1)
            {
                MenuSymbol = "☰";
            }
            else if (SettingsManager.MenuIcon == 2)
            {
                MenuSymbol = "≣";
            }
            else if (SettingsManager.MenuIcon == 3)
            {
                MenuSymbol = "𓃑";
            }

            TextBlock heart = new TextBlock
            {
                // Text = "♥",

                Name = "FrameMenuIcon", // New! Name
                Text = MenuSymbol,
                FontSize = 22,
                Foreground = System.Windows.Media.Brushes.White, // Match title and icon text color
                Margin = new Thickness(5, -3, 0, 0), // Position top-left, aligned with title
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                Cursor = Cursors.Hand,
                Opacity = (double)SettingsManager.MenuTintValue / 100 // 0.3 // Lower tint by default

            };
  
            _heartTextBlocks[frame] = heart;



            heart.MouseEnter += (s, e) =>
            {
                // Remove previous animation 
                heart.BeginAnimation(UIElement.OpacityProperty, null);

                heart.Opacity = 1.0;
            };

            heart.MouseLeave += (s, e) =>
            {
                double targetOpacity = (double)SettingsManager.MenuTintValue / 100;

                DoubleAnimation fadeBack = new DoubleAnimation
                {
                    From = 1.0,
                    To = targetOpacity,
                    Duration = TimeSpan.FromMilliseconds(300),
                    BeginTime = TimeSpan.FromMilliseconds(800)
                };

                heart.BeginAnimation(UIElement.OpacityProperty, fadeBack);
            };


            dp.Children.Add(heart);
            Panel.SetZIndex(heart, 100); // Ensure heart is above titleGrid to receive clicks

            // Store heart TextBlock reference for this frame
            _heartTextBlocks[frame] = heart;
            // Create and assign heart ContextMenu using centralized builder
      
            heart.ContextMenu = BuildHeartContextMenu(frame, false); // Normal menu by default
                                                                     //// Handle left-click to open heart context menu
                                                                     //heart.MouseLeftButtonDown += (s, e) =>
                                                                     //{
                                                                     // if (e.ChangedButton == MouseButton.Left && heart.ContextMenu != null)
                                                                     // {
                                                                     // heart.ContextMenu.IsOpen = true;
                                                                     // e.Handled = true;
                                                                     // }
                                                                     //};
                                                                     // Handle left-click to open heart context menu
            heart.MouseLeftButtonDown += (s, e) =>
            {


                // FIX: Directly call CommitRename logic
                if (titletb.IsVisible)
                {
                    CommitRename?.Invoke();
                    return;
                }


                if (e.ChangedButton == MouseButton.Left && heart.ContextMenu != null)
                {
                    // Check if Ctrl is pressed for extended menu
                    bool isCtrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                    // DEBUG: Log the Ctrl state
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Heart clicked - Ctrl pressed: {isCtrlPressed}");
                    // Rebuild context menu based on Ctrl state
                    heart.ContextMenu = BuildHeartContextMenu(frame, isCtrlPressed);
                    heart.ContextMenu.IsOpen = true;
                    e.Handled = true;
                }
            };
            // Add a protection symbol in top-right corner


            string LockSymbol = "🛡️";

            if (SettingsManager.LockIcon == 0)
            {
                LockSymbol = "🛡️";
            }
            else if (SettingsManager.LockIcon == 1)
            {
                LockSymbol = "🔑";
            }
            else if (SettingsManager.LockIcon == 2)
            {
                LockSymbol = "🔐";
            }
            else if (SettingsManager.LockIcon == 3)
            {
                LockSymbol = "🔒";
            }

            //MessageBox.Show(SettingsManager.LockIcon +" " + LockSymbol.ToString());


            TextBlock lockIcon = new TextBlock
           
            {

                //     Text = "🔐",
                //     Text = "🔑",
                //     Text = "🔒",
                //     Text = "🔓",
                Name = "FrameLockIcon", // New! Name
                Text = LockSymbol,//"🛡️",
                                FontSize = 14,
                Foreground = frame.IsLocked?.ToString().ToLower() == "true" ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.White,
                Margin = new Thickness(0, 3, 2, 0), // Adjusted for top-right positioning
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Cursor = Cursors.Hand,
                ToolTip = frame.IsLocked?.ToString().ToLower() == "true" ? "Frame is locked (click to unlock)" : "Frame is unlocked (click to lock)",
                Opacity = (double)SettingsManager.MenuTintValue / 100 // 0.3 // Lower tint by default
            };

  
            lockIcon.MouseEnter += (s, e) =>
            {
                // Remove previous animation
                lockIcon.BeginAnimation(UIElement.OpacityProperty, null);

                lockIcon.Opacity = 1.0;
            };

            lockIcon.MouseLeave += (s, e) =>
            {
                double targetOpacity = (double)SettingsManager.MenuTintValue / 100;

                DoubleAnimation fadeBack = new DoubleAnimation
                {
                    From = 1.0,
                    To = targetOpacity,
                    Duration = TimeSpan.FromMilliseconds(300),
                    BeginTime = TimeSpan.FromMilliseconds(800)
                };

                lockIcon.BeginAnimation(UIElement.OpacityProperty, fadeBack);
            };




            // Set initial state without saving to JSON
            UpdateLockState(lockIcon, frame, null, saveToJson: false);
            // Lock icon click handler
            lockIcon.MouseLeftButtonDown += (s, e) =>
            {
                if (e.ChangedButton == MouseButton.Left)
                {

                    // FIX: Directly call CommitRename logic
                    if (titletb.IsVisible)
                    {
                        CommitRename?.Invoke();
                        return;
                    }

                    // Get the NonActivatingWindow to find the frame by Id
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(lockIcon);
                    if (win == null)
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Could not find NonActivatingWindow for lock icon click in frame '{frame.Title}'");
                        return;
                    }
                    // Find the frame in FrameDataManager.FrameData using the window's Tag (Id)
                    string frameId = win.Tag?.ToString();
                    if (string.IsNullOrEmpty(frameId))
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Frame Id is missing for window '{win.Title}'");
                        return;
                    }
                    dynamic currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                    if (currentFrame == null)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Frame with Id '{frameId}' not found in FrameDataManager.FrameData");
                        return;
                    }
                    // Toggle the lock state
                    bool currentState = currentFrame.IsLocked?.ToString().ToLower() == "true";
                    bool newState = !currentState;
                    // Update UI and JSON on the main thread
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        UpdateLockState(lockIcon, currentFrame, newState, saveToJson: true);
                    });
                }
            };



            // Create a Grid for the titlebar - move here to ensure it is created before mouse handler
            Grid titleGrid = new Grid
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 0, 0))
            };
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(0, GridUnitType.Pixel) }); // Col 0: Spacer
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Col 1: Title
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });                      // Col 2: Filter Icon (Auto width)
            titleGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30, GridUnitType.Pixel) }); // Col 3: Lock Icon
                                                                                                                      // End of ctrl+click handler
            ContextMenu CnMnFramemanager = new ContextMenu();

            Style themedStyle = GetThemedContextMenuStyle(frame);
            if (themedStyle != null) CnMnFramemanager.Style = themedStyle;

            MenuItem miNewnoteFrame = new MenuItem { Header = Strings.MenuNewNoteFrame };

            // --- NEW: Auto Roll Menu Item ---
            MenuItem miAutoRoll = new MenuItem { Header = Strings.MenuAutoRoll, IsCheckable = true };
            miAutoRoll.IsChecked = frame.AutoRoll?.ToString().ToLower() == "true";
            miAutoRoll.Click += (s, e) =>
            {
                UpdateFrameProperty(frame, "AutoRoll", miAutoRoll.IsChecked.ToString().ToLower(), $"Set Auto Roll to {miAutoRoll.IsChecked}");

                // If turned off while currently auto-rolled, immediately unroll it
                string fId = frame.Id?.ToString();
                if (!miAutoRoll.IsChecked && !string.IsNullOrEmpty(fId) && _autoRolledFrames.Contains(fId))
                {
                    // Dynamically find the window since 'win' is out of scope here
                    var targetWin = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().FirstOrDefault(w => w.Tag?.ToString() == fId);
                    if (targetWin != null)
                    {
                        var args = new System.Windows.Input.MouseEventArgs(System.Windows.Input.Mouse.PrimaryDevice, 0) { RoutedEvent = System.Windows.Input.Mouse.MouseEnterEvent };
                        targetWin.RaiseEvent(args); // Simulate mouse enter to wake it up cleanly
                    }
                }
            };
            CnMnFramemanager.Items.Add(miAutoRoll);

            // --- NEW: Always On Top Menu Item ---
            MenuItem miAlwaysOnTop = new MenuItem { Header = Strings.MenuAlwaysOnTop, IsCheckable = true };
            miAlwaysOnTop.IsChecked = frame.AlwaysOnTop?.ToString().ToLower() == "true";
            miAlwaysOnTop.Click += (s, e) =>
            {
                UpdateFrameProperty(frame, "AlwaysOnTop", miAlwaysOnTop.IsChecked.ToString().ToLower(), $"Set Always on Top to {miAlwaysOnTop.IsChecked}");
            };
            CnMnFramemanager.Items.Add(miAlwaysOnTop);

            CnMnFramemanager.Items.Add(new Separator());
            // --------------------------------

            MenuItem miHide = new MenuItem { Header = Strings.MenuHideFrame }; // New Hide Frame item
                                                                   
            CnMnFramemanager.Items.Add(miHide); // Add Hide Frame
                                                // Add Note Frame specific context menu items if this is a Note Frame
            if (frame.ItemsType?.ToString() == "Note")
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                    $"Adding Note frame context menu items for '{frame.Title}'");
                // We'll add the TextBox reference after the window is created
                // For now, just mark that this will need Note menu items
            }
            NonActivatingWindow win = new NonActivatingWindow
            {
                ContextMenu = CnMnFramemanager,
                AllowDrop = true,
                AllowsTransparency = true,
                Background = System.Windows.Media.Brushes.Transparent,
                Title = frame.Title?.ToString() ?? "New Frame", // Handle null title
                ShowInTaskbar = false,
                WindowStyle = WindowStyle.None,
                Content = cborder,
                ResizeMode = frame.IsLocked?.ToString().ToLower() == "true" ? ResizeMode.NoResize : ResizeMode.CanResizeWithGrip,
                Topmost = frame.AlwaysOnTop?.ToString().ToLower() == "true", // --- NEW: Apply Always On Top ---
                // ResizeMode = ResizeMode.CanResizeWithGrip,
                Width = (double)frame.Width,
                Height = (double)frame.Height,
                Top = (double)frame.Y,
                Left = (double)frame.X,
                Tag = frame.Id?.ToString() ?? Guid.NewGuid().ToString() // Ensure ID exists
            };
            // Add Note frame specific context menu items after window creation
            if (frame.ItemsType?.ToString() == "Note")
            {
                // The TextBox will be created in InitContent(), so we need to add menu items after that
                // We'll modify the context menu after InitContent() is called
            }
            //Peek behind frame
            MenuItem miPeekBehind = new MenuItem { Header = Strings.MenuPeekBehind };
            CnMnFramemanager.Items.Add(miPeekBehind);
            miPeekBehind.Click += (s, e) =>
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Initiating Peek Behind for frame '{frame.Title}'");
                // Create a separate transparent window for the countdown
                var countdownWindow = new Window
                {
                    Width = 60,
                    Height = 40,
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = System.Windows.Media.Brushes.Transparent,
                    ShowInTaskbar = false,
                    Topmost = true,
                    Left = win.Left + (win.Width - 60) / 2, // Center horizontally
                    Top = win.Top + (win.Height - 40) / 2 // Center vertically
                };
                // Create countdown label
                Label countdownLabel = new Label
                {
                    Content = "10",
                    Foreground = System.Windows.Media.Brushes.White,
                    Background = System.Windows.Media.Brushes.Black,
                    Opacity = 0.7,
                    FontSize = 12,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Padding = new Thickness(4)
                };
                countdownWindow.Content = countdownLabel;
                // Show the countdown window
                countdownWindow.Show();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Created countdown window for frame '{frame.Title}' at ({countdownWindow.Left}, {countdownWindow.Top})");
                // Fade out animation for the frame
                var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.3))
                {
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
                };
                // Countdown timer
                int countdownSeconds = 10;
                var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
                timer.Tick += (timerSender, timerArgs) =>
                {
                    countdownSeconds--;
                    countdownLabel.Content = countdownSeconds.ToString();
                    if (countdownSeconds <= 0)
                    {
                        timer.Stop();
                        // Fade in animation for the frame
                        var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromSeconds(0.8))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                        };
                        win.BeginAnimation(UIElement.OpacityProperty, fadeIn);
                        // Close the countdown window
                        countdownWindow.Close();
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Peek Behind completed for frame '{frame.Title}'");
                    }
                };
                // Start fade out and timer
                win.BeginAnimation(UIElement.OpacityProperty, fadeOut);
                timer.Start();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Started Peek Behind fade-out and countdown for frame '{frame.Title}'");
            };



            // --- DYNAMIC MENU ITEMS ---

            // 1. Clear Dead Shortcuts (Initially Hidden)
            MenuItem miClearDeadShortcuts = null;
            Separator sepClearDead = null;

            if (frame.ItemsType?.ToString() == "Data")
            {
                sepClearDead = new Separator { Visibility = Visibility.Collapsed };
                CnMnFramemanager.Items.Add(sepClearDead);

                miClearDeadShortcuts = new MenuItem { Header = Strings.MenuClearDeadShortcuts, Visibility = Visibility.Collapsed };
                CnMnFramemanager.Items.Add(miClearDeadShortcuts);

                miClearDeadShortcuts.Click += (s, e) =>
                {
                    // --- BUG FIX: Fetch Live Frame ---
                    // The 'frame' closure variable becomes stale when Tabs are toggled/modified
                    string liveId = frame.Id?.ToString();
                    var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == liveId) ?? frame;

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Clear Dead Shortcuts clicked for frame '{liveFrame.Title}'");
                    try
                    {
                        int removedCount = FilePathUtilities.ClearDeadShortcutsFromFrame(liveFrame);
                        if (removedCount > 0)
                        {
                            FrameDataManager.SaveFrameData();
                            RefreshFrameUsingFormApproach(win, liveFrame);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameUpdate, $"Error in Clear Dead Shortcuts: {ex.Message}");
                    }
                };
            }

            // 2. Open Folder & Live View Toggle (Portal Only)
            if (frame.ItemsType?.ToString() == "Portal")
            {
                MenuItem miOpenFolder = new MenuItem { Header = Strings.MenuOpenFrameFolder };
                miOpenFolder.Click += (s, e) =>
                {
                    try
                    {
                        string folderPath = frame.Path?.ToString();
                        if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                            Process.Start("explorer.exe", folderPath);
                    }
                    catch { }
                };
                CnMnFramemanager.Items.Add(miOpenFolder);

                // --- LIVE VIEW MODE TOGGLE ---
                MenuItem miViewAsDetails = new MenuItem { Header = Strings.MenuViewAsDetails, IsCheckable = true };

                // --- ULTRA FIX: Re-evaluate checkmark state every single time the menu opens at runtime! ---
                CnMnFramemanager.Opened += (s, e) =>
                {
                    try
                    {
                        string fId = frame.Id?.ToString();
                        bool isDetails = false;

                        // 1. Read live running state directly from the active PortalFramemanager
                        var activePortalEntry = _portalFrames.FirstOrDefault(kvp => kvp.Key?.Id?.ToString() == fId);
                        if (activePortalEntry.Value != null)
                        {
                            isDetails = activePortalEntry.Value.IsDetailsView;
                        }
                        else
                        {
                            // 2. Fallback check from JSON if window isn't loaded yet
                            string currentViewMode = null;
                            try
                            {
                                IDictionary<string, object> viewModeDict = frame is IDictionary<string, object> vmd ? vmd : ((Newtonsoft.Json.Linq.JObject)frame).ToObject<IDictionary<string, object>>();
                                currentViewMode = viewModeDict.ContainsKey("ViewMode") ? viewModeDict["ViewMode"]?.ToString() : null;
                            }
                            catch { }
                            if (string.IsNullOrEmpty(currentViewMode)) currentViewMode = SettingsManager.DefaultPortalView;
                            isDetails = string.Equals(currentViewMode, "Details", StringComparison.OrdinalIgnoreCase);
                        }

                        miViewAsDetails.IsChecked = isDetails;
                    }
                    catch { }
                };

                miViewAsDetails.Click += (s, e) =>
                {
                    string newMode = miViewAsDetails.IsChecked ? "Details" : "Icons";
                    string clickId = frame.Id?.ToString();
                    var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == clickId) ?? frame;

                    UpdateFrameProperty(liveFrame, "ViewMode", newMode, $"Toggled Portal View Mode to {newMode}");

                    // Trigger instant live switch on the active portal manager without reloading the window
                    var portalEntry = _portalFrames.FirstOrDefault(kvp => kvp.Key?.Id?.ToString() == clickId);
                    if (portalEntry.Value != null)
                    {
                        portalEntry.Value.SetViewMode(newMode);
                    }
                };
                CnMnFramemanager.Items.Add(miViewAsDetails);
            }

            CnMnFramemanager.Items.Add(new Separator());

            // 3. Paste Item (Initially Hidden)
            MenuItem miPasteItem = new MenuItem { Header = Strings.MenuPasteItem, Visibility = Visibility.Collapsed };
            miPasteItem.Click += (s, e) => CopyPasteManager.PasteItem(frame);
            CnMnFramemanager.Items.Add(miPasteItem);




            //  CnMnFramemanager.Items.Add(new Separator());

            // --- NEW: Plugin Settings Context Menu ---
            if (frame.ItemsType?.ToString() == "Plugin")
            {
                MenuItem miPluginSettings = new MenuItem { Header = Strings.MenuPluginSettings };
                miPluginSettings.Click += (s, e) =>
                {
                    string fId = frame.Id?.ToString();
                    if (!string.IsNullOrEmpty(fId) && _activePlugins.TryGetValue(fId, out var plugin))
                    {
                        plugin.ShowSettingsWindow(win, frame);
                    }
                };
                CnMnFramemanager.Items.Add(miPluginSettings);
                CnMnFramemanager.Items.Add(new Separator());
            }
            // ---------------------------------------

            // CnMnFramemanager.Items.Add(miNewCustomize);
            MenuItem miCustomize = new MenuItem { Header = Strings.MenuCustomize };
            miCustomize.Click += (s, e) =>
            {
                try
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Opening CustomizeFrameFormManager for frame '{frame.Title}'");
                    var customizeForm = new CustomizeFrameFormManager(frame);
                    customizeForm.ShowDialog();
                    if (customizeForm.DialogResult)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"User saved changes in CustomizeFrameFormManager for frame '{frame.Title}'");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"User cancelled CustomizeFrameFormManager for frame '{frame.Title}'");
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error opening CustomizeFrameFormManager: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgOpenCustomizeFailed", ex.Message), Strings.DlgFormError);
                }
            };
            CnMnFramemanager.Items.Add(miCustomize);
            // CnMnFramemanager.Items.Add(new Separator());
            // CnMnFramemanager.Items.Add(new Separator());
            // CnMnFramemanager.Items.Add(miXT);
            // Handle both JObject and ExpandoObject access
            string isHiddenString = "false"; // Default value
            try
            {
                if (frame is JObject jFrame)
                {
                    // Existing Frame from JSON
                    isHiddenString = jFrame["IsHidden"]?.ToString().ToLower() ?? "false";
                }
                else
                {
                    // New ExpandoObject Frame
                    isHiddenString = (frame.IsHidden?.ToString() ?? "false").ToLower();
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error, $"Error reading IsHidden: {ex.Message}");
            }
            bool isHidden = isHiddenString == "true";
            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, $"Frame '{frame.Title}' IsHidden state: {isHidden}");
            // Adjust the frame position to ensure it fits within screen bounds
            AdjustFramePositionToScreen(win);
            win.Loaded += (s, e) =>
            {
                UpdateLockState(lockIcon, frame, null, saveToJson: false);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Applied lock state for frame '{frame.Title}' on load: IsLocked={frame.IsLocked?.ToString().ToLower()}");
                // Apply IsRolled state
                bool isRolled = frame.IsRolled?.ToString().ToLower() == "true";
                double targetHeight = 28; // Default for rolled-up state  //rolled height
                if (!isRolled)
                {
                    double unrolledHeight = (double)frame.Height; // Default to frame.Height
                    if (frame.UnrolledHeight != null)
                    {
                        if (double.TryParse(frame.UnrolledHeight.ToString(), out double parsedHeight))
                        {
                            if (parsedHeight > 0)
                            {
                                unrolledHeight = parsedHeight;
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"UnrolledHeight {parsedHeight} is invalid (non-positive) for frame '{frame.Title}', using Height={unrolledHeight}");
                            }
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Failed to parse UnrolledHeight '{frame.UnrolledHeight}' for frame '{frame.Title}', using Height={unrolledHeight}");
                        }
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"UnrolledHeight is null for frame '{frame.Title}', using Height={unrolledHeight}");
                    }
                    targetHeight = unrolledHeight;
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation, $"Applied rolled-down state for frame '{frame.Title}' on load: Height={targetHeight}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation, $"Applied rolled-up state for frame '{frame.Title}' on load: Height={targetHeight}");
                }
                win.Height = targetHeight;
                // Apply WrapPanel visibility
                var border = win.Content as Border;
                if (border != null)
                {
                    var dockPanel = border.Child as DockPanel;
                    if (dockPanel != null)
                    {
                        var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                        if (scrollViewer != null)
                        {
                            var wpcont = scrollViewer.Content as WrapPanel;
                            if (wpcont != null)
                            {
                                wpcont.Visibility = isRolled ? Visibility.Collapsed : Visibility.Visible;
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Set initial WrapPanel visibility to {(isRolled ? "Collapsed" : "Visible")} for frame '{frame.Title}'");

                                // --- MULTI-MONITOR INVISIBLE ICON FIX ---
                                if (!isRolled)
                                {
                                    wpcont.UpdateLayout();
                                    wpcont.InvalidateVisual();
                                    wpcont.Opacity = 0.99;
                                    wpcont.Dispatcher.BeginInvoke(new Action(() => { wpcont.Opacity = 1.0; }), System.Windows.Threading.DispatcherPriority.Render);
                                }
                            }
                        }
                    }
                }
                if (isHidden)
                {
                    win.Visibility = Visibility.Hidden;
                    TrayManager.AddHiddenFrame(win);
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation, $"Hid frame '{frame.Title}' after loading at startup");
                }
            };
            // Step 5
            if (isHidden)
            {
                TrayManager.AddHiddenFrame(win);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Added frame '{frame.Title}' to hidden list at startup");
            }
            // Hide click
            miHide.Click += (s, e) =>
            {
                UpdateFrameProperty(frame, "IsHidden", "true", $"Hid frame '{frame.Title}'");
                TrayManager.AddHiddenFrame(win);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Triggered Hide frame for '{frame.Title}'");
            };
            // Handle manual resize to update both Height and UnrolledHeight
            win.SizeChanged += (s, e) =>
            {
                // Get current frame reference by ID to avoid stale references
                string frameId = win.Tag?.ToString();
                if (string.IsNullOrEmpty(frameId))
                {

                    return;
                }

                // --- ACCORDION STACK CASCADE HOOK ---
                // Triggers on manual grip resizing AND smooth roll-up/roll-down animations
                if (win.IsLoaded && Math.Abs(e.NewSize.Height - e.PreviousSize.Height) > 0.1)
                {
                    double deltaY = e.NewSize.Height - e.PreviousSize.Height;
                    SnapManager.CascadeStack(frameId, deltaY);
                }
                // ------------------------------------

                // Skip updates if this frame is currently in a rollup/rolldown transition
                // --- NEW: Also skip if it is Auto-Rolled (so we don't save the rolled-up height) ---
                if (_framesInTransition.Contains(frameId) || _autoRolledFrames.Contains(frameId))
                {
                    return;
                }
                // Find the current frame in FrameDataManager.FrameData using ID
                // Find the current frame in FrameDataManager.FrameData using ID
                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame != null)
                {

                    double newHeight = e.NewSize.Height;
                    double newWidth = e.NewSize.Width;
                    double oldHeight = Convert.ToDouble(currentFrame.Height?.ToString() ?? "0");
                    double oldUnrolledHeight = Convert.ToDouble(currentFrame.UnrolledHeight?.ToString() ?? "0");
                    bool isRolled = currentFrame.IsRolled?.ToString().ToLower() == "true";

                    // Update Width and Height with the actual new values
                    currentFrame.Width = newWidth;
                    currentFrame.Height = newHeight;

                    // Handle UnrolledHeight update
                    if (!isRolled)
                    {

                        if (Math.Abs(newHeight - 28) > 5) // Only if height is significantly different from rolled-up height   //rolled height
                        {
                            double heightDifference = Math.Abs(newHeight - oldUnrolledHeight);
                            // DebugLog("LOGIC", frameId, $"Height difference from old UnrolledHeight: {heightDifference:F1}");
                            currentFrame.UnrolledHeight = newHeight;
                            // DebugLog("UPDATE", frameId, $"UPDATED UnrolledHeight from {oldUnrolledHeight:F1} to {newHeight:F1}");
                        }
                        else
                        {

                        }
                    }
                    else
                    {

                    }
                    // Save to JSON
                    //   MessageBox.Show("Debug: SizeChanged handler called. Saving frame data.");
                    FrameDataManager.SaveFrameData();
                    var verifyFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                    double verifyHeight = Convert.ToDouble(verifyFrame.Height?.ToString() ?? "0");
                    double verifyUnrolled = Convert.ToDouble(verifyFrame.UnrolledHeight?.ToString() ?? "0");

                }
                else
                {

                }
            };
            win.KeyDown += (sender, e) =>
            {
                if (IconDragDropManager.IsDragging && e.Key == Key.Escape)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Escape pressed during drag, cancelling operation");
                    IconDragDropManager.CancelDrag();
                    e.Handled = true;
                }


            };

            // --- NEW: Auto-Hide Interaction Hook ---
            win.PreviewMouseDown += (s, e) => { if (SettingsManager.AutoResetHideTimer) ResetAutoHideTimer(); };
            win.PreviewMouseWheel += (s, e) => { if (SettingsManager.AutoResetHideTimer) ResetAutoHideTimer(); };

            // --- NEW: Smart CTRL Overlay Engine (Event-Driven, Zero Polling) ---
            bool isOverlayActive = false;

            void ToggleOverlays(bool turnOn)
            {
                if (isOverlayActive == turnOn) return; // Prevent redundant drawing
                isOverlayActive = turnOn;

                var borderNode = win.Content as Border;
                var dockNode = borderNode?.Child as DockPanel;
                var scrollNode = dockNode?.Children.OfType<ScrollViewer>().FirstOrDefault();
                var wpNode = scrollNode?.Content as WrapPanel;

                if (wpNode == null) return;

                foreach (StackPanel itemPanel in wpNode.Children.OfType<StackPanel>())
                {
                    dynamic tagData = itemPanel.Tag;
                    if (tagData == null) continue;

                    string fp = tagData.GetType().GetProperty("FilePath")?.GetValue(tagData)?.ToString();

                    if (turnOn)
                    {
                        if (fp != null && fp.StartsWith("INTERNAL_BLANK_"))
                        {
                            // Spacer Overlay: Faint White/Gray
                            itemPanel.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 255, 255, 255));
                        }
                        else
                        {
                            // Real Icon Overlay: Faint Blue
                            itemPanel.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 100, 150, 255));
                        }
                    }
                    else
                    {
                        // Revert to transparent (but still hit-testable)
                        itemPanel.Background = System.Windows.Media.Brushes.Transparent;
                    }
                }
            }

            // --- BUG FIX: Native Keyboard Hooks (Replaces the 50ms Timer) ---
            if (frame.ItemsType?.ToString() != "Portal")
            {
                win.PreviewKeyDown += (s, e) =>
                {
                    if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
                    {
                        if (win.IsMouseOver) ToggleOverlays(true);
                    }
                };

                win.PreviewKeyUp += (s, e) =>
                {
                    if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl)
                    {
                        ToggleOverlays(false);
                    }
                };
            }
            // ----------------------------------------------------------------

            // --- NEW: Setup Auto Roll Timer ---
            string frameIdForTimer = frame.Id?.ToString() ?? win.Tag?.ToString();
            if (!string.IsNullOrEmpty(frameIdForTimer))
            {
                // Pull from options if it exists, otherwise strictly default to 2 seconds (hidden setting)
                int autoRollDelay = 2000;
                try { if (SettingsManager.AutoRollTime > 0) autoRollDelay = SettingsManager.AutoRollTime * 1000; } catch { }

                var autoRollTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(autoRollDelay) };
                _autoRollTimers[frameIdForTimer] = autoRollTimer;

                autoRollTimer.Tick += (s, e) =>
                {
                    autoRollTimer.Stop();

                    // --- BUG FIX: Prevent rolling up if the mouse is currently resting on the frame ---
                    if (win.IsMouseOver)
                    {
                        // We just abort. The MouseLeave event will restart the timer when they finally move away.
                        return;
                    }

                    var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameIdForTimer);
                    if (currentFrame == null) return;
                    bool isManuallyRolled = currentFrame.IsRolled?.ToString().ToLower() == "true";

                    // Don't auto-roll if it's already rolled manually, or caught in transition
                    if (isManuallyRolled || _framesInTransition.Contains(frameIdForTimer)) return;

                    _autoRolledFrames.Add(frameIdForTimer);
                    _framesInTransition.Add(frameIdForTimer);

                    double currentHeight = win.Height;
                    double targetHeight = 28;

                    var heightAnimation = new DoubleAnimation(currentHeight, targetHeight, TimeSpan.FromSeconds(0.3))
                    {
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                    };
                    heightAnimation.Completed += (animSender, animArgs) =>
                    {
                        var borderNode = win.Content as Border;
                        if (borderNode?.Child is DockPanel dockNode)
                        {
                            var scrollNode = dockNode.Children.OfType<ScrollViewer>().FirstOrDefault();
                            if (scrollNode?.Content is WrapPanel wpNode) wpNode.Visibility = Visibility.Collapsed;
                            else if (dockNode.Children.OfType<TextBox>().FirstOrDefault() is TextBox tbNode) tbNode.Visibility = Visibility.Collapsed;

                            // --- NEW: Plugin Pause & Hide ---
                            if (_activePlugins.TryGetValue(frameIdForTimer, out var plugin))
                            {
                                plugin.Pause();
                                foreach (UIElement child in dockNode.Children)
                                {
                                    if (child is FrameworkElement fe && fe.Name == "FrameMenuIcon") continue;
                                    if (DockPanel.GetDock(child) == Dock.Top) continue;
                                    child.Visibility = Visibility.Collapsed;
                                }
                            }
                        }
                        _framesInTransition.Remove(frameIdForTimer);
                        win.BeginAnimation(Window.HeightProperty, null);
                        win.Height = targetHeight;
                    };
                    win.BeginAnimation(Window.HeightProperty, heightAnimation);
                };

                // --- BUG FIX: Start timer immediately on startup if Auto Roll is enabled ---
                var currentFrameSetup = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameIdForTimer);
                if (currentFrameSetup != null)
                {
                    bool isAutoRollEnabled = currentFrameSetup.AutoRoll?.ToString().ToLower() == "true";
                    bool isManuallyRolled = currentFrameSetup.IsRolled?.ToString().ToLower() == "true";

                    if (isAutoRollEnabled && !isManuallyRolled)
                    {
                        autoRollTimer.Start();
                    }
                }
            }
            // ----------------------------------

            win.MouseEnter += (s, e) =>
            {
                if (SettingsManager.AutoResetHideTimer) ResetAutoHideTimer();

                // --- NEW: Auto Roll Wake Up ---
                if (!string.IsNullOrEmpty(frameIdForTimer))
                {
                    if (_autoRollTimers.ContainsKey(frameIdForTimer)) _autoRollTimers[frameIdForTimer].Stop();

                    if (_autoRolledFrames.Contains(frameIdForTimer))
                    {
                        _autoRolledFrames.Remove(frameIdForTimer);
                        _framesInTransition.Add(frameIdForTimer);

                        var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameIdForTimer);
                        double unrolledHeight = currentFrame != null ? Convert.ToDouble(currentFrame.UnrolledHeight?.ToString() ?? "130") : 130;

                        var heightAnimation = new DoubleAnimation(win.Height, unrolledHeight, TimeSpan.FromSeconds(0.3))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };
                        heightAnimation.Completed += (animSender, animArgs) =>
                        {
                            var borderNode = win.Content as Border;
                            if (borderNode?.Child is DockPanel dockNode)
                            {
                                var scrollNode = dockNode.Children.OfType<ScrollViewer>().FirstOrDefault();
                                if (scrollNode?.Content is WrapPanel wpNode)
                                {
                                    wpNode.Visibility = Visibility.Visible;

                                    // --- MULTI-MONITOR INVISIBLE ICON FIX ---
                                    wpNode.UpdateLayout();
                                    wpNode.InvalidateVisual();
                                    wpNode.Opacity = 0.99;
                                    wpNode.Dispatcher.BeginInvoke(new Action(() => { wpNode.Opacity = 1.0; }), System.Windows.Threading.DispatcherPriority.Render);
                                }
                                else if (dockNode.Children.OfType<TextBox>().FirstOrDefault() is TextBox tbNode)
                                {
                                    tbNode.Visibility = Visibility.Visible;
                                }

                                // --- NEW: Plugin Resume & Show ---
                                if (_activePlugins.TryGetValue(frameIdForTimer, out var plugin))
                                {
                                    foreach (UIElement child in dockNode.Children)
                                    {
                                        if (child is FrameworkElement fe && fe.Name == "FrameMenuIcon") continue;
                                        if (DockPanel.GetDock(child) == Dock.Top) continue;
                                        child.Visibility = Visibility.Visible;
                                    }
                                    plugin.Resume();
                                }
                            }
                            _framesInTransition.Remove(frameIdForTimer);
                            win.BeginAnimation(Window.HeightProperty, null);
                            win.Height = unrolledHeight;
                        };
                        win.BeginAnimation(Window.HeightProperty, heightAnimation);
                    }
                }
                // ------------------------------

                // --- BUG FIX: Event-Driven Overlay Trigger (Zero Polling) ---
                if (frame.ItemsType?.ToString() != "Portal")
                {
                    bool isCtrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                    if (isCtrl) ToggleOverlays(true);
                }
            };

            win.MouseLeave += (s, e) =>
            {
                if (frame.ItemsType?.ToString() != "Portal")
                {
                    ToggleOverlays(false); // Instantly hide if mouse leaves the frame
                }

                // --- NEW: Auto Roll Trigger ---
                if (!string.IsNullOrEmpty(frameIdForTimer))
                {
                    var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameIdForTimer);
                    if (currentFrame != null)
                    {
                        bool isAutoRollEnabled = currentFrame.AutoRoll?.ToString().ToLower() == "true";
                        bool isManuallyRolled = currentFrame.IsRolled?.ToString().ToLower() == "true";

                        // If Auto Roll is ON, it's not manually rolled up, and not already auto-rolled
                        if (isAutoRollEnabled && !isManuallyRolled && !_autoRolledFrames.Contains(frameIdForTimer))
                        {
                            if (_autoRollTimers.TryGetValue(frameIdForTimer, out var timer))
                            {
                                timer.Stop();
                                timer.Start();
                            }
                        }
                    }
                }
                // ------------------------------
            };



            // Make window focusable for key events during drag
            win.Focusable = true;
            win.Show();



            // --- FIX START ---

            // 1. Install the Message Hook (Enforces the 90% size limit)
            win.SourceInitialized += (s, e) =>
            {
                var source = HwndSource.FromHwnd(new WindowInteropHelper(win).Handle);
                source?.AddHook(WndProc);
            };
            // (Safety: if already initialized)
            if (PresentationSource.FromVisual(win) is HwndSource existingSource)
            {
                existingSource.AddHook(WndProc);
            }

            // 2. Reactive State Fix (The "Trick")
            // If Snap Assist sets state to Maximized, we immediately force it back to Normal.
            // This ensures the Resize Grip (bottom-right) remains visible and functional.
            win.StateChanged += (sender, args) =>
            {
                if (win.WindowState == WindowState.Maximized)
                {
                    // We use Dispatcher to let the OS finish its "snap" calculation first, 
                    // then we immediately override the mode back to Normal.
                    win.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        win.WindowState = WindowState.Normal;

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            "Snap Assist intercepted: Reverted Maximized state to Normal to preserve controls.");
                    }));
                }
            };


            // --- FIX END ---





            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
        $"Frame '{frame.Title}' successfully created and displayed");
  



            bool isBoldTitle = false;
            try
            {
                isBoldTitle = frame.BoldTitleText?.ToString().ToLower() == "true";
            }
            catch { /* Safe fallback */ }
            // Get title text color or use default white
            System.Windows.Media.Brush titleTextBrush = System.Windows.Media.Brushes.White; // Default
            try
            {
                string titleColorName = frame.TitleTextColor?.ToString();
                if (!string.IsNullOrEmpty(titleColorName))
                {
                    var titleColor = Utility.GetColorFromName(titleColorName);
                    titleTextBrush = new SolidColorBrush(titleColor);
                }
            }
            catch
            {
                titleTextBrush = System.Windows.Media.Brushes.White; // Fallback
            }
            // Get title text size from frame data
            double titleFontSize = 12; // Default Medium size
            try
            {
                string titleSizeValue = frame.TitleTextSize?.ToString() ?? "Medium";
                switch (titleSizeValue)
                {
                    case "Small":
                        titleFontSize = 10;
                        break;
                    case "Large":
                        titleFontSize = 16;
                        break;
                    default: // Medium
                        titleFontSize = 12;
                        break;
                }
            }
            catch
            {
                titleFontSize = 12; // Fallback to Medium
            }
            Label titlelabel = new Label
            {
                Content = frame.Title.ToString(),
                Foreground = titleTextBrush, // Changed from hardcoded White
                HorizontalContentAlignment = HorizontalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Cursor = Cursors.SizeAll,
                FontWeight = isBoldTitle ? FontWeights.Bold : FontWeights.Normal,
                FontSize = titleFontSize // Apply custom title text size
            };
            Grid.SetColumn(titlelabel, 1);
            titleGrid.Children.Add(titlelabel);

            // --- DYNAMIC STATE UPDATES (ON MENU OPEN) ---
            MenuItem miExportAllToDesktop = null;
            MenuItem miAddSpacerMenu = null; // REPLACED: Spacer Submenu
            Separator sepAfterSpacer = null; // NEW: Layout Separator
            MenuItem miNameAfterPath = null; // New: For Portal Renaming
            Separator sepNameAfterPath = null; // New: Separator for layout

            CnMnFramemanager.Opened += (contextSender, contextArgs) =>
            {
                // --- LIVE THEME FIX: Main Frame Context Menu ---
                try
                {
                    string id = frame.Id?.ToString();
                    var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == id) ?? frame;
                    Style liveStyle = GetThemedContextMenuStyle(liveFrame);
                    if (liveStyle != null) CnMnFramemanager.Style = liveStyle;
                }
                catch { }

                // A. Update Paste Visibility
                bool hasCopiedItem = CopyPasteManager.HasCopiedItem();
                miPasteItem.Visibility = hasCopiedItem ? Visibility.Visible : Visibility.Collapsed;

                // B. Update Clear Dead Shortcuts Visibility
                if (miClearDeadShortcuts != null)
                {
                    bool hasDead = HasDeadShortcuts(frame);
                    miClearDeadShortcuts.Visibility = hasDead ? Visibility.Visible : Visibility.Collapsed;
                    if (sepClearDead != null) sepClearDead.Visibility = hasDead ? Visibility.Visible : Visibility.Collapsed;
                }

                // --- CTRL + RIGHT CLICK LOGIC ---
                bool isCtrlPressed = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                bool isDataFrame = frame.ItemsType?.ToString() == "Data";
                bool isPortalFrame = frame.ItemsType?.ToString() == "Portal";



                // Clean up previous dynamic items to prevent duplicates
                if (miExportAllToDesktop != null && CnMnFramemanager.Items.Contains(miExportAllToDesktop))
                    CnMnFramemanager.Items.Remove(miExportAllToDesktop);

                if (miAddSpacerMenu != null && CnMnFramemanager.Items.Contains(miAddSpacerMenu))
                    CnMnFramemanager.Items.Remove(miAddSpacerMenu);

                if (sepAfterSpacer != null && CnMnFramemanager.Items.Contains(sepAfterSpacer))
                    CnMnFramemanager.Items.Remove(sepAfterSpacer);

                if (miNameAfterPath != null && CnMnFramemanager.Items.Contains(miNameAfterPath))


                    CnMnFramemanager.Items.Remove(miNameAfterPath);

                if (sepNameAfterPath != null && CnMnFramemanager.Items.Contains(sepNameAfterPath))
                    CnMnFramemanager.Items.Remove(sepNameAfterPath);

                // C. Export All & Spacer Menu (Data Frame + Ctrl)
                if (isCtrlPressed && isDataFrame)
                {
                    miExportAllToDesktop = new MenuItem { Header = Strings.MenuExportAllIcons };
                    miExportAllToDesktop.Click += (s, e) => ExportAllIconsToDesktop(frame);

                    miAddSpacerMenu = new MenuItem { Header = Strings.MenuAddSpacer };

                    MenuItem miSpacerBlank = new MenuItem { Header = Strings.MenuSpacerBlank };
                    MenuItem miSpacerDot = new MenuItem { Header = Strings.MenuSpacerDot };

                    // Centralized Spacer Creation Logic
                    Action<string> CreateSpacer = (prefix) =>
                    {
                        try
                        {
                            string id = frame.Id?.ToString();
                            var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == id);
                            if (liveFrame == null) return;

                            JArray items = null;
                            bool tabsEnabled = liveFrame.TabsEnabled?.ToString().ToLower() == "true";
                            if (tabsEnabled)
                            {
                                var tabs = liveFrame.Tabs as JArray ?? new JArray();
                                int currentTab = Convert.ToInt32(liveFrame.CurrentTab?.ToString() ?? "0");
                                if (currentTab >= 0 && currentTab < tabs.Count)
                                {
                                    var activeTab = tabs[currentTab] as JObject;
                                    items = activeTab?["Items"] as JArray ?? new JArray();
                                }
                            }
                            if (items == null) items = liveFrame.Items as JArray ?? new JArray();

                            int nextDisplayOrder = items.Count > 0 ? items.Max(i => i["DisplayOrder"]?.Value<int>() ?? 0) + 1 : 0;

                            dynamic newItem = new System.Dynamic.ExpandoObject();
                            IDictionary<string, object> newItemDict = newItem;
                            newItemDict["Filename"] = prefix + "_" + Guid.NewGuid().ToString();
                            newItemDict["DisplayName"] = "";
                            newItemDict["IsFolder"] = false;
                            newItemDict["IsLink"] = false;
                            newItemDict["IsNetwork"] = false;
                            newItemDict["AlwaysRunAsAdmin"] = false;
                            newItemDict["DisplayOrder"] = nextDisplayOrder;

                            items.Add(JObject.FromObject(newItem));

                            if (tabsEnabled && Convert.ToInt32(liveFrame.CurrentTab?.ToString() ?? "0") == 0)
                                TabManager.SynchronizeTab0Content(id, "tab0", "add");
                            else if (!tabsEnabled)
                                TabManager.SynchronizeTab0Content(id, "main", "add");

                            FrameDataManager.SaveFrameData();
                            RefreshFrameUsingFormApproach(win, liveFrame);
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error adding spacer: {ex.Message}");
                        }
                    };

                    miSpacerBlank.Click += (s, e) => CreateSpacer("INTERNAL_BLANK_EMPTY");
                    miSpacerDot.Click += (s, e) => CreateSpacer("INTERNAL_BLANK_DOT");

                    miAddSpacerMenu.Items.Add(miSpacerBlank);
                    miAddSpacerMenu.Items.Add(miSpacerDot);

                    sepAfterSpacer = new Separator();

                    // Insert before Customize (safe lookup)
                    int insertIndex = CnMnFramemanager.Items.Count - 1;
                    var customizeItem = CnMnFramemanager.Items.OfType<MenuItem>()
                        .FirstOrDefault(m => m.Header.ToString() == Strings.MenuCustomize);
                    if (customizeItem != null) insertIndex = CnMnFramemanager.Items.IndexOf(customizeItem);

                    CnMnFramemanager.Items.Insert(insertIndex, miExportAllToDesktop);
                    CnMnFramemanager.Items.Insert(insertIndex + 1, miAddSpacerMenu);
                    CnMnFramemanager.Items.Insert(insertIndex + 2, sepAfterSpacer);
                }
                // D. Name After Target (Portal Frame + Ctrl)
                if (isCtrlPressed && isPortalFrame)
                {
                    miNameAfterPath = new MenuItem { Header = Strings.MenuNameAfterPath };
                    miNameAfterPath.Click += (s, e) =>
                    {
                        // Get Base Path (Not navigation path)
                        string targetPath = frame.Path?.ToString();

                        if (!string.IsNullOrEmpty(targetPath))
                        {
                            // 1. Update Live Object (Global List)
                            string id = frame.Id?.ToString();
                            var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == id);

                            if (liveFrame != null)
                            {
                                if (liveFrame is JObject jFrame) jFrame["Title"] = targetPath;
                                else liveFrame.Title = targetPath;

                                frame.Title = targetPath; // Update local reference
                            }

                            // 2. Update UI
                            titlelabel.Content = targetPath;
                            win.Title = targetPath;
                            titletb.Text = targetPath; // Update hidden textbox too

                            // 3. Save
                            FrameDataManager.SaveFrameData();
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Renamed portal frame to target: {targetPath}");
                        }
                    };

                    sepNameAfterPath = new Separator();

                    // Insert before Customize
                    int insertIndex = CnMnFramemanager.Items.Count - 1;
                    var customizeItem = CnMnFramemanager.Items.OfType<MenuItem>()
                        .FirstOrDefault(m => m.Header.ToString() == Strings.MenuCustomize);

                    if (customizeItem != null)
                    {
                        insertIndex = CnMnFramemanager.Items.IndexOf(customizeItem);
                        // Insert Order: Name -> Separator -> Customize
                        CnMnFramemanager.Items.Insert(insertIndex, miNameAfterPath);
                        CnMnFramemanager.Items.Insert(insertIndex + 1, sepNameAfterPath);
                    }
                    else
                    {
                        CnMnFramemanager.Items.Add(miNameAfterPath);
                        CnMnFramemanager.Items.Add(sepNameAfterPath);
                    }
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Updated frame menu - Ctrl:{isCtrlPressed}");
            };


            Grid.SetColumn(titletb, 1);
            titleGrid.Children.Add(titletb);

            // --- FILTER UI START ---
            // Only add Filter UI for Portal Frames
            Grid filterBar = null;
            if (frame.ItemsType?.ToString() == "Portal")
            {
                // 1. The Filter Icon (Funnel)
                // 1. The Filter Icon (Funnel)
                string initialFilter = GetSafeProperty(frame, "FilterString");
                bool hasInitialFilter = !string.IsNullOrWhiteSpace(initialFilter);

                TextBlock filterIcon = new TextBlock
                {
                    Name = "FrameFilterIcon", // New! Name
                    Text = "❖",
                    FontSize = 18,
                    Foreground = hasInitialFilter ? System.Windows.Media.Brushes.Orange : System.Windows.Media.Brushes.White,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Cursor = Cursors.Hand,
                    Margin = new Thickness(0, 0, 5, 0),
                    ToolTip = Strings.TooltipFilterFiles,
                    Opacity = hasInitialFilter ? 1.0 : ((double)SettingsManager.MenuTintValue / 100)
                };

                filterIcon.MouseEnter += (s, e) =>
                {
                    filterIcon.BeginAnimation(UIElement.OpacityProperty, null);
                    filterIcon.Opacity = 1.0;
                };

                filterIcon.MouseLeave += (s, e) =>
                {
                    // Skip fade out if a filter is actively applied
                    if (filterIcon.Foreground != System.Windows.Media.Brushes.Orange)
                    {
                        double targetOpacity = (double)SettingsManager.MenuTintValue / 100;
                        DoubleAnimation fadeBack = new DoubleAnimation
                        {
                            From = 1.0,
                            To = targetOpacity,
                            Duration = TimeSpan.FromMilliseconds(300),
                            BeginTime = TimeSpan.FromMilliseconds(800)
                        };
                        filterIcon.BeginAnimation(UIElement.OpacityProperty, fadeBack);
                    }
                };

                Grid.SetColumn(filterIcon, 2);
                titleGrid.Children.Add(filterIcon);




                // 2. The Filter Bar (Stable Version)
                filterBar = new Grid
                {
                    Height = 26,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(60, 0, 0, 0)),
                    Visibility = Visibility.Collapsed,
                    Margin = new Thickness(0, 0, 0, 2)
                };

                filterBar.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                filterBar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                ComboBox cmbFilter = new ComboBox
                {
                    IsEditable = true,
                    Height = 24,
                    StaysOpenOnEdit = true,
                    VerticalContentAlignment = VerticalAlignment.Center,
                    Padding = new Thickness(5, 0, 5, 0),
                    Text = GetSafeProperty(frame, "FilterString")
                };

                // Helper: Only Repopulate when needed (prevents freeze loop)
                Action repopulateDropdown = () =>
                {
                    cmbFilter.Items.Clear();

                    // A. History
                    try
                    {
                        string currentId = frame.Id?.ToString();
                        var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == currentId);
                        if (liveFrame != null)
                        {
                            Newtonsoft.Json.Linq.JArray history = null;
                            if (liveFrame is Newtonsoft.Json.Linq.JObject jFrame)
                                history = jFrame["FilterHistory"] as Newtonsoft.Json.Linq.JArray;
                            else
                                history = liveFrame.GetType().GetProperty("FilterHistory")?.GetValue(liveFrame) as Newtonsoft.Json.Linq.JArray;

                            if (history != null)
                            {
                                foreach (var item in history)
                                {
                                    cmbFilter.Items.Add(new ComboBoxItem
                                    {
                                        Content = item.ToString(),
                                        Tag = item.ToString(),
                                        FontWeight = FontWeights.Bold
                                    });
                                }
                            }
                        }
                    }
                    catch { }

                    // B. Separator (if needed)
                    if (cmbFilter.Items.Count > 0) cmbFilter.Items.Add(new Separator());

                    // C. Standard Presets (From static dictionary)
                    foreach (var kvp in _standardPresets)
                    {
                        cmbFilter.Items.Add(new ComboBoxItem { Content = kvp.Key, Tag = kvp.Value });
                    }
                };

                // FIX 1: Only populate on OPEN, never during selection/typing
                cmbFilter.DropDownOpened += (s, e) => repopulateDropdown();

                // Clear Button
                Button btnClearFilter = new Button
                {
                    Content = "✕",
                    Background = System.Windows.Media.Brushes.Transparent,
                    Foreground = System.Windows.Media.Brushes.Gray,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    Width = 20,
                    ToolTip = Strings.TooltipClearFilter
                };

                filterBar.Children.Add(cmbFilter);
                filterBar.Children.Add(btnClearFilter);
                Grid.SetColumn(btnClearFilter, 1);

                // --- LOGIC SETUP ---

                // Helper to execute/save
                Action<string> commitFilter = (text) =>
                {
                    bool hasFilter = !string.IsNullOrWhiteSpace(text);
                    filterIcon.Foreground = hasFilter ? System.Windows.Media.Brushes.Orange : System.Windows.Media.Brushes.White;

                    filterIcon.BeginAnimation(UIElement.OpacityProperty, null);
                    filterIcon.Opacity = hasFilter ? 1.0 : ((double)SettingsManager.MenuTintValue / 100);

                    if (_portalFrames.ContainsKey(frame))
                        _portalFrames[frame].ApplyFilter(text);

                    UpdateFrameProperty(frame, "FilterString", text, "Updated filter");

                    // Update history (logic now handles ignoring presets)
                    UpdateFilterHistory(frame, text);
                };

                // Typing Timer
                System.Windows.Threading.DispatcherTimer debounceTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(500)
                };
                debounceTimer.Tick += (s, e) =>
                {
                    debounceTimer.Stop();
                    // Visual update only while typing
                    if (_portalFrames.ContainsKey(frame)) _portalFrames[frame].ApplyFilter(cmbFilter.Text);

                    bool hasFilter = !string.IsNullOrWhiteSpace(cmbFilter.Text);
                    filterIcon.Foreground = hasFilter ? System.Windows.Media.Brushes.Orange : System.Windows.Media.Brushes.White;

                    filterIcon.BeginAnimation(UIElement.OpacityProperty, null);
                    filterIcon.Opacity = hasFilter ? 1.0 : ((double)SettingsManager.MenuTintValue / 100);
                };

                // 1. STYLE FIX (Internal TextBox)
                cmbFilter.Loaded += (s, e) =>
                {
                    var textBox = (TextBox)cmbFilter.Template.FindName("PART_EditableTextBox", cmbFilter);
                    if (textBox != null)
                    {
                        textBox.Background = System.Windows.Media.Brushes.Transparent;
                        textBox.Foreground = System.Windows.Media.Brushes.White;
                        textBox.CaretBrush = System.Windows.Media.Brushes.White;
                        textBox.BorderThickness = new Thickness(0);
                    }
                };

                // 2. TYPING LOGIC
                cmbFilter.KeyUp += (s, e) =>
                {
                    if (e.Key == Key.Up || e.Key == Key.Down) return;

                    if (e.Key == Key.Enter)
                    {
                        commitFilter(cmbFilter.Text); // Save history on Enter
                        cmbFilter.IsDropDownOpen = false;
                        filterBar.Visibility = Visibility.Collapsed;
                        win.EndKeyboardInteractiveEdit();
                        e.Handled = true;
                        return;
                    }
                    else if (e.Key == Key.Escape)
                    {
                        filterBar.Visibility = Visibility.Collapsed;
                        win.EndKeyboardInteractiveEdit();
                        e.Handled = true;
                        return;
                    }

                    debounceTimer.Stop();
                    debounceTimer.Start();
                };

                // 3. SELECTION LOGIC (FIXED)
                // We do NOT repopulate here. We just apply the value.
                cmbFilter.SelectionChanged += (s, e) =>
                {
                    if (cmbFilter.SelectedItem is ComboBoxItem item && item.Tag != null)
                    {
                        string selectedFilter = item.Tag.ToString();

                        // Break the event loop by running this later
                        System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            cmbFilter.IsDropDownOpen = false;

                            // Important: Don't set SelectedItem = null here, it causes flickering.
                            // Just force the text and run logic.
                            cmbFilter.Text = selectedFilter;

                            debounceTimer.Stop();
                            commitFilter(selectedFilter);
                        }));
                    }
                };

                // 4. CLEAR LOGIC
                btnClearFilter.Click += (s, e) =>
                {
                    cmbFilter.Text = "";
                    if (_portalFrames.ContainsKey(frame)) _portalFrames[frame].ApplyFilter("");
                    UpdateFrameProperty(frame, "FilterString", "", "Cleared filter");
                    filterIcon.Foreground = System.Windows.Media.Brushes.White;
                    filterIcon.BeginAnimation(UIElement.OpacityProperty, null);
                    filterIcon.Opacity = (double)SettingsManager.MenuTintValue / 100;
                    filterBar.Visibility = Visibility.Collapsed;
                    win.EndKeyboardInteractiveEdit();
                };

                // 5. TOGGLE BAR
                filterIcon.MouseLeftButtonDown += (s, e) =>
                {
                    if (filterBar.Visibility == Visibility.Visible)
                    {
                        if (!string.IsNullOrWhiteSpace(cmbFilter.Text)) commitFilter(cmbFilter.Text); // Save on close
                        filterBar.Visibility = Visibility.Collapsed;
                        win.EndKeyboardInteractiveEdit();
                    }
                    else
                    {
                        filterBar.Visibility = Visibility.Visible;
                        win.BeginKeyboardInteractiveEdit(cmbFilter);
                    }
                    e.Handled = true;
                };

                // Initial Apply
                if (!string.IsNullOrEmpty(cmbFilter.Text))
                {
                    filterIcon.Foreground = System.Windows.Media.Brushes.Orange;
                    filterIcon.Opacity = 1.0; // Ensure it starts fully visible if active

                    win.Loaded += (s, e) =>
                    {
                        if (_portalFrames.ContainsKey(frame))
                            _portalFrames[frame].ApplyFilter(cmbFilter.Text);
                    };
                }
            }
            // --- FILTER UI END ---


            // --- STEP 4 FIX: Centralized Rename Logic ---

            // 1. Define the Logic ONCE
            CommitRename = () =>
            {
                // Only run if actually editing
                if (titletb.Visibility != Visibility.Visible) return;

                string originalTitle = frame.Title.ToString();
                string newTitle = titletb.Text;
                string finalTitle = InterCore.ProcessTitleChange(frame, newTitle, originalTitle);

                // Update Data
                string id = frame.Id?.ToString();
                var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == id);

                if (liveFrame != null)
                {
                    if (liveFrame is Newtonsoft.Json.Linq.JObject jFrame)
                        jFrame["Title"] = finalTitle;
                    else
                        liveFrame.Title = finalTitle;
                    frame.Title = finalTitle;
                }

                // Update UI
                titlelabel.Content = finalTitle;
                win.Title = finalTitle;
                titletb.Visibility = Visibility.Collapsed;
                titlelabel.Visibility = Visibility.Visible;

                // Save
                FrameDataManager.SaveFrameData();
                win.EndKeyboardInteractiveEdit();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Rename committed: {finalTitle}");
            };

            // 2. Wire up Events to use the central logic
            titletb.KeyDown += (sender, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    CommitRename(); // Call shared logic
                    e.Handled = true;
                }
                else if (e.Key == Key.Escape)
                {
                    // Cancel Logic
                    titletb.Text = frame.Title.ToString();
                    titletb.Visibility = Visibility.Collapsed;
                    titlelabel.Visibility = Visibility.Visible;
                    win.EndKeyboardInteractiveEdit();
                    e.Handled = true;
                }
            };

            titletb.LostFocus += (sender, e) =>
            {
                CommitRename(); // Call shared logic
            };



            //// --- STEP 4 START: Configure and Add Events ---
            //titletb.HorizontalContentAlignment = HorizontalAlignment.Center;
            //titletb.Visibility = Visibility.Collapsed;

            //// 1. Handle Keys (Enter = Save, Escape = Cancel)
            //titletb.KeyDown += (sender, e) =>
            //{
            //    if (e.Key == Key.Enter)
            //    {
            //        string originalTitle = frame.Title.ToString();
            //        string newTitle = titletb.Text;
            //        string finalTitle = InterCore.ProcessTitleChange(frame, newTitle, originalTitle);

            //        // Update LIVE Data
            //        string id = frame.Id?.ToString();
            //        var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == id);

            //        if (liveFrame != null)
            //        {
            //            if (liveFrame is Newtonsoft.Json.Linq.JObject jFrame)
            //                jFrame["Title"] = finalTitle;
            //            else
            //                liveFrame.Title = finalTitle;
            //            frame.Title = finalTitle;
            //        }

            //        titlelabel.Content = finalTitle;
            //        win.Title = finalTitle;
            //        titletb.Visibility = Visibility.Collapsed;
            //        titlelabel.Visibility = Visibility.Visible;

            //        FrameDataManager.SaveFrameData();
            //        win.EndKeyboardInteractiveEdit();
            //    }
            //    else if (e.Key == Key.Escape)
            //    {
            //        // ESCAPE: Cancel and Revert
            //        titletb.Text = frame.Title.ToString();
            //        titletb.Visibility = Visibility.Collapsed;
            //        titlelabel.Visibility = Visibility.Visible;
            //        win.EndKeyboardInteractiveEdit();
            //        e.Handled = true;

            //        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Rename cancelled via Escape");
            //        }
            //    }
            //    ;

            //// 2. Handle Focus Loss (Save when clicking away)
            //titletb.LostFocus += (sender, e) =>
            //{
            //    // If invisible, we already handled it (e.g. via Escape)
            //    if (titletb.Visibility != Visibility.Visible) return;

            //    string originalTitle = frame.Title.ToString();
            //    string newTitle = titletb.Text;
            //    string finalTitle = InterCore.ProcessTitleChange(frame, newTitle, originalTitle);

            //    string id = frame.Id?.ToString();
            //    var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == id);

            //    if (liveFrame != null)
            //    {
            //        if (liveFrame is Newtonsoft.Json.Linq.JObject jFrame)
            //            jFrame["Title"] = finalTitle;
            //        else
            //            liveFrame.Title = finalTitle;
            //        frame.Title = finalTitle;
            //    }

            //    titlelabel.Content = finalTitle;
            //    win.Title = finalTitle;
            //    titletb.Visibility = Visibility.Collapsed;
            //    titlelabel.Visibility = Visibility.Visible;

            //    FrameDataManager.SaveFrameData();
            //    win.EndKeyboardInteractiveEdit();
            //    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Rename saved via LostFocus: {finalTitle}");
            //};
            //// --- STEP 4 END ---


            // Move lockIcon to the Grid
            Grid.SetColumn(lockIcon, 3); // Moved to Column 3
            Grid.SetRow(lockIcon, 0);
            titleGrid.Children.Add(lockIcon);
            // Add the titleGrid to the DockPanel
            DockPanel.SetDock(titleGrid, Dock.Top);
            dp.Children.Add(titleGrid);
            if (filterBar != null)
            {
                DockPanel.SetDock(filterBar, Dock.Top);
                dp.Children.Add(filterBar);
            }




            // TABS FEATURE: Add TabStrip
            // We delegate to the shared method to ensure consistent "Scrollable" layout on startup.
            TabManager.RefreshTabStripUI(win, frame);

           string frameId = win.Tag?.ToString();
            if (string.IsNullOrEmpty(frameId))
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Frame Id is missing for window '{win.Title}'");
                return;
            }
            dynamic currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
            if (currentFrame == null)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Frame with Id '{frameId}' not found in FrameDataManager.FrameData");
                return;
            }
            bool isLocked = currentFrame.IsLocked?.ToString().ToLower() == "true";
            titlelabel.MouseDown += (sender, e) =>
            {
                // FIX: Directly call CommitRename logic
                if (titletb.IsVisible)
                {
                    CommitRename?.Invoke();
                    return;
                }

                if (e.ClickCount == 2)
                {
                    // Roll-up/roll-down logic (swapped from Ctrl+Click)
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(titlelabel);
                    string frameId = win?.Tag?.ToString();
                    // DebugLog("IMMEDIATE", frameId ?? "UNKNOWN", $"Ctrl+Click FIRST LINE - win.Height:{win?.Height ?? -1:F1}");
                    if (string.IsNullOrEmpty(frameId) || win == null)
                    {
                        // DebugLog("ERROR", frameId ?? "UNKNOWN", "Missing window or frameId in Ctrl+Click");
                        return;
                    }
              
                    dynamic currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                    if (currentFrame == null)
                    {
                        // DebugLog("ERROR", frameId, "Frame not found in FrameDataManager.FrameData for Ctrl+Click");
                        return;
                    }
                    bool isRolled = currentFrame.IsRolled?.ToString().ToLower() == "true";
                    // Get frame data height (always accurate from SizeChanged handler)
                    double frameHeight = Convert.ToDouble(currentFrame.Height?.ToString() ?? "130");
                    double windowHeight = win.Height;
                    // DebugLog("SYNC_CHECK", frameId, $"Before sync - frameHeight:{frameHeight:F1} | WindowHeight:{windowHeight:F1} | IsRolled:{isRolled}");
                    if (!isRolled)
                    {
                        // ROLLUP: Use frame data height (always current)
                        // DebugLog("ACTION", frameId, "Starting ROLLUP");
                        _framesInTransition.Add(frameId);
                        // DebugLog("TRANSITION", frameId, "Added to transition state");
                        // FINAL FIX: Use frame data height which is always accurate from SizeChanged handler
                        double currentHeight = frameHeight;

                        // --- BUG FIX: The "Stuck at 28" Safeguard ---
                        // If currentHeight is suspiciously small (like the title bar height), it means the 
                        // WPF layout data is stale. We reject it and use the last known good UnrolledHeight.
                        if (currentHeight <= 35)
                        {
                            currentHeight = Convert.ToDouble(currentFrame.UnrolledHeight?.ToString() ?? "130");
                            if (currentHeight <= 35) currentHeight = 130; // Ultimate fallback
                        }

                        // DebugLog("ROLLUP_HEIGHT_SOURCE", frameId, $"Using frame.Height:{frameHeight:F1} (win.Height was stale:{win.Height:F1})");
                        // Save current frame height as UnrolledHeight
                        IDictionary<string, object> frameDict = currentFrame as IDictionary<string, object> ?? ((JObject)currentFrame).ToObject<IDictionary<string, object>>();
                        frameDict["UnrolledHeight"] = currentHeight;
                        frameDict["IsRolled"] = "true";
                        int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                        if (frameIndex >= 0)
                        {
                            FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                        }
                        FrameDataManager.SaveFrameData();
                        // DebugLog("SAVE", frameId, $"Saved ROLLUP state | UnrolledHeight:{currentHeight:F1} | IsRolled:true");
                        // Roll up animation - starts from current height
                        double targetHeight = 28;   //rolled height
                        var heightAnimation = new DoubleAnimation(currentHeight, targetHeight, TimeSpan.FromSeconds(0.3))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };
                        heightAnimation.Completed += (animSender, animArgs) =>
                        {
                            // DebugLog("ANIMATION", frameId, "ROLLUP animation completed");
                            // Update WrapPanel visibility
                            var border = win.Content as Border;
                            if (border != null)
                            {
                                var dockPanel = border.Child as DockPanel;
                                if (dockPanel != null)
                                {
                                    var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                                    if (scrollViewer != null)
                                    {
                                        var wpcont = scrollViewer.Content as WrapPanel;
                                        if (wpcont != null)
                                        {
                                            wpcont.Visibility = Visibility.Collapsed;
                                            // DebugLog("UI", frameId, "Set WrapPanel visibility to Collapsed");
                                        }
                                    }

                                    // --- NEW: Plugin Pause & Hide ---
                                    if (_activePlugins.TryGetValue(frameId, out var plugin))
                                    {
                                        plugin.Pause();
                                        foreach (UIElement child in dockPanel.Children)
                                        {
                                            if (child is FrameworkElement fe && fe.Name == "FrameMenuIcon") continue;
                                            if (DockPanel.GetDock(child) == Dock.Top) continue;
                                            child.Visibility = Visibility.Collapsed;
                                        }
                                    }
                                }
                            }
                            _framesInTransition.Remove(frameId);

                            // FIX: Release animation lock so manual resizes track correctly later
                            win.BeginAnimation(Window.HeightProperty, null);
                            win.Height = targetHeight;
                        };
                        win.BeginAnimation(Window.HeightProperty, heightAnimation);
                        // DebugLog("ANIMATION", frameId, $"Started ROLLUP animation from {currentHeight:F1} to height {targetHeight:F1}");
                    }
                    else
                    {
                        // ROLLDOWN: Roll down to UnrolledHeight
                        double unrolledHeight = Convert.ToDouble(currentFrame.UnrolledHeight?.ToString() ?? "130");
                        // // DebugLog("ACTION", frameId, $"Starting ROLLDOWN to {unrolledHeight:F1}");
                        _framesInTransition.Add(frameId);
                        // DebugLog("TRANSITION", frameId, "Added to transition state");
                        IDictionary<string, object> frameDict = currentFrame as IDictionary<string, object> ?? ((JObject)currentFrame).ToObject<IDictionary<string, object>>();
                        frameDict["IsRolled"] = "false";
                        int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                        if (frameIndex >= 0)
                        {
                            FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                        }
                        FrameDataManager.SaveFrameData();
                        // DebugLog("SAVE", frameId, $"Saved ROLLDOWN state | IsRolled:false | TargetHeight:{unrolledHeight:F1}");
                        var heightAnimation = new DoubleAnimation(win.Height, unrolledHeight, TimeSpan.FromSeconds(0.3))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                        };
                        heightAnimation.Completed += (animSender, animArgs) =>
                        {
                            // DebugLog("ANIMATION", frameId, "ROLLDOWN animation completed");
                            // Update WrapPanel visibility
                            var border = win.Content as Border;
                            if (border != null)
                            {
                                var dockPanel = border.Child as DockPanel;
                                if (dockPanel != null)
                                {
                                    var scrollViewer = dockPanel.Children.OfType<ScrollViewer>().FirstOrDefault();
                                    if (scrollViewer != null)
                                    {
                                        var wpcont = scrollViewer.Content as WrapPanel;
                                        if (wpcont != null)
                                        {
                                            wpcont.Visibility = Visibility.Visible;

                                            // --- MULTI-MONITOR INVISIBLE ICON FIX ---
                                            wpcont.UpdateLayout();
                                            wpcont.InvalidateVisual();
                                            wpcont.Opacity = 0.99;
                                            wpcont.Dispatcher.BeginInvoke(new Action(() => { wpcont.Opacity = 1.0; }), System.Windows.Threading.DispatcherPriority.Render);
                                            // DebugLog("UI", frameId, "Set WrapPanel visibility to Visible");
                                        }
                                    }

                                    // --- NEW: Plugin Resume & Show ---
                                    if (_activePlugins.TryGetValue(frameId, out var plugin))
                                    {
                                        foreach (UIElement child in dockPanel.Children)
                                        {
                                            if (child is FrameworkElement fe && fe.Name == "FrameMenuIcon") continue;
                                            if (DockPanel.GetDock(child) == Dock.Top) continue;
                                            child.Visibility = Visibility.Visible;
                                        }
                                        plugin.Resume();
                                    }
                                }
                            }
                            _framesInTransition.Remove(frameId);
                            // FIX: Release animation lock and force real height so SizeChanged updates the JSON properly
                            win.BeginAnimation(Window.HeightProperty, null);
                            win.Height = unrolledHeight;
                        };
                        win.BeginAnimation(Window.HeightProperty, heightAnimation);
                        // DebugLog("ANIMATION", frameId, $"Started ROLLDOWN animation to height {unrolledHeight:F1}");
                    }
                    e.Handled = true;
                }
                else if (e.LeftButton == MouseButtonState.Pressed)
                {
                    if ((Keyboard.Modifiers & ModifierKeys.Control) != 0)
                    {
                        // Rename frame (swapped from double-click)
                        titletb.Text = titlelabel.Content.ToString();
                        titlelabel.Visibility = Visibility.Collapsed;
                        titletb.Visibility = Visibility.Visible;
                        win.BeginKeyboardInteractiveEdit(titletb);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Focus set to title textbox for frame: {frame.Title}");
                        e.Handled = true;
                    }
                    else
                    {
                        string frameId = win.Tag?.ToString();
                        if (string.IsNullOrEmpty(frameId))
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Frame Id is missing for window '{win.Title}' during MouseDown");
                            return;
                        }
                        dynamic currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                        if (currentFrame == null)
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Frame with Id '{frameId}' not found in FrameDataManager.FrameData during MouseDown");
                            return;
                        }
                        bool isLocked = currentFrame.IsLocked?.ToString().ToLower() == "true";
                        if (!isLocked)
                        {
                            SnapManager.StartDrag(win);
                            try
                            {
                                win.DragMove();
                            }
                            finally
                            {
                                SnapManager.EndDrag(win);
                            }
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Dragging frame '{currentFrame.Title}'");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"DragMove blocked for locked frame '{currentFrame.Title}'");
                        }
                    }
                }
            };



            //titletb.KeyDown += (sender, e) =>
            //{
            //    if (e.Key == Key.Enter)
            //    {
            //        string originalTitle = frame.Title.ToString();
            //        string newTitle = titletb.Text;

            //        // Process through InterCore for special triggers
            //        string finalTitle = InterCore.ProcessTitleChange(frame, newTitle, originalTitle);

            //        // --- FIX START: Update LIVE Data ---
            //        // Get the ID to find the fresh object in the global list
            //        string id = frame.Id?.ToString();
            //        var liveFrame = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == id);

            //        if (liveFrame != null)
            //        {
            //            // Update the live object in the list
            //            // Handle both JObject (JSON) and ExpandoObject (New frame)
            //            if (liveFrame is Newtonsoft.Json.Linq.JObject jFrame)
            //                jFrame["Title"] = finalTitle;
            //            else
            //                liveFrame.Title = finalTitle;

            //            // Also update the local reference just in case
            //            frame.Title = finalTitle;
            //        }
            //        // --- FIX END ---

            //        // Update UI
            //        titlelabel.Content = finalTitle;
            //        win.Title = finalTitle;
            //        titletb.Visibility = Visibility.Collapsed;
            //        titlelabel.Visibility = Visibility.Visible;

            //        // Save the global list which now contains the updated title
            //        FrameDataManager.SaveFrameData();

            //        win.ShowActivated = false;
            //        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Exited edit mode via Enter, final title for frame: {finalTitle}");
            //        win.Focus();
            //    }
            //    else if (e.Key == Key.Escape)
            //    {
            //        // FIX: ESCAPE Logic (Cancel)
            //        titletb.Text = frame.Title.ToString(); // Revert text
            //        titletb.Visibility = Visibility.Collapsed;
            //        titlelabel.Visibility = Visibility.Visible;

            //        Keyboard.ClearFocus(); // Drop focus
            //        win.ShowActivated = false;
            //        e.Handled = true;

            //        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Rename cancelled via Escape");
            //    }
            //};

            //// 2. Handle Focus Loss (Auto-Save)
            //titletb.LostFocus += (sender, e) =>
            //{
            //    // Don't save if we are cancelling (Escape key handles UI)
            //    if (titletb.Visibility != Visibility.Visible) return;

            //    string originalTitle = frame.Title.ToString();
            //    string newTitle = titletb.Text;
            //    string finalTitle = InterCore.ProcessTitleChange(frame, newTitle, originalTitle);

            //    string id = frame.Id?.ToString();
            //    var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == id);

            //    if (liveFrame != null)
            //    {
            //        IDictionary<string, object> frameDict = liveFrame as IDictionary<string, object> ??
            //            ((JObject)liveFrame).ToObject<IDictionary<string, object>>();
            //        frameDict["Title"] = finalTitle;

            //        int index = FrameDataManager.FrameData.IndexOf(liveFrame);
            //        if (index >= 0) FrameDataManager.FrameData[index] = JObject.FromObject(frameDict);

            //        frame.Title = finalTitle;
            //    }

            //    titlelabel.Content = finalTitle;
            //    win.Title = finalTitle;
            //    titletb.Visibility = Visibility.Collapsed;
            //    titlelabel.Visibility = Visibility.Visible;

            //    FrameDataManager.SaveFrameData();

            //    win.ShowActivated = false;
            //    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Rename saved via LostFocus: {finalTitle}");
            //};






            // WrapPanel wpcont = new WrapPanel();

            // --- FIX START ---
            WrapPanel wpcont = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                ItemWidth = double.NaN,
                ItemHeight = double.NaN,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0)
            };

            // CRITICAL: Tag the panel with the frame ID so AddIcon can find settings
            // even before the panel is attached to the window.
            wpcont.Tag = frame.Id?.ToString();
            // --- FIX END ---

            ScrollViewer wpcontscr = new ScrollViewer
            {
                Content = wpcont,
                VerticalScrollBarVisibility = SettingsManager.DisableFrameScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto,
                
            };

            // --- PERFORMANCE TWEAK: Hardware Acceleration & Safe UI Culling ---
            IconManager.OptimizeFramePanel(wpcont, wpcontscr);

            // Προσθήκη watermark για Portal frames
            if (frame.ItemsType?.ToString() == "Portal")
            {
                if (_options.ShowBackgroundImageOnPortalFences ?? true)
                {
                    double opacity = (SettingsManager.PortalBackgroundOpacity / 100.0);
                    wpcontscr.Background = new ImageBrush
                    {
                        ImageSource = new BitmapImage(new Uri("pack://application:,,,/Resources/portal.png")),
                        // Opacity = 0.2,
                        Opacity = opacity,
                        Stretch = Stretch.UniformToFill
                    };
                }
            }
            dp.Children.Add(wpcontscr);

            void InitContent()
            {
                // 1. Handle Note frames - they don't use WrapPanel
                if (frame.ItemsType?.ToString() == "Note")
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Creating Note frame content for '{frame.Title}'");
                    dp.Children.Remove(wpcontscr); // Remove ScrollViewer
                    TextBox noteTextBox = NoteFramemanager.CreateNoteContent(frame, dp);

                    bool isNoteRolled = frame.IsRolled?.ToString().ToLower() == "true";
                    noteTextBox.Visibility = isNoteRolled ? Visibility.Collapsed : Visibility.Visible;
                    return;
                }

                // --- NEW: Handle Plugin frames ---
                if (frame.ItemsType?.ToString() == "Plugin")
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation, $"Creating Plugin frame content for '{frame.Title}'");
                    dp.Children.Remove(wpcontscr); // Remove default Icon ScrollViewer

                    string pluginId = frame.PluginId?.ToString();
                    var pluginInstance = PluginManager.CreatePluginInstance(pluginId);

                    bool isPluginRolled = frame.IsRolled?.ToString().ToLower() == "true";

                    if (pluginInstance != null)
                    {
                        var visualElement = pluginInstance.CreateVisualElement();
                        if (visualElement != null)
                        {
                            visualElement.Visibility = isPluginRolled ? Visibility.Collapsed : Visibility.Visible;
                            dp.Children.Add(visualElement); // Inject plugin UI into frame

                            // --- CLEAN OPACITY FIX: Animate the Window, not the Element ---
                            // By fading the Window, the image fades to the desktop cleanly
                            // without the frame's background color bleeding through it.
                            if (SettingsManager.ApplyTintToIcons)
                            {
                              



                                
                                double targetOpacity = Math.Max(0.2, SettingsManager.TintValue / 100.0);
                                win.Opacity = targetOpacity;

                                win.MouseEnter += (s, ev) =>
                                {
                                    win.BeginAnimation(Window.OpacityProperty, null);
                                    win.Opacity = 1.0;
                                };

                                win.MouseLeave += (s, ev) =>
                                {
                                    System.Windows.Media.Animation.DoubleAnimation fadeBack = new System.Windows.Media.Animation.DoubleAnimation
                                    {
                                        From = 1.0,
                                        To = targetOpacity,
                                        Duration = TimeSpan.FromMilliseconds(300),
                                        BeginTime = TimeSpan.FromMilliseconds(500)
                                    };
                                    win.BeginAnimation(Window.OpacityProperty, fadeBack);
                                };
                            }
                            // --------------------------------------------------------------

                            // Extract settings safely into Dictionary
                            Dictionary<string, object> pluginSettings = new Dictionary<string, object>();
                            try
                            {
                                if (frame is JObject jFrame && jFrame["PluginSettings"] is JObject settingsObj)
                                {
                                    pluginSettings = settingsObj.ToObject<Dictionary<string, object>>();
                                }
                                else if (frame.PluginSettings != null)
                                {
                                    var settingsDict = frame.PluginSettings as IDictionary<string, object>;
                                    if (settingsDict != null) pluginSettings = new Dictionary<string, object>(settingsDict);
                                }
                            }
                            catch (Exception ex)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation, $"Failed to parse PluginSettings: {ex.Message}");
                            }

                            // Track active plugin instance and initialize it
                            string fId = frame.Id?.ToString();
                            if (!string.IsNullOrEmpty(fId))
                            {
                                _activePlugins[fId] = pluginInstance;
                            }

                            pluginInstance.Initialize(visualElement, pluginSettings);
                            if (isPluginRolled) pluginInstance.Pause();
                        }
                    }
                    else
                    {
                        // Safe fallback UI if plugin is missing/unregistered
                        TextBlock errorText = new TextBlock
                        {
                            Text = $"Plugin '{pluginId}' missing or failed to load.",
                            Foreground = System.Windows.Media.Brushes.Red,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            TextWrapping = TextWrapping.Wrap
                        };
                        errorText.Visibility = isPluginRolled ? Visibility.Collapsed : Visibility.Visible;
                        dp.Children.Add(errorText);
                    }
                    return;
                }
                // ---------------------------------

                // 2. Handle Data/Portal frames










                bool isRolled = frame.IsRolled?.ToString().ToLower() == "true";
                wpcont.Visibility = isRolled ? Visibility.Collapsed : Visibility.Visible;
                wpcont.Children.Clear();

                // 2a. Data frames (The complex logic)
                if (frame.ItemsType?.ToString() == "Data")
                {
                    JArray items = null;
                    bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";

                    // Load Items (Tabs vs Main)
                    if (tabsEnabled)
                    {
                        try
                        {
                            var tabs = frame.Tabs as JArray ?? new JArray();
                            int currentTab = Convert.ToInt32(frame.CurrentTab?.ToString() ?? "0");
                            if (currentTab >= 0 && currentTab < tabs.Count)
                            {
                                var activeTab = tabs[currentTab] as JObject;
                                items = activeTab?["Items"] as JArray ?? new JArray();
                            }
                            else
                            {
                                items = frame.Items as JArray ?? new JArray();
                            }
                        }
                        catch { items = frame.Items as JArray ?? new JArray(); }
                    }
                    else
                    {
                        items = frame.Items as JArray ?? new JArray();
                    }

                    if (items != null)
                    {
                        // Sort
                        var sortedItems = items.OfType<JObject>()
                            .OrderBy(item => item["DisplayOrder"]?.Type == JTokenType.Integer ? item["DisplayOrder"].Value<int>() : 0)
                            .ToList();

                        foreach (dynamic icon in sortedItems)
                        {
                            // FIX: Pass 'frame' context for customization
                            AddIcon(icon, wpcont, frame);

                            StackPanel sp = wpcont.Children[wpcont.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                // --- 1. DEFINE VARIABLES (Scope Fix) ---
                                // We extract these immediately so they are available for ALL blocks below
                                IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict
                                    ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();

                                string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";

                                // CRITICAL: Define these here to avoid CS0103 errors
                                bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
                                bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];

                                // --- FIX: MOVE NETWORK CHECK UP ---
                                bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
                                if (!isNetwork && !string.IsNullOrEmpty(filePath)) isNetwork = IsNetworkPath(filePath);

                                //// --- 2. Extract Arguments ---
                                //string arguments = null;
                                //if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
                                //{
                                //    // FIX: Do NOT use WshShell on network links during startup to prevent 30-second hangs!
                                //    if (!isNetwork)
                                //    {
                                //        try
                                //        {
                                //            WshShell shell = new WshShell();
                                //            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                                //            arguments = shortcut.Arguments;
                                //        }
                                //        catch { }
                                //    }
                                //}

                                // --- 2. Extract Arguments ---
                                // Completely removed WshShell from main thread to prevent local .lnk files 
                                // from auto-resolving dead network targets during startup.
                                string arguments = iconDict.ContainsKey("Arguments") ? (string)iconDict["Arguments"] : null;

                                // --- 3. Attach Click Event ---
                                ClickEventAdder(sp, filePath, isFolder, arguments);

                                // [TARGET VALIDATION ENGINE - INIT CONTENT]
                                bool allowNetworkChecking = _options.CheckNetworkPaths ?? false;
                                bool isShortcutFile = filePath != null && filePath.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase);

                                if (!isLink && (!isNetwork || allowNetworkChecking))
                                {
                                    string resolvedTarget = filePath;
                                    bool excludeFromCheck = false;

                                    if (isShortcutFile)
                                    {
                                        resolvedTarget = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath) ?? filePath;

                                        if (resolvedTarget.StartsWith(@"\\") ||
                                            resolvedTarget.StartsWith("steam://", StringComparison.OrdinalIgnoreCase) ||
                                            Utility.IsStoreAppShortcut(filePath))
                                        {
                                            excludeFromCheck = true;
                                        }
                                    }

                                    if (!excludeFromCheck)
                                    {
                                        targetChecker.AddCheckAction(filePath, () => UpdateIcon(sp, filePath, isFolder, resolvedTarget), isFolder);
                                        if (isNetwork && allowNetworkChecking) LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Monitoring network path {filePath}");
                                    }
                                    else
                                    {
                                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Excluded shortcut {filePath} from background checking to protect UI thread/avoid exclusions.");
                                    }
                                }
                                else if (isLink || (isNetwork && !allowNetworkChecking))
                                {
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Safety Exclusion: Skipped background check for {filePath}");
                                }

                                // --- 5. Attach Context Menu ---
                                AttachIconContextMenu(sp, icon, frame, win);
                            }
                        }
                    }
                }
                // 2b. Portal frames
                else if (frame.ItemsType?.ToString() == "Portal")
                {
                    try
                    {
                        var portalManager = new PortalFramemanager(frame, wpcont);
                        _portalFrames[frame] = portalManager;

                        // --- BULLETPROOF FIX: Decoupled Sort Cycle on CTRL + Click ---
                        // We use the Window's PreviewMouseDown to fire BEFORE any UI is destroyed.
                        // We trace the Visual Tree to ensure the click hit empty space, not an icon.
                        win.PreviewMouseLeftButtonDown += (s, e) =>
                        {
                            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                            {
                                // --- STEP 6 FIX: Abort CTRL+Click background sorting completely if in Details/List View ---
                                if (portalManager.IsDetailsView) return;

                                DependencyObject parent = e.OriginalSource as DependencyObject;
                                bool isBackground = true;

                                while (parent != null && parent != win)
                                {
                                    // Hit an icon? Abort and let the folder navigate!
                                    if (parent is StackPanel sp && sp.Tag != null) { isBackground = false; break; }
                                    // Hit a UI Control (Button, Scrollbar, Text, Title Label)? Abort!
                                    if (parent is Button || parent is System.Windows.Controls.Primitives.ScrollBar || parent is TextBox || parent is Label || parent is TextBlock) { isBackground = false; break; }
                                    // Hit the Nav Bar? Abort!
                                    if (parent is Grid g && g.Tag?.ToString() == "PORTAL_NAV_BAR") { isBackground = false; break; }

                                    // --- FIX: Safely step out of non-Visual text elements (like 'Run') ---
                                    if (parent is FrameworkContentElement fce)
                                    {
                                        parent = fce.Parent;
                                    }
                                    else if (parent is Visual || parent is System.Windows.Media.Media3D.Visual3D)
                                    {
                                        parent = VisualTreeHelper.GetParent(parent);
                                    }
                                    else
                                    {
                                        parent = null;
                                    }
                                }

                                if (isBackground)
                                {
                                    string newMode = portalManager.CycleSortMode();
                                    ShowPortalToast(win, $"Sorted by: {newMode}");
                                    e.Handled = true; // Stop event from triggering anything else
                                }
                            }
                        };
                    }
                    catch (Exception ex)
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgPortalInitFailed", ex.Message), Strings.DlgError);
                        FrameDataManager.FrameData.Remove(frame);
                        FrameDataManager.SaveFrameData();
                        win.Close();
                    }
                }
            }


            win.Drop += (sender, e) =>
            {
                e.Handled = true;
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] droppedFiles = null;
                    try
                    {
                        droppedFiles = (string[])e.Data.GetData(DataFormats.FileDrop);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation,
                            $"Drop handler received {droppedFiles?.Length ?? 0} files");
                    }
                    catch (Exception dataEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                            $"Error getting drop data: {dataEx.Message}");
                        return;
                    }
                    if (droppedFiles == null) return;

                    int portalCopiedCount = 0; // --- NEW: Track successful portal copies ---

                    foreach (string droppedFile in droppedFiles)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                            $"Processing dropped item: '{droppedFile}'");
                        try
                        {
                            // Path Validation (Existence / Folder Check)
                            bool fileExists = false;
                            bool directoryExists = false;
                            bool isFolderFlag = false;

                            try
                            {
                                FileAttributes attrs = System.IO.File.GetAttributes(droppedFile);
                                isFolderFlag = attrs.HasFlag(FileAttributes.Directory);
                                directoryExists = isFolderFlag;
                                fileExists = !isFolderFlag;
                            }
                            catch (Exception pathEx)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation, $"Path validation error: {pathEx.Message}");
                                continue;
                            }

                            if (!fileExists && !directoryExists)
                            {
                                MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgInvalidFileOrFolder", droppedFile), Strings.DlgError);
                                continue;
                            }

                     


                            // --- DATA FRAME LOGIC ---
                            if (frame.ItemsType?.ToString() == "Data")
                            {
                                if (!System.IO.Directory.Exists("Shortcuts")) System.IO.Directory.CreateDirectory("Shortcuts");
                                string baseShortcutName = System.IO.Path.Combine("Shortcuts", System.IO.Path.GetFileName(droppedFile));
                                string shortcutName = baseShortcutName;
                                int counter = 1;

                                bool isDroppedShortcut = System.IO.Path.GetExtension(droppedFile).ToLower() == ".lnk";
                                bool isDroppedUrlFile = System.IO.Path.GetExtension(droppedFile).ToLower() == ".url";

                                // FIX: Trust the extension. If it ends in .url, it IS a link.
                                // This bypasses content checks that fail on files with custom headers.
                                bool isWebLink = isDroppedUrlFile || CoreUtilities.IsWebLinkShortcut(droppedFile);

                                string targetPath;
                                bool isFolder = false;
                                string webUrl = null;

                                if (isWebLink)
                                {
                                    // Try to extract clean URL, but if it fails (weird header), fallback to file path
                                    try { webUrl = CoreUtilities.ExtractWebUrlFromFile(droppedFile); } catch { }

                                    targetPath = !string.IsNullOrEmpty(webUrl) ? webUrl : droppedFile;
                                    isFolder = false;
                                }
                                else
                                {
                                    if (isDroppedShortcut)
                                    {
                                        targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(droppedFile);
                                        if (string.IsNullOrEmpty(targetPath))
                                        {
                                            isFolder = false;
                                        }
                                        else
                                        {
                                            isFolder = System.IO.Directory.Exists(targetPath);
                                            if (!isFolder && string.IsNullOrEmpty(System.IO.Path.GetExtension(targetPath)))
                                            {
                                                isFolder = System.IO.Directory.Exists(targetPath);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        targetPath = droppedFile;
                                        isFolder = System.IO.Directory.Exists(targetPath);
                                    }
                                }

                                if (!isDroppedShortcut && !isDroppedUrlFile)
                                {
                                    // CASE A: Creating new shortcut from raw file/folder
                                    shortcutName = baseShortcutName + ".lnk";
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{System.IO.Path.GetFileNameWithoutExtension(droppedFile)} ({counter++}).lnk");
                                    }

                                    try
                                    {
                                        WshShell shell = new WshShell();
                                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutName);
                                        shortcut.TargetPath = droppedFile;
                                        if (isFolder) shortcut.WorkingDirectory = droppedFile;
                                        shortcut.Save();
                                    }
                                    catch { continue; }
                                }
                                else
                                {
                                    // CASE B: Copying existing shortcut (LNK or URL)
                                    // FIX: Determine correct extension based on type
                                    // If it's a Web Link, MUST remain .url. If it's a Shortcut, MUST remain .lnk.
                                    string ext = isWebLink ? ".url" : ".lnk";

                                    // Ensure base name has correct extension
                                    if (!baseShortcutName.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                                        baseShortcutName = System.IO.Path.ChangeExtension(baseShortcutName, ext);

                                    shortcutName = baseShortcutName;

                                    // Handle Duplicates while preserving extension
                                    while (System.IO.File.Exists(shortcutName))
                                    {
                                        string nameNoExt = System.IO.Path.GetFileNameWithoutExtension(droppedFile);
                                        shortcutName = System.IO.Path.Combine("Shortcuts", $"{nameNoExt} ({counter++}){ext}");
                                    }

                                    if (isWebLink)
                                    {
                                        // FIX: Never rewrite .url files on drop. This destroys custom game/app icons (Steam/Spotify).
                                        // ALWAYS copy the original file exactly as-is to preserve IconFile and IconIndex properties.
                                        System.IO.File.Copy(droppedFile, shortcutName, true);
                                    }
                                    else
                                    {
                                        System.IO.File.Copy(droppedFile, shortcutName, true);
                                    }
                                }

                                dynamic newItem = new System.Dynamic.ExpandoObject();
                                IDictionary<string, object> newItemDict = newItem;

                                newItemDict["Filename"] = shortcutName;
                                newItemDict["IsFolder"] = isFolder;
                                newItemDict["IsLink"] = isWebLink;
                                newItemDict["IsNetwork"] = IsNetworkPath(shortcutName);

                                // --- BUG FIX: Display Name for Network Roots & Folders ---
                                string displayFileName = System.IO.Path.GetFileNameWithoutExtension(droppedFile);

                                // If Path.GetFileName fails (happens for UNC roots like \\Server\Share or C:\)
                                if (string.IsNullOrWhiteSpace(displayFileName))
                                {
                                    displayFileName = droppedFile.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                                                                 .Split(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                                                                 .LastOrDefault();

                                    if (string.IsNullOrWhiteSpace(displayFileName)) displayFileName = droppedFile;
                                }

                                newItemDict["DisplayName"] = displayFileName;
                                newItemDict["AlwaysRunAsAdmin"] = false;

                                // TABS FEATURE: Get fresh framne data
                                string frameId = frame.Id?.ToString();
                                var freshFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                                JArray items = null;
                                bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";
                                bool addedToTab0 = false;

                                if (tabsEnabled && freshFrame != null)
                                {
                                    int currentTab = Convert.ToInt32(freshFrame.CurrentTab?.ToString() ?? "0");
                                    var tabs = freshFrame.Tabs as JArray ?? new JArray();
                                    if (currentTab >= 0 && currentTab < tabs.Count)
                                    {
                                        var activeTab = tabs[currentTab] as JObject;
                                        if (activeTab != null)
                                        {
                                            items = activeTab["Items"] as JArray ?? new JArray();
                                            addedToTab0 = (currentTab == 0);
                                        }
                                    }
                                    if (items == null) items = freshFrame.Items as JArray ?? new JArray();
                                }
                                else
                                {
                                    items = (freshFrame ?? frame).Items as JArray ?? new JArray();
                                }

                                int nextDisplayOrder = items.Count;
                                newItemDict["DisplayOrder"] = nextDisplayOrder;
                                items.Add(JObject.FromObject(newItem));

                                if (!string.IsNullOrEmpty(frameId))
                                {
                                    if (addedToTab0) TabManager.SynchronizeTab0Content(frameId, "tab0", "add");
                                    else if (!tabsEnabled) TabManager.SynchronizeTab0Content(frameId, "main", "add");
                                }

                                if (freshFrame != null)
                                {
                                    int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                                    if (frameIndex >= 0)
                                    {
                                        FrameDataManager.FrameData[frameIndex] = freshFrame;
                                        FrameDataManager.SaveFrameData();
                                    }
                                }

                                AddIcon(newItem, wpcont);
                                StackPanel sp = wpcont.Children[wpcont.Children.Count - 1] as StackPanel;
                                if (sp != null)
                                {
                                    // FIX: Extract arguments for the newly created shortcut
                                    string args = Utility.GetShortcutArguments(shortcutName);

                                    ClickEventAdder(sp, shortcutName, isFolder, args);

                                    // [TARGET VALIDATION ENGINE - DROP CONTENT]
                                    bool isRealNetwork = IsNetworkPath(shortcutName);
                                    bool allowNetworkChecking = _options.CheckNetworkPaths ?? false;
                                    bool isDroppedShortcutFile = shortcutName != null && shortcutName.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase);

                                    if (!isWebLink && (!isRealNetwork || allowNetworkChecking))
                                    {
                                        string resolvedTarget = shortcutName;
                                        bool excludeFromCheck = false;

                                        if (isDroppedShortcutFile)
                                        {
                                            resolvedTarget = FilePathUtilities.GetShortcutTargetUnicodeSafe(shortcutName) ?? shortcutName;

                                            if (resolvedTarget.StartsWith(@"\\") ||
                                                resolvedTarget.StartsWith("steam://", StringComparison.OrdinalIgnoreCase) ||
                                                Utility.IsStoreAppShortcut(shortcutName))
                                            {
                                                excludeFromCheck = true;
                                            }
                                        }

                                        if (!excludeFromCheck)
                                        {
                                            targetChecker.AddCheckAction(shortcutName, () => UpdateIcon(sp, shortcutName, isFolder, resolvedTarget), isFolder);
                                            if (isRealNetwork) LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Monitoring network path: {shortcutName}");
                                        }
                                        else
                                        {
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Skipped continuous background check for newly dropped shortcut/link: {shortcutName}");
                                        }
                                    }

                                    // Attach Context Menu
                                    AttachIconContextMenu(sp, newItem, frame, win);

                                    // --- HIDDEN TWEAK: Delete Original Shortcut On Drop ---
                                    if (SettingsManager.DeleteOriginalShortcutsOnDrop && isDroppedShortcut && !isFolder && !isWebLink)
                                    {
                                        try
                                        {
                                            string fileDir = System.IO.Path.GetDirectoryName(droppedFile);
                                            string userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                                            string commonDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);

                                            bool isFromDesktop = string.Equals(fileDir, userDesktop, StringComparison.OrdinalIgnoreCase) ||
                                                                 string.Equals(fileDir, commonDesktop, StringComparison.OrdinalIgnoreCase);

                                            if (isFromDesktop) System.IO.File.Delete(droppedFile);
                                        }
                                        catch { }
                                    }
                                }
                            }






                            // --- PORTAL FRAME LOGIC (RESTORED) ---
                            else if (frame.ItemsType?.ToString() == "Portal")
                            {
                                IDictionary<string, object> frameDict = frame is IDictionary<string, object> dict ? dict : ((JObject)frame).ToObject<IDictionary<string, object>>();
                                string destinationFolder = frameDict.ContainsKey("Path") ? frameDict["Path"]?.ToString() : null;

                                if (string.IsNullOrEmpty(destinationFolder))
                                {
                                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.MsgNoPortalDestination, Strings.DlgError);
                                    continue;
                                }
                                if (!System.IO.Directory.Exists(destinationFolder))
                                {
                                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgDestinationFolderGone", destinationFolder), Strings.DlgError);
                                    continue;
                                }

                                string destinationPath = System.IO.Path.Combine(destinationFolder, System.IO.Path.GetFileName(droppedFile));
                                int counter = 1;
                                string baseName = System.IO.Path.GetFileNameWithoutExtension(droppedFile);
                                string extension = System.IO.Path.GetExtension(droppedFile);

                                while (System.IO.File.Exists(destinationPath) || System.IO.Directory.Exists(destinationPath))
                                {
                                    destinationPath = System.IO.Path.Combine(destinationFolder, $"{baseName} ({counter++}){extension}");
                                }

                                if (System.IO.File.Exists(droppedFile))
                                {
                                    System.IO.File.Copy(droppedFile, destinationPath, false);
                                    portalCopiedCount++; // --- NEW ---
                                }
                                else if (System.IO.Directory.Exists(droppedFile))
                                {
                                    BackupManager.CopyDirectory(droppedFile, destinationPath);
                                    portalCopiedCount++; // --- NEW ---
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgAddItemFailed", droppedFile, ex.Message), Strings.DlgError);
                        }
                    }

                    // --- NEW: Visual Feedback for Portal Frames ---
                    if (frame.ItemsType?.ToString() == "Portal" && portalCopiedCount > 0)
                    {
                        ShowPortalToast(win, $"Copied {portalCopiedCount} item{(portalCopiedCount > 1 ? "s" : "")}");
                    }

                    FrameDataManager.SaveFrameData();
                }
                // --- URL DROP LOGIC (RESTORED) ----
                else if (e.Data.GetDataPresent(DataFormats.Text) ||
                         e.Data.GetDataPresent(DataFormats.Html) ||
                         e.Data.GetDataPresent("UniformResourceLocator"))
                {
                    try
                    {
                        string droppedUrl = ExtractUrlFromDropData(e.Data);
                        if (!string.IsNullOrEmpty(droppedUrl) && IsValidWebUrl(droppedUrl))
                        {
                            if (frame.ItemsType?.ToString() == "Data")
                            {
                                AddUrlShortcutToFrame(droppedUrl, frame, wpcont);
                            }
                            else
                            {
                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameCreation, "URL drops not supported for Portal frames");
                            }
                        }
                    }
                    catch (Exception urlEx)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation, $"Error processing URL drop: {urlEx.Message}");
                    }
                }
            };


      




            if (SettingsManager.EnableDimensionSnap)
            {
                win.SizeChanged += UpdateSizeFeedback;
            }






            win.LocationChanged += (s, e) =>
            {
                // --- THE UNIVERSAL DRAG GATEKEEPER ---
                // Only save coordinates to JSON if the USER is physically dragging this window with their mouse!
                // Completely ignores programmatic movement from Auto-Roll, Manual Roll-Up, and CascadeStack.
                if (SnapManager.ActiveDragWindow != win) return;
                // -------------------------------------

                string frameId = win.Tag?.ToString();
                if (string.IsNullOrEmpty(frameId)) return;

                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame != null)
                {
                    currentFrame.X = win.Left;
                    currentFrame.Y = win.Top;
                    FrameDataManager.SaveFrameData();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameUpdate, $"Position updated for frame '{currentFrame.Title}' to X={win.Left}, Y={win.Top}");
                }
            };











            InitContent();
            // Add Note frame specific context menu items after content is initialized
            if (frame.ItemsType?.ToString() == "Note")
            {
                // Use a small delay to ensure the TextBox is fully created and added to the visual tree
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        // Find the TextBox that was created in InitContent()
                        var border = win.Content as Border;
                        var dockPanel = border?.Child as DockPanel;
                        var noteTextBox = dockPanel?.Children.OfType<TextBox>().FirstOrDefault();
                        if (noteTextBox != null)
                        {
                            NoteFramemanager.AddNoteContextMenuItems(CnMnFramemanager, frame, noteTextBox);
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                                $"Added Note context menu items for frame '{frame.Title}'");
                        }
                        else
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                                $"CRITICAL: Could not find TextBox for Note frame '{frame.Title}' - checking DockPanel children");
                            // Debug: Log what children actually exist
                            if (dockPanel != null)
                            {
                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                                    $"DockPanel has {dockPanel.Children.Count} children:");
                                for (int i = 0; i < dockPanel.Children.Count; i++)
                                {
                                    var child = dockPanel.Children[i];
                                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameCreation,
                                        $" Child {i}: {child.GetType().Name}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FrameCreation,
                            $"Error adding Note context menu: {ex.Message}");
                    }
                }), System.Windows.Threading.DispatcherPriority.Loaded);
            }
            win.Show();







            // Check for persistent Legendary Mode (Nikos or >:)
            string fTitle = frame.Title?.ToString() ?? "";
            if (fTitle == "Nikos" || fTitle == "Nikos Georgousis" || fTitle.Contains(">:"))
            {
                // Defer slightly to ensure window is loaded
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    // We call ProcessTitleChange to trigger the visual effect
                    InterCore.ProcessTitleChange(frame, fTitle, "");
                }), System.Windows.Threading.DispatcherPriority.Background);
            }
            {
                // Defer slightly to ensure window is loaded
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    // We call ProcessTitleChange with same names just to trigger the logic
                    InterCore.ProcessTitleChange(frame, frame.Title.ToString(), "");
                }), System.Windows.Threading.DispatcherPriority.Background);
            }







            IDictionary<string, object> frameDict = frame is IDictionary<string, object> dict ? dict : ((JObject)frame).ToObject<IDictionary<string, object>>();
            SnapManager.AddSnapping(win, frameDict);
            // Apply custom color if present, otherwise let Utility handle the fallback
            string customColor = frame.CustomColor?.ToString();
            Utility.ApplyTintAndColorToFrame(win, customColor);
            targetChecker.Start();
        }

        public static void AddIcon(dynamic icon, WrapPanel wpcont, dynamic frameContext = null)
        {
            // 1. EXTRACT DATA
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict
                ? dict
                : ((JObject)icon).ToObject<IDictionary<string, object>>();

            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];
            bool isLink = iconDict.ContainsKey("IsLink") && (bool)iconDict["IsLink"];

            // Fail-Safe Network Detection
            bool isNetwork = iconDict.ContainsKey("IsNetwork") && (bool)iconDict["IsNetwork"];
            if (!isNetwork && !string.IsNullOrEmpty(filePath))
            {
                isNetwork = IsNetworkPath(filePath);
            }

            // --- STEP 1: Determine Settings Context ---
            dynamic settings = frameContext;
            if (settings == null)
            {
                try
                {
                    if (wpcont.Tag != null)
                    {
                        string frameId = wpcont.Tag.ToString();
                        settings = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == frameId);
                    }
                    if (settings == null)
                    {
                        NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(wpcont);
                        string frameId = win?.Tag?.ToString();
                        if (!string.IsNullOrEmpty(frameId))
                        {
                            settings = GetFrameData().FirstOrDefault(f => f.Id?.ToString() == frameId);
                        }
                    }
                }
                catch { }
            }

            // --- NEW: Identify frame Type (Portal vs Data) ---
            bool isPortal = false;
            if (settings != null)
            {
                try
                {
                    // Handle both JObject (JSON) and ExpandoObject (Runtime)
                    if (settings is JObject jObj)
                        isPortal = jObj["ItemsType"]?.ToString() == "Portal";
                    else
                        isPortal = settings.ItemsType?.ToString() == "Portal";
                }
                catch { }
            }

            // --- STEP 2: Read Customization Settings ---
            int iconSpacing = 5;
            double iconWidth = 40, iconHeight = 40;
            bool grayscale = false;
            string textColorName = null;
            bool disableShadow = false;

            if (settings != null)
            {
                try { iconSpacing = Convert.ToInt32(settings.IconSpacing?.ToString() ?? "5"); } catch { }

                string sizeVal = settings.IconSize?.ToString() ?? "Medium";
                switch (sizeVal)
                {
                    case "Tiny": iconWidth = iconHeight = 24; break;
                    case "Small": iconWidth = iconHeight = 32; break;
                    case "Large": iconWidth = iconHeight = 48; break;
                    case "Huge": iconWidth = iconHeight = 64; break;
                    default: iconWidth = iconHeight = 40; break;
                }

                try { grayscale = settings.GrayscaleIcons?.ToString().ToLower() == "true"; } catch { }
                try { textColorName = settings.TextColor?.ToString(); } catch { }
                try { disableShadow = settings.DisableTextShadow?.ToString().ToLower() == "true"; } catch { }
            }

            // --- STEP 3: Create UI Elements ---
            StackPanel sp = new StackPanel
            {
                Margin = new Thickness(iconSpacing),
                Width = 60 + (iconSpacing * 2)
            };

            System.Windows.Controls.Image ico = new System.Windows.Controls.Image
            {
                Width = iconWidth,
                Height = iconHeight,
                Margin = new Thickness(5)
            };
            if (SettingsManager.IconVisibilityEffect != IconVisibilityEffect.None)
            {
                ico.Effect = Utility.CreateIconEffect(SettingsManager.IconVisibilityEffect);
            }
            // --- NEW: Intercept Internal Blank Spacer ---
            if (filePath != null && filePath.StartsWith("INTERNAL_BLANK_"))
            {
                string resImage = filePath.StartsWith("INTERNAL_BLANK_DOT") ? "dot.png" : "empty.png";
                ico.Source = new BitmapImage(new Uri($"pack://application:,,,/Resources/{resImage}", UriKind.Absolute));
                sp.Children.Add(ico);

                // Add empty label to maintain exact grid alignment
                TextBlock blankLbl = new TextBlock { Text = " " };
                sp.Children.Add(blankLbl);

                // CRITICAL: Must be transparent to catch mouse clicks/drags
                sp.Background = System.Windows.Media.Brushes.Transparent;

                sp.Tag = new { FilePath = filePath, IsFolder = false, Arguments = (string)null };
                sp.ToolTip = new ToolTip { Content = Strings.TooltipBlankSpacer };
                wpcont.Children.Add(sp);
                return; // Skip standard extraction
            }


            // --- LAZY LOAD PREPARATION ---
            bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
            bool isUrlFile = System.IO.Path.GetExtension(filePath).ToLower() == ".url";

            // 1. Apply base visual effects immediately so the UI draws fast
            double targetIconOpacity = 1.0;
            if (SettingsManager.ApplyTintToIcons)
            {
                targetIconOpacity = Math.Max(0.2, SettingsManager.TintValue / 100.0);
                ico.Opacity = targetIconOpacity;
            }

            if (grayscale)
            {
                ico.Opacity = 0.6;
                ico.Effect = new DropShadowEffect { Color = Colors.Gray, BlurRadius = 0, ShadowDepth = 0 };
            }

            //// 2. Set an instant placeholder to prevent blank spaces during load
            //ico.Source = isFolder
            //    ? new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"))
            //    : new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));

            // 2. Set an instant placeholder to prevent layout collapsing during load
            // OPTION A: We use your existing transparent 'empty.png' (or null).
            // This reserves the physical space for the grid and catches mouse clicks, 
            // but remains completely invisible until the real icon seamlessly fades in.
            ico.Source = new BitmapImage(new Uri("pack://application:,,,/Resources/placeholder.png", UriKind.Absolute));


            // 3. Assemble Network Overlay
            if (isNetwork)
            {
                Grid iconGrid = new Grid { Width = iconWidth + 8, Height = iconHeight + 8 };
                iconGrid.Children.Add(ico);

                TextBlock networkIndicator = new TextBlock
                {
                    Text = "🔗",
                    FontSize = 14,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(65, 135, 225)),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(2, 2, 0, 0),
                    Effect = new DropShadowEffect { Color = Colors.Black, BlurRadius = 2, Opacity = 0.7 }
                };
                iconGrid.Children.Add(networkIndicator);
                sp.Children.Add(iconGrid);
            }
            else
            {
                sp.Children.Add(ico);
            }

            // 4. Assemble Text Label
            string displayName = (!iconDict.ContainsKey("DisplayName") || iconDict["DisplayName"] == null)
                ? System.IO.Path.GetFileNameWithoutExtension(filePath)
                : (string)iconDict["DisplayName"];

            if (string.IsNullOrWhiteSpace(displayName) && !string.IsNullOrWhiteSpace(filePath))
            {
                displayName = filePath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                                      .Split(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                                      .LastOrDefault();
                if (string.IsNullOrWhiteSpace(displayName)) displayName = "Unknown";
            }

            string originalDisplayName = displayName;
            if (displayName.Length > SettingsManager.MaxDisplayNameLength)
                displayName = displayName.Substring(0, SettingsManager.MaxDisplayNameLength) + "...";

            System.Windows.Media.Brush textBrush = System.Windows.Media.Brushes.White;
            if (!string.IsNullOrEmpty(textColorName))
            {
                try { textBrush = new SolidColorBrush(Utility.GetColorFromName(textColorName)); } catch { }
            }

            TextBlock lbl = new TextBlock
            {
                Text = displayName,
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                Foreground = textBrush,
                MaxWidth = 70,
                Opacity = targetIconOpacity
            };

            if (!disableShadow)
                lbl.Effect = new DropShadowEffect { Color = Colors.Black, Direction = 315, ShadowDepth = 2, BlurRadius = 3, Opacity = 0.8 };

            sp.Children.Add(lbl);
            sp.Tag = new { FilePath = filePath, IsFolder = isFolder, Arguments = (string)(iconDict.ContainsKey("Arguments") ? iconDict["Arguments"] : null) };
            sp.ToolTip = new ToolTip { Content = originalDisplayName + "\nLoading target..." };

            wpcont.Children.Add(sp); // ADD TO UI INSTANTLY!

            // --- 5. BACKGROUND RESOLVER TASK ---
            // This offloads the 30-second WshShell and Directory.Exists network hangs!
            Action backgroundLoader = () =>
            {
                ImageSource shortcutIcon = null;
                string targetPath = null;
                bool targetIsUncRoot = false;

                // Helper to safely create unfrozen background images for the UI
                BitmapImage CreateSafeBitmap(string uriPath)
                {
                    var bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.UriSource = new Uri(uriPath);
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.EndInit();
                    bmp.Freeze(); // CRITICAL FIX: Allows passing to UI thread
                    return bmp;
                }

                if (isShortcut || isLink || isUrlFile)
                {
                    if (isUrlFile || isLink)
                    {
                        var urlIcon = GetUrlCustomIcon(filePath);
                        if (urlIcon.Path != null) shortcutIcon = IconManager.ExtractIconFromFile(urlIcon.Path, urlIcon.Index);
                    }

                    if (shortcutIcon == null)
                    {
                        try
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(System.IO.Path.GetFullPath(filePath));
                            targetPath = shortcut.TargetPath;

                            if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                            {
                                string[] iconParts = shortcut.IconLocation.Split(',');
                                string iconPath = iconParts[0];
                                int iconIndex = 0;
                                if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex)) iconIndex = parsedIndex;
                                if (System.IO.File.Exists(iconPath)) shortcutIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                            }
                        }
                        catch { }
                    }
                }

                if (shortcutIcon == null)
                {
                    if (isLink || isUrlFile)
                    {
                        string urlTarget = targetPath;
                        if (string.IsNullOrEmpty(urlTarget))
                        {
                            try
                            {
                                var lines = System.IO.File.ReadAllLines(filePath);
                                var urlLine = lines.FirstOrDefault(l => l.StartsWith("URL=", StringComparison.OrdinalIgnoreCase));
                                if (urlLine != null) urlTarget = urlLine.Substring(4).Trim();
                            }
                            catch { }
                        }

                        if (!string.IsNullOrEmpty(urlTarget) && !urlTarget.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && !urlTarget.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                        {
                            shortcutIcon = IconManager.GetIconForFile(targetPath, filePath, isFolder, isLink, isShortcut, iconDict);
                        }

                        if (shortcutIcon == null) shortcutIcon = CreateSafeBitmap("pack://application:,,,/Resources/link-White.png");
                    }
                    else if (isShortcut)
                    {
                        try
                        {
                            if (string.IsNullOrEmpty(targetPath)) targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                            if (string.IsNullOrEmpty(targetPath))
                            {
                                try { WshShell shell = new WshShell(); IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(System.IO.Path.GetFullPath(filePath)); targetPath = shortcut.TargetPath; } catch { }
                            }
                            if (IsUncRoot(targetPath)) { targetIsUncRoot = true; isFolder = true; isNetwork = true; }
                        }
                        catch { }

                        bool targetIsFolder = false;
                        bool targetExists = false;
                        if (!string.IsNullOrEmpty(targetPath))
                        {
                            targetIsFolder = System.IO.Directory.Exists(targetPath) || targetIsUncRoot;
                            targetExists = targetIsFolder || System.IO.File.Exists(targetPath);
                        }

                        if (!targetExists && ScrapeLnkForNetworkRoot(filePath))
                        {
                            targetIsFolder = true; isNetwork = true; isFolder = true; targetIsUncRoot = true;
                        }

                        if (!isPortal && (isFolder || targetIsFolder)) shortcutIcon = null;
                        else if (targetExists) shortcutIcon = Utility.GetShellIcon(targetPath, targetIsFolder);
                        else shortcutIcon = Utility.GetShellIcon(filePath, isFolder);
                    }
                    else
                    {
                        if (!isPortal && isFolder) shortcutIcon = null;
                        else shortcutIcon = Utility.GetShellIcon(filePath, isFolder);
                    }
                }

                if (shortcutIcon == null)
                {
                    if (isFolder)
                    {
                        bool isUncRoot = targetIsUncRoot || IsUncRoot(filePath);
                        bool valid = isUncRoot || FilePathUtilities.DoesFolderExist(filePath, isFolder);
                        if (!valid && isShortcut && !string.IsNullOrEmpty(targetPath)) valid = IsUncRoot(targetPath) || System.IO.Directory.Exists(targetPath);
                        shortcutIcon = valid ? CreateSafeBitmap("pack://application:,,,/Resources/folder-White.png") : CreateSafeBitmap("pack://application:,,,/Resources/folder-WhiteX.png");
                    }
                    else shortcutIcon = CreateSafeBitmap("pack://application:,,,/Resources/file-WhiteX.png");
                }

                // CRITICAL FIX: Freeze dynamically extracted shell icons as well
                if (shortcutIcon != null && shortcutIcon.CanFreeze && !shortcutIcon.IsFrozen)
                {
                    shortcutIcon.Freeze();
                }

                // [Tooltip Parsing]
                string toolTipText = originalDisplayName;
                string resolvedTarget = filePath;

                if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk") resolvedTarget = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                else if (System.IO.Path.GetExtension(filePath).ToLower() == ".url")
                {
                    try { string content = System.IO.File.ReadAllText(filePath); var match = System.Text.RegularExpressions.Regex.Match(content, @"URL=([^\r\n]+)"); if (match.Success) resolvedTarget = match.Groups[1].Value.Trim(); } catch { }
                }
                if (string.IsNullOrEmpty(resolvedTarget)) resolvedTarget = filePath;

                string displayTarget = resolvedTarget;
                if (!string.IsNullOrEmpty(displayTarget) && displayTarget.Contains("!") && !displayTarget.Contains(":\\"))
                {
                    string packageId = displayTarget.Split('!')[0]; int hashIndex = packageId.IndexOf('_'); if (hashIndex > 0) packageId = packageId.Substring(0, hashIndex); displayTarget = $"Windows App ({packageId})";
                }
                if (resolvedTarget != filePath) toolTipText += $"\nTarget: {displayTarget}"; else toolTipText += $"\nLocation: {displayTarget}";

                // 6. PUSH TO UI
                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (shortcutIcon != null)
                    {
                        if (grayscale && shortcutIcon is BitmapSource bmp)
                        {
                            var formatBmp = new FormatConvertedBitmap(bmp, PixelFormats.Gray8, BitmapPalettes.Gray256, 0);
                            formatBmp.Freeze(); // CRITICAL FIX
                            ico.Source = formatBmp;
                        }
                        else ico.Source = shortcutIcon;
                    }
                    sp.ToolTip = new ToolTip { Content = toolTipText };
                }), System.Windows.Threading.DispatcherPriority.Background);
            };





            // --- SAFE EXECUTION WRAPPER ---
            Action safeBackgroundExecution = () =>
            {
                try { backgroundLoader(); }
                catch { /* Catch all unexpected errors so they NEVER crash the app */ }
            };

            if (isNetwork)
            {
                // CRITICAL FIX: COM Objects (WshShell) MUST run on an STA thread. Task.Run is MTA!
                System.Threading.Thread staThread = new System.Threading.Thread(new System.Threading.ThreadStart(safeBackgroundExecution));
                staThread.SetApartmentState(System.Threading.ApartmentState.STA); // Set to Single Threaded Apartment
                staThread.IsBackground = true;
                staThread.Start();
            }
            else
            {
                // Local files are fast, run immediately to avoid visual "pop in" delays
                safeBackgroundExecution();
            }
        }


        //    // --- ICON EXTRACTION LOGIC ---
        //    ImageSource shortcutIcon = null;
        //    bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
        //    bool isUrlFile = System.IO.Path.GetExtension(filePath).ToLower() == ".url";

        //    // Variables to track target state
        //    string targetPath = null;
        //    bool targetIsUncRoot = false;

        //    // FIX: Unified Custom Icon Extraction for Links AND Shortcuts
        //    // We check for a custom icon FIRST. If one exists, we use it immediately.
        //    if (isShortcut || isLink || isUrlFile)
        //    {
        //        // NEW: Special handling for .url files using manual parser
        //        if (isUrlFile || isLink)
        //        {
        //            var urlIcon = GetUrlCustomIcon(filePath);
        //            if (urlIcon.Path != null)
        //            {
        //                shortcutIcon = IconManager.ExtractIconFromFile(urlIcon.Path, urlIcon.Index);
        //            }
        //        }

        //        // Existing .lnk logic
        //        if (shortcutIcon == null)
        //        {


        //            try
        //            {
        //                WshShell shell = new WshShell();
        //                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(System.IO.Path.GetFullPath(filePath));
        //                targetPath = shortcut.TargetPath;

        //                // 1. Try Custom Icon (Properties -> Change Icon)
        //                // Check if IconLocation is valid and NOT ",0" (which implies default)
        //                if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
        //                {
        //                    string[] iconParts = shortcut.IconLocation.Split(',');
        //                    string iconPath = iconParts[0];
        //                    int iconIndex = 0;
        //                    if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
        //                        iconIndex = parsedIndex;

        //                    if (System.IO.File.Exists(iconPath))
        //                    {
        //                        shortcutIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
        //                    }
        //                }
        //            }


        //            catch { }
        //        }
        //        }

        //    // 2. Fallback: Determine Icon if no custom icon found
        //    if (shortcutIcon == null)
        //    {
        //        if (isLink || isUrlFile)
        //        {
        //            // --- NEW: Smart Protocol Detection ---
        //            string urlTarget = targetPath; // Inherit from .lnk extraction if available

        //            if (string.IsNullOrEmpty(urlTarget))
        //            {
        //                try
        //                {
        //                    // Extract the actual URL from the .url file
        //                    var lines = System.IO.File.ReadAllLines(filePath);
        //                    var urlLine = lines.FirstOrDefault(l => l.StartsWith("URL=", StringComparison.OrdinalIgnoreCase));
        //                    if (urlLine != null) urlTarget = urlLine.Substring(4).Trim();
        //                }
        //                catch { }
        //            }

        //            // If it's a custom app protocol (e.g., spotify:, steam:), let IconManager grab the specific internal icon
        //            if (!string.IsNullOrEmpty(urlTarget) &&
        //                !urlTarget.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
        //                !urlTarget.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        //            {
        //                shortcutIcon = IconManager.GetIconForFile(targetPath, filePath, isFolder, isLink, isShortcut, iconDict);
        //            }

        //            // If it's a standard web link, or the custom extraction failed, enforce the unified white theme
        //            if (shortcutIcon == null)
        //            {
        //                shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png")); //check
        //            }
        //        }
        //        else if (isShortcut)
        //        {
        //            try
        //            {
        //                // FIX: Use the robust Unicode-Safe reader (IShellLink)
        //                // This handles tricky paths better than WshShell
        //                if (string.IsNullOrEmpty(targetPath))
        //                    targetPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);

        //                // Fallback to WshShell only if utility failed and we haven't tried yet
        //                if (string.IsNullOrEmpty(targetPath))
        //                {
        //                    try
        //                    {
        //                        WshShell shell = new WshShell();
        //                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(System.IO.Path.GetFullPath(filePath));
        //                        targetPath = shortcut.TargetPath;
        //                    }
        //                    catch { }
        //                }

        //                // FIX 1: RUNTIME CORRECTION (Server Root)
        //                // If target is \\192.168.1.10, force isFolder=true immediately.
        //                if (IsUncRoot(targetPath))
        //                {
        //                    targetIsUncRoot = true;
        //                    isFolder = true;   // Force folder treatment!
        //                    isNetwork = true;  // Force network flag
        //                }
        //            }
        //            catch { }

        //            // A. Standard Checks
        //            bool targetIsFolder = false;
        //            bool targetExists = false;

        //            if (!string.IsNullOrEmpty(targetPath))
        //            {
        //                targetIsFolder = System.IO.Directory.Exists(targetPath) || targetIsUncRoot;
        //                targetExists = targetIsFolder || System.IO.File.Exists(targetPath);
        //            }

        //            // B. SCRAPING FALLBACK
        //            // If standard checks failed, scan the file content for \\Server pattern
        //            if (!targetExists)
        //            {
        //                if (ScrapeLnkForNetworkRoot(filePath))
        //                {
        //                    targetIsFolder = true;
        //                    isNetwork = true;      // Network Symbol
        //                    isFolder = true;       // Folder Shape
        //                    targetIsUncRoot = true; // CRITICAL: Force Validity (No White X)
        //                }
        //            }

        //            // C. Icon Selection Logic
        //            if (!isPortal && (isFolder || targetIsFolder))
        //            {
        //                shortcutIcon = null; // Force fall-through to White Theme
        //            }
        //            else if (targetExists)
        //            {
        //                shortcutIcon = Utility.GetShellIcon(targetPath, targetIsFolder);
        //            }
        //            else
        //            {
        //                shortcutIcon = Utility.GetShellIcon(filePath, isFolder);
        //            }
        //        }
        //        else
        //        {
        //            // Standard File/Folder (Not a shortcut/link)
        //            if (!isPortal && isFolder)
        //            {
        //                shortcutIcon = null;
        //            }
        //            else
        //            {
        //                shortcutIcon = Utility.GetShellIcon(filePath, isFolder);
        //            }
        //        }
        //    }

        //    // Final Fallback (The White Theme Logic)
        //    if (shortcutIcon == null)
        //    {
        //        if (isFolder)
        //        {
        //            // FIX 2: VALIDATION OVERRIDE
        //            // If it is a UNC Root (either detected via shortcut target or direct path), consider it "Valid"
        //            bool isUncRoot = targetIsUncRoot || IsUncRoot(filePath);
        //            bool valid = isUncRoot || FilePathUtilities.DoesFolderExist(filePath, isFolder);

        //            // Double check target if shortcut
        //            if (!valid && isShortcut && !string.IsNullOrEmpty(targetPath))
        //            {
        //                valid = IsUncRoot(targetPath) || System.IO.Directory.Exists(targetPath);
        //            }

        //            if (valid)
        //                shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
        //            else
        //                shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
        //        }
        //        else
        //        {
        //            shortcutIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
        //        }
        //    }

        //    ico.Source = shortcutIcon;










        //    // Apply Grayscale
        //    if (grayscale && shortcutIcon is BitmapSource bmp)
        //    {
        //        ico.Source = new FormatConvertedBitmap(bmp, PixelFormats.Gray8, BitmapPalettes.Gray256, 0);
        //    }
        //    else if (grayscale)
        //    {
        //        ico.Opacity = 0.6;
        //        ico.Effect = new DropShadowEffect { Color = Colors.Gray, BlurRadius = 0, ShadowDepth = 0 };
        //    }

        //    // --- NEW: Apply Icon Tint ---
        //    double targetIconOpacity = 1.0;
        //    if (SettingsManager.ApplyTintToIcons)
        //    {
        //        // Convert TintValue (0-100) to Opacity (0.0-1.0)
        //        // We use a slight offset because 0% tint would make icons invisible.
        //        // A minimum of 20% opacity ensures they remain clickable.
        //        targetIconOpacity = Math.Max(0.2, SettingsManager.TintValue / 100.0);
        //        ico.Opacity = targetIconOpacity;
        //    }

        //    // --- STEP 4: Network Overlay ---
        //    if (isNetwork)
        //    {
        //        Grid iconGrid = new Grid { Width = iconWidth + 8, Height = iconHeight + 8 };
        //        iconGrid.Children.Add(ico);

        //        TextBlock networkIndicator = new TextBlock
        //        {
        //            Text = "🔗",
        //            FontSize = 14,
        //            Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(65, 135, 225)),
        //            HorizontalAlignment = HorizontalAlignment.Left,
        //            VerticalAlignment = VerticalAlignment.Top,
        //            Margin = new Thickness(2, 2, 0, 0),
        //            Effect = new DropShadowEffect { Color = Colors.Black, BlurRadius = 2, Opacity = 0.7 }
        //        };

        //        iconGrid.Children.Add(networkIndicator);
        //        sp.Children.Add(iconGrid);
        //    }
        //    else
        //    {
        //        sp.Children.Add(ico);
        //    }

        //    // --- STEP 5: Text Label ---
        //    string displayName = (!iconDict.ContainsKey("DisplayName") || iconDict["DisplayName"] == null)
        //        ? System.IO.Path.GetFileNameWithoutExtension(filePath)
        //        : (string)iconDict["DisplayName"];

        //    // --- BUG FIX: Recover missing names for UNC Roots and Drives ---
        //    if (string.IsNullOrWhiteSpace(displayName) && !string.IsNullOrWhiteSpace(filePath))
        //    {
        //        displayName = filePath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
        //                              .Split(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
        //                              .LastOrDefault();
        //        if (string.IsNullOrWhiteSpace(displayName)) displayName = "Unknown";
        //    }

        //    if (displayName.Length > SettingsManager.MaxDisplayNameLength)
        //        displayName = displayName.Substring(0, SettingsManager.MaxDisplayNameLength) + "...";

        //    System.Windows.Media.Brush textBrush = System.Windows.Media.Brushes.White;
        //    if (!string.IsNullOrEmpty(textColorName))
        //    {
        //        try { textBrush = new SolidColorBrush(Utility.GetColorFromName(textColorName)); } catch { }
        //    }

        //    TextBlock lbl = new TextBlock
        //    {
        //        Text = displayName,
        //        TextWrapping = TextWrapping.Wrap,
        //        TextAlignment = TextAlignment.Center,
        //        HorizontalAlignment = HorizontalAlignment.Center,
        //        Foreground = textBrush,
        //        MaxWidth = 70,
        //        Opacity = SettingsManager.ApplyTintToIcons ? Math.Max(0.2, SettingsManager.TintValue / 100.0) : 1.0
        //    };

        //    if (!disableShadow)
        //    {
        //        lbl.Effect = new DropShadowEffect { Color = Colors.Black, Direction = 315, ShadowDepth = 2, BlurRadius = 3, Opacity = 0.8 };
        //    }

        //    sp.Children.Add(lbl);
        //    sp.Tag = new { FilePath = filePath, IsFolder = isFolder, Arguments = (string)(iconDict.ContainsKey("Arguments") ? iconDict["Arguments"] : null) };

        //    // FIX: Show the final target in the tooltip instead of the internal shortcut path
        //    string toolTipText = displayName;
        //    string resolvedTarget = filePath;

        //    if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
        //    {
        //        resolvedTarget = FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
        //    }
        //    else if (System.IO.Path.GetExtension(filePath).ToLower() == ".url")
        //    {
        //        try
        //        {
        //            // Try to extract the true web URL or protocol from the file
        //            string content = System.IO.File.ReadAllText(filePath);
        //            var match = System.Text.RegularExpressions.Regex.Match(content, @"URL=([^\r\n]+)");
        //            if (match.Success) resolvedTarget = match.Groups[1].Value.Trim();
        //        }
        //        catch { }
        //    }

        //    // Fallback in case extraction fails
        //    if (string.IsNullOrEmpty(resolvedTarget)) resolvedTarget = filePath;

        //    string displayTarget = resolvedTarget;

        //    // Clean up MS Store App AUMID targets (e.g. Microsoft.WindowsCalculator_8wekyb3d8bbwe!App)
        //    if (!string.IsNullOrEmpty(displayTarget) && displayTarget.Contains("!") && !displayTarget.Contains(":\\"))
        //    {
        //        string packageId = displayTarget.Split('!')[0];
        //        int hashIndex = packageId.IndexOf('_');
        //        if (hashIndex > 0) packageId = packageId.Substring(0, hashIndex);
        //        displayTarget = $"Windows App ({packageId})";
        //    }

        //    if (resolvedTarget != filePath)
        //    {
        //        toolTipText += $"\nTarget: {displayTarget}";
        //    }
        //    else
        //    {
        //        toolTipText += $"\nLocation: {displayTarget}";
        //    }

        //    sp.ToolTip = new ToolTip { Content = toolTipText };
        //    wpcont.Children.Add(sp);
        //}



        private static void CreateNewFrame(string title, string itemsType, double x = 20, double y = 20, string customColor = null, string customLaunchEffect = null, string pluginId = null)
        {
            // Generate random name instead of using the passed title
            string frameName = CoreUtilities.GenerateRandomName();

            dynamic newFrame = new System.Dynamic.ExpandoObject();
            newFrame.Id = Guid.NewGuid().ToString();
            IDictionary<string, object> newframeDict = newFrame;
            // newframeDict["Title"] = title;


            //  newframeDict["Title"] = newFrame; // Use random name
            //Option to set Portal frames fodler nane
            // Only use random name for non-Portal frames
            if (itemsType != "Portal")
            {
                newframeDict["Title"] = frameName; // Use random name
            }

            newframeDict["X"] = x;
            newframeDict["Y"] = y;
            newframeDict["Width"] = 230;
            newframeDict["Height"] = 130;
            newframeDict["ItemsType"] = itemsType;

            // --- NEW: Plugin Configuration ---
            if (itemsType == "Plugin")
            {
                newframeDict["PluginId"] = pluginId;
                newframeDict["PluginSettings"] = new JObject();
                newframeDict["Items"] = new JArray(); // Safe default
            }
            else
            {
                newframeDict["Items"] = itemsType == "Portal" ? "" : new JArray();
            }

            newframeDict["CustomColor"] = customColor; // Use passed value
            newframeDict["IsHidden"] = false; // Use passed value
            newframeDict["IsLocked"] = false; // Init ISLocked

            // Initialize ALL frame properties with defaults to match JSON structure
            newframeDict["IsLocked"] = "false";
            newframeDict["IsHidden"] = "false";
            newframeDict["CustomColor"] = customColor;
            newframeDict["CustomLaunchEffect"] = customLaunchEffect;
            newframeDict["IsRolled"] = "false";
            newframeDict["AutoRoll"] = "false"; // --- NEW ---
            newframeDict["AlwaysOnTop"] = "false"; // --- NEW ---
            newframeDict["UnrolledHeight"] = 130;
            newframeDict["TextColor"] = null;
            newframeDict["BoldTitleText"] = "false";
            newframeDict["TitleTextColor"] = null;
            newframeDict["DisableTextShadow"] = "false";
            newframeDict["IconSize"] = "Medium";
            newframeDict["GrayscaleIcons"] = "false";
            newframeDict["IconSpacing"] = 5;
            newframeDict["TitleTextSize"] = "Medium";
            newframeDict["FrameBorderColor"] = null;
            newframeDict["FrameBorderThickness"] = 2;
            // TABS FEATURE: Initialize tab properties for new frames
            newframeDict["TabsEnabled"] = "false";  // Default to no tabs
            newframeDict["CurrentTab"] = 0;         // Default to first tab
            newframeDict["Tabs"] = new JArray();    // Empty tabs array

            if (itemsType == "Portal")
                if (itemsType == "Portal")
                {
                    using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                    {
                        dialog.Description = "Select the folder to monitor for this Portal Frame";
                        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            newframeDict["Path"] = dialog.SelectedPath;

                            // Get the folder name from the selected path
                            // should add and option to be checked for the below
                            // Set title to folder name for Portal frames
                            string folderName = System.IO.Path.GetFileName(dialog.SelectedPath);
                            newframeDict["Title"] = folderName;

                        }
                        else
                        {
                            return;
                        }
                    }
                }
            FrameDataManager.FrameData.Add(newFrame);
            FrameDataManager.SaveFrameData();
            CreateFrame(newFrame, new TargetChecker(1000));
        }


        /// <summary>
        /// HELPER: Extracts a Network Root path (e.g. \\Server) from a .lnk file by reading raw content.
        /// Used when Windows APIs fail to resolve the target.
        /// </summary>
        private static string GetScrapedNetworkPath(string lnkPath)
        {
            try
            {
                if (string.IsNullOrEmpty(lnkPath) || !System.IO.File.Exists(lnkPath)) return null;

                string fileContent;
                // Read with Share.ReadWrite to avoid locking issues
                using (var fs = new System.IO.FileStream(lnkPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                using (var sr = new System.IO.StreamReader(fs, System.Text.Encoding.Default))
                {
                    fileContent = sr.ReadToEnd();
                }

                // Regex to capture \\Servername or \\1.2.3.4
                // Matches start of line or null-preceded, followed by \\, then valid host chars, ending with null or newline
                var regex = new System.Text.RegularExpressions.Regex(@"(^|\0)\\\\([a-zA-Z0-9\.\-_]+)(\x00|$)", System.Text.RegularExpressions.RegexOptions.Multiline);
                var match = regex.Match(fileContent);

                if (match.Success)
                {
                    // Clean up the match (remove leading nulls if any) to get just "\\Server"
                    string raw = match.Value.Trim('\0', '\r', '\n');
                    if (raw.StartsWith(@"\\")) return raw;
                }

                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Scrape path failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// SAFE FALLBACK: Scans the raw binary content of a .lnk file for Network Root patterns (e.g. \\192.168.1.10).
        /// This bypasses Windows API validation failures for server roots or offline locations.
        /// </summary>
        private static bool ScrapeLnkForNetworkRoot(string lnkPath)
        {
            try
            {
                if (string.IsNullOrEmpty(lnkPath) || !System.IO.File.Exists(lnkPath)) return false;

                // 1. Read file safely with a timeout mechanism (handled by standard stream opening)
                // Use default encoding to capture ANSI/ASCII strings embedded in the binary
                string fileContent;
                using (var fs = new System.IO.FileStream(lnkPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite))
                using (var sr = new System.IO.StreamReader(fs, System.Text.Encoding.Default))
                {
                    fileContent = sr.ReadToEnd();
                }

                // 2. Define Regex for UNC Root: Matches \\Server or \\192.168.1.1
                // Explanation:
                // \\\\       -> Literal "\\"
                // [^\\]+     -> Any character that is NOT a backslash (The server name/IP)
                // (\x00|$)   -> Ends with a null byte or end of string (typical in binary formats)
                var networkRootRegex = new System.Text.RegularExpressions.Regex(@"^\\\\([a-zA-Z0-9\.\-_]+)(\x00|$)", System.Text.RegularExpressions.RegexOptions.Multiline);

                // 3. Scan the content
                // We scan for specific "unc" prefixes often found in LNK structures or just the raw path
                if (networkRootRegex.IsMatch(fileContent)) return true;

                // Fallback simple string check for the user's specific case (\\IP) if regex is too strict on binary noise
                // We look for "\\" followed by a digit (IP) that appears in the file
                int index = fileContent.IndexOf(@"\\");
                if (index >= 0 && index + 3 < fileContent.Length)
                {
                    char nextChar = fileContent[index + 2];
                    // If it looks like \\1... or \\a... it's likely a path. 
                    // To be safe, we check if it DOESN'T look like a file path (no ":\")
                    bool hasColon = fileContent.IndexOf(@":\", index) == index + 2; // e.g. C:\ check
                    if (!hasColon) return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                // SANITIZATION: Never let a scraping error break the UI loading
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Lnk scraping suppressed error: {ex.Message}");
                return false;
            }
        }



        private static void UpdateLockState(TextBlock lockIcon, dynamic frame, bool? forceState = null, bool saveToJson = true)
        {
            // Get the actual frame from FrameDataManager.FrameData using Id to ensure correct reference
            string frameId = frame.Id?.ToString();
            if (string.IsNullOrEmpty(frameId))
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Frame '{frame.Title}' has no Id, cannot update lock state");
                return;
            }

            int index = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
            if (index < 0)
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Frame '{frame.Title}' not found in FrameDataManager.FrameData, cannot update lock state");
                return;
            }

            dynamic actualFrame = FrameDataManager.FrameData[index];
            bool isLocked = forceState ?? (actualFrame.IsLocked?.ToString().ToLower() == "true");

            // Only update JSON if explicitly requested (e.g., during toggle, not initialization)
            if (saveToJson)
            {
                UpdateFrameProperty(actualFrame, "IsLocked", isLocked.ToString().ToLower(), $"Frame {(isLocked ? "locked" : "unlocked")}");
            }

            // Update UI on the main thread
            System.Windows.Application.Current.Dispatcher.Invoke(() =>


            {
                // Update lock icon
                lockIcon.Foreground = isLocked ? System.Windows.Media.Brushes.DeepPink : System.Windows.Media.Brushes.White;
                lockIcon.ToolTip = isLocked ? "Frame is locked (click to unlock)" : "frame is unlocked (click to lock)";

                // Find the NonActivatingWindow
                NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(lockIcon);
                if (win == null)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Could not find NonActivatingWindow for frame '{actualFrame.Title}'");
                    return;
                }

                // Update ResizeMode
                win.ResizeMode = isLocked ? ResizeMode.NoResize : ResizeMode.CanResizeWithGrip;
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FrameUpdate, $"Set ResizeMode to {win.ResizeMode} for frame '{actualFrame.Title}'");
            });

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Updated lock state for frame '{actualFrame.Title}': IsLocked={isLocked}");
        }



        // --- HELPER: Manual Icon Extraction for .url Files ---
        private static (string Path, int Index) GetUrlCustomIcon(string urlPath)
        {
            try
            {
                if (!System.IO.File.Exists(urlPath)) return (null, 0);
                var lines = System.IO.File.ReadAllLines(urlPath);
                string iconFile = null;
                int iconIndex = 0;

                foreach (var line in lines)
                {
                    string trimmed = line.Trim();
                    if (trimmed.StartsWith("IconFile=", StringComparison.OrdinalIgnoreCase))
                        iconFile = trimmed.Substring(9).Trim();
                    else if (trimmed.StartsWith("IconIndex=", StringComparison.OrdinalIgnoreCase))
                        int.TryParse(trimmed.Substring(10).Trim(), out iconIndex);
                }

                if (!string.IsNullOrEmpty(iconFile)) return (iconFile, iconIndex);
            }
            catch { }
            return (null, 0);
        }




        public static void EditItem(dynamic icon, dynamic frame, NonActivatingWindow win)
        {
            IDictionary<string, object> iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
            string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
            string displayName = iconDict.ContainsKey("DisplayName") ? (string)iconDict["DisplayName"] : System.IO.Path.GetFileNameWithoutExtension(filePath);

            // --- BUG FIX: Handle empty display names for UNC roots in editor ---
            if (string.IsNullOrWhiteSpace(displayName) && !string.IsNullOrWhiteSpace(filePath))
            {
                displayName = filePath.TrimEnd('\\', '/').Split('\\', '/').LastOrDefault();
                if (string.IsNullOrWhiteSpace(displayName)) displayName = "Unknown";
            }

            // Allow .lnk and .url
            string ext = System.IO.Path.GetExtension(filePath).ToLower();
            bool isEditable = ext == ".lnk" || ext == ".url";

            if (!isEditable)
            {
                MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.MsgEditNotAvailable, Strings.DlgInfo);
                return;
            }

            var editWindow = new EditShortcutWindow(filePath, displayName);

            // Note: The EditWindow calls 'UpdateFrameDataForIcon' internally when Save is clicked.
            // We don't need to duplicate the saving logic here.
            if (editWindow.ShowDialog() == true)
            {
                string newDisplayName = editWindow.NewDisplayName;

                // Update local memory reference immediately for responsiveness
                iconDict["DisplayName"] = newDisplayName;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"EditItem: Local display name updated to {newDisplayName}");

                // UI Refresh Logic
                if (win == null) return;

                WrapPanel wpcont = null;
                var border = win.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();
                if (scrollViewer != null) wpcont = scrollViewer.Content as WrapPanel;

                if (wpcont == null) wpcont = FrameUtilities.FindWrapPanel(win);

                if (wpcont != null)
                {
                    var sp = wpcont.Children.OfType<StackPanel>()
                        .FirstOrDefault(s => s.Tag != null && s.Tag.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString() == filePath);

                    if (sp != null)
                    {
                        RefreshSingleIconComplete(sp, filePath, newDisplayName, win);
                    }
                    else
                    {
                        // Fallback: Rebuild if not found (e.g. Tab switch might have hidden it)
                        // In most cases RefreshSingleIconComplete handles it, or a full frame refresh will catch it later.
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, "StackPanel not found for immediate refresh, data is saved though.");
                    }
                }
            }
        }



        



        private static void BackupOrRestoreShortcut(string filePath, bool targetExists, bool isFolder)
        {
            // FIX: Use ProfileManager to get the path relative to the active profile
            string tempShortcutsDir = ProfileManager.GetProfileFilePath("Temp Shortcuts");

            string backupFileName = System.IO.Path.GetFileName(filePath);
            string backupPath = System.IO.Path.Combine(tempShortcutsDir, backupFileName);

            try
            {
                // Ensure TempShortcuts directory exists
                if (!Directory.Exists(tempShortcutsDir))
                {
                    Directory.CreateDirectory(tempShortcutsDir);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Created TempShortcuts directory: {tempShortcutsDir}");
                }

                if (!targetExists)
                {
                    // Backup shortcut if target is missing and not already backed up
                    if (System.IO.File.Exists(filePath) && !System.IO.File.Exists(backupPath))
                    {
                        System.IO.File.Copy(filePath, backupPath, true);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Backed up shortcut {filePath} to {backupPath}");
                    }
                }
                else
                {
                    // Restore shortcut from backup if target exists
                    if (System.IO.File.Exists(backupPath))
                    {
                        // Verify the backup has a custom icon before restoring
                        WshShell shell = new WshShell();
                        IWshShortcut backupShortcut = (IWshShortcut)shell.CreateShortcut(backupPath);
                        if (!string.IsNullOrEmpty(backupShortcut.IconLocation))
                        {
                            System.IO.File.Copy(backupPath, filePath, true);
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Restored shortcut {filePath} from {backupPath} with custom icon");
                        }
                        // Delete backup
                        System.IO.File.Delete(backupPath);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Deleted backup {backupPath} after restoration");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error in BackupOrRestoreShortcut for {filePath}: {ex.Message}");
            }
        }

       /// <summary>
        /// Refresh frame using the same approach as CustomizeFrameForm
        /// This ensures consistent sizing and behavior
        /// </summary>
        public static void RefreshFrameUsingFormApproach(NonActivatingWindow win, dynamic frame)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Refreshing frame using form approach for '{frame.Title}'");

                // Find the WrapPanel
                var wrapPanel = FindWrapPanel(win);
                if (wrapPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot find WrapPanel for frame refresh");
                    return;
                }

                // Clear existing icons
                wrapPanel.Children.Clear();

                // TABS FEATURE: Check if tabs are enabled and load from appropriate source
                bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";
                JArray items = null;

                if (tabsEnabled)
                {
                    // Load from current tab
                    var tabs = frame.Tabs as JArray ?? new JArray();
                    int currentTab = Convert.ToInt32(frame.CurrentTab?.ToString() ?? "0");

                    if (currentTab >= 0 && currentTab < tabs.Count)
                    {
                        var activeTab = tabs[currentTab] as JObject;
                        if (activeTab != null)
                        {
                            items = activeTab["Items"] as JArray ?? new JArray();
                            string tabName = activeTab["TabName"]?.ToString() ?? $"Tab {currentTab}";
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                $"Refreshing icons from tab '{tabName}' for frame '{frame.Title}'");
                        }
                    }
                }
                else
                {
                    // Load from main Items array (existing behavior)
                    items = frame.Items as JArray;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        $"Refreshing icons from main Items for frame '{frame.Title}'");
                }

                if (items != null)
                {
                    // Sort items by DisplayOrder and add them (SAME AS FORM)
                    var sortedItems = items
                        .OfType<JObject>()
                        .OrderBy(item => item["DisplayOrder"]?.Value<int>() ?? 0)
                        .ToList();

                    foreach (dynamic item in sortedItems)
                    {
                        // Use the SAME method as the working form
                        AddIcon(item, wrapPanel);

                        // Add basic event handlers (SAME AS FORM)
                        if (wrapPanel.Children.Count > 0)
                        {
                            var sp = wrapPanel.Children[wrapPanel.Children.Count - 1] as StackPanel;
                            if (sp != null)
                            {
                                string filePath = item.Filename?.ToString() ?? "Unknown";
                                bool isFolder = item.IsFolder?.ToString().ToLower() == "true";
                                ClickEventAdder(sp, filePath, isFolder);

                                // FIX: Attach the missing Context Menu
                                AttachIconContextMenu(sp, item, frame, win);
                            }
                        }

                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Successfully refreshed {sortedItems.Count} icons for frame '{frame.Title}' using form approach");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing frame using form approach: {ex.Message}");
            }
        }

        /// <summary>
        /// Helper method to find WrapPanel in frame window
        /// </summary>
        private static WrapPanel FindWrapPanel(NonActivatingWindow win)
        {
            try
            {
                var border = win.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                var scrollViewer = dockPanel?.Children.OfType<ScrollViewer>().FirstOrDefault();
                return scrollViewer?.Content as WrapPanel;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error finding WrapPanel: {ex.Message}");
                return null;
            }
        }



        // [TARGET VALIDATION ENGINE - START]
        private static void UpdateIcon(StackPanel sp, string filePath, bool isFolder, string resolvedTargetPath = null)
        {
            if (System.Windows.Application.Current == null) return;

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                if (filePath != null && filePath.StartsWith("INTERNAL_BLANK_")) return;

                // 1. OPTIMIZATION: Cache Check (Ensures local proxy file exists)
                bool fileExists = System.IO.File.Exists(filePath) || System.IO.Directory.Exists(filePath);
                if (!fileExists) return;

                DateTime currentLastWrite;
                try { currentLastWrite = System.IO.File.GetLastWriteTime(filePath); }
                catch { return; }

                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";
                bool isUrlFile = System.IO.Path.GetExtension(filePath).ToLower() == ".url";
                bool targetValid = true;

                // 2. TARGET VALIDATION LOGIC (Native File Polling, No COM)
                if (isShortcut && !Utility.IsStoreAppShortcut(filePath))
                {
                    string tPath = resolvedTargetPath ?? FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath);
                    targetValid = !string.IsNullOrEmpty(tPath) &&
                                 (System.IO.File.Exists(tPath) || System.IO.Directory.Exists(tPath));
                }
                else if (!isShortcut)
                {
                    targetValid = System.IO.File.Exists(filePath) || System.IO.Directory.Exists(filePath);
                }
                // [TARGET VALIDATION ENGINE - END]

                bool isNowBroken = !targetValid;

                if (_iconStates.TryGetValue(filePath, out var lastState))
                {
                    if (lastState.LastWrite == currentLastWrite && lastState.IsBroken == isNowBroken)
                        return; // Cache hit
                }
                _iconStates[filePath] = (currentLastWrite, isNowBroken);

                // 2. Identify Web Link (Do NOT return early anymore)
                bool isWebLink = isUrlFile;
                if (!isWebLink)
                {
                    try
                    {
                        foreach (var frame in FrameDataManager.FrameData)
                        {
                            if (frame.ItemsType?.ToString() == "Data")
                            {
                                bool CheckList(JArray list)
                                {
                                    if (list == null) return false;
                                    foreach (var item in list)
                                    {
                                        string itemPath = item["Filename"]?.ToString();
                                        if (!string.IsNullOrEmpty(itemPath) && string.Equals(itemPath, filePath, StringComparison.OrdinalIgnoreCase))
                                            return item["IsLink"]?.ToObject<bool>() ?? false;
                                    }
                                    return false;
                                }
                                if (CheckList(frame.Items as JArray)) { isWebLink = true; break; }
                                var tabs = frame.Tabs as JArray;
                                if (tabs != null)
                                {
                                    foreach (var tab in tabs) if (CheckList(tab["Items"] as JArray)) { isWebLink = true; goto FoundLink; }
                                }
                            }
                        }
                    FoundLink:;
                    }
                    catch { }
                }

                System.Windows.Controls.Image ico = sp.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                if (ico == null) return;

                // 3. MAIN ICON LOGIC
                ImageSource newIcon = null;

                // [TARGET VALIDATION ENGINE - SECONDARY PATCH]
                string targetPath = isShortcut ? (resolvedTargetPath ?? FilePathUtilities.GetShortcutTargetUnicodeSafe(filePath)) : filePath;

                bool targetExists = System.IO.File.Exists(targetPath) || System.IO.Directory.Exists(targetPath);
                bool isTargetFolder = System.IO.Directory.Exists(targetPath);
                if (isShortcut && isTargetFolder) isFolder = true;

                if (isShortcut) BackupOrRestoreShortcut(filePath, targetExists, isFolder);

                
                // --- PRIORITY LOGIC START ---
                // CASE A: WEB LINKS
                if (isWebLink)
                {
                    // 1. Try Custom Icon first
                    var urlIcon = GetUrlCustomIcon(filePath); // <--- USE NEW HELPER
                    if (urlIcon.Path != null)
                    {
                        newIcon = IconManager.ExtractIconFromFile(urlIcon.Path, urlIcon.Index);
                    }
                    else
                    {
                        // Try WshShell only if manual failed (legacy fallback)
                        try { /* existing WshShell logic */ } catch { }
                    }

                    // 2. Fallback to Theme or Custom Protocol
                    if (newIcon == null)
                    {
                        // Recheck the actual target inside the file
                        string fileText = "";
                        try { fileText = System.IO.File.ReadAllText(filePath).ToLower(); } catch { }

                        if (fileText.Contains("spotify:") || fileText.Contains("spotify.com"))
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/spotify-White.png"));
                        }
                        else if (fileText.Contains("steam://"))
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/steam-White.png"));
                        }
                        else
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                        }
                    }
                }
                // CASE B: FOLDERS
                else if (isFolder)
                {
                    // 1. Broken Target -> White X
                    if (!FilePathUtilities.DoesFolderExist(filePath, isFolder))
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                    }

                    // 2. Custom Icon
                    if (newIcon == null && isShortcut)
                    {
                        try
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                            if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                            {
                                string[] iconParts = shortcut.IconLocation.Split(',');
                                string iconPath = iconParts[0];
                                int iconIndex = 0;
                                if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex))
                                    iconIndex = parsedIndex;

                                if (System.IO.File.Exists(iconPath))
                                    newIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                            }
                        }
                        catch { }
                    }

                    // 3. Theme Fallback
                    if (newIcon == null)
                    {
                        NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);
                        string frameId = win?.Tag?.ToString();
                        var frame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                        bool isPortal = frame != null && frame.ItemsType?.ToString() == "Portal";

                        if (isPortal) newIcon = Utility.GetShellIcon(filePath, true);
                        else newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    }
                }
                // CASE C: BROKEN FILES
                else if (!targetExists)
                {
                    if (Utility.IsStoreAppShortcut(filePath))
                    {
                        // --- BUG FIX: Remove arrows from Store Apps ---
                        // GetShellIcon forces an arrow on .lnk files. ExtractAssociatedIcon grabs the raw cached PE icon.
                        try { newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource(); }
                        catch { newIcon = Utility.GetShellIcon(filePath, isFolder); }
                    }
                    else
                    {
                        newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    }
                }
                // CASE D: SHORTCUT FILES
                else if (isShortcut)
                {
                    try
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                        if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                        {
                            string[] iconParts = shortcut.IconLocation.Split(',');
                            string iconPath = iconParts[0];
                            int iconIndex = 0;
                            if (iconParts.Length == 2 && int.TryParse(iconParts[1], out int parsedIndex)) iconIndex = parsedIndex;

                            if (System.IO.File.Exists(iconPath))
                                newIcon = IconManager.ExtractIconFromFile(iconPath, iconIndex);
                        }
                    }
                    catch { }

                    if (newIcon == null)
                    {
                        // --- BUG FIX: Remove arrows from Folders and URIs ---
                        // 1. targetExists checks BOTH File and Directory (Fixes Folders like "Downloads")
                        if (targetPath != null && targetExists)
                        {
                            newIcon = Utility.GetShellIcon(targetPath, isTargetFolder);
                        }
                        // 2. If forced to fallback to the .lnk, bypass the Shell's arrow overlay natively
                        else
                        {
                            try { newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource(); }
                            catch { newIcon = Utility.GetShellIcon(filePath, false); }
                        }
                    }
                }
                // CASE E: STANDARD FILES
                //else
                //{
                //    try { newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource(); }
                //    catch { newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png")); }
                //}
                // CASE E: STANDARD FILES
                else
                {
                    try { newIcon = System.Drawing.Icon.ExtractAssociatedIcon(filePath).ToImageSource(); }
                    catch
                    {
                        // FIX: Office/AppX Icon Crash (.pptx, etc.). Fallback to Shell API.
                        newIcon = Utility.GetShellIcon(filePath, false);
                        if (newIcon == null)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        }
                    }
                }


                // Apply
                if (ico.Source != newIcon && newIcon != null)
                {
                    ico.Source = newIcon;
                    if (IconManager.IconCache.ContainsKey(filePath))
                    {
                        IconManager.IconCache[filePath] = newIcon;
                    }
                }
            });
        }

        



        // Safety method to ensure no frames are stuck in transition state
        public static void ClearAllTransitionStates()
        {
            if (_framesInTransition.Count > 0)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FrameUpdate, $"Clearing {_framesInTransition.Count} stuck transition states");
                _framesInTransition.Clear();
            }
        }

        public static void UpdateOptionsAndClickEvents()
        {


            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, $"Updating options, new singleClickToLaunch={SettingsManager.SingleClickToLaunch}");

            // Update _options
            _options = new
            {
                IsSnapEnabled = SettingsManager.IsSnapEnabled,
                ShowBackgroundImageOnPortalFences = SettingsManager.ShowBackgroundImageOnPortalFrames,
                Showintray = SettingsManager.ShowInTray,
                EnableSounds = SettingsManager.EnableSounds,
                TintValue = SettingsManager.TintValue,
                MenuTintValue = SettingsManager.MenuTintValue,
                MenuIcon = SettingsManager.MenuIcon,
                LockIcon = SettingsManager.LockIcon,
                SelectedColor = SettingsManager.SelectedColor,
                IsLogEnabled = SettingsManager.IsLogEnabled,
                singleClickToLaunch = SettingsManager.SingleClickToLaunch,
                LaunchEffect = SettingsManager.LaunchEffect,
                CheckNetworkPaths = false // Keep this as is
            };


            if (System.Windows.Application.Current != null)
            {

                // Force UI update on the main thread
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                int updatedItems = 0;
                foreach (var win in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>())
                {
                    var wpcont = ((win.Content as Border)?.Child as DockPanel)?.Children
                        .OfType<ScrollViewer>().FirstOrDefault()?.Content as WrapPanel;
                    if (wpcont != null)
                    {
                        foreach (var sp in wpcont.Children.OfType<StackPanel>())
                        {
                            string path = sp.Tag as string;
                            if (!string.IsNullOrEmpty(path))
                            {
                                bool isFolder = Directory.Exists(path) ||
                                    (System.IO.Path.GetExtension(path).ToLower() == ".lnk" &&
                                     Directory.Exists(Utility.GetShortcutTarget(path)));
                                string arguments = null;
                                if (System.IO.Path.GetExtension(path).ToLower() == ".lnk")
                                {
                                    WshShell shell = new WshShell();
                                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(path);
                                    arguments = shortcut.Arguments;
                                }
                                ClickEventAdder(sp, path, isFolder, arguments);
                                updatedItems++;
                            }
                        }
                    }
                }
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, $"Updated click events for {updatedItems} items");
            });
            }
            else
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings, "Application.Current is null, cannot update icon.");
            }
        }


        public static void ClickEventAdder(StackPanel sp, string path, bool isFolder, string arguments = null)
        {
            // Store only path, isFolder, and arguments in Tag
            sp.Tag = new { FilePath = path, IsFolder = isFolder, Arguments = arguments };

            // --- HANG FIX 1 (Startup): Removed synchronous Utility.GetShortcutTarget(path) from here. ---
            // The folder double-check is deferred until the exact moment of clicking to prevent startup freezing.
            bool isShortcut = System.IO.Path.GetExtension(path).ToLower() == ".lnk";

            // --- NAMED LOCAL FUNCTIONS FOR EVENTS ---
            void MouseDownHandler(object sender, MouseButtonEventArgs e)
            {
                if (e.ChangedButton != MouseButton.Left) return;

                // Runtime Correction for Extension Mismatch
                if (!System.IO.File.Exists(path) && !System.IO.Directory.Exists(path))
                {
                    if (path.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase))
                    {
                        string potentialUrlPath = System.IO.Path.ChangeExtension(path, ".url");
                        if (System.IO.File.Exists(potentialUrlPath))
                        {
                            path = potentialUrlPath;
                        }
                    }
                }

                // CTRL + CLICK LOGIC
                if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);
                    string frameId = win?.Tag?.ToString();
                    dynamic frame = null;
                    if (!string.IsNullOrEmpty(frameId))
                        frame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);

                    if (frame != null && frame.ItemsType?.ToString() == "Portal")
                    {
                        // Safe resolution for Portal frames
                        bool isPortalFolder = isFolder;
                        if (!isPortalFolder && isShortcut)
                        {
                            try { isPortalFolder = System.IO.Directory.Exists(Utility.GetShortcutTarget(path)); } catch { }
                        }

                        if (isPortalFolder)
                        {
                            NavigatePortalFrame(frame, path);
                            e.Handled = true;
                            return;
                        }
                        else
                        {
                            e.Handled = true;
                            return;
                        }
                    }

                    System.Windows.Point mousePosition = e.GetPosition(sp);
                    IconDragDropManager.StartIconDrag(sp, mousePosition);
                    e.Handled = true;
                    return;
                }

                bool singleClickToLaunch = SettingsManager.SingleClickToLaunch;

                if ((singleClickToLaunch && e.ClickCount == 1) || (!singleClickToLaunch && e.ClickCount == 2))
                {
                    e.Handled = true; // Mark handled immediately so the UI thread is freed

                    // --- HANG FIX 2 (Click): Move network resolution to an STA Background Thread ---
                    System.Threading.Thread launchThread = new System.Threading.Thread(() =>
                    {
                        try
                        {
                            bool isShortcutLocal = System.IO.Path.GetExtension(path).ToLower() == ".lnk";
                            bool targetExists;
                            string resolvedPath = path;
                            bool dynamicIsFolder = isFolder;

                            if (isShortcutLocal)
                            {
                                resolvedPath = FilePathUtilities.GetShortcutTargetUnicodeSafe(path);
                                if (string.IsNullOrEmpty(resolvedPath))
                                {
                                    targetExists = false;
                                }
                                else
                                {
                                    targetExists = dynamicIsFolder ? System.IO.Directory.Exists(resolvedPath) : System.IO.File.Exists(resolvedPath);
                                    if (!dynamicIsFolder && System.IO.Directory.Exists(resolvedPath))
                                    {
                                        dynamicIsFolder = true;
                                        targetExists = true;
                                    }
                                }
                            }
                            else
                            {
                                targetExists = dynamicIsFolder ? System.IO.Directory.Exists(path) : System.IO.File.Exists(path);
                            }

                            bool isStoreApp = false;
                            string scrapedPath = null;

                            if (!targetExists && isShortcutLocal)
                            {
                                isStoreApp = Utility.IsStoreAppShortcut(path);
                                if (!isStoreApp)
                                {
                                    try { scrapedPath = GetScrapedNetworkPath(path); } catch { }
                                    if (!string.IsNullOrEmpty(scrapedPath))
                                    {
                                        targetExists = true;
                                    }
                                }
                            }


                            if (!targetExists && !isStoreApp) return;
                            // NEEDS TO BE TESTED AND VERIFIED
                            // Safely send the launch command back to the main UI thread once resolved
                            System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                // --- THE SMART WORKAROUND: SELF-HEALING UI ---
                                // If the user clicks a dead icon and it successfully resolves, 
                                // we force the UI to instantly refresh the icon image to the "Alive" state!
                                try { UpdateIcon(sp, path, dynamicIsFolder, resolvedPath); } catch { }

                                LaunchItem(sp, path, dynamicIsFolder, arguments);
                            }));
                        }
                        catch (Exception ex)
                        {
                            // FOR SAFETY REASONS WE KEEP IN COMMENT PREVIOUS APPROACH 

                            //    if (!targetExists && !isStoreApp) return;

                            //    // Safely send the launch command back to the main UI thread once resolved
                            //    System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                            //    {
                            //        LaunchItem(sp, path, dynamicIsFolder, arguments);
                            //    }));
                            //}
                            //catch (Exception ex)
                            //{
                         
                            
                            
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error checking target existence: {ex.Message}");
                        }
                    });

                    launchThread.SetApartmentState(System.Threading.ApartmentState.STA);
                    launchThread.IsBackground = true;
                    launchThread.Start();
                }
            }

            void MouseMoveHandler(object sender, MouseEventArgs e)
            {
                if (IconDragDropManager.IsDragging)
                {
                    try
                    {
                        System.Windows.Point screenPosition = sp.PointToScreen(e.GetPosition(sp));
                        IconDragDropManager.HandleDragMove(screenPosition);
                    }
                    catch { }
                }
            }

            void MouseUpHandler(object sender, MouseButtonEventArgs e)
            {
                if (IconDragDropManager.IsDragging)
                {
                    try
                    {
                        WrapPanel wrapPanel = FindVisualParent<WrapPanel>(sp);
                        if (wrapPanel != null)
                        {
                            System.Windows.Point finalPosition = e.GetPosition(wrapPanel);
                            IconDragDropManager.CompleteDrag(finalPosition);
                        }
                        else IconDragDropManager.CancelDrag();
                        e.Handled = true;
                    }
                    catch
                    {
                        IconDragDropManager.CancelDrag();
                    }
                }
            }

            void KeyUpHandler(object sender, KeyEventArgs e)
            {
                if (IconDragDropManager.IsDragging &&
                    (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl) &&
                    !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
                {
                    IconDragDropManager.CancelDrag();
                    e.Handled = true;
                }
            }

            // --- SAFELY REMOVE PREVIOUS HANDLERS (WPF APPROACH) ---
            sp.RemoveHandler(UIElement.MouseLeftButtonDownEvent, new MouseButtonEventHandler(MouseDownHandler));
            sp.RemoveHandler(UIElement.MouseMoveEvent, new MouseEventHandler(MouseMoveHandler));
            sp.RemoveHandler(UIElement.MouseLeftButtonUpEvent, new MouseButtonEventHandler(MouseUpHandler));
            sp.RemoveHandler(UIElement.KeyUpEvent, new KeyEventHandler(KeyUpHandler));

            // --- ATTACH FRESH HANDLERS ---
            sp.MouseLeftButtonDown += MouseDownHandler;
            sp.MouseMove += MouseMoveHandler;
            sp.MouseLeftButtonUp += MouseUpHandler;
            sp.KeyUp += KeyUpHandler;
        }


        // --- NEW: Unified Manual Hide (Used by Tray Icon) ---
        public static void ForceHideFrames()
        {
            _areFramesAutoHidden = true;

            // --- NEW: Hide Native Desktop Icons if synced ---
            if (SettingsManager.HideDesktopElementsOnAllFramesHide)
            {
                DesktopIconManager.SetDesktopIconsVisible(false);
            }

            var frames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
            foreach (var frame in frames)
            {
                // --- ANTI-FLICKER SAFEGUARD ---
                // Skip frames that are already hidden or currently invisible
                if (frame.Visibility != Visibility.Visible || frame.Opacity == 0.0) continue;

                // Animate the fade out
                var fadeOut = new System.Windows.Media.Animation.DoubleAnimation
                {
                    From = 1.0,
                    To = 0.0,
                    Duration = TimeSpan.FromMilliseconds(250)
                };

                // When the animation finishes, physically hide the window
                fadeOut.Completed += (s, e) =>
                {
                    // --- DWM CACHE FIX ---
                    // DO NOT strip the animation here! Let WPF's FillBehavior.HoldEnd keep it firmly at 0.0.
                    // Stripping the animation here is what causes Windows DWM to snapshot the window at 1.0 opacity.
                    frame.Visibility = Visibility.Hidden;
                };

                frame.BeginAnimation(UIElement.OpacityProperty, fadeOut);
            }

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Frames faded out and hidden.");
        }






        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return parent as T;
        }


        public static void LaunchItem(StackPanel sp, string path, bool isFolder, string arguments = null)
        {
            try
            {
                // --- NEW: Ignore clicks on internal placeholders ---
                if (path != null && path.StartsWith("INTERNAL_BLANK_")) return;
                // 1. Visual Feedback
                NonActivatingWindow win = FindVisualParent<NonActivatingWindow>(sp);

                // FIX: Use ID (Tag) lookup instead of Title (Robust)
                string frameId = win?.Tag?.ToString();
                dynamic frame = null;

                if (!string.IsNullOrEmpty(frameId))
                {
                    frame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                }

                // Fallback to Title only if ID lookup failed
                if (frame == null)
                {
                    frame = FrameDataManager.FrameData.FirstOrDefault(f => f.Title == win?.Title);
                }

                // FIX: Read LIVE Global Setting directly from SettingsManager
                // This bypasses the stale _options cache.
                LaunchEffectsManager.LaunchEffect effect = SettingsManager.LaunchEffect;

                // Check for frame-Specific Override
                string customEffect = frame?.CustomLaunchEffect?.ToString();
                if (!string.IsNullOrEmpty(customEffect))
                {
                    try
                    {
                        effect = (LaunchEffectsManager.LaunchEffect)Enum.Parse(typeof(LaunchEffectsManager.LaunchEffect), customEffect, true);
                    }
                    catch { }
                }

                LaunchEffectsManager.ExecuteLaunchEffect(sp, effect);

                // 2. Path Resolution
                string fullPath = path;
                try { fullPath = System.IO.Path.GetFullPath(path); } catch { }

                string extension = System.IO.Path.GetExtension(fullPath).ToLower();
                bool isUrlFile = extension == ".url";
                bool isLnk = extension == ".lnk";

                string targetPath = fullPath;
                string workingDirectory = "";
                string finalArguments = arguments ?? "";

                // 3. Resolve Target
                if (isUrlFile)
                {
                    targetPath = ExtractUrlFromUrlFile(fullPath);
                    if (string.IsNullOrEmpty(targetPath))
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgInvalidUrlFile", path), Strings.DlgError);
                        return;
                    }
                }
                else if (isLnk)
                {
                    string resolved = FilePathUtilities.GetShortcutTargetUnicodeSafe(fullPath);
                    if (string.IsNullOrEmpty(resolved)) resolved = GetScrapedNetworkPath(fullPath);

                    if (!string.IsNullOrEmpty(resolved))
                    {
                        targetPath = resolved;
                        try
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(fullPath);
                            string rawTarget = shortcut.TargetPath?.ToLower() ?? "";
                            if (!rawTarget.EndsWith("explorer.exe"))
                            {
                                if (string.IsNullOrEmpty(finalArguments)) finalArguments = shortcut.Arguments;
                                workingDirectory = shortcut.WorkingDirectory;
                            }
                        }
                        catch { }
                    }
                }

                // 4. Classification
                bool isWebUrl = IsWebUrl(targetPath);
                bool isStoreApp = isLnk && Utility.IsStoreAppShortcut(fullPath);
                bool isNetworkRoot = !string.IsNullOrEmpty(targetPath) && targetPath.StartsWith(@"\\") && !targetPath.Contains(@":\");
                bool isSpecialPath = IsSpecialWindowsPath(targetPath);

                bool isTargetFolder = false;
                bool targetExists = false;

                if (!isWebUrl && !isSpecialPath && !isNetworkRoot && !isStoreApp)
                {
                    isTargetFolder = System.IO.Directory.Exists(targetPath);
                    targetExists = isTargetFolder || System.IO.File.Exists(targetPath);
                }

                // 5. Validation Guard
                if (!targetExists && !isWebUrl && !isSpecialPath && !isStoreApp && !isNetworkRoot)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Target not found: {targetPath}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgTargetNotFound", targetPath), Strings.DlgLaunchError);
                    return;
                }

                // 6. Admin & Different User Logic
                bool alwaysRunAsAdmin = false;
                bool alwaysRunAsDifferentUser = false;
                try
                {
                    if (frame != null && frame.ItemsType?.ToString() == "Data")
                    {
                        var items = frame.Items as JArray ?? new JArray();
                        bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";

                        if (tabsEnabled)
                        {
                            var tabs = frame.Tabs as JArray ?? new JArray();
                            int currentTabIndex = Convert.ToInt32(frame.CurrentTab?.ToString() ?? "0");
                            if (currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                            {
                                var currentTab = tabs[currentTabIndex] as JObject;
                                items = currentTab?["Items"] as JArray ?? items;
                            }
                        }

                        var matchingItem = items.FirstOrDefault(i => string.Equals(
                            System.IO.Path.GetFullPath(i["Filename"]?.ToString() ?? ""),
                            fullPath, StringComparison.OrdinalIgnoreCase));

                        if (matchingItem != null)
                        {
                            alwaysRunAsAdmin = Convert.ToBoolean(matchingItem["AlwaysRunAsAdmin"] ?? false);
                            alwaysRunAsDifferentUser = Convert.ToBoolean(matchingItem["AlwaysRunAsDifferentUser"] ?? false);
                        }
                    }
                }
                catch { }

                // 7. Execution
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.UseShellExecute = true;

                if (!isWebUrl && !isSpecialPath && !isTargetFolder && !isNetworkRoot)
                {
                    if (alwaysRunAsAdmin)
                        psi.Verb = "runas";
                    else if (alwaysRunAsDifferentUser)
                        psi.Verb = "runasuser";
                }

                if (isStoreApp)
                {
                    psi.FileName = "explorer.exe";
                    psi.Arguments = $"\"{fullPath}\"";
                    psi.WorkingDirectory = "";
                }
                else if (isWebUrl)
                {
                    psi.FileName = targetPath;
                }
                else if (isTargetFolder || isNetworkRoot)
                {
                    psi.FileName = "explorer.exe";
                    psi.Arguments = $"\"{targetPath}\"";
                }
                else
                {
                    psi.FileName = targetPath;
                    if (!string.IsNullOrEmpty(finalArguments)) psi.Arguments = finalArguments;

                    if (!string.IsNullOrEmpty(workingDirectory) && System.IO.Directory.Exists(workingDirectory))
                        psi.WorkingDirectory = workingDirectory;
                    else
                    {
                        string dir = System.IO.Path.GetDirectoryName(targetPath);
                        if (!string.IsNullOrEmpty(dir) && System.IO.Directory.Exists(dir))
                            psi.WorkingDirectory = dir;
                    }
                }

                Process.Start(psi);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Launched: {targetPath}");
            }
            catch (Exception ex)
            {
                if (!ex.Message.Contains("canceled"))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Launch Error: {ex.Message}");
                    if (!SettingsManager.SuppressLaunchWarnings)
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.Get("MsgLaunchFailed", ex.Message), Strings.DlgLaunchError);
                }
            }
        }



        //// Determines if a path is a web URL
        //private static bool IsWebUrl(string path)
        //{
        //    if (string.IsNullOrEmpty(path)) return false;

        //    return path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
        //           path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
        //           path.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase) ||
        //           path.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) ||
        //   path.StartsWith("steam://", StringComparison.OrdinalIgnoreCase);
        //}

        // Determines if a path is a web URL
        private static bool IsWebUrl(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;

            if (Uri.TryCreate(path, UriKind.Absolute, out Uri uriResult))
            {
                string scheme = uriResult.Scheme.ToLower();
                return scheme == Uri.UriSchemeHttp ||
                       scheme == Uri.UriSchemeHttps ||
                       scheme == Uri.UriSchemeFtp ||
                       scheme == Uri.UriSchemeMailto;
            }
            return false;
        }


        //// Determines if a path is a special Windows path
        //private static bool IsSpecialWindowsPath(string path)
        //{
        //    if (string.IsNullOrEmpty(path)) return false;

        //    return path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) ||
        //           path.StartsWith("ms-settings:", StringComparison.OrdinalIgnoreCase) ||
        //           path.StartsWith("ms-", StringComparison.OrdinalIgnoreCase) ||
        //           path.Contains("::") ||
        //           path.StartsWith("control", StringComparison.OrdinalIgnoreCase);
        //}


        // Determines if a path is a special Windows path
        private static bool IsSpecialWindowsPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;

            // Support Custom App Protocols globally (steam://, spotify:, discord:, etc.)
            if (Uri.TryCreate(path, UriKind.Absolute, out Uri uriResult))
            {
                string scheme = uriResult.Scheme.ToLower();
                if (scheme != Uri.UriSchemeFile &&
                    scheme != Uri.UriSchemeHttp &&
                    scheme != Uri.UriSchemeHttps &&
                    scheme != Uri.UriSchemeFtp &&
                    scheme != Uri.UriSchemeMailto)
                {
                    return true;
                }
            }

            return path.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("ms-settings:", StringComparison.OrdinalIgnoreCase) ||
                   path.StartsWith("ms-", StringComparison.OrdinalIgnoreCase) ||
                   path.Contains("::") ||
                   path.StartsWith("control", StringComparison.OrdinalIgnoreCase);
        }


        // Extracts the URL from a .url file
        private static string ExtractUrlFromUrlFile(string urlFilePath)
        {
            try
            {
                if (!System.IO.File.Exists(urlFilePath))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"URL file not found: {urlFilePath}");
                    return null;
                }

                string[] lines = System.IO.File.ReadAllLines(urlFilePath);
                foreach (string line in lines)
                {
                    if (line.StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                    {
                        string url = line.Substring(4).Trim(); // Remove "URL=" prefix
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Extracted URL from {urlFilePath}: {url}");
                        return url;
                    }
                }

                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"No URL= line found in .url file: {urlFilePath}");
                return null;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error reading .url file {urlFilePath}: {ex.Message}");
                return null;
            }
        }

        // Helper method for network path detection
        // Helper method for network path detection
        public static bool IsNetworkPath(string filePath)
        {
            try
            {
                bool isShortcut = System.IO.Path.GetExtension(filePath).ToLower() == ".lnk";

                if (isShortcut)
                {
                    // --- HANG FIX: Bypassed Utility.GetShortcutTarget(filePath) ---
                    // We use a fast binary scan to detect UNC paths without triggering Windows Network timeouts.
                    if (System.IO.File.Exists(filePath))
                    {
                        byte[] fileContent = System.IO.File.ReadAllBytes(filePath);
                        string asciiContent = System.Text.Encoding.ASCII.GetString(fileContent);
                        string unicodeContent = System.Text.Encoding.Unicode.GetString(fileContent);

                        bool isUncPath = asciiContent.Contains("\\\\") || unicodeContent.Contains("\\\\");

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Shortcut {filePath} binary scan for UNC, IsUNC: {isUncPath}");
                        return isUncPath;
                    }
                }
                else
                {
                    // For direct paths, check if it's UNC
                    bool isUncPath = filePath.StartsWith("\\\\");
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling, $"Direct path {filePath}, IsUNC: {isUncPath}");
                    return isUncPath;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.IconHandling, $"Error checking if {filePath} is network path: {ex.Message}");
            }
            return false;
        }

        // Public method to refresh an icon's click handlers after shortcut editing
        // Called by EditShortcutWindow to ensure immediate argument updates
        public static void RefreshIconClickHandlers(string shortcutPath, string newDisplayName)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Refreshing click handlers for edited shortcut: {shortcutPath}");

                // Find all frame windows and locate the icon
                var windows = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>();
                bool iconFound = false;

                foreach (var window in windows)
                {
                    // Find the WrapPanel containing icons
                    var wrapPanel = FrameUtilities.FindWrapPanel(window);

                    if (wrapPanel == null) continue;

                    // Find the specific icon StackPanel
                    foreach (StackPanel iconPanel in wrapPanel.Children.OfType<StackPanel>())
                    {
                        var tagData = iconPanel.Tag;
                        if (tagData != null)
                        {
                            string filePath = tagData.GetType().GetProperty("FilePath")?.GetValue(tagData)?.ToString();
                            if (!string.IsNullOrEmpty(filePath) &&
                                string.Equals(System.IO.Path.GetFullPath(filePath), System.IO.Path.GetFullPath(shortcutPath), StringComparison.OrdinalIgnoreCase))
                            {
                                // Found it! Refresh this icon completely
                                RefreshSingleIconComplete(iconPanel, shortcutPath, newDisplayName, window);
                                iconFound = true;
                                break;
                            }
                        }
                    }
                    if (iconFound) break;
                }

                if (!iconFound)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                        $"Could not find icon to refresh for: {shortcutPath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error refreshing icon click handlers: {ex.Message}");
            }
        }

        // Completely refreshes a single icon with fresh data from the .lnk fil
        private static void RefreshSingleIconComplete(StackPanel iconPanel, string shortcutPath, string newDisplayName, NonActivatingWindow parentWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Performing complete refresh of icon: {shortcutPath}");

                // 1. Read fresh data
                string freshTargetPath = shortcutPath;
                string freshArguments = "";
                bool isFolder = false;
                string workingDirectory = "";

                // Identify Type
                bool isUrl = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url";
                bool isLnk = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".lnk";

                if (isLnk)
                {
                    try
                    {
                        WshShell shell = new WshShell();
                        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                        freshTargetPath = shortcut.TargetPath ?? shortcutPath;
                        freshArguments = shortcut.Arguments ?? "";
                        workingDirectory = shortcut.WorkingDirectory ?? "";
                        isFolder = System.IO.Directory.Exists(freshTargetPath);
                    }
                    catch { }
                }
                else if (isUrl)
                {
                    // For URLs, we can try to extract the clean URL, but IsFolder is always false
                    string url = CoreUtilities.ExtractWebUrlFromFile(shortcutPath);
                    if (!string.IsNullOrEmpty(url)) freshTargetPath = url;
                }

                // 2. Update Tag
                iconPanel.Tag = new
                {
                    FilePath = shortcutPath,
                    IsFolder = isFolder,
                    Arguments = freshArguments
                };

                // 3. Update Text
                if (!string.IsNullOrEmpty(newDisplayName))
                {
                    var textBlock = iconPanel.Children.OfType<TextBlock>().FirstOrDefault();
                    if (textBlock != null)
                    {
                        string displayText = newDisplayName.Length > SettingsManager.MaxDisplayNameLength
                            ? newDisplayName.Substring(0, SettingsManager.MaxDisplayNameLength) + "..."
                            : newDisplayName;
                        textBlock.Text = displayText;
                    }
                }

                // 4. Update ToolTip
                string displayTarget = freshTargetPath;

                // Clean up MS Store App targets for the tooltip
                if (!string.IsNullOrEmpty(displayTarget) && displayTarget.Contains("!") && !displayTarget.Contains(":\\"))
                {
                    string packageId = displayTarget.Split('!')[0];
                    int hashIndex = packageId.IndexOf('_');
                    if (hashIndex > 0) packageId = packageId.Substring(0, hashIndex);
                    displayTarget = $"Windows App ({packageId})";
                }

                string toolTipText = $"{newDisplayName ?? System.IO.Path.GetFileNameWithoutExtension(shortcutPath)}\nTarget: {displayTarget}";
                if (!string.IsNullOrEmpty(freshArguments)) toolTipText += $"\nArguments: {freshArguments}";
                iconPanel.ToolTip = new ToolTip { Content = toolTipText };

                // 5. UPDATE ICON IMAGE (The Critical Fix)
                var ico = iconPanel.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                if (ico != null)
                {
                    ImageSource newIcon = null;

                    // A. Try Manual Extraction for .url (Fixes Runtime Update)
                    if (isUrl)
                    {
                        var urlIcon = GetUrlCustomIcon(shortcutPath);
                        if (urlIcon.Path != null)
                        {
                            newIcon = IconManager.ExtractIconFromFile(urlIcon.Path, urlIcon.Index);
                        }
                    }
                    // B. Try WshShell for .lnk
                    else if (isLnk)
                    {
                        try
                        {
                            WshShell shell = new WshShell();
                            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                            if (!string.IsNullOrEmpty(shortcut.IconLocation) && shortcut.IconLocation != ",0")
                            {
                                string[] parts = shortcut.IconLocation.Split(',');
                                if (System.IO.File.Exists(parts[0]))
                                {
                                    int idx = 0;
                                    if (parts.Length > 1) int.TryParse(parts[1], out idx);
                                    newIcon = IconManager.ExtractIconFromFile(parts[0], idx);
                                }
                            }
                        }
                        catch { }
                    }

                    // C. Fallback to Theme/Shell Icon
                    if (newIcon == null)
                    {
                        if (isUrl)
                        {
                            newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                        }
                        else if (isFolder)
                        {
                            // Check for broken target
                            if (!FilePathUtilities.DoesFolderExist(shortcutPath, true))
                                newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-WhiteX.png"));
                            else
                                newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                        }
                        else
                        {
                            // Standard file fallback
                            newIcon = Utility.GetShellIcon(freshTargetPath, false);
                            if (newIcon == null) newIcon = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                        }
                    }

                    // Apply and Update Cache
                    if (newIcon != null)
                    {
                        ico.Source = newIcon;
                        IconManager.IconCache[shortcutPath] = newIcon; // Force update cache
                    }
                }

                // 6. Re-attach Events and Update Data
                ClickEventAdder(iconPanel, shortcutPath, isFolder, freshArguments);
                UpdateFrameDataForIcon(shortcutPath, newDisplayName, parentWindow);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Complete icon refresh successful for: {shortcutPath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error in complete icon refresh: {ex.Message}");
            }
        }



        // Updates frame JSON data after icon refresh
        private static void UpdateFrameDataForIcon(string shortcutPath, string newDisplayName, NonActivatingWindow parentWindow)
        {
            try
            {
                if (string.IsNullOrEmpty(newDisplayName)) return;

                string frameId = parentWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(frameId)) return;

                var frame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (frame == null) return;

                // FIX: Search for the item in BOTH Main Items and ALL Tabs
                bool found = false;

                // Helper function to search and update a list
                bool SearchList(JArray list)
                {
                    if (list == null) return false;
                    foreach (var item in list)
                    {
                        string itemFilename = item["Filename"]?.ToString();
                        if (!string.IsNullOrEmpty(itemFilename) &&
                            string.Equals(System.IO.Path.GetFullPath(itemFilename), System.IO.Path.GetFullPath(shortcutPath), StringComparison.OrdinalIgnoreCase))
                        {
                            item["DisplayName"] = newDisplayName;
                            return true; // Found and updated
                        }
                    }
                    return false;
                }

                // 1. Check Main Items
                if (SearchList(frame.Items as JArray)) found = true;

                // 2. Check Tabs (Always check tabs to keep Tab 0 and Main Items completely synced)
                var tabs = frame.Tabs as JArray;
                if (tabs != null)
                {
                    foreach (var tab in tabs)
                    {
                        if (SearchList(tab["Items"] as JArray))
                        {
                            found = true;
                            // We purposefully do NOT break here to ensure all instances across all tabs are updated
                        }
                    }
                }

                if (found)
                {
                    FrameDataManager.SaveFrameData();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Updated JSON DisplayName for: {shortcutPath}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not find item in JSON to update name: {shortcutPath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error updating frame JSON: {ex.Message}");
            }
        }







        // Size feedback during resizing
        private static void ShowSizeFeedback(double width, double height)
        {
            if (_sizeFeedbackWindow == null)
            {
                _sizeFeedbackWindow = new Window
                {
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = System.Windows.Media.Brushes.Transparent,
                    Width = 100,
                    Height = 30,
                    ShowInTaskbar = false,
                    Topmost = true
                };

                var label = new Label
                {
                    Content = "",
                    Foreground = System.Windows.Media.Brushes.White,
                    Background = System.Windows.Media.Brushes.Black,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                _sizeFeedbackWindow.Content = label;
            }

            var labelContent = (Label)_sizeFeedbackWindow.Content;
            labelContent.Content = $"{Math.Round(width)} x {Math.Round(height)}";

            var mousePos = System.Windows.Forms.Cursor.Position;
            _sizeFeedbackWindow.Left = mousePos.X + 10;
            _sizeFeedbackWindow.Top = mousePos.Y + 10;

            _sizeFeedbackWindow.Show();

            // --- BUG FIX: Unified Debounce Timer ---
            // This guarantees the indicator will ALWAYS disappear 1.5 seconds 
            // after the last call, preventing orphaned windows on screen.
            if (_hideTimer == null)
            {
                _hideTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(1500)
                };
                _hideTimer.Tick += (s, e) => HideSizeFeedback();
            }
            _hideTimer.Stop();
            _hideTimer.Start();
        }

        private static void HideSizeFeedback()
        {
            if (_sizeFeedbackWindow != null)
            {
                _sizeFeedbackWindow.Hide();
            }
            if (_hideTimer != null)
            {
                _hideTimer.Stop();
            }
        }

        public static void OnResizingStarted(NonActivatingWindow frame)
        {
            if (SettingsManager.EnableDimensionSnap)
            {
                frame.SizeChanged += UpdateSizeFeedback;
                ShowSizeFeedback(frame.Width, frame.Height);
            }
        }

        public static void OnResizingEnded(NonActivatingWindow frame)
        {
            if (SettingsManager.EnableDimensionSnap)
            {
                frame.SizeChanged -= UpdateSizeFeedback;

                double snappedWidth = Math.Round(frame.Width / 10.0) * 10;
                double snappedHeight = Math.Round(frame.Height / 10.0) * 10;

                frame.Width = snappedWidth;
                frame.Height = snappedHeight;

                dynamic FrameData = GetFrameData().FirstOrDefault(f => f.Title == frame.Title);
                if (FrameData != null)
                {
                    FrameData.Width = snappedWidth;
                    FrameData.Height = snappedHeight;
                    FrameDataManager.SaveFrameData();
                }

                // Show one last time. The unified timer in ShowSizeFeedback will clean it up automatically.
                ShowSizeFeedback(snappedWidth, snappedHeight);
            }
        }





        // --- HELPER: Efficient Dead Shortcut Check using Cached States ---
        private static bool HasDeadShortcuts(dynamic frame)
        {
            try
            {
                // --- BUG FIX: Fetch Live frame ---
                // The passed 'frame' object may be an orphaned JSON node if Tabs were recently modified.
                string frameId = frame.Id?.ToString();
                var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId) ?? frame;

                // Only Data frames track dead shortcuts
                if (liveFrame.ItemsType?.ToString() != "Data") return false;

                bool tabsEnabled = liveFrame.TabsEnabled?.ToString().ToLower() == "true";
                if (tabsEnabled)
                {
                    var tabs = liveFrame.Tabs as JArray;
                    if (tabs != null)
                    {
                        foreach (var tab in tabs)
                        {
                            var items = tab["Items"] as JArray;
                            if (items != null && CheckItemsForDead(items)) return true;
                        }
                    }
                }
                else
                {
                    var items = liveFrame.Items as JArray;
                    if (items != null && CheckItemsForDead(items)) return true;
                }
            }
            catch { }
            return false;
        }

        private static bool CheckItemsForDead(JArray items)
        {
            foreach (var item in items)
            {
                string path = item["Filename"]?.ToString();
                if (!string.IsNullOrEmpty(path))
                {
                    // Check our cache which is updated by the background TargetChecker
                    // This avoids scanning the disk on the UI thread
                    if (_iconStates.TryGetValue(path, out var state))
                    {
                        if (state.IsBroken) return true;
                    }
                }
            }
            return false;
        }










        // Add this static method inside your Framemanager class
        private static IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_GETMINMAXINFO = 0x0024;

            if (msg == WM_GETMINMAXINFO)
            {
                // Get the screen information for the current monitor
                var screen = System.Windows.Forms.Screen.FromHandle(hwnd);
                if (screen != null)
                {
                    // Define the limit: 90% of the working area
                    int maxWidth = (int)(screen.WorkingArea.Width * 0.90);
                    int maxHeight = (int)(screen.WorkingArea.Height * 0.90);

                    // Marshal the structure
                    MINMAXINFO mmi = (MINMAXINFO)Marshal.PtrToStructure(lParam, typeof(MINMAXINFO));

                    // Force the "Maximized" size to be 90%, not 100%
                    mmi.ptMaxSize.x = maxWidth;
                    mmi.ptMaxSize.y = maxHeight;

                    // Force the user-draggable limit to be 90%
                    mmi.ptMaxTrackSize.x = maxWidth;
                    mmi.ptMaxTrackSize.y = maxHeight;

                    Marshal.StructureToPtr(mmi, lParam, true);
                    handled = true;
                }
            }
            return IntPtr.Zero;
        }



        private static string GetSafeProperty(dynamic obj, string propName)
        {
            try
            {
                if (obj is JObject jObj && jObj[propName] != null) return jObj[propName].ToString();
                return obj.GetType().GetProperty(propName)?.GetValue(obj, null)?.ToString() ?? "";
            }
            catch { return ""; }
        }



        private static void UpdateSizeFeedback(object sender, SizeChangedEventArgs e)
        {
            var frame = sender as NonActivatingWindow;
            if (frame != null)
            {
                // --- BUG FIX: Ignore programmatic animations ---
                // Do not show the resizing indicator if the frame is just rolling up or down
                string frameId = frame.Tag?.ToString();
                if (!string.IsNullOrEmpty(frameId) && _framesInTransition.Contains(frameId)) return;

                ShowSizeFeedback(frame.Width, frame.Height);
            }
        }


    }

}
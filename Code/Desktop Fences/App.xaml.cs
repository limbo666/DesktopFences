using DesktopFences;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace Desktop_Fences
{
    public partial class App : Application
    {
        private TrayManager _trayManager;
        private TargetChecker _targetChecker;
        private static bool _desktopIsShown = false;
        private static Mutex _mutex;
        private const string UNIQUE_APP_NAME = "Global\\DesktopFences_Mutex_UniqueId_v2";



        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // --- 1. INITIALIZE PROFILES & SETTINGS FIRST ---
            // This ensures we know the user's true DisableSingleInstance preference immediately
            try
            {
                ProfileManager.Initialize();
                System.IO.Directory.SetCurrentDirectory(ProfileManager.CurrentProfileDir);
                SettingsManager.LoadSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Profile Initialization Error: {ex.Message}", "Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
                return;
            }

            // --- 2. SINGLE INSTANCE PROTECTION START ---
            bool isNewInstance;
            _mutex = new Mutex(true, UNIQUE_APP_NAME, out isNewInstance);

            // Only exit if it's not a new instance AND the user hasn't explicitly disabled the check
            if (!isNewInstance && !SettingsManager.DisableSingleInstance)
            {
                // --- DEBUGGING GHOSTS START ---
                try
                {
                    string debugLog = $"[{DateTime.Now}] Instance 2 Started.\n";
                    debugLog += $"Args (e.Args): {string.Join(" | ", e.Args)}\n";
                    debugLog += $"Args (Environment): {string.Join(" | ", Environment.GetCommandLineArgs())}\n";

                    bool isDrawCommand = e.Args.Any(arg => arg.IndexOf("-create", StringComparison.OrdinalIgnoreCase) >= 0)
                                         || Environment.GetCommandLineArgs().Any(arg => arg.IndexOf("-create", StringComparison.OrdinalIgnoreCase) >= 0);

                    if (isDrawCommand)
                    {
                        RegistryHelper.WriteTrigger($"CMD_DRAW|{Guid.NewGuid()}");
                    }
                    else
                    {
                        RegistryHelper.WriteTrigger(null);
                    }
                }
                catch { }
                // --- DEBUGGING GHOSTS END ---

                Shutdown();
                return;
            }
            // --- SINGLE INSTANCE PROTECTION END ---

            try
            {
                // --- NEW: Sanitize Registry on Startup ---
                RegistryHelper.DeleteTrigger();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    $"Startup: Working Directory set to {ProfileManager.CurrentProfileDir}");

                // 3. Continue Normal Startup
                {
                    // Initialize settings (Now loads from Profile/options.json)
                    SettingsManager.LoadSettings();

                    // --- NEW: Self-Heal Context Menu Path ---
                    // Ensures the registry key points to the current EXE location
                    RegistryHelper.RefreshContextMenuPath();

                    // Initialize InterCore system
                    InterCore.Initialize();

                    // --- CHAMELEON ENGINE ---
                    WallpaperColorManager.Initialize();
                    // --- AUTO-ORGANIZE ENGINE ---
                    AutoOrganizeManager.Initialize();
                    WallpaperColorManager.WallpaperColorChanged += (s, ev) =>
                    {
                        // Magically update all visuals live when Windows wallpaper changes
                        System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            Utility.UpdateFenceVisuals();
                        }));
                    };

                    // Initialize TrayManager BEFORE fences
                    _trayManager = new TrayManager();
                    _trayManager.InitializeTray();

         

                    // Initialize TargetChecker
                    _targetChecker = new TargetChecker(1000);
                    _targetChecker.Start();

                    // Start the Background Icon Loader Engine
                    LazyIconLoader.Start();

                    // Load fences (Now loads from Profile/fences.json)
                    FenceManager.LoadAndCreateFences(_targetChecker);

                    // --- PRODUCTION START LOGIC ---
                    if (SettingsManager.EnableProfileAutomation)
                    {
                        AutomationManager.Start();
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "Startup: Profile Automation Engine ignited.");
                    }

                    // Ensure UI reflects the current state of profiles and automation
                    _trayManager.UpdateProfilesMenu();
                    _trayManager.UpdateTrayIcon();
                    _trayManager.UpdateHiddenFencesMenu();

                    // Initialize global hotkey monitoring
                    try
                    {
                        GlobalHotkeyManager.WindowsPlusDDetected += OnWindowsPlusDDetected;
                        GlobalHotkeyManager.StartMonitoring();
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                            "GlobalHotkeyManager: Successfully initialized hotkey monitoring");
                    }
                    catch (System.Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                            $"GlobalHotkeyManager: Failed to initialize: {ex.Message}");
                    }



                    // --- NEW: Start Desktop Double-Click Listener ---
                    try
                    {
                        DesktopMouseHook.Start();
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "DesktopMouseHook started.");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to start DesktopMouseHook: {ex.Message}");
                    }



                    // --- NEW: Direct Draw Mode Check ---
                    // If this MAIN instance was started via Context Menu, trigger draw mode now.
                    // Use the same robust check as above.
                    var allArgs = Environment.GetCommandLineArgs();
                    bool isDrawStartup = e.Args.Any(arg => arg.IndexOf("-create", StringComparison.OrdinalIgnoreCase) >= 0)
                                         || allArgs.Any(arg => arg.IndexOf("-create", StringComparison.OrdinalIgnoreCase) >= 0);

                    if (isDrawStartup)
                    {
                        // Wait 500ms for UI to settle, then draw
                        Task.Delay(500).ContinueWith(t => Dispatcher.Invoke(() => FenceManager.StartDrawMode()));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Critical Startup Error: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            InterCore.Cleanup();
            try
            {
                GlobalHotkeyManager.StopMonitoring();
            }
            catch { }

            // --- NEW: Stop Desktop Double-Click Listener ---
            try
            {
                DesktopMouseHook.Stop();
            }
            catch { }

            _trayManager?.Dispose();
            base.OnExit(e);
        }

        private static void OnWindowsPlusDDetected(object sender, System.EventArgs e)
        {
            try
            {
                _desktopIsShown = !_desktopIsShown;
                if (_desktopIsShown)
                {
                    var restoreTimer = new System.Windows.Threading.DispatcherTimer
                    {
                        Interval = TimeSpan.FromMilliseconds(800)
                    };
                    restoreTimer.Tick += (timerSender, timerArgs) =>
                    {
                        restoreTimer.Stop();
                        RestoreAllFenceWindows();
                    };
                    restoreTimer.Start();
                }

                var resetTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(10)
                };
                resetTimer.Tick += (timerSender, timerArgs) =>
                {
                    resetTimer.Stop();
                    _desktopIsShown = false;
                };
                resetTimer.Start();
            }
            catch { }
        }

        private static void RestoreAllFenceWindows()
        {
            try
            {
                var fenceWindows = System.Windows.Application.Current.Windows
                    .OfType<NonActivatingWindow>()
                    .ToList();

                foreach (var fenceWindow in fenceWindows)
                {
                    try
                    {
                        if (fenceWindow.WindowState == WindowState.Minimized)
                            fenceWindow.WindowState = WindowState.Normal;

                        if (!fenceWindow.IsVisible)
                            fenceWindow.Show();

                        fenceWindow.Topmost = true;
                        fenceWindow.Topmost = false;
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
}
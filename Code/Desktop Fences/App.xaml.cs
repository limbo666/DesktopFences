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
            // --- 1. SINGLE INSTANCE PROTECTION START ---
            bool isNewInstance;
            _mutex = new Mutex(true, UNIQUE_APP_NAME, out isNewInstance);

            if (!isNewInstance)
            {
                // --- DEBUGGING GHOSTS START ---
                try
                {
                    string debugLog = $"[{DateTime.Now}] Instance 2 Started.\n";
                    debugLog += $"Args (e.Args): {string.Join(" | ", e.Args)}\n";
                    debugLog += $"Args (Environment): {string.Join(" | ", Environment.GetCommandLineArgs())}\n";

                    // Robust Check
                    bool isDrawCommand = e.Args.Any(arg => arg.IndexOf("-create", StringComparison.OrdinalIgnoreCase) >= 0)
                                         || Environment.GetCommandLineArgs().Any(arg => arg.IndexOf("-create", StringComparison.OrdinalIgnoreCase) >= 0);

                    debugLog += $"Detected -create? {isDrawCommand}\n";

                    if (isDrawCommand)
                    {
                        string cmd = $"CMD_DRAW|{Guid.NewGuid()}";
                        RegistryHelper.WriteTrigger(cmd);
                        debugLog += $"Action: Sent {cmd}";
                    }
                    else
                    {
                        RegistryHelper.WriteTrigger(null);
                        debugLog += "Action: Sent NULL (Wake Up)";
                    }

                }
                catch { }
                // --- DEBUGGING GHOSTS END ---

                // Close this second instance immediately
                Shutdown();
                return;
            }
            // --- SINGLE INSTANCE PROTECTION END ---

            try
            {
                // 1. PHASE 1: INITIALIZE PROFILE SYSTEM (CRITICAL)
                ProfileManager.Initialize();

                // --- NEW: Sanitize Registry on Startup ---
                // Prevents "Ghost" commands from previous sessions triggering automatically
                RegistryHelper.DeleteTrigger();
                // -----------------------------------------

                // 2. SET WORKING DIRECTORY TO PROFILE
                System.IO.Directory.SetCurrentDirectory(ProfileManager.CurrentProfileDir);

                // Debug Log
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

                    // Initialize TrayManager BEFORE fences
                    _trayManager = new TrayManager();
                    _trayManager.InitializeTray();

                    // Initialize TargetChecker
                    _targetChecker = new TargetChecker(1000);
                    _targetChecker.Start();

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
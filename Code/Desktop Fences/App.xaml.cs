using System;
using System.Linq;
using System.Windows;
namespace Desktop_Fences
{
    public partial class App : Application
    {
        private TrayManager _trayManager;
        private TargetChecker _targetChecker;
        private static bool _desktopIsShown = false; // Add this line




        private void Application_Startup(object sender, StartupEventArgs e)
        {
          
            // Fix Working Directory so Registry Startup finds json files correctly
            System.IO.Directory.SetCurrentDirectory(System.AppDomain.CurrentDomain.BaseDirectory);
           
            {
                // Initialize settings FIRST
                SettingsManager.LoadSettings();

                // NEW: Initialize InterCore system
                InterCore.Initialize();

                // Initialize TrayManager BEFORE fences
                _trayManager = new TrayManager();
                _trayManager.InitializeTray();

                // Initialize TargetChecker
                _targetChecker = new TargetChecker(1000);
                _targetChecker.Start();

                // Load fences (hidden ones will auto-register with TrayManager)
                FenceManager.LoadAndCreateFences(_targetChecker);

                // Force tray icon update after all fences are loaded
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
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // NEW: Cleanup InterCore system
            InterCore.Cleanup();
            // Cleanup global hotkey monitoring
            try
            {
                GlobalHotkeyManager.StopMonitoring();
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "GlobalHotkeyManager: Successfully stopped hotkey monitoring");
            }
            catch (System.Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"GlobalHotkeyManager: Error during cleanup: {ex.Message}");
            }
            _trayManager?.Dispose();
            base.OnExit(e);
        }
       



        private static void OnWindowsPlusDDetected(object sender, System.EventArgs e)
        {
            try
            {
                // Toggle the state
                _desktopIsShown = !_desktopIsShown;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    $"Windows+D detected - desktop shown: {_desktopIsShown}");

                if (_desktopIsShown)
                {
                    // Desktop is shown - restore fences after delay
                    var restoreTimer = new System.Windows.Threading.DispatcherTimer
                    {
                        Interval = TimeSpan.FromMilliseconds(800)
                    };

                    restoreTimer.Tick += (timerSender, timerArgs) =>
                    {
                        try
                        {
                            restoreTimer.Stop();
                            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                                "Restoring fences after Windows+D");
                            RestoreAllFenceWindows();
                        }
                        catch (System.Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                                $"Error restoring fences: {ex.Message}");
                        }
                    };
                    restoreTimer.Start();
                }

                // Auto-reset state after 10 seconds in case of missed detection
                var resetTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(10)
                };
                resetTimer.Tick += (timerSender, timerArgs) =>
                {
                    resetTimer.Stop();
                    _desktopIsShown = false;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                        "Auto-reset Windows+D state");
                };
                resetTimer.Start();
            }
            catch (System.Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"Error handling Windows+D detection: {ex.Message}");
            }
        }
        /// <summary>
        /// Restores visibility of all fence windows after Windows+D
        /// </summary>
        private static void RestoreAllFenceWindows()
        {
            try
            {
                var fenceWindows = System.Windows.Application.Current.Windows
                    .OfType<NonActivatingWindow>()
                    .ToList();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    $"Found {fenceWindows.Count} fence windows to restore");

                foreach (var fenceWindow in fenceWindows)
                {
                    try
                    {
                        // Restore window if it was minimized/hidden
                        if (fenceWindow.WindowState == WindowState.Minimized)
                        {
                            fenceWindow.WindowState = WindowState.Normal;
                        }

                        // Ensure window is visible
                        if (!fenceWindow.IsVisible)
                        {
                            fenceWindow.Show();
                        }

                        // Bring to front (but don't activate to avoid stealing focus)
                        fenceWindow.Topmost = true;
                        fenceWindow.Topmost = false;

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                            $"Restored fence window: {fenceWindow.Tag}");
                    }
                    catch (System.Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                            $"Error restoring individual fence window: {ex.Message}");
                    }
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                    "Completed fence window restoration");
            }
            catch (System.Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"Error in RestoreAllFenceWindows: {ex.Message}");
            }
        }



    }
}
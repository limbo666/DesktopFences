using System;
using System.Windows;

namespace Desktop_Fences
{
    public partial class App : Application
    {
        private TrayManager _trayManager;
        private TargetChecker _targetChecker;

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Initialize settings
            SettingsManager.LoadSettings();

            // Initialize tray icon and context menu
            _trayManager = new TrayManager();
            _trayManager.InitializeTray();

            // Initialize TargetChecker for periodic checks
            _targetChecker = new TargetChecker(1000); // Check every 1 second
            _targetChecker.Start();

            // Load and create fences
            FenceManager.LoadAndCreateFences(_targetChecker);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _trayManager?.Dispose();
            base.OnExit(e);
        }
    }
}
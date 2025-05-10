//using System;
//using System.Windows;

//namespace Desktop_Fences
//{
//    public partial class App : Application
//    {
//        private TrayManager _trayManager;
//        private TargetChecker _targetChecker;

//        private void Application_Startup(object sender, StartupEventArgs e)
//        {
//            // Initialize settings
//            SettingsManager.LoadSettings();

//            // Initialize tray icon and context menu
//            _trayManager = new TrayManager();
//            _trayManager.InitializeTray();

//            // Initialize TargetChecker for periodic checks
//            _targetChecker = new TargetChecker(1000); // Check every 1 second
//            _targetChecker.Start();

//            // Load and create fences
//            FenceManager.LoadAndCreateFences(_targetChecker);
//        }

//        protected override void OnExit(ExitEventArgs e)
//        {
//            _trayManager?.Dispose();
//            base.OnExit(e);
//        }
//    }
//}

using System.Windows;

namespace Desktop_Fences
{
    public partial class App : Application
    {
        private TrayManager _trayManager;
        private TargetChecker _targetChecker;

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            // Initialize settings FIRST
            SettingsManager.LoadSettings();

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
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _trayManager?.Dispose();
            base.OnExit(e);
        }
    }
}
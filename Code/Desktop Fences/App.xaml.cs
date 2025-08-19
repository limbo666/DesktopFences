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
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // NEW: Cleanup InterCore system
            InterCore.Cleanup();

            _trayManager?.Dispose();
            base.OnExit(e);
        }
    }
}
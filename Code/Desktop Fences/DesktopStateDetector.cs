using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Desktop_Fences
{
    /// <summary>
    /// Detects Windows+D desktop state using DWM cloaking detection
    /// Most reliable method for determining if desktop is currently shown
    /// </summary>
    public static class DesktopStateDetector
    {
        #region Win32 API - DWM Cloaking Detection
        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out bool pvAttribute, int cbAttribute);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetShellWindow();

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // DWM Window Attributes
        private const int DWMWA_CLOAKED = 14;
        #endregion

        #region Public Methods
        /// <summary>
        /// Detects if desktop is currently shown (Windows+D first press state)
        /// Uses DWM cloaking detection - most accurate method
        /// </summary>
        /// <returns>True if desktop is shown, False if windows are visible normally</returns>
        public static bool IsDesktopCurrentlyShown()
        {
            try
            {
                int visibleWindows = 0;
                int cloakedWindows = 0;
                int totalRelevantWindows = 0;
                IntPtr shellWindow = GetShellWindow(); // Desktop window handle

                EnumWindows((hwnd, lParam) =>
                {
                    try
                    {
                        // Skip non-windows and desktop window
                        if (!IsWindow(hwnd) || hwnd == shellWindow)
                            return true;

                        // Only check windows with titles (user windows)
                        if (GetWindowTextLength(hwnd) == 0)
                            return true;

                        // Only check visible windows (not minimized normally)
                        if (!IsWindowVisible(hwnd))
                            return true;

                        totalRelevantWindows++;

                        // Check if window is cloaked by DWM (Win+D hiding)
                        bool isCloaked = IsWindowCloaked(hwnd);

                        if (isCloaked)
                        {
                            cloakedWindows++;
                        }
                        else if (!IsIconic(hwnd)) // Not minimized normally
                        {
                            visibleWindows++;
                        }
                    }
                    catch
                    {
                        // Skip problematic windows
                    }

                    return true; // Continue enumeration
                }, IntPtr.Zero);

                // Analysis logic
                bool desktopIsShown = false;

                if (totalRelevantWindows > 0)
                {
                    // If most windows are cloaked, desktop is shown
                    double cloakedPercentage = (double)cloakedWindows / totalRelevantWindows;
                    double visiblePercentage = (double)visibleWindows / totalRelevantWindows;

                    // Desktop is shown if:
                    // 1. Most windows are cloaked (>70%)
                    // 2. Very few windows are visible (<20%)
                    desktopIsShown = cloakedPercentage > 0.7 || visiblePercentage < 0.2;
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"Desktop state: Total={totalRelevantWindows}, Visible={visibleWindows}, Cloaked={cloakedWindows}, DesktopShown={desktopIsShown}");

                return desktopIsShown;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"Error detecting desktop state: {ex.Message}");
                return false; // Default to "not shown" on error
            }
        }

   
        #endregion

        #region Private Helper Methods
        /// <summary>
        /// Checks if a window is cloaked by DWM (hidden by Win+D)
        /// </summary>
        private static bool IsWindowCloaked(IntPtr hWnd)
        {
            try
            {
                bool cloaked;
                int result = DwmGetWindowAttribute(hWnd, DWMWA_CLOAKED, out cloaked, sizeof(bool));
                return result == 0 && cloaked;
            }
            catch
            {
                return false;
            }
        }
        #endregion
    }
}
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;

namespace Desktop_Frames
{
    public static class DesktopIconManager
    {
        #region P/Invoke Definitions
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int GWL_EXSTYLE = -20;

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        #endregion
        private static bool _originalDesktopState = true;
        private static bool _isInitialized = false;
        private static bool _isCurrentlyHidden = false;
        private static bool _isShuttingDown = false; // --- NEW: Shutdown Lock ---
        private static Window _dotWindow;

        public static void Initialize()
        {
            if (_isInitialized) return;

            // Store the user's actual desktop state before we mess with it
            IntPtr listView = GetDesktopListView();
            if (listView != IntPtr.Zero)
            {
                _originalDesktopState = IsWindowVisible(listView);
            }

            // Safety net: Ensure we restore icons if the program crashes
            AppDomain.CurrentDomain.ProcessExit += (s, e) => RestoreOriginalState();

            // --- BUG FIX: Graceful Shutdown Trap ---
            // Catch normal WPF application shutdown (like quitting from the tray icon)
            if (Application.Current != null)
            {
                Application.Current.Exit += (s, e) => RestoreOriginalState();
            }

            _isInitialized = true;
        }

        private static IntPtr GetDesktopListView()
        {
            IntPtr progman = FindWindow("Progman", null);
            IntPtr shelldll = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);

            if (shelldll == IntPtr.Zero)
            {
                IntPtr workerW = IntPtr.Zero;
                do
                {
                    workerW = FindWindowEx(IntPtr.Zero, workerW, "WorkerW", null);
                    shelldll = FindWindowEx(workerW, IntPtr.Zero, "SHELLDLL_DefView", null);
                } while (shelldll == IntPtr.Zero && workerW != IntPtr.Zero);
            }

            if (shelldll != IntPtr.Zero)
            {
                return FindWindowEx(shelldll, IntPtr.Zero, "SysListView32", null);
            }

            return IntPtr.Zero;
        }

        public static void SetDesktopIconsVisible(bool visible)
        {
            // --- BUG FIX: Prevent late-firing hide commands from executing during app teardown ---
            if (_isShuttingDown) return;

            if (!_isInitialized) Initialize();

            IntPtr listView = GetDesktopListView();
            if (listView != IntPtr.Zero)
            {
                ShowWindow(listView, visible ? SW_SHOW : SW_HIDE);
                _isCurrentlyHidden = !visible;
                UpdateDotVisibility();
            }
        }

        public static void ToggleDesktopIcons()
        {
            SetDesktopIconsVisible(_isCurrentlyHidden); // Flip the state
        }

        public static void RestoreOriginalState()
        {
            if (_isInitialized)
            {
                _isShuttingDown = true; // Lock the visibility engine immediately

                IntPtr listView = GetDesktopListView();
                if (listView != IntPtr.Zero)
                {
                    // --- BUG FIX: State Pollution Override ---
                    // If a previous crash left the desktop hidden, _originalDesktopState was recorded wrong.
                    // If the user's settings dictate icons should be visible while running, we MUST guarantee they are visible on exit!
                    bool finalState = _originalDesktopState;
                    if (!SettingsManager.HideDesktopElementsOnStart)
                    {
                        finalState = true;
                    }

                    ShowWindow(listView, finalState ? SW_SHOW : SW_HIDE);
                }
                HideDot();
            }
        }

        #region Floating Dot Logic

        public static void UpdateDotVisibility()
        {
            bool shouldShowDot = _isCurrentlyHidden && SettingsManager.ShowDesktopDot;

            if (shouldShowDot)
            {
                ShowDot();
            }
            else
            {
                HideDot();
            }
        }

        private static void ShowDot()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (_dotWindow == null)
                {
                    Color accent = Utility.GetColorFromName(SettingsManager.SelectedColor);

                    _dotWindow = new Window
                    {
                        Width = 24,
                        Height = 24,
                        WindowStyle = WindowStyle.None,
                        AllowsTransparency = true,
                        Background = Brushes.Transparent,
                        ShowInTaskbar = false,
                        Topmost = false,
                        ResizeMode = ResizeMode.NoResize,
                        ShowActivated = false
                    };

                    Border dot = new Border
                    {
                        Background = new SolidColorBrush(accent),
                        CornerRadius = new CornerRadius(12), // Makes it a perfect circle (half of 24)
                        Opacity = 0.5,
                        Cursor = Cursors.Hand
                    };

                    dot.MouseEnter += (s, e) => dot.Opacity = 1.0;
                    dot.MouseLeave += (s, e) => dot.Opacity = 0.5;
                    dot.MouseLeftButtonUp += (s, e) => ToggleDesktopIcons();

                    _dotWindow.Content = dot;

                    // Make it a non-activating window so clicking it doesn't steal focus from frames
                    _dotWindow.SourceInitialized += (s, e) =>
                    {
                        var hwnd = new WindowInteropHelper(_dotWindow).Handle;
                        SetWindowLong(hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) | WS_EX_NOACTIVATE);
                    };
                }

                // Simplified placement: Bottom center, roughly 50px above the bottom edge
                _dotWindow.Left = (SystemParameters.PrimaryScreenWidth - _dotWindow.Width) / 2;
                _dotWindow.Top = SystemParameters.PrimaryScreenHeight - 65; // ~24px dot + 40px taskbar

                if (!_dotWindow.IsVisible)
                {
                    _dotWindow.Show();
                }

                // Keep color updated if they changed it in options
                if (_dotWindow.Content is Border b)
                {
                    b.Background = new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor));
                }
            });
        }

        private static void HideDot()
        {
            // --- BUG FIX: Safety check for early startup or late shutdown ---
            if (Application.Current == null || Application.Current.Dispatcher == null) return;

            Application.Current.Dispatcher.Invoke(() =>
            {
                if (_dotWindow != null && _dotWindow.IsVisible)
                {
                    _dotWindow.Hide();
                }
            });
        }

        #endregion
    }
}
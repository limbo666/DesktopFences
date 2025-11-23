using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

namespace Desktop_Fences
{
    /// <summary>
    /// Manages system-wide hotkey detection using low-level keyboard hooks
    /// Detects Windows+D combination and other global hotkeys
    /// </summary>
    public static class GlobalHotkeyManager
    {
        #region Win32 API Constants and Structures
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;

        // Virtual key codes
        private const int VK_LWIN = 0x5B;
        private const int VK_RWIN = 0x5C;
        private const int VK_D = 0x44;



        private const int VK_OEM_3 = 0xC0; // Tilde/Backtick key `




        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        #endregion

        #region Win32 API Imports
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);
        #endregion

        #region Private Fields
        private static IntPtr _hookID = IntPtr.Zero;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static bool _isWindowsKeyPressed = false;
        private static bool _isDKeyPressed = false;
        private static bool _hotkeyDetected = false;
        #endregion

        #region Public Events
        /// <summary>
        /// Fired when Windows+D combination is detected
        /// </summary>
        public static event EventHandler WindowsPlusDDetected;
        #endregion

        #region Public Methods
        /// <summary>
        /// Starts monitoring for global hotkeys (call this from your main application)
        /// </summary>
        public static void StartMonitoring()
        {
            try
            {
                if (_hookID != IntPtr.Zero)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                        "GlobalHotkeyManager: Hook already installed, skipping");
                    return;
                }

                _hookID = SetHook(_proc);
                if (_hookID != IntPtr.Zero)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "GlobalHotkeyManager: Started monitoring global hotkeys");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                        "GlobalHotkeyManager: Failed to install keyboard hook");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"GlobalHotkeyManager: Error starting monitoring: {ex.Message}");
            }
        }

        /// <summary>
        /// Stops monitoring for global hotkeys (call this on application exit)
        /// </summary>
        public static void StopMonitoring()
        {
            try
            {
                if (_hookID != IntPtr.Zero)
                {
                    bool result = UnhookWindowsHookEx(_hookID);
                    if (result)
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                            "GlobalHotkeyManager: Stopped monitoring global hotkeys");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                            "GlobalHotkeyManager: Warning - Failed to unhook keyboard hook");
                    }
                    _hookID = IntPtr.Zero;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"GlobalHotkeyManager: Error stopping monitoring: {ex.Message}");
            }
        }

        /// <summary>
        /// Test method to simulate Windows+D detection (for testing purposes)
        /// </summary>
        public static void TestWindowsPlusD()
        {
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                "GlobalHotkeyManager: Test Windows+D triggered");
            OnWindowsPlusDDetected();
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Sets up the low-level keyboard hook
        /// </summary>
        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                    GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        /// <summary>
        /// Hook callback procedure that processes all keyboard input system-wide
        /// </summary>
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0)
                {
                    KBDLLHOOKSTRUCT hookStruct = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));
                    uint vkCode = hookStruct.vkCode;

                    bool isKeyDown = (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN);
                    bool isKeyUp = (wParam == (IntPtr)WM_KEYUP || wParam == (IntPtr)WM_SYSKEYUP);

                    // Track Windows key state
                    if (vkCode == VK_LWIN || vkCode == VK_RWIN)
                    {
                        if (isKeyDown)
                        {
                            _isWindowsKeyPressed = true;
                        }
                        else if (isKeyUp)
                        {
                            _isWindowsKeyPressed = false;
                            _hotkeyDetected = false; // Reset detection
                        }
                    }

                    // Track D key state
                    if (vkCode == VK_D)
                    {
                        if (isKeyDown)
                        {
                            _isDKeyPressed = true;
                        }
                        else if (isKeyUp)
                        {
                            _isDKeyPressed = false;
                        }
                    }

                    // Detect Windows+D combination (only trigger once per key combination)
                    if (_isWindowsKeyPressed && _isDKeyPressed && !_hotkeyDetected)
                    {
                        _hotkeyDetected = true;

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                            "GlobalHotkeyManager: Windows+D combination detected!");

                        // Fire the event on the UI thread
                        System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            OnWindowsPlusDDetected();
                        }));
                    }

                    // Detect Ctrl + ` (Tilde) for Search
                    bool isCtrlDown = (GetAsyncKeyState(0x11) & 0x8000) != 0; // VK_CONTROL
                    if (vkCode == VK_OEM_3 && isCtrlDown && isKeyDown)
                    {
                        // Prevent typing the ` character
                        // Fire search toggle
                        System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            SearchFormManager.ToggleSearch();
                        }));
                        return (IntPtr)1; // Swallow the key
                    }

                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"GlobalHotkeyManager: Error in hook callback: {ex.Message}");
            }

            // Always call next hook in chain
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        /// <summary>
        /// Triggers the Windows+D detected event
        /// </summary>
        private static void OnWindowsPlusDDetected()
        {
            try
            {
                WindowsPlusDDetected?.Invoke(null, EventArgs.Empty);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"GlobalHotkeyManager: Error firing Windows+D event: {ex.Message}");
            }
        }
        #endregion
    }
}
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input; // For Key enum

namespace Desktop_Fences
{
    /// <summary>
    /// Centralized Global Hotkey Manager
    /// Handles low-level keyboard hooking for all application shortcuts (Win+D, Search, Easter Eggs)
    /// </summary>
    public static class GlobalHotkeyManager
    {
        #region Win32 API Constants
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int WM_SYSKEYUP = 0x0105;

        // Virtual key codes
        private const int VK_LWIN = 0x5B;
        private const int VK_RWIN = 0x5C;
        private const int VK_D = 0x44;
        private const int VK_G = 0x47; // Used for Gravity Drop

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

        // State tracking for specific complex combos if needed
        private static bool _isWindowsKeyPressed = false;
        private static bool _isDKeyPressed = false;
        private static bool _winDDetected = false;
        private static bool _searchHotkeyDetected = false;
        #endregion

        #region Public Events
        // 1. Windows + D (Toggle Desktop/Fences)
        public static event EventHandler WindowsPlusDDetected;

        // 2. Dance Party (Ctrl + Alt + D) - Easter Egg from InterCore
        public static event EventHandler DancePartyTriggered;

        // 3. Gravity Drop (Ctrl + Shift + G) - Easter Egg from InterCore
        public static event EventHandler GravityDropTriggered;
        #endregion

        #region Public Methods
        public static void StartMonitoring()
        {
            try
            {
                if (_hookID != IntPtr.Zero) return;
                _hookID = SetHook(_proc);
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "GlobalHotkeyManager: Hook started.");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"GlobalHotkeyManager Error: {ex.Message}");
            }
        }

        public static void StopMonitoring()
        {
            try
            {
                if (_hookID != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookID);
                    _hookID = IntPtr.Zero;
                }
            }
            catch { }
        }
        #endregion

        #region Private Methods
        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

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

                    // ------------------------------------------------------------
                    // 1. Windows + D Detection (System Toggle)
                    // ------------------------------------------------------------
                    if (vkCode == VK_LWIN || vkCode == VK_RWIN)
                    {
                        if (isKeyDown) _isWindowsKeyPressed = true;
                        else if (isKeyUp) { _isWindowsKeyPressed = false; _winDDetected = false; }
                    }
                    if (vkCode == VK_D)
                    {
                        if (isKeyDown) _isDKeyPressed = true;
                        else if (isKeyUp) _isDKeyPressed = false;
                    }

                    if (_isWindowsKeyPressed && _isDKeyPressed && !_winDDetected)
                    {
                        _winDDetected = true;
                        System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            WindowsPlusDDetected?.Invoke(null, EventArgs.Empty);
                        }));
                    }

                    // ------------------------------------------------------------
                    // 2. SpotSearch Detection (Dynamic Key + Modifier)
                    // ------------------------------------------------------------
                    if (SettingsManager.EnableSpotSearchHotkey)
                    {
                        int triggerKey = SettingsManager.SpotSearchKey;
                        // Convert friendly setting to vkCode if needed, but SettingsManager stores int now.

                        if (vkCode == triggerKey)
                        {
                            string mod = SettingsManager.SpotSearchModifier?.ToLower();
                            bool isModPressed = false;

                            if (mod == "control") isModPressed = (GetAsyncKeyState(0x11) & 0x8000) != 0;
                            else if (mod == "alt") isModPressed = (GetAsyncKeyState(0x12) & 0x8000) != 0;
                            else if (mod == "shift") isModPressed = (GetAsyncKeyState(0x10) & 0x8000) != 0;
                            else if (mod == "none") isModPressed = true;

                            if (isKeyDown && isModPressed)
                            {
                                if (!_searchHotkeyDetected)
                                {
                                    _searchHotkeyDetected = true;
                                    System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        SearchFormManager.ToggleSearch();
                                    }));
                                }
                                return (IntPtr)1; // Swallow key
                            }
                            else if (isKeyUp)
                            {
                                _searchHotkeyDetected = false;
                            }
                        }
                    }

                    // ------------------------------------------------------------
                    // 3. Dance Party Detection (Ctrl + Alt + D) - InterCore
                    // ------------------------------------------------------------
                    if (vkCode == VK_D && isKeyDown)
                    {
                        bool ctrl = (GetAsyncKeyState(0x11) & 0x8000) != 0; // VK_CONTROL
                        bool alt = (GetAsyncKeyState(0x12) & 0x8000) != 0;  // VK_MENU

                        if (ctrl && alt)
                        {
                            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                DancePartyTriggered?.Invoke(null, EventArgs.Empty);
                            }));
                        }
                    }

                    // ------------------------------------------------------------
                    // 4. Gravity Drop Detection (Ctrl + Shift + G) - InterCore
                    // ------------------------------------------------------------
                    if (vkCode == VK_G && isKeyDown)
                    {
                        bool ctrl = (GetAsyncKeyState(0x11) & 0x8000) != 0;
                        bool shift = (GetAsyncKeyState(0x10) & 0x8000) != 0; // VK_SHIFT

                        if (ctrl && shift)
                        {
                            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                GravityDropTriggered?.Invoke(null, EventArgs.Empty);
                            }));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Hook Error: {ex.Message}");
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
        #endregion
    }
}
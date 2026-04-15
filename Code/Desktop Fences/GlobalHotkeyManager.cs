using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Input; // For Key enum

namespace Desktop_Fences
{
    /// <summary>
    /// Centralized Global Hotkey Manager
    /// Handles low-level keyboard hooking for all application shortcuts (Win+D, Search, Profiles)
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
        private const int VK_CONTROL = 0x11;
        private const int VK_MENU = 0x12; // Alt
        private const int VK_SHIFT = 0x10;

        // Triggers
        private const int VK_D = 0x44;
        private const int VK_G = 0x47; // Gravity
        private const int VK_Z = 0x5A; // Focus Fence
        private const int VK_0 = 0x30; // 0 key
        private const int VK_9 = 0x39; // 9 key
        private const int VK_OEM_COMMA = 0xBC; // , < key
        private const int VK_OEM_PERIOD = 0xBE; // . > key

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

        // State tracking
        private static bool _isWindowsKeyPressed = false;
        private static bool _isDKeyPressed = false;
        private static bool _winDDetected = false;
        private static bool _searchHotkeyDetected = false;
        #endregion

        #region Public Events
        public static event EventHandler WindowsPlusDDetected;
        public static event EventHandler DancePartyTriggered;
        public static event EventHandler GravityDropTriggered;
        #endregion

        #region Public Methods
        public static void StartMonitoring()
        {
            try
            {
                if (_hookID != IntPtr.Zero) return;

                _hookID = SetHook(_proc);

                // Enhanced Verification Logging
                if (_hookID == IntPtr.Zero)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                        $"GlobalHotkeyManager: CRITICAL FAILURE. Hook failed to start. Error Code: {errorCode}");
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        $"GlobalHotkeyManager: Hook started successfully. ID: {_hookID}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"GlobalHotkeyManager Exception: {ex.Message}");
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
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, "GlobalHotkeyManager: Hook stopped.");
                }
            }
            catch { }
        }
        #endregion

        #region Private Methods
        /// <summary>
        /// Strictly parses comma-separated modifiers (e.g. "Control, Alt") 
        /// and ensures ONLY those modifiers are pressed.
        /// </summary>
        private static bool CheckModifiersStrict(string modifierString)
        {
            if (string.IsNullOrWhiteSpace(modifierString)) return false;
            string mod = modifierString.ToLower();

            bool requireCtrl = mod.Contains("ctrl") || mod.Contains("control");
            bool requireAlt = mod.Contains("alt");
            bool requireShift = mod.Contains("shift");
            bool requireWin = mod.Contains("win");

            bool isCtrl = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
            bool isAlt = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
            bool isShift = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
            bool isWin = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0 || (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;

            return (requireCtrl == isCtrl) && (requireAlt == isAlt) && (requireShift == isShift) && (requireWin == isWin);
        }

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

                    // ============================================================
                    // ============================================================
                    // 1. PROFILE SWITCHING (Dynamic Customizable Modifiers)
                    // ============================================================
                    if (isKeyDown && SettingsManager.EnableProfileHotkeys)
                    {
                        // A. Configurable Keys -> Switch by ID
                        if (CheckModifiersStrict(SettingsManager.ProfileSwitchModifier))
                        {
                            for (int i = 0; i < SettingsManager.ProfileSwitchKeys.Length; i++)
                            {
                                if (vkCode == SettingsManager.ProfileSwitchKeys[i])
                                {
                                    int profileId = i;
                                    System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                                    {
                                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Hotkey: Switching to Profile ID {profileId}");
                                        ProfileManager.SwitchToProfileById(profileId);
                                        TrayManager.Instance?.UpdateProfilesMenu();
                                        TrayManager.Instance?.UpdateTrayIcon();
                                    }));
                                    return (IntPtr)1; // Swallow key
                                }
                            }
                        }

                        // B. Previous Profile
                        if (vkCode == SettingsManager.ProfilePrevKey && CheckModifiersStrict(SettingsManager.ProfilePrevModifier))
                        {
                            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                ProfileManager.SwitchToPreviousProfile();
                                TrayManager.Instance?.UpdateProfilesMenu();
                                TrayManager.Instance?.UpdateTrayIcon();
                            }));
                            return (IntPtr)1;
                        }

                        // C. Next Profile
                        if (vkCode == SettingsManager.ProfileNextKey && CheckModifiersStrict(SettingsManager.ProfileNextModifier))
                        {
                            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                ProfileManager.SwitchToNextProfile();
                                TrayManager.Instance?.UpdateProfilesMenu();
                                TrayManager.Instance?.UpdateTrayIcon();
                            }));
                            return (IntPtr)1;
                        }
                    }

                    // ============================================================
                    // 2. Windows + D Detection
                    // ============================================================
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

                    // ============================================================
                    // 3. SpotSearch (Ctrl + `)
                    // ============================================================
                    if (SettingsManager.EnableSpotSearchHotkey)
                    {
                        int triggerKey = SettingsManager.SpotSearchKey;
                        if (vkCode == triggerKey)
                        {
                            string mod = SettingsManager.SpotSearchModifier?.ToLower();
                            bool isModPressed = false;

                            if (mod == "control") isModPressed = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                            else if (mod == "alt") isModPressed = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
                            else if (mod == "shift") isModPressed = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
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
                                return (IntPtr)1;
                            }
                            else if (isKeyUp) _searchHotkeyDetected = false;
                        }
                    }

                    // ============================================================
                    // 4. Focus Fence (Dynamic Configurable)
                    // ============================================================
                    if (SettingsManager.EnableFocusFenceHotkey && vkCode == SettingsManager.FocusFenceKey && isKeyDown)
                    {
                        if (CheckModifiersStrict(SettingsManager.FocusFenceModifier))
                        {
                            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                FenceFocusFormManager focusManager = new FenceFocusFormManager();
                                focusManager.ShowDialog();
                            }));
                            return (IntPtr)1; // Swallow the key so other apps don't process it
                        }
                    }

                    // ============================================================
                    // 5. Easter Eggs (InterCore)
                    // ============================================================
                    // Dance Party (Ctrl + Alt + Shift + D)
                    if (vkCode == VK_D && isKeyDown)
                    {
                        bool isC = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                        bool isA = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
                        bool isS = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
                        bool isW = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0 || (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;

                        if (isC && isA && isS && !isW)
                        {
                            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                            {
                                DancePartyTriggered?.Invoke(null, EventArgs.Empty);
                            }));
                        }
                    }

                    // Gravity Drop (Ctrl + Alt + Shift + G)
                    if (vkCode == VK_G && isKeyDown)
                    {
                        bool isC = (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0;
                        bool isA = (GetAsyncKeyState(VK_MENU) & 0x8000) != 0;
                        bool isS = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
                        bool isW = (GetAsyncKeyState(VK_LWIN) & 0x8000) != 0 || (GetAsyncKeyState(VK_RWIN) & 0x8000) != 0;

                        if (isC && isA && isS && !isW)
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
                Debug.WriteLine($"Hook Error: {ex.Message}");
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
        #endregion
    }
}
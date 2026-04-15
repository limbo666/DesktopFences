using Desktop_Fences;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;

namespace DesktopFences
{
    /// <summary>
    /// Installs a WH_MOUSE_LL hook and fires FenceManager.WakeUpFences()
    /// when the user double-clicks the bare Windows desktop (Progman / WorkerW).
    /// </summary>
    public static class DesktopMouseHook
    {
        // ── Win32 ────────────────────────────────────────────────────────────
        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;

        // ListView Messages
        private const uint LVM_FIRST = 0x1000;
        private const uint LVM_GETSELECTEDCOUNT = LVM_FIRST + 50;

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int x, y; }

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetDoubleClickTime();

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT p);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        // ── State ────────────────────────────────────────────────────────────
        private static IntPtr _hookHandle = IntPtr.Zero;
        private static LowLevelMouseProc? _hookProc;

        private static uint _lastClickTime = 0;
        private static POINT _lastClickPoint = default;
        private const int CLICK_RADIUS = 4;

        // ── Public API ───────────────────────────────────────────────────────
        public static void Start()
        {
            if (_hookHandle != IntPtr.Zero) return;

            _hookProc = HookCallback;

            using var curProcess = Process.GetCurrentProcess();
            using var curModule = curProcess.MainModule ?? throw new InvalidOperationException("Cannot obtain main module.");

            IntPtr hMod = GetModuleHandle(curModule.ModuleName);
            _hookHandle = SetWindowsHookEx(WH_MOUSE_LL, _hookProc, hMod, 0);

            if (_hookHandle == IntPtr.Zero)
                throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "Failed to install low-level mouse hook.");
        }

        public static void Stop()
        {
            if (_hookHandle == IntPtr.Zero) return;
            UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
            _hookProc = null;
        }

        // ── Hook callback ────────────────────────────────────────────────────
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (int)wParam == WM_LBUTTONDOWN)
            {
                var data = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                HandleClick(data.pt, data.time);
            }
            return CallNextHookEx(_hookHandle, nCode, wParam, lParam);
        }

        private static void HandleClick(POINT pt, uint time)
        {
            uint dblClickMs = GetDoubleClickTime();
            uint elapsed = time - _lastClickTime;
            bool withinTime = elapsed <= dblClickMs;
            bool withinArea = WithinRadius(pt, _lastClickPoint, CLICK_RADIUS);

            if (withinTime && withinArea)
            {
                _lastClickTime = 0;
                _lastClickPoint = default;

                IntPtr hWnd = WindowFromPoint(pt);

                if (IsDesktopWindow(hWnd) && !HasSelectedDesktopIcon(hWnd))
                {
                    // FIX: Use BeginInvoke so the mouse hook returns instantly.
                    // This stops the UI from feeling like it is "following/lagging" your click.
                    Application.Current?.Dispatcher.BeginInvoke(new Action(() => FenceManager.WakeUpFences()));
                }
            }
            else
            {
                _lastClickTime = time;
                _lastClickPoint = pt;
            }
        }

        // ── Helpers ──────────────────────────────────────────────────────────
        private static bool IsDesktopWindow(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return false;

            var sb = new System.Text.StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);
            string cls = sb.ToString();

            return cls == "Progman" || cls == "WorkerW" || cls == "SysListView32" || cls == "SHELLDLL_DefView";
        }

        /// <summary>
        /// Sends a direct message to the desktop asking if any icons are selected.
        /// When double-clicking an icon, the first click selects it. 
        /// Therefore, at the exact moment of the double-click, this will be > 0.
        /// </summary>
        private static bool HasSelectedDesktopIcon(IntPtr hWnd)
        {
            var sb = new System.Text.StringBuilder(256);
            GetClassName(hWnd, sb, sb.Capacity);

            // We only query the list view itself
            if (sb.ToString() == "SysListView32")
            {
                int selectedCount = (int)SendMessage(hWnd, LVM_GETSELECTEDCOUNT, IntPtr.Zero, IntPtr.Zero);
                return selectedCount > 0;
            }

            return false;
        }

        private static bool WithinRadius(POINT a, POINT b, int radius)
        {
            int dx = a.x - b.x;
            int dy = a.y - b.y;
            return (dx * dx + dy * dy) <= (radius * radius);
        }
    }
}
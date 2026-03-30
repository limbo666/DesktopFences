using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Threading;
using System.IO;
using System.Diagnostics;

namespace Desktop_Fences
{
    public static class AutomationManager
    {
        private static DispatcherTimer _timer;
        private static string _lastDetectedProcess = "";
        private static DateTime _matchStartTime;
        private static bool _isCurrentlyAutomated = false;

        // Win32 Imports
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)] private static extern bool QueryFullProcessImageName(IntPtr hProcess, int dwFlags, [Out] StringBuilder lpExeName, ref int lpdwSize);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);

        public static void Start()
        {
            if (_timer != null) return;

            // Run on UI thread to interact safely with Profile/Tray
            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(1000) };
            _timer.Tick += (s, e) => Tick();
            _timer.Start();
        }

        public static void Stop()
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer = null;
            }
            _isCurrentlyAutomated = false;
            _lastDetectedProcess = "";
        }

        private static void Tick()
        {
            // SAFETY 1: Global Try-Catch to prevent Timer death
            try
            {
                // Respect the Master Toggle
                if (!SettingsManager.EnableProfileAutomation)
                {
                    Stop();
                    return;
                }

                string currentProc = GetActiveProcessName();

                // If we can't detect a process (e.g. UAC prompt, lock screen), do nothing this tick
                if (string.IsNullOrEmpty(currentProc)) return;

                // Match against rules
                var rule = ProfileManager.AutomationRules.FirstOrDefault(r =>
                    currentProc.Equals(r.ProcessName, StringComparison.OrdinalIgnoreCase));

                if (rule != null)
                {
                    // If process changed or we just started tracking this match
                    if (!string.Equals(_lastDetectedProcess, currentProc, StringComparison.OrdinalIgnoreCase))
                    {
                        _lastDetectedProcess = currentProc;
                        _matchStartTime = DateTime.Now;
                        return; // Wait for delay in next ticks
                    }

                    // Delay Check
                    if ((DateTime.Now - _matchStartTime).TotalSeconds >= rule.DelaySeconds)
                    {
                        // Check if we are already on the correct profile
                        if (ProfileManager.CurrentProfileName != rule.TargetProfile)
                        {
                            // "Manual Override" Check:
                            // Only switch if we are currently on the "Home" profile OR the "Automated" profile.
                            // If user manually switched to a 3rd profile, don't interrupt them.
                            bool isSafeToSwitch = (ProfileManager.CurrentProfileName == ProfileManager.ManualBaseProfile) || _isCurrentlyAutomated;

                            if (isSafeToSwitch)
                            {
                                if (rule.IsPersisted)
                                {
                                    // Persistent rules change the "Home" so we don't revert
                                    ProfileManager.SetManualBaseProfile(rule.TargetProfile);
                                    _isCurrentlyAutomated = false;
                                }
                                else
                                {
                                    _isCurrentlyAutomated = true;
                                }

                                PerformProfileSwitch(rule.TargetProfile);
                            }
                        }
                    }
                }
                else if (_isCurrentlyAutomated)
                {
                    // REVERT LOGIC: Process closed/lost focus, and we were in automated mode
                    _isCurrentlyAutomated = false;
                    _lastDetectedProcess = "";

                    // Only revert if we are still on the profile we automated TO (don't revert if user moved away)
                    // Actually, revert anyway to be safe and restore order
                    PerformProfileSwitch(ProfileManager.ManualBaseProfile);
                }
                else
                {
                    // No rule match, not automated. Just track process to avoid flapping
                    _lastDetectedProcess = currentProc;
                }
            }
            catch
            {
                // Swallow exceptions so the Timer keeps ticking.
                // In production, you might log this, but for now, we just ensure survival.
            }
        }

        private static void PerformProfileSwitch(string profileName)
        {
            try
            {
                ProfileManager.SwitchToProfile(profileName);

                // Update UI safely
                if (System.Windows.Application.Current != null)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            TrayManager.Instance?.UpdateTrayIcon();
                            TrayManager.Instance?.UpdateProfilesMenu();
                        }
                        catch { }
                    });
                }
            }
            catch { }
        }

        private static string GetActiveProcessName()
        {
            try
            {
                IntPtr hwnd = GetForegroundWindow();
                if (hwnd == IntPtr.Zero) return "";

                uint pid;
                GetWindowThreadProcessId(hwnd, out pid);
                if (pid == 0) return "";

                // Use a higher permission request if possible, or fallback? 
                // 0x1000 is PROCESS_QUERY_LIMITED_INFORMATION (usually enough)
                IntPtr processHandle = OpenProcess(0x1000, false, pid);
                if (processHandle == IntPtr.Zero) return "";

                try
                {
                    int capacity = 2048;
                    StringBuilder sb = new StringBuilder(capacity);
                    if (QueryFullProcessImageName(processHandle, 0, sb, ref capacity))
                    {
                        return Path.GetFileNameWithoutExtension(sb.ToString());
                    }
                }
                finally
                {
                    CloseHandle(processHandle);
                }
            }
            catch { }
            return "";
        }
    }
}
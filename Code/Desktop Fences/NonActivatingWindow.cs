using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using Desktop_Fences;

public class NonActivatingWindow : Window
{
    private const int WM_SYSCOMMAND = 0x0112;
    private const int SC_MAXIMIZE = 0xF030;
    private const int SC_RESTORE = 0xF120;

    private const int WM_MOUSEACTIVATE = 0x0021;
    private const int MA_NOACTIVATE = 3;
    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_NOACTIVATE = 0x08000000;
    private bool _focusPreventionEnabled = true;

    protected override void OnSourceInitialized(EventArgs e)
    {
        base.OnSourceInitialized(e);
        var hwndSource = (HwndSource)PresentationSource.FromVisual(this);
        hwndSource.AddHook(WndProc);
        SetWindowLong(new WindowInteropHelper(this).Handle, GWL_EXSTYLE, GetWindowLong(new WindowInteropHelper(this).Handle, GWL_EXSTYLE) | WS_EX_NOACTIVATE);
    }

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {

        const int WM_ENTERSIZEMOVE = 0x0231; // Resizing starts
        const int WM_EXITSIZEMOVE = 0x0232;  // Resizing ends

        if (msg == WM_ENTERSIZEMOVE)
        {
            FenceManager.OnResizingStarted(this);
        }
        else if (msg == WM_EXITSIZEMOVE)
        {
            FenceManager.OnResizingEnded(this);
        }

        // Handle existing focus prevention

        if (_focusPreventionEnabled && msg == WM_MOUSEACTIVATE)
        {
            handled = true;
            return new IntPtr(MA_NOACTIVATE);
        }

        // Block Aero Snap maximize/restore commands
        if (msg == WM_SYSCOMMAND)
        {
            int command = wParam.ToInt32() & 0xFFF0;
            if (command == SC_MAXIMIZE || command == SC_RESTORE)
            {
                handled = true;
                return IntPtr.Zero;
            }
        }

        return IntPtr.Zero;
    }

    public void EnableFocusPrevention(bool enable)
    {
        _focusPreventionEnabled = enable;
        if (enable)
        {
            SetWindowLong(new WindowInteropHelper(this).Handle, GWL_EXSTYLE, GetWindowLong(new WindowInteropHelper(this).Handle, GWL_EXSTYLE) | WS_EX_NOACTIVATE);
        }
        else
        {
            SetWindowLong(new WindowInteropHelper(this).Handle, GWL_EXSTYLE, GetWindowLong(new WindowInteropHelper(this).Handle, GWL_EXSTYLE) & ~WS_EX_NOACTIVATE);
        }
    }

    [DllImport("user32.dll")]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll")]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_SHOWNOACTIVATE = 4;

    public void ShowWithoutActivation()
    {
        var hwnd = new WindowInteropHelper(this).Handle;
        ShowWindow(hwnd, SW_SHOWNOACTIVATE);
    }
}
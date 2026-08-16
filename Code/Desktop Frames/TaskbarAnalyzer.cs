using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Desktop_Frames
{
    /// <summary>
    /// Highly optimized, stateless helper to evaluate Z-order and classes at a specific pixel.
    /// </summary>
    public static class TaskbarAnalyzer
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT Point);

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        public static bool IsShowDesktopButtonClick(int x, int y)
        {
            try
            {
                IntPtr hWnd = WindowFromPoint(new POINT { X = x, Y = y });

                // Walk up the visual tree. WindowFromPoint often returns nested child elements.
                // We keep checking parents until we hit the root taskbar or run out of parents.
                while (hWnd != IntPtr.Zero)
                {
                    StringBuilder className = new StringBuilder(256);
                    GetClassName(hWnd, className, className.Capacity);
                    string cls = className.ToString();

                    // Direct Hit on primary monitor button
                    if (cls == "TrayShowDesktopButtonWClass") return true;

                    // Hit on the general taskbar wrapper (Primary or Secondary)
                    if (cls == "Shell_TrayWnd" || cls == "Shell_SecondaryTrayWnd")
                    {
                        if (GetWindowRect(hWnd, out RECT rect))
                        {
                            bool isHorizontal = (rect.Right - rect.Left) > (rect.Bottom - rect.Top);
                            if (isHorizontal)
                            {
                                // Expanded hitbox to 20px for Win11 compatibility
                                if (x >= rect.Right - 20) return true;
                            }
                            else
                            {
                                if (y >= rect.Bottom - 20) return true;
                            }
                        }
                    }

                    // Move up to the parent element and check again
                    hWnd = GetParent(hWnd);
                }
            }
            catch
            {
                // Fails silently to protect the global mouse hook chain
            }
            return false;
        }
    }
}
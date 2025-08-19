using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms; // For Screen.AllScreens

namespace Desktop_Fences
{
    public static class SnapManager
    {
        private const double SnapDistance = 40;
        private const double SnapThreshold = 30;
        private const double InternalSnapThreshold = 15;
        private const double MinGap = 15;

        public static void AddSnapping(NonActivatingWindow win, IDictionary<string, object> fenceData)
        {
            // Snapping is optional, triggered on LocationChanged
            win.LocationChanged += (sender, e) =>
            {
                var (newLeft, newTop) = SnapToClosestFence(win, System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList());
                win.Left = newLeft;
                win.Top = newTop;
                fenceData["X"] = newLeft;
                fenceData["Y"] = newTop;
                FenceManager.SaveFenceData();
            };
        }

        private static (double, double) SnapToClosestFence(NonActivatingWindow currentFence, List<NonActivatingWindow> allFences)
        {
            if (!SettingsManager.IsSnapEnabled)
            {
                return (currentFence.Left, currentFence.Top);
            }

            double initialX = currentFence.Left;
            double initialY = currentFence.Top;
            double snapX = initialX;
            double snapY = initialY;

            // Get virtual desktop bounds
            double virtualLeft = SystemParameters.VirtualScreenLeft;
            double virtualTop = SystemParameters.VirtualScreenTop;
            double virtualWidth = SystemParameters.VirtualScreenWidth;
            double virtualHeight = SystemParameters.VirtualScreenHeight;

            List<string> nearbyFences = new List<string>();
            List<string> causingSnapFences = new List<string>();

            // Snap to other fences
            foreach (var fence in allFences)
            {
                if (fence == currentFence) continue;

                double currentLeft = currentFence.Left;
                double currentRight = currentFence.Left + currentFence.Width;
                double currentTop = currentFence.Top;
                double currentBottom = currentFence.Top + currentFence.Height;

                double otherLeft = fence.Left;
                double otherRight = fence.Left + fence.Width;
                double otherTop = fence.Top;
                double otherBottom = fence.Top + fence.Height;

                double horizontalThreshold = (currentLeft >= otherLeft && currentRight <= otherRight) ? InternalSnapThreshold : SnapThreshold;
                double verticalThreshold = (currentTop >= otherTop && currentBottom <= otherBottom) ? InternalSnapThreshold : SnapThreshold;

                if (Math.Abs(currentRight - otherLeft) <= horizontalThreshold)
                {
                    snapX = otherLeft - currentFence.Width - MinGap;
                    causingSnapFences.Add(fence.Title);
                }
                else if (Math.Abs(currentLeft - otherRight) <= horizontalThreshold)
                {
                    snapX = otherRight + MinGap;
                    causingSnapFences.Add(fence.Title);
                }

                if (Math.Abs(currentBottom - otherTop) <= verticalThreshold)
                {
                    snapY = otherTop - currentFence.Height - MinGap;
                    causingSnapFences.Add(fence.Title);
                }
                else if (Math.Abs(currentTop - otherBottom) <= verticalThreshold)
                {
                    snapY = otherBottom + MinGap;
                    causingSnapFences.Add(fence.Title);
                }

                nearbyFences.Add(fence.Title);
            }

            // Snap to screen edges (all monitors)
            foreach (var screen in Screen.AllScreens)
            {
                double screenLeft = screen.Bounds.Left;
                double screenRight = screen.Bounds.Right;
                double screenTop = screen.Bounds.Top;
                double screenBottom = screen.Bounds.Bottom;

                if (Math.Abs(currentFence.Left - screenLeft) <= SnapThreshold)
                {
                    snapX = screenLeft;
                    causingSnapFences.Add($"ScreenEdge({screen.DeviceName}, Left)");
                }
                else if (Math.Abs((currentFence.Left + currentFence.Width) - screenRight) <= SnapThreshold)
                {
                    snapX = screenRight - currentFence.Width;
                    causingSnapFences.Add($"ScreenEdge({screen.DeviceName}, Right)");
                }

                if (Math.Abs(currentFence.Top - screenTop) <= SnapThreshold)
                {
                    snapY = screenTop;
                    causingSnapFences.Add($"ScreenEdge({screen.DeviceName}, Top)");
                }
                else if (Math.Abs((currentFence.Top + currentFence.Height) - screenBottom) <= SnapThreshold)
                {
                    snapY = screenBottom - currentFence.Height;
                    causingSnapFences.Add($"ScreenEdge({screen.DeviceName}, Bottom)");
                }
            }

            // Clamp to virtual desktop bounds
            snapX = Math.Max(virtualLeft, Math.Min(snapX, virtualLeft + virtualWidth - currentFence.Width));
            snapY = Math.Max(virtualTop, Math.Min(snapY, virtualTop + virtualHeight - currentFence.Height));

            // Log details
            if (SettingsManager.IsLogEnabled)
            {
                // Collect monitor details
                var monitorDetails = Screen.AllScreens.Select(s =>
                    $"{s.DeviceName}: ({s.Bounds.Left}, {s.Bounds.Top}, {s.Bounds.Width}x{s.Bounds.Height})");
                string logMessage = $"FenceDragged: {currentFence.Title}, " +
                                   $"NearbyFences: {string.Join(", ", nearbyFences)}, " +
                                   $"CausingSnapFences: {string.Join(", ", causingSnapFences)}, " +
                                   $"FenceDraggedPosition: ({initialX}, {initialY}), " +
                                   $"FenceSnappedPosition: ({snapX}, {snapY}), " +
                                   $"VirtualBounds: Left={virtualLeft}, Top={virtualTop}, Width={virtualWidth}, Height={virtualHeight}, " +
                                   $"Monitors: {string.Join("; ", monitorDetails)}";
                LogSnapDetails(logMessage);
            }

            return (snapX, snapY);
        }

        private static void LogSnapDetails(string message)
        {
            if (!SettingsManager.IsLogEnabled) return;

            string logFilePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Desktop_Fences_Snap.log");
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(logFilePath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }
    }
}
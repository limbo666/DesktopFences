using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace Desktop_Fences
{
    public static class SnapManager
    {
        private const double SnapDistance = 20;
        private const double SnapThreshold = 30;
        private const double InternalSnapThreshold = 15;
        private const double MinGap = 8;

        public static void AddSnapping(NonActivatingWindow win, IDictionary<string, object> fenceData)
        {
            // Snapping is now optional since dragging is handled in titlelabel.MouseDown
            win.LocationChanged += (sender, e) =>
            {
                var (newLeft, newTop) = SnapToClosestFence(win, Application.Current.Windows.OfType<NonActivatingWindow>().ToList());
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

            List<string> nearbyFences = new List<string>();
            List<string> causingSnapFences = new List<string>();

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

            snapX = Math.Max(0, snapX);
            snapY = Math.Max(0, snapY);

            if (SettingsManager.IsLogEnabled)
            {
                string logMessage = $"FenceDragged: {currentFence.Title}, NearbyFences: {string.Join(", ", nearbyFences)}, " +
                                   $"CausingSnapFences: {string.Join(", ", causingSnapFences)}, " +
                                   $"FenceDraggedPosition: ({initialX}, {initialY}), " +
                                   $"FenceSnappedPosition: ({snapX}, {snapY}), " +
                                   $"ScreenSize: ({SystemParameters.PrimaryScreenWidth}, {SystemParameters.PrimaryScreenHeight})";
                LogSnapDetails(logMessage);
            }

            return (snapX, snapY);
        }

        private static void LogSnapDetails(string message)
        {
            if (!SettingsManager.IsLogEnabled) return;

            string logFilePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Deskop_Fences_Snap.log");
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(logFilePath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }
    }
}
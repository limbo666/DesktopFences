using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms; // For Screen.AllScreens
using System.Windows.Media; // For VisualTreeHelper

namespace Desktop_Frames
{
    public static class SnapManager
    {
        private const double SnapThreshold = 20; // Reduced slightly for tighter feel
        private const double MinGap = 10;        // Gap between snapped frames

        // Recursion guard to prevent the "fighting" loop
        private static bool _isSnapping = false;




        public static NonActivatingWindow ActiveDragWindow = null;

        public static void StartDrag(NonActivatingWindow win)
        {
            ActiveDragWindow = win;
        }

        public static void EndDrag(NonActivatingWindow win)
        {
            try
            {
                if (ActiveDragWindow != win) return;

                string myId = Framemanager.GetFrameIdFromWindow(win);
                if (myId != null && FrameDataManager.DockingMap.TryGetValue(myId, out List<string> parentIds))
                {
                    Framemanager.AnimateSnapConfirmation(win);
                    FrameDataManager.UpdateDockedRelationships(myId, parentIds);
                }
                else if (myId != null)
                {
                    Framemanager.ShowSnapPreview(win, false);
                    FrameDataManager.UpdateDockedRelationships(myId, null);
                }
            }
            finally
            {
                ActiveDragWindow = null;
            }
        }

        public static void AddSnapping(NonActivatingWindow win, IDictionary<string, object> FrameData)
        {
            string myId = FrameData.ContainsKey("Id") ? FrameData["Id"].ToString() : null;

            win.PreviewMouseLeftButtonUp += (sender, e) =>
            {
                if (myId != null && FrameDataManager.DockingMap.TryGetValue(myId, out List<string> parentIds))
                {
                    Framemanager.AnimateSnapConfirmation(win);
                    FrameDataManager.UpdateDockedRelationships(myId, parentIds);
                }
                else if (myId != null)
                {
                    Framemanager.ShowSnapPreview(win, false);
                    FrameDataManager.UpdateDockedRelationships(myId, null);
                }
            };

            win.LocationChanged += (sender, e) =>
            {
                if (_isSnapping) return;
                if (ActiveDragWindow != win) return;

                _isSnapping = true;
                try
                {
                    var allFrames = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList();
                    var (newLeft, newTop) = CalculateSnapPosition(win, allFrames);

                    if (Math.Abs(win.Left - newLeft) > 0.1 || Math.Abs(win.Top - newTop) > 0.1)
                    {
                        win.Left = newLeft;
                        win.Top = newTop;
                        FrameData["X"] = newLeft;
                        FrameData["Y"] = newTop;
                        FrameDataManager.SaveFrameData();

                        if (myId != null)
                        {
                            // Find ALL co-parents sitting above this window that overlap horizontally by at least 20%
                            var parents = allFrames.Where(f =>
                            {
                                if (f == win) return false;
                                bool verticalMatch = Math.Abs(f.Top + f.Height + MinGap - newTop) < SnapThreshold;
                                double overlapWidth = Math.Min(f.Left + f.Width, newLeft + win.Width) - Math.Max(f.Left, newLeft);
                                return verticalMatch && overlapWidth > (Math.Min(f.Width, win.Width) * 0.2);
                            }).ToList();

                            var parentIds = parents.Select(p => Framemanager.GetFrameIdFromWindow(p)).Where(id => id != null).ToList();

                            if (parentIds.Count > 0)
                            {
                                FrameDataManager.DockingMap[myId] = parentIds;
                                Framemanager.ShowSnapPreview(win, true);
                            }
                            else
                            {
                                FrameDataManager.DockingMap.Remove(myId);
                                Framemanager.ShowSnapPreview(win, false);
                            }
                        }
                    }
                }
                finally
                {
                    _isSnapping = false;
                }
            };
        }



        private static (double, double) CalculateSnapPosition(NonActivatingWindow current, List<NonActivatingWindow> allFrames)
        {
            if (!SettingsManager.IsSnapEnabled) return (current.Left, current.Top);

            double currentLeft = current.Left;
            double currentTop = current.Top;
            double currentRight = currentLeft + current.Width;
            double currentBottom = currentTop + current.Height;

            // We look for the SMALLEST adjustment needed to snap
            double minDeltaX = double.MaxValue;
            double minDeltaY = double.MaxValue;

            // 1. Snap to Other Frames
            foreach (var other in allFrames)
            {
                if (other == current) continue;

                double otherLeft = other.Left;
                double otherTop = other.Top;
                double otherRight = otherLeft + other.Width;
                double otherBottom = otherTop + other.Height;

                // Horizontal Checks
                // Snap Right Side to Other's Left
                CheckSnap(currentRight, otherLeft - MinGap, ref minDeltaX);
                // Snap Left Side to Other's Right
                CheckSnap(currentLeft, otherRight + MinGap, ref minDeltaX);
                // Align Lefts
                CheckSnap(currentLeft, otherLeft, ref minDeltaX);
                // Align Rights
                CheckSnap(currentRight, otherRight, ref minDeltaX);

                // Vertical Checks
                // Snap Bottom to Other's Top
                CheckSnap(currentBottom, otherTop - MinGap, ref minDeltaY);
                // Snap Top to Other's Bottom
                CheckSnap(currentTop, otherBottom + MinGap, ref minDeltaY);
                // Align Tops
                CheckSnap(currentTop, otherTop, ref minDeltaY);
                // Align Bottoms
                CheckSnap(currentBottom, otherBottom, ref minDeltaY);
            }

            // 2. Snap to Screen Edges (DPI Aware)
            // We need to get the DPI scale factor. Assuming uniform scaling for simplicity, 
            // but ideally should be per-monitor.
            double dpiScale = GetDpiScale(current);

            foreach (var screen in Screen.AllScreens)
            {
                // Convert Pixel bounds to WPF Coordinates
                double sLeft = screen.Bounds.Left / dpiScale;
                double sTop = screen.Bounds.Top / dpiScale;
                double sRight = screen.Bounds.Right / dpiScale;
                double sBottom = screen.Bounds.Bottom / dpiScale;

                // Horizontal Screen Snaps
                CheckSnap(currentLeft, sLeft, ref minDeltaX);
                CheckSnap(currentRight, sRight, ref minDeltaX);

                // Vertical Screen Snaps
                CheckSnap(currentTop, sTop, ref minDeltaY);
                CheckSnap(currentBottom, sBottom, ref minDeltaY);
            }

            // 3. Apply the smallest valid delta found
            double finalX = (Math.Abs(minDeltaX) < double.MaxValue) ? currentLeft + minDeltaX : currentLeft;
            double finalY = (Math.Abs(minDeltaY) < double.MaxValue) ? currentTop + minDeltaY : currentTop;

            return (finalX, finalY);
        }

        // Helper to check if a snap point is closer than the current best
        private static void CheckSnap(double currentPos, double targetPos, ref double minDelta)
        {
            double delta = targetPos - currentPos;

            // Check if within threshold AND closer than any previous match
            if (Math.Abs(delta) <= SnapThreshold && Math.Abs(delta) < Math.Abs(minDelta))
            {
                minDelta = delta;
            }
        }

        // Helper to get DPI scaling
        private static double GetDpiScale(Visual visual)
        {
            try
            {
                var source = PresentationSource.FromVisual(visual);
                if (source != null && source.CompositionTarget != null)
                {
                    return source.CompositionTarget.TransformToDevice.M11;
                }
            }
            catch { }
            return 1.0; // Default if fails
        }
    }
}
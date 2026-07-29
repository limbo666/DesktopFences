using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
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

                string myId = GetFrameIdFromWindow(win);
                if (myId != null && FrameDataManager.DockingMap.TryGetValue(myId, out List<string> parentIds))
                {
                    AnimateSnapConfirmation(win);
                    FrameDataManager.UpdateDockedRelationships(myId, parentIds);
                }
                else if (myId != null)
                {
                    ShowSnapPreview(win, false);
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
                    AnimateSnapConfirmation(win);
                    FrameDataManager.UpdateDockedRelationships(myId, parentIds);
                }
                else if (myId != null)
                {
                    ShowSnapPreview(win, false);
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

                            var parentIds = parents.Select(p => GetFrameIdFromWindow(p)).Where(id => id != null).ToList();

                            if (parentIds.Count > 0)
                            {
                                FrameDataManager.DockingMap[myId] = parentIds;
                                ShowSnapPreview(win, true);
                            }
                            else
                            {
                                FrameDataManager.DockingMap.Remove(myId);
                                ShowSnapPreview(win, false);
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
        // --- CONSTRAINT-BASED MULTI-PARENT STACK RESOLVER ---
        public static void CascadeStack(string parentId, double deltaY)
        {
            // Find all children that list this parentId in their co-parent list
            var childrenIds = FrameDataManager.DockingMap
                .Where(kvp => kvp.Value != null && kvp.Value.Contains(parentId))
                .Select(kvp => kvp.Key)
                .ToList();

            foreach (var childId in childrenIds)
            {
                var win = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>()
                    .FirstOrDefault(w => GetFrameIdFromWindow(w) == childId);

                if (win != null && FrameDataManager.DockingMap.TryGetValue(childId, out List<string> parentIds))
                {
                    // Find all currently active window instances for all co-parents of this child
                    var activeParents = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>()
                        .Where(w => parentIds.Contains(GetFrameIdFromWindow(w)))
                        .ToList();

                    if (activeParents.Count > 0)
                    {
                        // The golden geometric rule: Anchor below the lowest unrolled bottom edge among all co-parents
                        double maxParentBottom = activeParents.Max(p => p.Top + p.Height);
                        double targetTop = maxParentBottom + 10.0; // Standard 10px snap gap

                        if (Math.Abs(win.Top - targetTop) > 0.5)
                        {
                            double actualDeltaY = targetTop - win.Top;
                            win.Top = targetTop;

                            // Recursively cascade downstream to any frames docked beneath this child
                            CascadeStack(childId, actualDeltaY);
                        }
                    }
                }
            }
        }
        public static string GetFrameIdFromWindow(NonActivatingWindow win)
        {
            return win?.Tag?.ToString();
        }
        // --- ACCORDION SNAP FEEDBACK ENGINE ---
        // Holds a vibrant border pulse for 350ms before smoothly fading out over 650ms (1000ms total)
        // --- INTERACTIVE SNAP FEEDBACK ENGINE ---
        private static readonly Dictionary<NonActivatingWindow, System.Windows.Media.Brush> _snapOrigBrushes = new Dictionary<NonActivatingWindow, System.Windows.Media.Brush>();
        private static readonly Dictionary<NonActivatingWindow, Thickness> _snapOrigThicknesses = new Dictionary<NonActivatingWindow, Thickness>();
        private static readonly HashSet<NonActivatingWindow> _inSnapPreview = new HashSet<NonActivatingWindow>();
        public static void ShowSnapPreview(NonActivatingWindow win, bool isSnapped)
        {
            try
            {
                if (win?.Content is not Border border) return;

                if (isSnapped)
                {
                    if (!_inSnapPreview.Contains(win))
                    {
                        _inSnapPreview.Add(win);
                        _snapOrigBrushes[win] = border.BorderBrush;
                        _snapOrigThicknesses[win] = border.BorderThickness;

                        border.BeginAnimation(Border.BorderBrushProperty, null);
                        border.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 210, 255));
                        border.BorderThickness = new Thickness(Math.Max(3, _snapOrigThicknesses[win].Top + 2));
                    }
                }
                else if (_inSnapPreview.Contains(win))
                {
                    _inSnapPreview.Remove(win);
                    border.BeginAnimation(Border.BorderBrushProperty, null);
                    if (_snapOrigBrushes.TryGetValue(win, out System.Windows.Media.Brush origB)) border.BorderBrush = origB;
                    if (_snapOrigThicknesses.TryGetValue(win, out Thickness origT)) border.BorderThickness = origT;
                }
            }
            catch { }
        }
        public static void AnimateSnapConfirmation(NonActivatingWindow win)
        {
            try
            {
                if (win?.Content is not Border border || !_inSnapPreview.Contains(win)) return;
                _inSnapPreview.Remove(win);

                System.Windows.Media.Brush origBrush = _snapOrigBrushes.ContainsKey(win) ? _snapOrigBrushes[win] : border.BorderBrush;
                Thickness origThick = _snapOrigThicknesses.ContainsKey(win) ? _snapOrigThicknesses[win] : border.BorderThickness;

                var pulseBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 210, 255));
                border.BorderBrush = pulseBrush;
                border.BorderThickness = origThick;

                var fadePulse = new System.Windows.Media.Animation.DoubleAnimation
                {
                    From = 1.0,
                    To = 0.0,
                    Duration = TimeSpan.FromMilliseconds(500),
                    EasingFunction = new System.Windows.Media.Animation.QuadraticEase { EasingMode = System.Windows.Media.Animation.EasingMode.EaseOut }
                };

                fadePulse.Completed += (s, e) =>
                {
                    pulseBrush.BeginAnimation(System.Windows.Media.Brush.OpacityProperty, null);
                    border.BorderBrush = origBrush;
                };

                pulseBrush.BeginAnimation(System.Windows.Media.Brush.OpacityProperty, fadePulse);
            }
            catch { }
        }
    }
}
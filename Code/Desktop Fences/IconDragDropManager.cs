using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{
    /// <summary>
    /// Manages drag and drop operations for icon reordering within Data fences.
    /// Updated to be TAB-AWARE (supports reordering inside specific tabs).
    /// </summary>
    public static class IconDragDropManager
    {
        #region Win32 API for cursor position
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }
        #endregion

        #region Private Fields
        // Drag and drop state management for icon reordering
        private static bool _isDragging = false;
        private static StackPanel _draggedIcon = null;
        private static System.Windows.Point _dragStartPoint;

        private static dynamic _draggedItem = null;
        private static dynamic _sourceFence = null;
        private static JArray _sourceItemsList = null; // FIX: The specific list we are editing (Main or Tab)

        private static WrapPanel _sourceWrapPanel = null;
        private static Window _dragPreviewWindow = null;
        private static System.Windows.Point _lastDropIndicatorPosition = new System.Windows.Point(-1, -1);
        private static int _lastDropIndicatorIndex = -1;
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets whether a drag operation is currently in progress
        /// </summary>
        public static bool IsDragging => _isDragging;
        #endregion

        #region Public Methods
        /// <summary>
        /// Starts a drag operation for icon reordering
        /// </summary>
        /// <param name="iconStackPanel">The icon being dragged</param>
        /// <param name="startPoint">The starting point of the drag</param>
        public static void StartIconDrag(StackPanel iconStackPanel, System.Windows.Point startPoint)
        {
            try
            {
                // Only allow dragging in Data fences, not Portal fences
                NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(iconStackPanel);
                if (parentWindow == null) return;

                string fenceId = parentWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                var fenceData = FenceManager.GetFenceData();
                dynamic fence = fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (fence == null || fence.ItemsType?.ToString() != "Data") return;

                // Find the WrapPanel containing the icons
                WrapPanel wrapPanel = FindWrapPanel(parentWindow);
                if (wrapPanel == null) return;

                // Get the dragged item data from the icon's Tag
                var tagData = iconStackPanel.Tag;
                if (tagData == null) return;

                string filePath = tagData.GetType().GetProperty("FilePath")?.GetValue(tagData)?.ToString();
                if (string.IsNullOrEmpty(filePath)) return;

                // --- FIX: TAB-AWARE LIST SELECTION ---
                // Determine which JArray we are modifying (Main Items vs Active Tab Items)
                JArray targetList = null;
                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                if (tabsEnabled)
                {
                    var tabs = fence.Tabs as JArray;
                    int currentTabIndex = Convert.ToInt32(fence.CurrentTab?.ToString() ?? "0");

                    if (tabs != null && currentTabIndex >= 0 && currentTabIndex < tabs.Count)
                    {
                        var activeTab = tabs[currentTabIndex] as JObject;
                        targetList = activeTab?["Items"] as JArray;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Drag started in Tab {currentTabIndex}");
                    }
                }

                // Fallback to Main Items if tabs disabled or invalid
                if (targetList == null)
                {
                    targetList = fence.Items as JArray;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Drag started in Main Items");
                }

                if (targetList == null) return;

                // Find the specific item in the specific list
                dynamic draggedItem = null;
                foreach (var item in targetList)
                {
                    if (item["Filename"]?.ToString() == filePath)
                    {
                        draggedItem = item;
                        break;
                    }
                }

                if (draggedItem == null)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Cannot start drag: Item {filePath} not found in the active list.");
                    return;
                }

                // Set drag state
                _isDragging = true;
                _draggedIcon = iconStackPanel;
                _dragStartPoint = startPoint;
                _draggedItem = draggedItem;
                _sourceFence = fence;
                _sourceItemsList = targetList; // Store the specific list reference!
                _sourceWrapPanel = wrapPanel;

                // Capture mouse
                iconStackPanel.CaptureMouse();

                // Focus parent for Key events (Escape)
                if (parentWindow.Focusable) parentWindow.Focus();

                // Create visual drag preview
                CreateDragPreview(iconStackPanel);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Started drag for {filePath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error starting drag: {ex.Message}");
                CancelDrag();
            }
        }

        /// <summary>
        /// Cancels the current drag operation
        /// </summary>
        public static void CancelDrag()
        {
            try
            {
                if (_isDragging)
                {
                    if (_draggedIcon != null) _draggedIcon.ReleaseMouseCapture();

                    if (_dragPreviewWindow != null)
                    {
                        _dragPreviewWindow.Close();
                        _dragPreviewWindow = null;
                    }

                    if (_sourceWrapPanel != null) RemoveDropZoneIndicators(_sourceWrapPanel);

                    _isDragging = false;
                    _draggedIcon = null;
                    _draggedItem = null;
                    _sourceFence = null;
                    _sourceItemsList = null;
                    _sourceWrapPanel = null;
                    _lastDropIndicatorPosition = new System.Windows.Point(-1, -1);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error cancelling drag: {ex.Message}");
                // Force reset state
                _isDragging = false;
            }
        }

        /// <summary>
        /// Handles mouse move during drag operation
        /// </summary>
        public static void HandleDragMove(System.Windows.Point screenPosition)
        {
            if (!_isDragging || _draggedIcon == null) return;

            try
            {
                UpdateDragPreviewPosition(screenPosition);

                if (_sourceWrapPanel != null)
                {
                    System.Windows.Point wrapPanelPosition = _sourceWrapPanel.PointFromScreen(screenPosition);
                    ShowDropZoneIndicators(_sourceWrapPanel, wrapPanelPosition);
                }
            }
            catch { }
        }

        /// <summary>
        /// Completes the drag operation and performs reordering
        /// </summary>
        public static void CompleteDrag(System.Windows.Point finalPosition)
        {
            if (!_isDragging || _draggedIcon == null || _sourceWrapPanel == null) return;

            try
            {
                // Calculate where to drop the item relative to the panel
                int dropPosition = CalculateDropPosition(_sourceWrapPanel, finalPosition);

                // Perform the reordering on the specific list
                ReorderFenceItems(dropPosition);

                CancelDrag();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error completing drag: {ex.Message}");
                CancelDrag();
            }
        }
        #endregion

        #region Reordering Logic (The Core Fix)

        private static int CalculateDropPosition(WrapPanel wrapPanel, System.Windows.Point mousePosition)
        {
            try
            {
                if (wrapPanel == null || _draggedIcon == null) return 0;

                var iconPanels = wrapPanel.Children.OfType<StackPanel>().Where(sp => sp != _draggedIcon).ToList();
                if (iconPanels.Count == 0) return 0;

                double closestDistance = double.MaxValue;
                int bestInsertIndex = 0;

                // We assume items in the WrapPanel match the order in _sourceItemsList
                // But we must be careful if the visual list and data list are out of sync.
                // Best bet is to find the index of the closest icon in the source list.

                for (int i = 0; i < iconPanels.Count; i++)
                {
                    var iconPanel = iconPanels[i];
                    try
                    {
                        var iconPosition = iconPanel.TranslatePoint(new System.Windows.Point(0, 0), wrapPanel);
                        var iconCenter = new System.Windows.Point(
                            iconPosition.X + iconPanel.ActualWidth / 2,
                            iconPosition.Y + iconPanel.ActualHeight / 2
                        );

                        double distance = Math.Sqrt(Math.Pow(mousePosition.X - iconCenter.X, 2) + Math.Pow(mousePosition.Y - iconCenter.Y, 2));

                        if (distance < closestDistance)
                        {
                            closestDistance = distance;
                            bool insertBefore = mousePosition.X < iconCenter.X;

                            // Find this visual icon's corresponding data index
                            var tagData = iconPanel.Tag;
                            string filePath = tagData?.GetType().GetProperty("FilePath")?.GetValue(tagData)?.ToString();

                            int dataIndex = -1;
                            if (_sourceItemsList != null && !string.IsNullOrEmpty(filePath))
                            {
                                for (int k = 0; k < _sourceItemsList.Count; k++)
                                {
                                    if (_sourceItemsList[k]["Filename"]?.ToString() == filePath)
                                    {
                                        dataIndex = k;
                                        break;
                                    }
                                }
                            }

                            if (dataIndex != -1)
                            {
                                bestInsertIndex = insertBefore ? dataIndex : dataIndex + 1;
                            }
                            else
                            {
                                // Fallback: Use visual index if data match fails
                                bestInsertIndex = insertBefore ? i : i + 1;
                            }
                        }
                    }
                    catch { }
                }

                // Bounds Check
                int maxCount = _sourceItemsList?.Count ?? 0;
                return Math.Max(0, Math.Min(bestInsertIndex, maxCount));
            }
            catch
            {
                return 0;
            }
        }

        private static void ReorderFenceItems(int newPosition)
        {
            try
            {
                // FIX: Use _sourceItemsList instead of _sourceFence.Items
                if (_sourceItemsList == null || _draggedItem == null) return;

                int currentPosition = -1;
                for (int i = 0; i < _sourceItemsList.Count; i++)
                {
                    if (_sourceItemsList[i]["Filename"]?.ToString() == _draggedItem["Filename"]?.ToString())
                    {
                        currentPosition = i;
                        break;
                    }
                }

                if (currentPosition == -1) return;
                if (currentPosition == newPosition || (currentPosition + 1 == newPosition)) return;

                // Move logic
                var itemToMove = _sourceItemsList[currentPosition];
                _sourceItemsList.RemoveAt(currentPosition);

                int adjustedPosition = newPosition;
                if (currentPosition < newPosition) adjustedPosition--;

                adjustedPosition = Math.Max(0, Math.Min(adjustedPosition, _sourceItemsList.Count));
                _sourceItemsList.Insert(adjustedPosition, itemToMove);

                // Update DisplayOrder
                for (int i = 0; i < _sourceItemsList.Count; i++)
                {
                    _sourceItemsList[i]["DisplayOrder"] = i;
                }

                FenceDataManager.SaveFenceData();

                // Refresh UI
                RefreshFenceUI();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error reordering items: {ex.Message}");
            }
        }

        private static void RefreshFenceUI()
        {
            try
            {
                if (_sourceFence == null || _sourceWrapPanel == null) return;

                NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(_sourceWrapPanel);
                if (parentWindow == null) return;

                Application.Current.Dispatcher.Invoke(() =>
                {
                    // FIX: Delegate to FenceManager's robust refresh logic
                    // This handles Tabs vs Main logic automatically
                    FenceManager.RefreshFenceUsingFormApproach(parentWindow, _sourceFence);
                });
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing fence UI: {ex.Message}");
            }
        }

        #endregion

        #region Private Helper Methods (Visuals & Utils)

        // ... (Standard FindVisualParent, FindWrapPanel, GetCursorPos, DragPreview logic remains same) ...
        // Included for completeness to ensure the file compiles without missing refs

        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return parent as T;
        }

        private static WrapPanel FindWrapPanel(DependencyObject parent, int depth = 0, int maxDepth = 10)
        {
            if (parent == null || depth > maxDepth) return null;
            if (parent is WrapPanel wrapPanel) return wrapPanel;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var result = FindWrapPanel(VisualTreeHelper.GetChild(parent, i), depth + 1, maxDepth);
                if (result != null) return result;
            }
            return null;
        }

        private static System.Windows.Point GetCursorPosition()
        {
            POINT point;
            GetCursorPos(out point);
            return new System.Windows.Point(point.X, point.Y);
        }

        private static double GetDpiScaleFactor(Window window)
        {
            var screen = System.Windows.Forms.Screen.FromPoint(new System.Drawing.Point((int)window.Left, (int)window.Top));
            using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                return graphics.DpiX / 96.0;
            }
        }

        private static void CreateDragPreview(StackPanel originalIcon)
        {
            try
            {
                if (_dragPreviewWindow != null)
                {
                    _dragPreviewWindow.Close();
                    _dragPreviewWindow = null;
                }

                NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(originalIcon);
                double dpiScale = parentWindow != null ? GetDpiScaleFactor(parentWindow) : 1.0;

                _dragPreviewWindow = new Window
                {
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = System.Windows.Media.Brushes.Transparent,
                    ShowInTaskbar = false,
                    Topmost = true,
                    Width = originalIcon.ActualWidth > 0 ? originalIcon.ActualWidth : 60,
                    Height = originalIcon.ActualHeight > 0 ? originalIcon.ActualHeight : 80,
                    IsHitTestVisible = false,
                    WindowStartupLocation = WindowStartupLocation.Manual
                };

                // Clone visual content for preview
                StackPanel previewContent = new StackPanel { Width = originalIcon.Width, Margin = originalIcon.Margin, Opacity = 0.7 };

                // Copy Image/Grid
                var originalGrid = originalIcon.Children.OfType<Grid>().FirstOrDefault();
                var originalImage = originalIcon.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();

                if (originalGrid != null)
                {
                    // Clone Grid (Network Icon)
                    Grid previewGrid = new Grid { Width = originalGrid.Width, Height = originalGrid.Height, Margin = originalGrid.Margin };
                    var gridImage = originalGrid.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                    if (gridImage != null) previewGrid.Children.Add(new System.Windows.Controls.Image { Source = gridImage.Source, Width = gridImage.Width, Height = gridImage.Height, Margin = gridImage.Margin });

                    var netInd = originalGrid.Children.OfType<TextBlock>().FirstOrDefault();
                    if (netInd != null) previewGrid.Children.Add(new TextBlock { Text = netInd.Text, FontSize = netInd.FontSize, Foreground = netInd.Foreground, Margin = netInd.Margin });

                    previewContent.Children.Add(previewGrid);
                }
                else if (originalImage != null)
                {
                    previewContent.Children.Add(new System.Windows.Controls.Image { Source = originalImage.Source, Width = originalImage.Width, Height = originalImage.Height, Margin = originalImage.Margin });
                }

                // Copy Label
                var originalLabel = originalIcon.Children.OfType<TextBlock>().FirstOrDefault();
                if (originalLabel != null)
                {
                    previewContent.Children.Add(new TextBlock
                    {
                        Text = originalLabel.Text,
                        TextWrapping = originalLabel.TextWrapping,
                        TextAlignment = originalLabel.TextAlignment,
                        Foreground = originalLabel.Foreground,
                        Width = originalLabel.Width
                    });
                }

                _dragPreviewWindow.Content = previewContent;

                System.Windows.Point cursorPos = GetCursorPosition();
                _dragPreviewWindow.Left = (cursorPos.X / dpiScale) + 10;
                _dragPreviewWindow.Top = (cursorPos.Y / dpiScale) - 10;
                _dragPreviewWindow.Show();
            }
            catch { }
        }

        private static void UpdateDragPreviewPosition(System.Windows.Point screenPosition)
        {
            if (_dragPreviewWindow != null)
            {
                double dpiScale = 1.0; // Simplified for speed, usually sufficient
                _dragPreviewWindow.Left = (screenPosition.X / dpiScale) + 10;
                _dragPreviewWindow.Top = (screenPosition.Y / dpiScale) - 10;
            }
        }

        private static void ShowDropZoneIndicators(WrapPanel wrapPanel, System.Windows.Point mousePosition)
        {
            // Simple optimization
            if ((mousePosition - _lastDropIndicatorPosition).Length < 15) return;
            _lastDropIndicatorPosition = mousePosition;

            RemoveDropZoneIndicators(wrapPanel);

            var iconPanels = wrapPanel.Children.OfType<StackPanel>().Where(sp => sp != _draggedIcon).ToList();
            if (iconPanels.Count == 0) return;

            StackPanel closestIcon = null;
            double closestDist = double.MaxValue;
            bool insertBefore = true;

            foreach (var panel in iconPanels)
            {
                var pos = panel.TranslatePoint(new System.Windows.Point(0, 0), wrapPanel);
                var center = new System.Windows.Point(pos.X + panel.ActualWidth / 2, pos.Y + panel.ActualHeight / 2);
                double dist = (mousePosition - center).Length;

                if (dist < closestDist)
                {
                    closestDist = dist;
                    closestIcon = panel;
                    insertBefore = mousePosition.X < center.X;
                }
            }

            if (closestIcon != null)
            {
                var indicator = new Border
                {
                    Width = 3,
                    Height = closestIcon.ActualHeight,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(200, 0, 150, 255)),
                    CornerRadius = new CornerRadius(1.5),
                    Tag = "DropIndicator",
                    Margin = new Thickness(2, 5, 2, 5),
                    Effect = new DropShadowEffect { Color = System.Windows.Media.Color.FromRgb(0, 150, 255), BlurRadius = 8, ShadowDepth = 0 }
                };

                int idx = wrapPanel.Children.IndexOf(closestIcon);
                if (insertBefore) wrapPanel.Children.Insert(idx, indicator);
                else wrapPanel.Children.Insert(idx + 1, indicator);
            }
        }

        private static void RemoveDropZoneIndicators(WrapPanel wrapPanel)
        {
            if (wrapPanel == null) return;
            var toRemove = wrapPanel.Children.OfType<Border>().Where(b => "DropIndicator".Equals(b.Tag?.ToString())).ToList();
            foreach (var item in toRemove) wrapPanel.Children.Remove(item);
        }

        #endregion
    }
}
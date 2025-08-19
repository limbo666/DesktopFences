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
    /// Manages drag and drop operations for icon reordering within Data fences
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
                if (parentWindow == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cannot start drag: parent window not found");
                    return;
                }

                // Find the fence data for this window
                string fenceId = parentWindow.Tag?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cannot start drag: fence ID not found");
                    return;
                }

                var fenceData = FenceManager.GetFenceData();
                dynamic fence = fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                if (fence == null || fence.ItemsType?.ToString() != "Data")
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Cannot start drag: fence not found or not Data type");
                    return;
                }

                // Find the WrapPanel containing the icons
                WrapPanel wrapPanel = FindWrapPanel(parentWindow);
                if (wrapPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cannot start drag: WrapPanel not found");
                    return;
                }

                // Get the dragged item data from the icon's Tag
                var tagData = iconStackPanel.Tag;
                if (tagData == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cannot start drag: icon Tag data not found");
                    return;
                }

                // Extract file path from Tag (which is an anonymous object with FilePath property)
                string filePath = tagData.GetType().GetProperty("FilePath")?.GetValue(tagData)?.ToString();
                if (string.IsNullOrEmpty(filePath))
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cannot start drag: file path not found in Tag");
                    return;
                }

                // Find the corresponding item in fence data
                var items = fence.Items as JArray;
                if (items == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cannot start drag: fence items not found");
                    return;
                }

                dynamic draggedItem = null;
                foreach (var item in items)
                {
                    if (item["Filename"]?.ToString() == filePath)
                    {
                        draggedItem = item;
                        break;
                    }
                }

                if (draggedItem == null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Cannot start drag: item {filePath} not found in fence data");
                    return;
                }

                // Set drag state
                _isDragging = true;
                _draggedIcon = iconStackPanel;
                _dragStartPoint = startPoint;
                _draggedItem = draggedItem;
                _sourceFence = fence;
                _sourceWrapPanel = wrapPanel;

                // Capture mouse for drag tracking
                iconStackPanel.CaptureMouse();

                // Ensure parent window can receive key events
                if (parentWindow != null && parentWindow.Focusable)
                {
                    parentWindow.Focus();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Set focus to parent window for Escape key handling");
                }

                // Get cursor position first
                System.Windows.Point cursorPosition = GetCursorPosition();

                // Create visual drag preview (it will position itself at cursor)
                CreateDragPreview(iconStackPanel);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Initial drag preview setup: cursor({cursorPosition.X:F1}, {cursorPosition.Y:F1})");

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Started drag operation for icon {filePath} in fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error starting drag operation: {ex.Message}");
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
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Cancelling drag operation");

                    // Release mouse capture
                    if (_draggedIcon != null)
                    {
                        _draggedIcon.ReleaseMouseCapture();
                    }

                    // Clean up drag preview window
                    if (_dragPreviewWindow != null)
                    {
                        _dragPreviewWindow.Close();
                        _dragPreviewWindow = null;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Closed drag preview window");
                    }

                    // Remove drop zone indicators
                    if (_sourceWrapPanel != null)
                    {
                        RemoveDropZoneIndicators(_sourceWrapPanel);
                    }

                    // Reset drag state
                    _isDragging = false;
                    _draggedIcon = null;
                    _draggedItem = null;
                    _sourceFence = null;
                    _sourceWrapPanel = null;
                    _lastDropIndicatorPosition = new System.Windows.Point(-1, -1);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error cancelling drag operation: {ex.Message}");
            }
            finally
            {
                // Ensure state is reset even if there's an error
                _isDragging = false;
                _draggedIcon = null;
                _draggedItem = null;
                _sourceFence = null;
                _sourceWrapPanel = null;
            }
        }

        /// <summary>
        /// Handles mouse move during drag operation
        /// </summary>
        /// <param name="screenPosition">Current screen position</param>
        public static void HandleDragMove(System.Windows.Point screenPosition)
        {
            if (!_isDragging || _draggedIcon == null) return;

            try
            {
                // Update drag preview position to follow cursor
                UpdateDragPreviewPosition(screenPosition);

                // Show drop zone indicators
                if (_sourceWrapPanel != null)
                {
                    System.Windows.Point wrapPanelPosition = _sourceWrapPanel.PointFromScreen(screenPosition);
                    ShowDropZoneIndicators(_sourceWrapPanel, wrapPanelPosition);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error during drag move: {ex.Message}");
            }
        }

        /// <summary>
        /// Completes the drag operation and performs reordering
        /// </summary>
        /// <param name="finalPosition">Final drop position</param>
        public static void CompleteDrag(System.Windows.Point finalPosition)
        {
            if (!_isDragging || _draggedIcon == null || _sourceWrapPanel == null) return;

            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Completing drag operation for {_draggedItem?["Filename"]} at position ({finalPosition.X:F1}, {finalPosition.Y:F1})");

                // Calculate where to drop the item
                int dropPosition = CalculateDropPosition(_sourceWrapPanel, finalPosition);

                // Perform the reordering
                ReorderFenceItems(dropPosition);

                // Clean up visual elements
                CancelDrag();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error completing drag operation: {ex.Message}");
                CancelDrag();
            }
        }
        #endregion

        #region Private Helper Methods
        /// <summary>
        /// Finds a visual parent of the specified type
        /// </summary>
        private static T FindVisualParent<T>(DependencyObject child) where T : DependencyObject
        {
            DependencyObject parent = VisualTreeHelper.GetParent(child);
            while (parent != null && !(parent is T))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return parent as T;
        }

        /// <summary>
        /// Finds the WrapPanel containing icons in a window
        /// </summary>
        private static WrapPanel FindWrapPanel(DependencyObject parent, int depth = 0, int maxDepth = 10)
        {
            // Prevent infinite recursion
            if (parent == null || depth > maxDepth)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"FindWrapPanel: Reached max depth {maxDepth} or null parent at depth {depth}");
                return null;
            }

            // Check if current element is a WrapPanel
            if (parent is WrapPanel wrapPanel)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"FindWrapPanel: Found WrapPanel at depth {depth}");
                return wrapPanel;
            }

            // Recurse through visual tree
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                var result = FindWrapPanel(child, depth + 1, maxDepth);
                if (result != null)
                {
                    return result;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets accurate cursor position
        /// </summary>
        private static System.Windows.Point GetCursorPosition()
        {
            POINT point;
            GetCursorPos(out point);
            return new System.Windows.Point(point.X, point.Y);
        }

        /// <summary>
        /// Creates drag preview window
        /// </summary>
        private static void CreateDragPreview(StackPanel originalIcon)
        {
            try
            {
                if (_dragPreviewWindow != null)
                {
                    _dragPreviewWindow.Close();
                    _dragPreviewWindow = null;
                }

                // Find parent window to get DPI scale factor
                NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(originalIcon);
                double dpiScale = parentWindow != null ? GetDpiScaleFactor(parentWindow) : 1.0;

                // Create a semi-transparent copy of the dragged icon
                _dragPreviewWindow = new Window
                {
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = System.Windows.Media.Brushes.Transparent,
                    ShowInTaskbar = false,
                    Topmost = true,
                    Width = originalIcon.ActualWidth > 0 ? originalIcon.ActualWidth : 60,
                    Height = originalIcon.ActualHeight > 0 ? originalIcon.ActualHeight : 80,
                    IsHitTestVisible = false, // Allow mouse events to pass through
                    WindowStartupLocation = WindowStartupLocation.Manual // Prevent automatic positioning
                };

                // Create visual copy of the original icon
                StackPanel previewContent = new StackPanel
                {
                    Width = originalIcon.Width,
                    Margin = originalIcon.Margin,
                    Opacity = 0.7 // Semi-transparent
                };

                // Copy the icon image
                var originalImage = originalIcon.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                var originalGrid = originalIcon.Children.OfType<Grid>().FirstOrDefault(); // For network icons with overlay

                if (originalGrid != null)
                {
                    // Handle icons with network overlay (Grid containing Image + TextBlock)
                    Grid previewGrid = new Grid
                    {
                        Width = originalGrid.Width,
                        Height = originalGrid.Height,
                        Margin = originalGrid.Margin
                    };

                    // Copy the main image from the grid
                    var gridImage = originalGrid.Children.OfType<System.Windows.Controls.Image>().FirstOrDefault();
                    if (gridImage != null)
                    {
                        var previewImage = new System.Windows.Controls.Image
                        {
                            Source = gridImage.Source,
                            Width = gridImage.Width,
                            Height = gridImage.Height,
                            Margin = gridImage.Margin
                        };
                        previewGrid.Children.Add(previewImage);
                    }

                    // Copy the network indicator
                    var networkIndicator = originalGrid.Children.OfType<TextBlock>().FirstOrDefault();
                    if (networkIndicator != null)
                    {
                        var previewIndicator = new TextBlock
                        {
                            Text = networkIndicator.Text,
                            FontSize = networkIndicator.FontSize,
                            Foreground = networkIndicator.Foreground,
                            HorizontalAlignment = networkIndicator.HorizontalAlignment,
                            VerticalAlignment = networkIndicator.VerticalAlignment,
                            Margin = networkIndicator.Margin,
                            Effect = networkIndicator.Effect
                        };
                        previewGrid.Children.Add(previewIndicator);
                    }

                    previewContent.Children.Add(previewGrid);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Created drag preview with network indicator");
                }
                else if (originalImage != null)
                {
                    // Handle regular icons (just Image)
                    var previewImage = new System.Windows.Controls.Image
                    {
                        Source = originalImage.Source,
                        Width = originalImage.Width,
                        Height = originalImage.Height,
                        Margin = originalImage.Margin
                    };
                    previewContent.Children.Add(previewImage);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Created drag preview with regular image");
                }

                // Copy the text label
                var originalLabel = originalIcon.Children.OfType<TextBlock>().FirstOrDefault();
                if (originalLabel != null)
                {
                    var previewLabel = new TextBlock
                    {
                        Text = originalLabel.Text,
                        TextWrapping = originalLabel.TextWrapping,
                        TextTrimming = originalLabel.TextTrimming,
                        HorizontalAlignment = originalLabel.HorizontalAlignment,
                        Foreground = originalLabel.Foreground,
                        MaxWidth = originalLabel.MaxWidth,
                        Width = originalLabel.Width,
                        TextAlignment = originalLabel.TextAlignment,
                        Effect = originalLabel.Effect
                    };
                    previewContent.Children.Add(previewLabel);
                }

                _dragPreviewWindow.Content = previewContent;

                // Get current cursor position in physical pixels
                System.Windows.Point cursorPosPixels = GetCursorPosition();

                // Convert to device-independent units (DIUs) using DPI scale
                double cursorX_DIU = cursorPosPixels.X / dpiScale;
                double cursorY_DIU = cursorPosPixels.Y / dpiScale;

                // Set position in DIUs with offset
                _dragPreviewWindow.Left = cursorX_DIU + (10 / dpiScale);
                _dragPreviewWindow.Top = cursorY_DIU - (10 / dpiScale);

                // Show the window
                _dragPreviewWindow.Show();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Created drag preview window: {_dragPreviewWindow.Width}x{_dragPreviewWindow.Height}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error creating drag preview: {ex.Message}");
                if (_dragPreviewWindow != null)
                {
                    _dragPreviewWindow.Close();
                    _dragPreviewWindow = null;
                }
            }
        }

        /// <summary>
        /// Gets DPI scale factor for proper positioning
        /// </summary>
        private static double GetDpiScaleFactor(Window window)
        {
            // Get the screen where the window is located based on its position
            var screen = System.Windows.Forms.Screen.FromPoint(
                new System.Drawing.Point((int)window.Left, (int)window.Top));

            // Use Graphics to get the screen's DPI
            using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = graphics.DpiX; // Horizontal DPI
                return dpiX / 96.0; // Standard DPI is 96, so scale factor = dpiX / 96
            }
        }

        /// <summary>
        /// Updates drag preview position
        /// </summary>
        private static void UpdateDragPreviewPosition(System.Windows.Point screenPosition)
        {
            try
            {
                if (_dragPreviewWindow != null)
                {
                    // Get DPI scale factor (use cached from drag state or find from current window)
                    double dpiScale = 1.0;
                    if (_draggedIcon != null)
                    {
                        NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(_draggedIcon);
                        if (parentWindow != null)
                        {
                            dpiScale = GetDpiScaleFactor(parentWindow);
                        }
                    }

                    // Convert screen position (physical pixels) to DIUs
                    double newLeft = (screenPosition.X / dpiScale) + (10 / dpiScale);
                    double newTop = (screenPosition.Y / dpiScale) - (10 / dpiScale);

                    _dragPreviewWindow.Left = newLeft;
                    _dragPreviewWindow.Top = newTop;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error updating drag preview position: {ex.Message}");
            }
        }

        /// <summary>
        /// Shows drop zone indicators (optimized)
        /// </summary>
        private static void ShowDropZoneIndicators(WrapPanel wrapPanel, System.Windows.Point mousePosition)
        {
            try
            {
                // Skip update if mouse hasn't moved significantly (performance optimization)
                double distanceFromLast = Math.Sqrt(
                    Math.Pow(mousePosition.X - _lastDropIndicatorPosition.X, 2) +
                    Math.Pow(mousePosition.Y - _lastDropIndicatorPosition.Y, 2)
                );

                if (distanceFromLast < 15) // Only update if moved more than 15 pixels
                {
                    return;
                }

                _lastDropIndicatorPosition = mousePosition;

                // Remove existing drop indicators
                RemoveDropZoneIndicators(wrapPanel);

                // Calculate which icons are near the mouse position
                var iconPanels = wrapPanel.Children.OfType<StackPanel>().Where(sp => sp != _draggedIcon).ToList();

                if (iconPanels.Count == 0)
                {
                    return;
                }

                // Find the best drop position based on mouse coordinates
                StackPanel closestIcon = null;
                double closestDistance = double.MaxValue;
                bool insertBefore = true;

                foreach (var iconPanel in iconPanels)
                {
                    try
                    {
                        var iconPosition = iconPanel.TranslatePoint(new System.Windows.Point(0, 0), wrapPanel);
                        var iconCenter = new System.Windows.Point(
                            iconPosition.X + iconPanel.ActualWidth / 2,
                            iconPosition.Y + iconPanel.ActualHeight / 2
                        );

                        double distance = Math.Sqrt(
                            Math.Pow(mousePosition.X - iconCenter.X, 2) +
                            Math.Pow(mousePosition.Y - iconCenter.Y, 2)
                        );

                        if (distance < closestDistance)
                        {
                            closestDistance = distance;
                            closestIcon = iconPanel;
                            // Determine if we should insert before or after this icon
                            insertBefore = mousePosition.X < iconCenter.X;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Error calculating icon position: {ex.Message}");
                    }
                }

                // Create drop indicator near the closest icon
                if (closestIcon != null)
                {
                    CreateDropIndicator(wrapPanel, closestIcon, insertBefore);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing drop zone indicators: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a single drop indicator
        /// </summary>
        private static void CreateDropIndicator(WrapPanel wrapPanel, StackPanel targetIcon, bool insertBefore)
        {
            try
            {
                // Create visual drop indicator (vertical line)
                var dropIndicator = new Border
                {
                    Width = 3,
                    Height = targetIcon.ActualHeight > 0 ? targetIcon.ActualHeight : 60,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(200, 0, 150, 255)), // Semi-transparent blue
                    CornerRadius = new CornerRadius(1.5),
                    Tag = "DropIndicator", // Tag to identify drop indicators for removal
                    Margin = new Thickness(2, 5, 2, 5)
                };

                // Add glow effect to make it more visible
                dropIndicator.Effect = new DropShadowEffect
                {
                    Color = System.Windows.Media.Color.FromRgb(0, 150, 255),
                    BlurRadius = 8,
                    ShadowDepth = 0,
                    Opacity = 0.8
                };

                // Find the index where to insert the indicator
                int targetIndex = wrapPanel.Children.IndexOf(targetIcon);
                if (targetIndex >= 0)
                {
                    int insertIndex = insertBefore ? targetIndex : targetIndex + 1;
                    if (insertIndex <= wrapPanel.Children.Count)
                    {
                        wrapPanel.Children.Insert(insertIndex, dropIndicator);
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error creating drop indicator: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes all drop zone indicators
        /// </summary>
        private static void RemoveDropZoneIndicators(WrapPanel wrapPanel)
        {
            try
            {
                if (wrapPanel == null) return;

                // Remove all elements tagged as drop indicators
                var indicatorsToRemove = wrapPanel.Children.OfType<Border>()
                    .Where(b => "DropIndicator".Equals(b.Tag?.ToString()))
                    .ToList();

                foreach (var indicator in indicatorsToRemove)
                {
                    wrapPanel.Children.Remove(indicator);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error removing drop indicators: {ex.Message}");
            }
        }

        /// <summary>
        /// Calculates where to drop the dragged item
        /// </summary>
        private static int CalculateDropPosition(WrapPanel wrapPanel, System.Windows.Point mousePosition)
        {
            try
            {
                if (wrapPanel == null || _draggedIcon == null)
                {
                    return 0;
                }

                // Get all StackPanel children except the dragged one and drop indicators
                var iconPanels = wrapPanel.Children.OfType<StackPanel>()
                    .Where(sp => sp != _draggedIcon)
                    .ToList();

                if (iconPanels.Count == 0)
                {
                    return 0;
                }

                // Find the closest icon to determine drop position
                double closestDistance = double.MaxValue;
                int bestInsertIndex = 0;
                bool insertBefore = true;

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

                        double distance = Math.Sqrt(
                            Math.Pow(mousePosition.X - iconCenter.X, 2) +
                            Math.Pow(mousePosition.Y - iconCenter.Y, 2)
                        );

                        if (distance < closestDistance)
                        {
                            closestDistance = distance;
                            insertBefore = mousePosition.X < iconCenter.X;

                            // Calculate the actual index in the JSON array
                            var tagData = iconPanel.Tag;
                            if (tagData != null)
                            {
                                string filePath = tagData.GetType().GetProperty("FilePath")?.GetValue(tagData)?.ToString();
                                if (!string.IsNullOrEmpty(filePath))
                                {
                                    // Find this icon in the fence data
                                    var items = _sourceFence.Items as JArray;
                                    if (items != null)
                                    {
                                        for (int j = 0; j < items.Count; j++)
                                        {
                                            if (items[j]["Filename"]?.ToString() == filePath)
                                            {
                                                bestInsertIndex = insertBefore ? j : j + 1;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Error processing icon {i} for drop calculation: {ex.Message}");
                    }
                }

                // Ensure index is within bounds
                var totalItems = (_sourceFence.Items as JArray)?.Count ?? 0;
                bestInsertIndex = Math.Max(0, Math.Min(bestInsertIndex, totalItems));

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Calculated drop position: {bestInsertIndex} (insertBefore: {insertBefore}, distance: {closestDistance:F1})");
                return bestInsertIndex;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error calculating drop position: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Reorders items in fence data and updates UI
        /// </summary>
        private static void ReorderFenceItems(int newPosition)
        {
            try
            {
                if (_sourceFence == null || _draggedItem == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot reorder: missing source fence or dragged item");
                    return;
                }

                var items = _sourceFence.Items as JArray;
                if (items == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot reorder: fence items is null");
                    return;
                }

                // Find current position of dragged item
                int currentPosition = -1;
                for (int i = 0; i < items.Count; i++)
                {
                    if (items[i]["Filename"]?.ToString() == _draggedItem["Filename"]?.ToString())
                    {
                        currentPosition = i;
                        break;
                    }
                }

                if (currentPosition == -1)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Cannot find dragged item {_draggedItem["Filename"]} in fence data");
                    return;
                }

                // Don't reorder if dropping in the same position
                if (currentPosition == newPosition || (currentPosition + 1 == newPosition))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Item already in correct position ({currentPosition} -> {newPosition}), no reordering needed");
                    return;
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Reordering item '{_draggedItem["Filename"]}' from position {currentPosition} to {newPosition}");

                // Remove item from current position
                var itemToMove = items[currentPosition];
                items.RemoveAt(currentPosition);

                // Adjust insertion position if we removed an item before it
                int adjustedPosition = newPosition;
                if (currentPosition < newPosition)
                {
                    adjustedPosition = newPosition - 1;
                }

                // Insert item at new position
                adjustedPosition = Math.Max(0, Math.Min(adjustedPosition, items.Count));
                items.Insert(adjustedPosition, itemToMove);

                // Update DisplayOrder for all items to reflect new order
                for (int i = 0; i < items.Count; i++)
                {
                    items[i]["DisplayOrder"] = i;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Updated DisplayOrder: {items[i]["Filename"]} -> {i}");
                }

                // Save changes to JSON
                FenceManager.SaveFenceData();

                // Refresh the UI to reflect new order
                RefreshFenceUI();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Successfully reordered item to position {adjustedPosition}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error reordering fence items: {ex.Message}");
            }
        }

        /// <summary>
        /// Refreshes fence UI after reordering
        /// </summary>
        private static void RefreshFenceUI()
        {
            try
            {
                if (_sourceFence == null || _sourceWrapPanel == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot refresh UI: missing source fence or wrap panel");
                    return;
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Refreshing UI for fence '{_sourceFence.Title}'");

                // Find the parent window
                NonActivatingWindow parentWindow = FindVisualParent<NonActivatingWindow>(_sourceWrapPanel);
                if (parentWindow == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot refresh UI: parent window not found");
                    return;
                }

                // Use Application.Current.Dispatcher to ensure UI updates happen on the main thread
                Application.Current.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        // Clear current icons from WrapPanel
                        _sourceWrapPanel.Children.Clear();

                        // Recreate icons in new order
                        var items = _sourceFence.Items as JArray;
                        if (items != null)
                        {
                            // Sort items by DisplayOrder
                            var sortedItems = items
                                .OfType<JObject>()
                                .OrderBy(item => item["DisplayOrder"]?.Value<int>() ?? 0)
                                .ToList();

                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Recreating {sortedItems.Count} icons in new display order");

                            // Use FenceManager's methods to recreate the icons with proper event handlers
                            foreach (dynamic icon in sortedItems)
                            {
                                // Add the icon using FenceManager's AddIcon method
                                FenceManager.AddIcon(icon, _sourceWrapPanel);

                                // Get the newly added StackPanel
                                StackPanel sp = _sourceWrapPanel.Children[_sourceWrapPanel.Children.Count - 1] as StackPanel;
                                if (sp != null)
                                {
                                    // Extract icon properties
                                    var iconDict = icon is IDictionary<string, object> dict ? dict : ((JObject)icon).ToObject<IDictionary<string, object>>();
                                    string filePath = iconDict.ContainsKey("Filename") ? (string)iconDict["Filename"] : "Unknown";
                                    bool isFolder = iconDict.ContainsKey("IsFolder") && (bool)iconDict["IsFolder"];

                                    // Get arguments if it's a shortcut
                                    string arguments = null;
                                    if (System.IO.Path.GetExtension(filePath).ToLower() == ".lnk")
                                    {
                                        try
                                        {
                                            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                                            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(filePath);
                                            arguments = shortcut.Arguments;
                                        }
                                        catch (Exception ex)
                                        {
                                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Could not read shortcut arguments for {filePath}: {ex.Message}");
                                        }
                                    }

                                    // Add click event handler using FenceManager's method
                                    FenceManager.ClickEventAdder(sp, filePath, isFolder, arguments);

                                    // Create context menu (simplified version for drag/drop refresh)
                                    CreateBasicContextMenu(sp, icon, filePath);
                                }
                            }
                        }

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Successfully refreshed UI for fence '{_sourceFence.Title}'");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error in UI refresh dispatcher action: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing fence UI: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a basic context menu for refreshed icons
        /// </summary>
        private static void CreateBasicContextMenu(StackPanel sp, dynamic icon, string filePath)
        {
            try
            {
                ContextMenu mn = new ContextMenu();
                MenuItem miEdit = new MenuItem { Header = "Edit..." };
                MenuItem miRemove = new MenuItem { Header = "Remove" };

                mn.Items.Add(miEdit);
                mn.Items.Add(miRemove);

                // Add basic event handlers that delegate to FenceManager
                miEdit.Click += (sender, e) =>
                {
                    try
                    {
                        NonActivatingWindow parentWin = FindVisualParent<NonActivatingWindow>(sp);
                        if (parentWin != null)
                        {
                            // Find the current fence data
                            string fenceId = parentWin.Tag?.ToString();
                            if (!string.IsNullOrEmpty(fenceId))
                            {
                                var fenceData = FenceManager.GetFenceData();
                                dynamic currentFence = fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                                if (currentFence != null)
                                {
                                    // Use reflection to call FenceManager's EditItem method
                                    var editMethod = typeof(FenceManager).GetMethod("EditItem",
                                        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                                    if (editMethod != null)
                                    {
                                        editMethod.Invoke(null, new object[] { icon, currentFence, parentWin });
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error in edit handler: {ex.Message}");
                    }
                };

                miRemove.Click += (sender, e) =>
                {
                    try
                    {
                        // Find current fence safely
                        NonActivatingWindow parentWin = FindVisualParent<NonActivatingWindow>(sp);
                        if (parentWin != null)
                        {
                            string fenceId = parentWin.Tag?.ToString();
                            if (!string.IsNullOrEmpty(fenceId))
                            {
                                var fenceData = FenceManager.GetFenceData();
                                dynamic currentFence = fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);
                                if (currentFence != null)
                                {
                                    var itemsArray = currentFence.Items as JArray;
                                    if (itemsArray != null)
                                    {
                                        var itemToRemove = itemsArray.FirstOrDefault(i => i["Filename"]?.ToString() == filePath);
                                        if (itemToRemove != null)
                                        {
                                            // Remove with fade animation
                                            var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.3));
                                            fade.Completed += (s, a) =>
                                            {
                                                try
                                                {
                                                    itemsArray.Remove(itemToRemove);
                                                    WrapPanel wrapPanel = FindVisualParent<WrapPanel>(sp);
                                                    if (wrapPanel != null)
                                                    {
                                                        wrapPanel.Children.Remove(sp);
                                                    }
                                                    FenceManager.SaveFenceData();
                                                }
                                                catch (Exception ex)
                                                {
                                                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error during item removal: {ex.Message}");
                                                }
                                            };
                                            sp.BeginAnimation(UIElement.OpacityProperty, fade);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error in remove handler: {ex.Message}");
                    }
                };

                sp.ContextMenu = mn;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error creating basic context menu: {ex.Message}");
            }
        }
        #endregion
    }
}
using Desktop_Frames.Localization;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace Desktop_Frames
{
    /// <summary>
    /// Tab strip UI, lifecycle (add/rename/delete/switch/move), Tab0 sync, and tab button styling.
    /// Extracted from Framemanager (cut #1 of the FrameManager decomposition roadmap).
    /// Pure move + call-site qualifier changes; no behavior change.
    /// </summary>
    public static class TabManager
    {
        // TABS FEATURE: Dynamically refresh the tab strip UI (Ghost Arrows)
        // v2.5. 4.181: Adds smart left/right chevrons that appear only when tabs are hidden.
        public static void RefreshTabStripUI(NonActivatingWindow frameWindow, dynamic frame)
        {
            try
            {
                var border = frameWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                // 1. CAPTURE SCROLL STATE
                double previousScrollOffset = 0;
                var oldContainer = dockPanel.Children.OfType<Grid>()
                    .FirstOrDefault(g => g.Tag?.ToString() == "TAB_STRIP_CONTAINER");
                if (oldContainer != null)
                {
                    var oldScroll = oldContainer.Children.OfType<ScrollViewer>().FirstOrDefault();
                    if (oldScroll != null) previousScrollOffset = oldScroll.HorizontalOffset;
                }

                // 2. CLEANUP
                var existingStrips = dockPanel.Children.OfType<FrameworkElement>()
                    .Where(c => c is Grid g && g.Tag?.ToString() == "TAB_STRIP_CONTAINER" ||
                                c is StackPanel sp && sp.Height == 20)
                    .ToList();
                foreach (var oldStrip in existingStrips) dockPanel.Children.Remove(oldStrip);

                // 3. CHECK ENABLED
                bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";
                if (!tabsEnabled) return;

                // 4. COLOR ANALYSIS (For Arrow Visibility)
                string frameColorName = frame.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                string effectiveColor = !string.IsNullOrEmpty(frameColorName) ? frameColorName : SettingsManager.SelectedColor;

                System.Windows.Media.Color baseColor = System.Windows.Media.Colors.Gray;
                try
                {
                    var drawingColor = Utility.GetColorFromName(effectiveColor);
                    baseColor = System.Windows.Media.Color.FromArgb(255, drawingColor.R, drawingColor.G, drawingColor.B);
                }
                catch { }

                string c = effectiveColor?.ToLower() ?? "";
                bool isExplicitDark = c.Contains("blue") || c.Contains("teal") || c.Contains("black") ||
                                      c.Contains("red") || c.Contains("green") || c.Contains("purple") ||
                                      c.Contains("bismark") || c.Contains("fuchsia") || c.Contains("default");

                double brightness = Math.Sqrt(
                    (0.299 * baseColor.R * baseColor.R) +
                    (0.587 * baseColor.G * baseColor.G) +
                    (0.114 * baseColor.B * baseColor.B)
                );

                bool isDarkTheme = isExplicitDark || brightness < 160;
                SolidColorBrush arrowBrush = isDarkTheme ? System.Windows.Media.Brushes.White : new SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 50));

                // 5. CREATE NEW GRID STRUCTURE
                // Cols: [LeftArrow] [ScrollViewer*] [RightArrow] [PlusButton]
                Grid containerGrid = new Grid
                {
                    Tag = "TAB_STRIP_CONTAINER",
                    Height = 20,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 0, 0)),
                    Margin = new Thickness(0, 1, 0, 0),
                    VerticalAlignment = VerticalAlignment.Top
                };

                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // 0: Left Arrow
                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // 1: Tabs
                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // 2: Right Arrow
                containerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // 3: Plus Button

                // 6. CREATE SCROLLVIEWER
                ScrollViewer scrollViewer = new ScrollViewer
                {
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                    PanningMode = PanningMode.HorizontalOnly,
                    CanContentScroll = true
                };

                scrollViewer.PreviewMouseWheel += (s, e) =>
                {
                    if (e.Delta > 0) scrollViewer.LineLeft();
                    else scrollViewer.LineRight();
                    e.Handled = true;
                };

                StackPanel tabStack = new StackPanel { Orientation = Orientation.Horizontal };
                scrollViewer.Content = tabStack;

                // 7. CREATE GHOST ARROWS
                TextBlock leftArrow = new TextBlock
                {
                    Text = "‹", // Elegant chevron
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Foreground = arrowBrush,
                    Opacity = 0.6,
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, -2, 2, 0), // Slight nudge up
                    Visibility = Visibility.Collapsed // Hidden by default
                };
                leftArrow.MouseLeftButtonDown += (s, e) => { scrollViewer.LineLeft(); e.Handled = true; };

                TextBlock rightArrow = new TextBlock
                {
                    Text = "›",
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Foreground = arrowBrush,
                    Opacity = 0.6,
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(2, -2, 2, 0),
                    Visibility = Visibility.Collapsed
                };
                rightArrow.MouseLeftButtonDown += (s, e) => { scrollViewer.LineRight(); e.Handled = true; };

                // 8. WIRE UP SCROLL LOGIC
                // This updates visibility whenever the scroll position or size changes
                scrollViewer.ScrollChanged += (s, e) =>
                {
                    // Only show left if we have scrolled right
                    leftArrow.Visibility = scrollViewer.HorizontalOffset > 0
                        ? Visibility.Visible : Visibility.Collapsed;

                    // Only show right if there is scrollable content remaining
                    rightArrow.Visibility = scrollViewer.HorizontalOffset < scrollViewer.ScrollableWidth
                        ? Visibility.Visible : Visibility.Collapsed;
                };

                // Add Elements to Grid
                containerGrid.Children.Add(leftArrow); Grid.SetColumn(leftArrow, 0);
                containerGrid.Children.Add(scrollViewer); Grid.SetColumn(scrollViewer, 1);
                containerGrid.Children.Add(rightArrow); Grid.SetColumn(rightArrow, 2);

                // 9. POPULATE TABS
                var tabs = frame.Tabs as JArray ?? new JArray();
                int currentTab = Convert.ToInt32(frame.CurrentTab?.ToString() ?? "0");

                for (int i = 0; i < tabs.Count; i++)
                {
                    var tab = tabs[i] as JObject;
                    if (tab == null) continue;

                    string tabName = tab["TabName"]?.ToString() ?? $"Tab {i + 1}";
                    bool isActiveTab = (i == currentTab);
                    int capturedIndex = i;

                    Button tabButton = new Button
                    {
                        Content = tabName,
                        Tag = i,
                        Height = 18,
                        MinWidth = 50,
                        Margin = new Thickness(1, 0, 1, 0),
                        Padding = new Thickness(10, 2, 10, 2),
                        FontSize = 10,
                        FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                        Cursor = Cursors.Hand,
                        Focusable = false
                    };

                    ApplyTabStyle(tabButton, isActiveTab, frameColorName);

                    tabButton.PreviewMouseLeftButtonDown += (s, e) =>
                    {
                        frameWindow.Focus();
                        System.Windows.Input.Keyboard.ClearFocus();

                        if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                        {
                            e.Handled = true;
                            RenameTab(frame, capturedIndex, frameWindow);
                        }
                        else
                        {
                            SwitchTabByFrame(frame, capturedIndex, frameWindow);
                            e.Handled = true;
                        }
                    };

                    tabButton.PreviewMouseRightButtonDown += (s, e) => SwitchTabByFrame(frame, capturedIndex, frameWindow);

                    // Context Menu
                    ContextMenu tabContextMenu = new ContextMenu();

                    Style themedStyle = Framemanager.GetThemedContextMenuStyle(frame);
                    if (themedStyle != null) tabContextMenu.Style = themedStyle;

                    MenuItem miAddTab = new MenuItem { Header = Strings.TabAddNew };
                    MenuItem miRenameTab = new MenuItem { Header = Strings.TabRename };
                    MenuItem miDeleteTab = new MenuItem { Header = Strings.DeleteTabTitle };
                    MenuItem miMoveLeft = new MenuItem { Header = Strings.TabMoveLeft };
                    MenuItem miMoveRight = new MenuItem { Header = Strings.TabMoveRight };

                    tabContextMenu.Items.Add(miAddTab);
                    tabContextMenu.Items.Add(new Separator());
                    tabContextMenu.Items.Add(miRenameTab);
                    tabContextMenu.Items.Add(miDeleteTab);
                    tabContextMenu.Items.Add(new Separator());
                    tabContextMenu.Items.Add(miMoveLeft);
                    tabContextMenu.Items.Add(miMoveRight);

                    miAddTab.Click += (s, e) => AddNewTab(frame, frameWindow);
                    miRenameTab.Click += (s, e) => RenameTab(frame, capturedIndex, frameWindow);
                    miDeleteTab.Click += (s, e) => DeleteTab(frame, capturedIndex, frameWindow);
                    miMoveLeft.Click += (s, e) => MoveTab(frame, capturedIndex, -1, frameWindow);
                    miMoveRight.Click += (s, e) => MoveTab(frame, capturedIndex, 1, frameWindow);

                    tabContextMenu.Opened += (s, e) =>
                    {
                        // --- LIVE THEME FIX ---
                        try
                        {
                            Style liveStyle = Framemanager.GetThemedContextMenuStyle(frame);
                            if (liveStyle != null) tabContextMenu.Style = liveStyle;
                        }
                        catch { }

                        miMoveLeft.IsEnabled = capturedIndex > 0;
                        miMoveRight.IsEnabled = capturedIndex < tabs.Count - 1;
                    };

                    tabButton.ContextMenu = tabContextMenu;
                    tabStack.Children.Add(tabButton);
                }

                // 10. POPULATE PINNED [+] BUTTON
                Button addTabButton = new Button
                {
                    Content = "+",
                    Tag = "ADD_TAB",
                    Height = 18,
                    Width = 25,
                    Margin = new Thickness(3, 0, 1, 0),
                    FontSize = 12,
                    FontWeight = FontWeights.Bold,
                    FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                    BorderThickness = new Thickness(1),
                    Cursor = Cursors.Hand,
                    ToolTip = Strings.TabAddNewHint, // Updated Tooltip
                    Focusable = false
                };

                ApplyTabStyle(addTabButton, false, frameColorName, true);

                bool isAddingTab = false;
                addTabButton.PreviewMouseLeftButtonDown += async (s, e) =>
                {
                    frameWindow.Focus();
                    e.Handled = true;

                    // --- NEW IMPORT LOGIC ---
                    if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                    {
                        // Advanced: Import Tab
                        ImportTabManager.HandleImportRequest(frame, frameWindow);
                        return;
                    }
                    // ------------------------

                    if (isAddingTab) return;
                    isAddingTab = true;
                    try { AddNewTab(frame, frameWindow); }
                    finally { await System.Threading.Tasks.Task.Delay(500); isAddingTab = false; }
                };

                containerGrid.Children.Add(addTabButton);
                Grid.SetColumn(addTabButton, 3); // Col 3 is for the Button

                // 11. SURGICAL INSERTION
                DockPanel.SetDock(containerGrid, Dock.Top);

                int insertIndex = 0;
                bool titleFound = false;
                for (int i = 0; i < dockPanel.Children.Count; i++)
                {
                    if (dockPanel.Children[i] is Grid g)
                    {
                        if (g.Children.OfType<TextBlock>().Any(tb => tb.Name == "FrameLockIcon"))
                        {
                            insertIndex = i + 1;
                            titleFound = true;
                            break;
                        }
                    }
                }
                if (!titleFound) insertIndex = 0;

                if (insertIndex < dockPanel.Children.Count &&
                    dockPanel.Children[insertIndex] is Grid potentialFilter &&
                    potentialFilter.Children.OfType<ComboBox>().Any())
                {
                    insertIndex++;
                }

                if (insertIndex > dockPanel.Children.Count) insertIndex = dockPanel.Children.Count;
                dockPanel.Children.Insert(insertIndex, containerGrid);

                // 12. RESTORE SCROLL (Async to allow layout pass)
                if (previousScrollOffset > 0)
                {
                    frameWindow.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Loaded, new Action(() =>
                    {
                        scrollViewer.ScrollToHorizontalOffset(previousScrollOffset);
                    }));
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Refreshed Tab Strip with Ghost Arrows");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error refreshing tab strip UI: {ex.Message}");
            }
        }      // TABS FEATURE: Add new tab with random herb name
        public static void AddNewTab(dynamic frame, NonActivatingWindow frameWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"AddNewTab called for frame '{frame.Title}'");

                // Get fresh frame data
                string frameId = frame.Id?.ToString();
                if (string.IsNullOrEmpty(frameId))
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "Cannot add tab: frame ID missing");
                    return;
                }

                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Cannot add tab: frame with ID '{frameId}' not found");
                    return;
                }

                var tabs = currentFrame.Tabs as JArray ?? new JArray();

                // Generate random herb name with tab index
                string herbName = FrameUtilities.GenerateRandomHerbName();
                string newTabName = $"{tabs.Count}. {herbName}";

                // Create new tab object
                var newTab = new JObject();
                newTab["TabName"] = newTabName;
                newTab["Items"] = new JArray();

                // Add to tabs array
                tabs.Add(newTab);

                // Update frame data properly
                int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                if (frameIndex >= 0)
                {
                    IDictionary<string, object> frameDict = currentFrame is IDictionary<string, object> dict ?
                        dict : ((JObject)currentFrame).ToObject<IDictionary<string, object>>();

                    frameDict["Tabs"] = tabs; // Store JArray directly
                    frameDict["CurrentTab"] = tabs.Count - 1; // Switch to new tab

                    FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                    FrameDataManager.SaveFrameData();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Added new tab '{newTabName}' to frame '{currentFrame.Title}'");

                    // Get updated frame and refresh the display
                    var updatedFrame = FrameDataManager.FrameData[frameIndex];
                    int newTabIndex = tabs.Count - 1;

                    // Refresh content and styling
                    Framemanager.RefreshFrameContentSimple(frameWindow, updatedFrame, newTabIndex);
                    RefreshTabStyling(frameWindow, newTabIndex);

                    // Refresh the entire tab strip to show new tab
                    RefreshTabStripUI(frameWindow, updatedFrame);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "New tab added successfully");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error adding new tab: {ex.Message}");
            }
        }


        // TABS FEATURE: Rename tab with inline editing (in-button, focus-enabled)
        // v2.5.4.187: Fixed Visual Tree traversal to support ScrollViewer/Ghost Arrows structure
        public static void RenameTab(dynamic frame, int tabIndex, NonActivatingWindow frameWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"RenameTab called for frame '{frame.Title}', tab {tabIndex}");

                // 1. Find the main DockPanel
                var border = frameWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                // 2. FIX: Find the new Grid Container first (Tag: TAB_STRIP_CONTAINER)
                var containerGrid = dockPanel.Children.OfType<Grid>()
                    .FirstOrDefault(g => g.Tag?.ToString() == "TAB_STRIP_CONTAINER");

                if (containerGrid == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "RenameTab: Tab strip container not found.");
                    return;
                }

                // 3. FIX: Find the ScrollViewer inside the Grid
                var scrollViewer = containerGrid.Children.OfType<ScrollViewer>().FirstOrDefault();
                if (scrollViewer == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "RenameTab: ScrollViewer not found.");
                    return;
                }

                // 4. FIX: Get the StackPanel from the ScrollViewer content
                var tabStrip = scrollViewer.Content as StackPanel;
                if (tabStrip == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, "RenameTab: Tab StackPanel not found.");
                    return;
                }

                // 5. Find the button for this tab index
                Button targetButton = null;
                foreach (Button btn in tabStrip.Children.OfType<Button>())
                {
                    if (btn.Tag is int buttonTabIndex && buttonTabIndex == tabIndex)
                    {
                        targetButton = btn;
                        break;
                    }
                }

                if (targetButton == null) return;

                // Get current tab data
                string frameId = frame.Id?.ToString();
                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame == null) return;

                var tabs = currentFrame.Tabs as JArray ?? new JArray();
                if (tabIndex < 0 || tabIndex >= tabs.Count) return;

                var tab = tabs[tabIndex] as JObject;
                if (tab == null) return;

                string currentName = tab["TabName"]?.ToString() ?? $"Tab {tabIndex}";

                // Temporarily increase button height to accommodate TextBox properly
                double originalHeight = targetButton.Height;
                targetButton.Height = 22; // Slightly taller for TextBox

                // Create TextBox for inline editing
                TextBox editTextBox = new TextBox
                {
                    Text = currentName,
                    FontSize = targetButton.FontSize,
                    FontFamily = targetButton.FontFamily,
                    FontWeight = FontWeights.Normal,
                    Background = System.Windows.Media.Brushes.White,
                    Foreground = System.Windows.Media.Brushes.Black,
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 70, 130, 180)),
                    BorderThickness = new Thickness(1),
                    Padding = new Thickness(4, 2, 4, 2),
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Stretch,
                    VerticalContentAlignment = VerticalAlignment.Center
                };

                // Store original button properties
                object originalContent = targetButton.Content;
                var originalBackground = targetButton.Background;
                var originalForeground = targetButton.Foreground;
                var originalBorderBrush = targetButton.BorderBrush;

                targetButton.Content = editTextBox;
                targetButton.Background = System.Windows.Media.Brushes.White;
                targetButton.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(255, 70, 130, 180));

                // CRITICAL: Disable NonActivatingWindow focus prevention during editing and force focus
                frameWindow.BeginKeyboardInteractiveEdit(editTextBox);
                bool editingComplete = false;

                // Action to complete editing and restore NonActivatingWindow behavior
                Action<bool> completeEditing = (save) =>
                {
                    if (editingComplete) return;
                    editingComplete = true;

                    try
                    {
                        if (save && !string.IsNullOrWhiteSpace(editTextBox.Text))
                        {
                            string newName = editTextBox.Text.Trim();

                            // Validate name length
                            if (newName.Length > 30)
                            {
                                newName = newName.Substring(0, 30);
                            }

                            // Update tab name in data
                            tab["TabName"] = newName;

                            // Save to JSON
                            int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                            if (frameIndex >= 0)
                            {
                                IDictionary<string, object> frameDict = currentFrame is IDictionary<string, object> dict ?
                                    dict : ((JObject)currentFrame).ToObject<IDictionary<string, object>>();

                                frameDict["Tabs"] = tabs;
                                FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                                FrameDataManager.SaveFrameData();

                                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                                    $"Renamed tab from '{currentName}' to '{newName}'");

                                // Update button with new name
                                targetButton.Content = newName;
                            }
                        }
                        else
                        {
                            // Cancel - restore original content
                            targetButton.Content = originalContent;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Tab rename cancelled");
                        }

                        // Restore original button properties
                        targetButton.Background = originalBackground;
                        targetButton.Foreground = originalForeground;
                        targetButton.BorderBrush = originalBorderBrush;
                        targetButton.Height = originalHeight;

                        // CRITICAL: Re-enable NonActivatingWindow focus prevention
                        frameWindow.EndKeyboardInteractiveEdit();
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                            $"Error completing tab rename: {ex.Message}");

                        // Restore everything on error
                        targetButton.Content = originalContent;
                        targetButton.Background = originalBackground;
                        targetButton.Foreground = originalForeground;
                        targetButton.BorderBrush = originalBorderBrush;
                        targetButton.Height = originalHeight;
                        frameWindow.EndKeyboardInteractiveEdit();
                    }
                };

                // Handle Enter key (save) and Escape (cancel)
                editTextBox.KeyDown += (s, e) =>
                {
                    if (e.Key == Key.Enter)
                    {
                        completeEditing(true);
                        e.Handled = true;
                    }
                    else if (e.Key == Key.Escape)
                    {
                        completeEditing(false);
                        e.Handled = true;
                    }
                };

                // Handle focus loss (save)
                editTextBox.LostFocus += (s, e) =>
                {
                    System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(() => completeEditing(true))
                    );
                };

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Started inline editing for tab '{currentName}' with focus enabled");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error starting tab rename: {ex.Message}");

                // Ensure focus prevention is restored on any error
                frameWindow.EnableFocusPrevention(true);
            }
        }

        // TABS FEATURE: Delete tab with confirmation, Auto-Export, and Sync Fix
        public static void DeleteTab(dynamic frame, int tabIndex, NonActivatingWindow frameWindow)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"DeleteTab called for frame '{frame.Title}', tab {tabIndex}");

                // 1. Get Fresh Data
                string frameId = frame.Id?.ToString();
                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame == null) return;

                var tabs = currentFrame.Tabs as JArray ?? new JArray();

                // 2. Validation: Don't allow deleting the last tab
                if (tabs.Count <= 1)
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm(Strings.MsgCannotDeleteLastTab, Strings.DlgDeleteTab);
                    return;
                }

                if (tabIndex < 0 || tabIndex >= tabs.Count) return;

                var tab = tabs[tabIndex] as JObject;
                if (tab == null) return;

                string tabName = tab["TabName"]?.ToString() ?? $"Tab {tabIndex}";
                var items = tab["Items"] as JArray ?? new JArray();

                // 3. Confirmation
                if (items.Count > 0)
                {
                    bool result = MessageBoxesManager.ShowTabDeleteConfirmationForm(tabName, items.Count);
                    if (!result)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"User cancelled tab deletion for '{tabName}'");
                        return;
                    }
                }

                // --- Auto-Export Logic (Tabs follow Frame Settings) ---
                if (SettingsManager.ExportShortcutsOnFrameDeletion && items.Count > 0)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Auto-exporting {items.Count} items from tab '{tabName}' before deletion.");

                    int exportCount = 0;
                    foreach (var item in items)
                    {
                        try
                        {
                            CopyPasteManager.SendToDesktop(item);
                            exportCount++;
                        }
                        catch (Exception ex)
                        {
                            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Failed to export item from tab: {ex.Message}");
                        }
                    }
                }
                // -----------------------------------------------------------

                // 4. Remove Tab
                tabs.RemoveAt(tabIndex);

                // 5. Calculate New Active Tab
                int currentTab = Convert.ToInt32(currentFrame.CurrentTab?.ToString() ?? "0");
                int newCurrentTab = currentTab;

                // Shift logic:
                // If we deleted the active tab -> Go to previous (or 0)
                // If we deleted a tab BEFORE the active one -> Active tab shifts down by 1
                // If we deleted a tab AFTER the active one -> Active index stays same
                if (tabIndex <= currentTab)
                {
                    newCurrentTab = Math.Max(0, currentTab - 1);
                }

                // 6. Save Data
                int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                if (frameIndex >= 0)
                {
                    IDictionary<string, object> frameDict = currentFrame is IDictionary<string, object> dict ?
                        dict : ((JObject)currentFrame).ToObject<IDictionary<string, object>>();

                    frameDict["Tabs"] = tabs;
                    frameDict["CurrentTab"] = newCurrentTab;

                    FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                    FrameDataManager.SaveFrameData();

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Deleted tab '{tabName}' from frame '{currentFrame.Title}'. New Active Tab: {newCurrentTab}");

                    // --- BUG FIX: FORCE SYNC WITH MAIN ITEMS ---
                    // Now that the tabs have shifted, "Tab 0" might be different.
                    // We must force 'Main.Items' to mirror the NEW 'Tab 0' to prevent "Ghost Icons" 
                    // from appearing if the user later disables tabs.
                    if (tabs.Count > 0)
                    {
                        SynchronizeTab0Content(frameId, "tab0", "full");
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Forced Tab0 synchronization after deletion.");
                    }
                    // -------------------------------------------

                    // 7. REFRESH UI
                    var updatedFrame = FrameDataManager.FrameData[frameIndex];

                    // A. Refresh the Icons (Content)
                    Framemanager.RefreshFrameContentSimple(frameWindow, updatedFrame, newCurrentTab);

                    // B. Refresh the Tab Buttons (Strip)
                    RefreshTabStripUI(frameWindow, updatedFrame);

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Tab deletion UI refresh complete");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error deleting tab: {ex.Message}");
            }
        }


        // TABS FEATURE: Switch tab using frame ID
        // FIXED: Now calls RefreshTabStripUI instead of RefreshTabStyling to force a clean redraw.
        public static void SwitchTabByFrame(dynamic frame, int newTabIndex, NonActivatingWindow frameWindow)
        {
            try
            {
                // Get fresh frame data by ID to avoid stale references
                string frameId = frame.Id?.ToString();
                if (string.IsNullOrEmpty(frameId)) return;

                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame == null) return;

                // Validate tab index
                var tabs = currentFrame.Tabs as JArray ?? new JArray();
                if (newTabIndex < 0 || newTabIndex >= tabs.Count) return;

                // Check current tab
                int currentTabIndex = Convert.ToInt32(currentFrame.CurrentTab?.ToString() ?? "0");
                if (currentTabIndex == newTabIndex) return; // Already there

                // Update Data
                IDictionary<string, object> frameDict = currentFrame as IDictionary<string, object> ??
                    ((JObject)currentFrame).ToObject<IDictionary<string, object>>();
                frameDict["CurrentTab"] = newTabIndex;

                // Save
                int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                if (frameIndex >= 0)
                {
                    FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                    FrameDataManager.SaveFrameData();
                }

                var freshFrame = FrameDataManager.FrameData[frameIndex];

                // 1. Refresh Icons (Content)
                Framemanager.RefreshFrameContentSimple(frameWindow, freshFrame, newTabIndex);

                // 2. Refresh Tabs (UI Redraw) - THE FIX
                // We use RefreshTabStripUI instead of RefreshTabStyling.
                // This destroys the old buttons and creates new ones with the correct "Active" state baked in.
                RefreshTabStripUI(frameWindow, freshFrame);

                string tabName = tabs[newTabIndex]["TabName"]?.ToString() ?? $"Tab {newTabIndex}";
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Switched to tab '{tabName}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error switching tab: {ex.Message}");
            }
        }







        // TABS FEATURE: Move Tab (Left/Right)
        public static void MoveTab(dynamic frame, int fromIndex, int direction, NonActivatingWindow frameWindow)
        {
            try
            {
                // Direction: -1 for Left, +1 for Right
                int toIndex = fromIndex + direction;

                // 1. Get Fresh Data
                string frameId = frame.Id?.ToString();
                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame == null) return;

                var tabs = currentFrame.Tabs as JArray;
                if (tabs == null) return;

                // 2. Validate Bounds
                if (toIndex < 0 || toIndex >= tabs.Count) return;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Moving tab {fromIndex} to {toIndex} for frame '{currentFrame.Title}'");

                // 3. Swap Tabs
                var tempTab = tabs[fromIndex];
                tabs[fromIndex] = tabs[toIndex];
                tabs[toIndex] = tempTab;

                // 4. Update CurrentTab Index if necessary
                // If we moved the active tab, follow it. 
                // If we moved a tab into the active slot, update the index to stay on the same content.
                int currentTabIdx = Convert.ToInt32(currentFrame.CurrentTab?.ToString() ?? "0");

                if (currentTabIdx == fromIndex)
                {
                    currentFrame.CurrentTab = toIndex; // Follow the moved tab
                }
                else if (currentTabIdx == toIndex)
                {
                    currentFrame.CurrentTab = fromIndex; // The other tab swapped into our slot
                }

                // 5. Save & Refresh
                int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                if (frameIndex >= 0)
                {
                    IDictionary<string, object> frameDict = currentFrame is IDictionary<string, object> dict ?
                        dict : ((JObject)currentFrame).ToObject<IDictionary<string, object>>();

                    frameDict["Tabs"] = tabs;
                    frameDict["CurrentTab"] = currentFrame.CurrentTab; // Updated index

                    FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                    FrameDataManager.SaveFrameData();

                    // Refresh Strip (Buttons)
                    RefreshTabStripUI(frameWindow, FrameDataManager.FrameData[frameIndex]);

                    // Refresh Content (Icons) - just in case the active tab index changed logic
                    Framemanager.RefreshFrameContentSimple(frameWindow, FrameDataManager.FrameData[frameIndex], (int)frameDict["CurrentTab"]);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error moving tab: {ex.Message}");
            }
        }






        // TABS FEATURE: Tab0-frame Content Synchronization Manager
        private static bool _isSynchronizing = false; // Prevent circular sync operations

        /// <summary>
        /// Synchronizes content between Tab0 and main Items to ensure they remain identical
        /// Called whenever items are added/removed from either location
        /// </summary>
        /// <param name="frameId">The frame ID to synchronize</param>
        /// <param name="sourceLocation">Where the change originated: "tab0" or "main"</param>
        /// <param name="operationType">Type of operation: "add", "remove", "full"</param>
        public static void SynchronizeTab0Content(string frameId, string sourceLocation, string operationType = "full")
        {
            if (_isSynchronizing)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    "Sync already in progress, skipping to prevent circular operation");
                return;
            }

            try
            {
                _isSynchronizing = true;

                var frame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (frame == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                        $"Cannot sync: frame {frameId} not found");
                    return;
                }

                bool tabsEnabled = frame.TabsEnabled?.ToString().ToLower() == "true";
                var tabs = frame.Tabs as JArray ?? new JArray();
                var mainItems = frame.Items as JArray ?? new JArray();

                // Only sync if tabs are enabled and Tab0 exists
                if (!tabsEnabled || tabs.Count == 0) return;

                var tab0 = tabs[0] as JObject;
                if (tab0 == null) return;

                var tab0Items = tab0["Items"] as JArray ?? new JArray();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Synchronizing {operationType} from {sourceLocation} for frame '{frame.Title}' - Tab0: {tab0Items.Count} items, Main: {mainItems.Count} items");

                bool syncPerformed = false;

                // Determine sync direction and perform synchronization
                if (sourceLocation == "tab0")
                {
                    // Tab0 changed - sync to main Items
                    if (!AreItemArraysEqual(tab0Items, mainItems))
                    {
                        frame.Items = JArray.FromObject(tab0Items.ToArray());
                        syncPerformed = true;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Synced {tab0Items.Count} items from Tab0 to main Items for frame '{frame.Title}'");
                    }
                }
                else if (sourceLocation == "main")
                {
                    // Main Items changed - sync to Tab0
                    if (!AreItemArraysEqual(mainItems, tab0Items))
                    {
                        tab0["Items"] = JArray.FromObject(mainItems.ToArray());
                        syncPerformed = true;
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Synced {mainItems.Count} items from main Items to Tab0 for frame '{frame.Title}'");
                    }
                }

                // Save changes if synchronization was performed
                if (syncPerformed)
                {
                    int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);
                    if (frameIndex >= 0)
                    {
                        FrameDataManager.FrameData[frameIndex] = frame;
                        FrameDataManager.SaveFrameData();
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in Tab0 synchronization: {ex.Message}");
            }
            finally
            {
                _isSynchronizing = false;
            }
        }

        /// <summary>
        /// Helper method to compare two JArrays for equality
        /// </summary>
        private static bool AreItemArraysEqual(JArray array1, JArray array2)
        {
            if (array1.Count != array2.Count) return false;

            for (int i = 0; i < array1.Count; i++)
            {
                var item1 = array1[i] as JObject;
                var item2 = array2[i] as JObject;

                if (item1 == null || item2 == null) return false;

                // Compare essential properties (Filename is the key identifier)
                string filename1 = item1["Filename"]?.ToString();
                string filename2 = item2["Filename"]?.ToString();

                if (filename1 != filename2) return false;
            }

            return true;
        }

        // TABS FEATURE: Update tab button styling
        // Updated to support the new ScrollViewer structure
        public static void RefreshTabStyling(NonActivatingWindow frameWindow, int activeTabIndex)
        {
            try
            {
                var border = frameWindow.Content as Border;
                var dockPanel = border?.Child as DockPanel;
                if (dockPanel == null) return;

                // 1. Find the Container (Grid)
                var containerGrid = dockPanel.Children.OfType<Grid>()
                    .FirstOrDefault(g => g.Tag?.ToString() == "TAB_STRIP_CONTAINER");

                if (containerGrid == null) return;

                // 2. Find the Tab StackPanel (Inside ScrollViewer)
                var scrollViewer = containerGrid.Children.OfType<ScrollViewer>().FirstOrDefault();
                var tabStack = scrollViewer?.Content as StackPanel;

                // 3. Find the [+] Button (Direct child of Grid)
                var plusButton = containerGrid.Children.OfType<Button>()
                    .FirstOrDefault(b => b.Tag?.ToString() == "ADD_TAB");

                string frameId = frameWindow.Tag?.ToString();
                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame == null) return;

                string frameColorName = currentFrame.CustomColor?.ToString();
                if (string.IsNullOrEmpty(frameColorName)) frameColorName = SettingsManager.SelectedColor;

                // 4. Update Tab Buttons
                if (tabStack != null)
                {
                    foreach (var child in tabStack.Children)
                    {
                        if (child is Button tabButton && tabButton.Tag is int idx)
                        {
                            bool isActive = (idx == activeTabIndex);
                            ApplyTabStyle(tabButton, isActive, frameColorName);
                        }
                    }
                }

                // 5. Update [+] Button
                if (plusButton != null)
                {
                    ApplyTabStyle(plusButton, false, frameColorName, true);
                }
            }
            catch { }
        }

        public static void RefreshTabColors(NonActivatingWindow frameWindow, string newColorName)
        {
            try
            {
                var border = frameWindow.Content as Border;
                if (border == null) return;
                var dockPanel = border.Child as DockPanel;
                if (dockPanel == null) return;
                var tabStrip = dockPanel.Children.OfType<StackPanel>().FirstOrDefault(sp => sp.Orientation == Orientation.Horizontal && sp.Height == 20);
                if (tabStrip == null) return;

                string frameId = frameWindow.Tag?.ToString();
                var currentFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == frameId);
                if (currentFrame == null) return;

                int currentTab = Convert.ToInt32(currentFrame.CurrentTab?.ToString() ?? "0");

                foreach (var child in tabStrip.Children)
                {
                    if (child is Button tabButton)
                    {
                        if (tabButton.Tag is int idx)
                        {
                            bool isActive = (idx == currentTab);
                            ApplyTabStyle(tabButton, isActive, newColorName);
                        }
                        else if (tabButton.Tag?.ToString() == "ADD_TAB")
                        {
                            ApplyTabStyle(tabButton, false, newColorName, true);
                        }
                    }
                }
            }
            catch { }
        }

        // --- PROFESSIONAL UX FINAL v13: The "Sorcery" Fix ---
        // 1. Reverts Padding to 'TemplateBinding' so the [+] button isn't crushed.
        // 2. Ensures the [+] button (Width 25) gets its 0 padding, while Text tabs get 10.
        private static void ApplyTabStyle(Button btn, bool isActive, string colorName, bool isPlusButton = false)
        {
            try
            {
                // 1. RESET
                btn.Style = null;
                btn.FocusVisualStyle = null;
                btn.Focusable = false;

                btn.ClearValue(Button.BackgroundProperty);
                btn.ClearValue(Button.ForegroundProperty);
                btn.ClearValue(Button.BorderBrushProperty);
                btn.ClearValue(Button.FontWeightProperty);
                btn.ClearValue(Button.PaddingProperty); // Clear local padding to be safe

                btn.MouseEnter -= Tab_MouseEnter_Lambda;
                btn.MouseLeave -= Tab_MouseLeave_Lambda;

                // 2. COLOR CALCULATION
                string effectiveColor = !string.IsNullOrEmpty(colorName) ? colorName : SettingsManager.SelectedColor;
                System.Windows.Media.Color baseColor = System.Windows.Media.Colors.Gray;
                try
                {
                    var drawingColor = Utility.GetColorFromName(effectiveColor);
                    baseColor = System.Windows.Media.Color.FromArgb(255, drawingColor.R, drawingColor.G, drawingColor.B);
                }
                catch { }

                string c = effectiveColor?.ToLower() ?? "";
                bool isExplicitDark = c.Contains("blue") || c.Contains("teal") || c.Contains("black") ||
                                      c.Contains("red") || c.Contains("green") || c.Contains("purple") ||
                                      c.Contains("bismark") || c.Contains("fuchsia") || c.Contains("default");

                double brightness = Math.Sqrt(
                    (0.299 * baseColor.R * baseColor.R) +
                    (0.587 * baseColor.G * baseColor.G) +
                    (0.114 * baseColor.B * baseColor.B)
                );

                bool isDarkTheme = isExplicitDark || brightness < 160;

                // 3. PALETTE DEFINITION
                SolidColorBrush bgActive, bgInactive, bgHover;
                SolidColorBrush textActive, textInactive;
                SolidColorBrush borderActive, borderInactive;

                if (isDarkTheme)
                {
                    bgActive = new SolidColorBrush(baseColor);
                    bgInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 255, 255, 255));
                    bgHover = new SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 255, 255, 255));

                    textActive = System.Windows.Media.Brushes.White;
                    textInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(180, 255, 255, 255));

                    borderActive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(120, 255, 255, 255));
                    borderInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 255, 255, 255));
                }
                else
                {
                    bgActive = System.Windows.Media.Brushes.White;
                    bgInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(20, 0, 0, 0));
                    bgHover = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 0, 0));

                    textActive = new SolidColorBrush(System.Windows.Media.Color.FromRgb(20, 20, 20));
                    textInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(160, 0, 0, 0));

                    borderActive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(80, 0, 0, 0));
                    borderInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(40, 0, 0, 0));
                }

                // 4. BUTTON CONFIGURATION
                if (isPlusButton)
                {
                    // Special Case: [+] Button needs 0 padding to center the text in 25px width
                    btn.Padding = new Thickness(0);

                    if (!isActive)
                    {
                        bgInactive = new SolidColorBrush(System.Windows.Media.Color.FromArgb(
                            (byte)(bgInactive.Color.A / 2),
                            bgInactive.Color.R,
                            bgInactive.Color.G,
                            bgInactive.Color.B));
                    }
                }
                else
                {
                    // Standard Tab: Needs padding for breathing room
                    btn.Padding = new Thickness(10, 2, 10, 2);
                }

                // 5. TEMPLATE GENERATION
                ControlTemplate template = new ControlTemplate(typeof(Button));
                FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
                border.Name = "Border";
                border.SetValue(Border.CornerRadiusProperty, new CornerRadius(4, 4, 0, 0));

                // CRITICAL FIX: Bind Padding to the Button's property instead of hardcoding it.
                // This allows the [+] button to have 0 padding and Text tabs to have 10 padding.
                border.SetValue(Border.PaddingProperty, new TemplateBindingExtension(Button.PaddingProperty));

                FrameworkElementFactory content = new FrameworkElementFactory(typeof(ContentPresenter));
                content.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                content.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

                border.AppendChild(content);
                template.VisualTree = border;

                // 6. APPLY STYLE
                Style style = new Style(typeof(Button));
                style.Setters.Add(new Setter(Button.TemplateProperty, template));

                if (isActive)
                {
                    // ACTIVE
                    style.Setters.Add(new Setter(Button.BackgroundProperty, bgActive));
                    style.Setters.Add(new Setter(Button.ForegroundProperty, textActive));
                    style.Setters.Add(new Setter(Button.BorderBrushProperty, borderActive));
                    style.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1, 1, 1, 0)));
                    style.Setters.Add(new Setter(Button.FontWeightProperty, FontWeights.Bold));
                    style.Setters.Add(new Setter(Button.OpacityProperty, 1.0));

                    Trigger staticTrigger = new Trigger { Property = Button.IsEnabledProperty, Value = true };
                    staticTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgActive, "Border"));
                    staticTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, borderActive, "Border"));
                    staticTrigger.Setters.Add(new Setter(Border.BorderThicknessProperty, new Thickness(1, 1, 1, 0), "Border"));
                    template.Triggers.Add(staticTrigger);

                    if (isDarkTheme)
                        style.Setters.Add(new Setter(Button.EffectProperty, new DropShadowEffect { BlurRadius = 4, ShadowDepth = 1, Direction = 270, Color = System.Windows.Media.Colors.Black, Opacity = 0.5 }));
                }
                else
                {
                    // INACTIVE
                    style.Setters.Add(new Setter(Button.BackgroundProperty, bgInactive));
                    style.Setters.Add(new Setter(Button.ForegroundProperty, textInactive));
                    style.Setters.Add(new Setter(Button.BorderBrushProperty, borderInactive));
                    style.Setters.Add(new Setter(Button.BorderThicknessProperty, new Thickness(1, 1, 1, 1)));
                    style.Setters.Add(new Setter(Button.FontWeightProperty, FontWeights.Normal));

                    Trigger baseTrigger = new Trigger { Property = Button.IsEnabledProperty, Value = true };
                    baseTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgInactive, "Border"));
                    baseTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, borderInactive, "Border"));
                    baseTrigger.Setters.Add(new Setter(Border.BorderThicknessProperty, new Thickness(1, 1, 1, 1), "Border"));
                    template.Triggers.Add(baseTrigger);

                    Trigger hoverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
                    hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, bgHover, "Border"));
                    hoverTrigger.Setters.Add(new Setter(Border.BorderBrushProperty, borderActive, "Border"));
                    hoverTrigger.Setters.Add(new Setter(Button.ForegroundProperty, textActive));
                    hoverTrigger.Setters.Add(new Setter(Button.CursorProperty, Cursors.Hand));
                    template.Triggers.Add(hoverTrigger);
                }

                btn.Style = style;
            }
            catch { }
        }


        // Dummy handlers to allow -= syntax (prevents compiler errors if we were using named methods)
        // Since we use lambdas above, we don't strictly need these, but good for safety if refactoring.
        private static void Tab_MouseEnter_Lambda(object sender, System.Windows.Input.MouseEventArgs e) { }
        private static void Tab_MouseLeave_Lambda(object sender, System.Windows.Input.MouseEventArgs e) { }
    }
}

using Desktop_Frames.Localization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using Newtonsoft.Json.Linq;

namespace Desktop_Frames
{
    /// <summary>
    /// Handles importing tabs from other frames.
    /// Rebuilt to match the visual style and UX of ItemMoveDialog.
    /// </summary>
    public static class ImportTabManager
    {
        #region Public Entry Point

        public static void HandleImportRequest(dynamic targetFrame, NonActivatingWindow targetWindow)
        {
            try
            {
                // 1. Window Setup (Matches ItemMoveDialog)
                var importWindow = new Window
                {
                    Title = Strings.ImportTabTitle,
                    Width = 480,
                    Height = 600,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize,
                    WindowStyle = WindowStyle.None,
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250)),
                    AllowsTransparency = true
                };

                // Main Container Card
                Border mainBorder = new Border
                {
                    Background = System.Windows.Media.Brushes.White,
                    CornerRadius = new CornerRadius(12),
                    Margin = new Thickness(8),
                    Effect = new DropShadowEffect
                    {
                        Color = System.Windows.Media.Colors.Black,
                        Direction = 315,
                        ShadowDepth = 2,
                        BlurRadius = 8,
                        Opacity = 0.2
                    }
                };

                Grid rootGrid = new Grid();
                rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
                rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
                rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Footer

                // --- HEADER ---
                Border headerBorder = new Border
                {
                    Background = GetAccentColorBrush(),
                    CornerRadius = new CornerRadius(0),
                    Padding = new Thickness(20, 15, 15, 15)
                };

                Grid headerGrid = new Grid();
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                StackPanel titlePanel = new StackPanel();
                TextBlock mainTitle = new TextBlock
                {
                    Text = Strings.ImportTabTitle,
                    FontSize = 16,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = System.Windows.Media.Brushes.White,
                    Margin = new Thickness(0, 0, 0, 2)
                };
                TextBlock subTitle = new TextBlock
                {
                    Text = $"Target: {targetFrame.Title}",
                    FontSize = 12,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromArgb(200, 255, 255, 255)),
                    FontStyle = FontStyles.Italic
                };
                titlePanel.Children.Add(mainTitle);
                titlePanel.Children.Add(subTitle);

                Button closeButton = new Button
                {
                    Content = "✕",
                    Width = 30,
                    Height = 30,
                    FontSize = 16,
                    FontWeight = FontWeights.Normal,
                    Foreground = System.Windows.Media.Brushes.White,
                    Background = System.Windows.Media.Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Top
                };
                closeButton.Click += (s, e) => importWindow.Close();

                headerGrid.Children.Add(titlePanel);
                headerGrid.Children.Add(closeButton); Grid.SetColumn(closeButton, 1);
                headerBorder.Child = headerGrid;

                // Allow Drag
                headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) importWindow.DragMove(); };

                Grid.SetRow(headerBorder, 0);
                rootGrid.Children.Add(headerBorder);

                // --- CONTENT ---
                Border contentBorder = new Border { Background = System.Windows.Media.Brushes.White };

                StackPanel contentStack = new StackPanel();
                contentStack.Children.Add(new TextBlock
                {
                    Text = Strings.ImportTabPrompt,
                    FontSize = 13,
                    FontWeight = FontWeights.Medium,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(95, 99, 104)),
                    Margin = new Thickness(20, 15, 20, 10)
                });

                ScrollViewer scrollViewer = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    MaxHeight = 400,
                    Margin = new Thickness(20, 0, 20, 0),
                    CanContentScroll = true
                };

                StackPanel targetsPanel = new StackPanel();

                // Build the list using the same style as ItemMoveDialog
                BuildSourceList(targetsPanel, targetFrame, targetWindow, importWindow);

                scrollViewer.Content = targetsPanel;
                contentStack.Children.Add(scrollViewer);

                contentBorder.Child = contentStack;
                Grid.SetRow(contentBorder, 1);
                rootGrid.Children.Add(contentBorder);

                // --- FOOTER ---
                Border footerBorder = new Border
                {
                    Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250)),
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 220, 224)),
                    BorderThickness = new Thickness(0, 1, 0, 0),
                    Padding = new Thickness(20, 15, 20, 15),
                    CornerRadius = new CornerRadius(0, 0, 12, 12)
                };

                Button cancelButton = new Button
                {
                    Content = "Cancel",
                    Width = 100,
                    Height = 36,
                    FontSize = 13,
                    FontWeight = FontWeights.Medium,
                    Background = System.Windows.Media.Brushes.White,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(95, 99, 104)),
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 220, 224)),
                    BorderThickness = new Thickness(1),
                    Cursor = Cursors.Hand,
                    HorizontalAlignment = HorizontalAlignment.Right
                };
                cancelButton.Click += (s, e) => importWindow.Close();

                footerBorder.Child = cancelButton;
                Grid.SetRow(footerBorder, 2);
                rootGrid.Children.Add(footerBorder);

                mainBorder.Child = rootGrid;
                importWindow.Content = mainBorder;
                importWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Import UI Error: {ex.Message}");
            }
        }

        #endregion

        #region UI Helpers & Builders

        private static void BuildSourceList(StackPanel targetsPanel, dynamic targetFrame, NonActivatingWindow targetWindow, Window importWindow)
        {
            string targetId = targetFrame.Id?.ToString();
            var FrameData = Framemanager.GetFrameData();
            bool foundSources = false;

            foreach (var frame in FrameData)
            {
                // Logic: 
                // 1. Skip Self
                // 2. Skip non-Data frames (Portal/Notes can't be imported as tabs)
                if (frame.Id?.ToString() == targetId) continue;
                if (frame.ItemsType?.ToString() != "Data") continue;

                foundSources = true;
                string frameTitle = frame.Title?.ToString() ?? "Unnamed frame";
                bool hasTabs = frame.TabsEnabled?.ToString().ToLower() == "true";

                if (!hasTabs)
                {
                    // Case A: Import entire non-tabbed frame as a new tab
                    CreateImportButton(targetsPanel, frameTitle, "🏠", frame, null, targetFrame, targetWindow, importWindow, false);
                }
                else
                {
                    // Case B: Tabbed frame - Show Header + Individual Tabs

                    // Create Group Header (Non-clickable visual only)
                    Border groupHeader = new Border
                    {
                        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250)),
                        BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 220, 224)),
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(6),
                        Margin = new Thickness(0, 5, 0, 2),
                        Padding = new Thickness(12, 8, 12, 8)
                    };
                    StackPanel headerStack = new StackPanel { Orientation = Orientation.Horizontal };
                    headerStack.Children.Add(new TextBlock { Text = "📂", FontSize = 16, Margin = new Thickness(0, 0, 10, 0) });
                    headerStack.Children.Add(new TextBlock
                    {
                        Text = frameTitle,
                        FontSize = 13,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 33, 36)),
                        VerticalAlignment = VerticalAlignment.Center
                    });

                    groupHeader.Child = headerStack;
                    targetsPanel.Children.Add(groupHeader);

                    // Add individual tabs as sources
                    var tabs = frame.Tabs as JArray ?? new JArray();
                    for (int i = 0; i < tabs.Count; i++)
                    {
                        var tab = tabs[i] as JObject;
                        if (tab != null)
                        {
                            string tabName = tab["TabName"]?.ToString() ?? $"Tab {i + 1}";
                            CreateImportButton(targetsPanel, $"📄 {tabName}", "↳", frame, i, targetFrame, targetWindow, importWindow, true);
                        }
                    }

                    // Separator
                    targetsPanel.Children.Add(new Border
                    {
                        Height = 1,
                        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(230, 230, 230)),
                        Margin = new Thickness(0, 8, 0, 8)
                    });
                }
            }

            if (!foundSources)
            {
                targetsPanel.Children.Add(new TextBlock
                {
                    Text = Strings.ImportTabNoFrames,
                    Foreground = Brushes.Gray,
                    FontStyle = FontStyles.Italic,
                    Margin = new Thickness(5)
                });
            }
        }

        private static void CreateImportButton(StackPanel parent, string text, string prefix,
            dynamic sourceFrame, int? sourceTabIndex, dynamic targetFrame, NonActivatingWindow targetWindow, Window importWindow, bool isIndented)
        {
            Button btn = new Button
            {
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Margin = new Thickness(isIndented ? 20 : 0, 2, 0, 2),
                Padding = new Thickness(12, 10, 12, 10),
                Cursor = Cursors.Hand,
                MinHeight = 40
            };

            StackPanel btnContent = new StackPanel { Orientation = Orientation.Horizontal };
            if (isIndented)
            {
                btnContent.Children.Add(new TextBlock
                {
                    Text = prefix,
                    FontSize = 12,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 120, 120)),
                    Margin = new Thickness(0, 0, 8, 0),
                    VerticalAlignment = VerticalAlignment.Center
                });
            }
            btnContent.Children.Add(new TextBlock
            {
                Text = text,
                FontSize = 13,
                FontWeight = isIndented ? FontWeights.Normal : FontWeights.Medium,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 33, 36)),
                VerticalAlignment = VerticalAlignment.Center
            });

            btn.Content = btnContent;

            // Hover Styling
            btn.MouseEnter += (s, e) => { btn.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(243, 244, 246)); };
            btn.MouseLeave += (s, e) => { btn.Background = System.Windows.Media.Brushes.White; };

            // Logic
            btn.Click += (s, e) =>
            {
                // Confirmation Dialog inside the click
                if (MessageBoxesManager.ShowCustomYesNoMessageBox(
                    $"Import contents of '{text}' into a new tab?",
                    "Confirm Import"))
                {
                    PerformImport(targetFrame, targetWindow, sourceFrame, sourceTabIndex);
                    importWindow.Close();
                }
            };

            parent.Children.Add(btn);
        }

        private static SolidColorBrush GetAccentColorBrush()
        {
            try
            {
                var mediaColor = Utility.GetColorFromName(SettingsManager.SelectedColor);
                return new SolidColorBrush(mediaColor);
            }
            catch { return new SolidColorBrush(System.Windows.Media.Color.FromRgb(66, 133, 244)); }
        }

        #endregion

        #region Import Logic

        private static void PerformImport(dynamic targetFrame, NonActivatingWindow targetWindow, dynamic sourceFrame, int? sourceTabIndex)
        {
            try
            {
                JArray itemsToCopy = new JArray();
                string newTabName = "Imported";

                // 1. Extract Data
                if (sourceTabIndex.HasValue)
                {
                    // Source is a specific tab
                    var tabs = sourceFrame.Tabs as JArray;
                    if (tabs != null && sourceTabIndex.Value < tabs.Count)
                    {
                        var tabData = tabs[sourceTabIndex.Value] as JObject;
                        newTabName = tabData?["TabName"]?.ToString() ?? "Imported Tab";
                        itemsToCopy = tabData?["Items"] as JArray ?? new JArray();
                    }
                }
                else
                {
                    // Source is an entire non-tabbed frame
                    newTabName = sourceFrame.Title.ToString();
                    itemsToCopy = sourceFrame.Items as JArray ?? new JArray();
                }

                // 2. Prepare Target
                var targetTabs = targetFrame.Tabs as JArray ?? new JArray();

                // Deep Clone Items (Critical for safety)
                JArray clonedItems = JArray.FromObject(itemsToCopy);

                // Create New Tab
                JObject newTab = new JObject();
                newTab["TabName"] = newTabName + " (Copy)"; // Append Copy to distinguish
                newTab["Items"] = clonedItems;

                targetTabs.Add(newTab);

                // 3. Save
                string frameId = targetFrame.Id?.ToString();
                int frameIndex = FrameDataManager.FrameData.FindIndex(f => f.Id?.ToString() == frameId);

                if (frameIndex >= 0)
                {
                    IDictionary<string, object> frameDict = targetFrame is IDictionary<string, object> dict ?
                        dict : ((JObject)targetFrame).ToObject<IDictionary<string, object>>();

                    frameDict["Tabs"] = targetTabs;
                    frameDict["CurrentTab"] = targetTabs.Count - 1; // Switch to new tab

                    FrameDataManager.FrameData[frameIndex] = JObject.FromObject(frameDict);
                    FrameDataManager.SaveFrameData();

                    // 4. Refresh UI
                    var updatedFrame = FrameDataManager.FrameData[frameIndex];
                    Framemanager.RefreshFrameContentSimple(targetWindow, updatedFrame, targetTabs.Count - 1);
                    TabManager.RefreshTabStyling(targetWindow, targetTabs.Count - 1);
                    TabManager.RefreshTabStripUI(targetWindow, updatedFrame);

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.ImportExport, $"Imported tab '{newTabName}' into frame '{targetFrame.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.ImportExport, $"Import Failed: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Import failed: {ex.Message}", "Error");
            }
        }

        #endregion
    }
}
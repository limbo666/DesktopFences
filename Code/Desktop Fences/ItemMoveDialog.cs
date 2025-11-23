using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using IWshRuntimeLibrary;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{
    /// <summary>
    /// Handles all item move operations with modern, tab-aware dialog interface.
    /// Separated from FenceManager for better code organization and maintainability.
    /// </summary>
    public static class ItemMoveDialog
    {
        #region Public Methods

        /// <summary>
        /// Shows the modern tab-aware Move dialog and handles the complete move operation
        /// </summary>
        /// <param name="item">The item to move</param>
        /// <param name="sourceFence">The source fence containing the item</param>
        /// <param name="dispatcher">Dispatcher for UI operations</param>



        private static SolidColorBrush GetAccentColorBrush()
        {
            try
            {
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                return new SolidColorBrush(mediaColor);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error getting accent color: {ex.Message}");
                // Fallback to blue
                return new SolidColorBrush(System.Windows.Media.Color.FromRgb(66, 133, 244));
            }
        }


        public static void ShowMoveDialog(dynamic item, dynamic sourceFence, Dispatcher dispatcher)
        {
            // Create modern hierarchical Move dialog
            var moveWindow = new Window
            {
                Title = "Move Item To...",
                Width = 480,
                Height = 600,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.None,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250)),
                AllowsTransparency = true
            };

            // Get item details for display
            IDictionary<string, object> itemDict = item is IDictionary<string, object> dict ?
                dict : ((JObject)item).ToObject<IDictionary<string, object>>();
            string itemName = itemDict.ContainsKey("DisplayName") ?
                itemDict["DisplayName"].ToString() :
                System.IO.Path.GetFileNameWithoutExtension(itemDict.ContainsKey("Filename") ? itemDict["Filename"].ToString() : "Unknown");

            // Main container with modern styling
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

            // HEADER: Modern title bar with close button
       

            Border headerBorder = new Border
            {
                Background = GetAccentColorBrush(), // Use user's selected theme color
                CornerRadius = new CornerRadius(0, 0, 0, 0), // Match the window's corner radius
                Padding = new Thickness(20, 15, 15, 15)
            };

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // Title with item name
            StackPanel titlePanel = new StackPanel();

            TextBlock mainTitle = new TextBlock
            {
                Text = "Move Item To...",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Foreground = System.Windows.Media.Brushes.White,
                Margin = new Thickness(0, 0, 0, 2)
            };

            TextBlock subTitle = new TextBlock
            {
                Text = $"Moving: {itemName}",
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
            closeButton.Click += (s, e) => moveWindow.Close();

            headerGrid.Children.Add(titlePanel);
            headerGrid.Children.Add(closeButton);
            Grid.SetColumn(closeButton, 1);
            // Add drag functionality to header area only (like EditShortcutWindow)
            headerBorder.MouseLeftButtonDown += (sender, e) =>
            {
                if (e.ButtonState == MouseButtonState.Pressed)
                {
                    moveWindow.DragMove();
                }
            };

            headerBorder.Child = headerGrid;
            Grid.SetRow(headerBorder, 0);
            rootGrid.Children.Add(headerBorder);

            // CONTENT: Hierarchical fence and tab list
            Border contentBorder = new Border
            {
                Padding = new Thickness(0),
                Background = System.Windows.Media.Brushes.White
            };

            // Instructions
            TextBlock instructionsText = new TextBlock
            {
                Text = "Select a destination fence or tab:",
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(95, 99, 104)),
                Margin = new Thickness(20, 15, 20, 10)
            };

            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                MaxHeight = 400, // Explicit height limit to ensure scrolling activates
                Margin = new Thickness(20, 0, 20, 0),
                    CanContentScroll = true // Improve scrolling performance
            };

            StackPanel targetsPanel = new StackPanel();

            // Build hierarchical target list
            BuildTargetsList(targetsPanel, sourceFence, item, moveWindow);

            scrollViewer.Content = targetsPanel;

            StackPanel contentStack = new StackPanel();
            contentStack.Children.Add(instructionsText);
            contentStack.Children.Add(scrollViewer);
            contentBorder.Child = contentStack;
            Grid.SetRow(contentBorder, 1);
            rootGrid.Children.Add(contentBorder);

            // FOOTER: Cancel button
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
               // CornerRadius = new CornerRadius(6),
                Cursor = Cursors.Hand,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            cancelButton.Click += (s, e) => moveWindow.Close();

            footerBorder.Child = cancelButton;
            Grid.SetRow(footerBorder, 2);
            rootGrid.Children.Add(footerBorder);

            mainBorder.Child = rootGrid;
            // Add drag functionality like EditShortcutWindow
            moveWindow.MouseLeftButtonDown += (sender, e) =>
            {
                if (e.ButtonState == MouseButtonState.Pressed)
                {
                    moveWindow.DragMove();
                }
            };

            moveWindow.Content = mainBorder;
            moveWindow.ShowDialog();
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// Builds the hierarchical list of move targets (fences and tabs)
        /// </summary>
        private static void BuildTargetsList(StackPanel targetsPanel, dynamic sourceFence, dynamic item, Window moveWindow)
        {
            string sourceFenceId = sourceFence.Id?.ToString();
            var fenceData = FenceManager.GetFenceData();

            foreach (var fence in fenceData)
            {
                // Skip source fence and Portal fences
                if (fence.Id?.ToString() == sourceFenceId || fence.ItemsType?.ToString() == "Portal" || fence.ItemsType?.ToString() == "Note")
                    continue;

                bool fenceHasTabs = fence.TabsEnabled?.ToString().ToLower() == "true";
                string fenceTitle = fence.Title?.ToString() ?? "Unnamed Fence";

                if (!fenceHasTabs)
                {
                    // Regular fence - single target
                    CreateMoveTargetButton(targetsPanel, fenceTitle, "🏠", fence, null, item, sourceFence, moveWindow, false);
                }
                else
                {
                    // Tabbed fence - show parent fence + tabs

                    // Parent fence header (non-clickable, just shows fence name)
                    Border fenceHeaderBorder = new Border
                    {
                        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250)),
                        BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 220, 224)),
                        BorderThickness = new Thickness(1),
                        CornerRadius = new CornerRadius(6),
                        Margin = new Thickness(0, 5, 0, 2),
                        Padding = new Thickness(12, 8, 12, 8)
                    };

                    StackPanel fenceHeaderPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal
                    };

                    TextBlock fenceIcon = new TextBlock
                    {
                        Text = "📂",
                        FontSize = 16,
                        Margin = new Thickness(0, 0, 10, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    TextBlock fenceNameText = new TextBlock
                    {
                        Text = fenceTitle,
                        FontSize = 13,
                        FontWeight = FontWeights.SemiBold,
                        Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 33, 36)),
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    TextBlock tabCountText = new TextBlock
                    {
                        Text = $"({GetTabCount(fence)} tabs)",
                        FontSize = 11,
                        Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(95, 99, 104)),
                        Margin = new Thickness(8, 0, 0, 0),
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    fenceHeaderPanel.Children.Add(fenceIcon);
                    fenceHeaderPanel.Children.Add(fenceNameText);
                    fenceHeaderPanel.Children.Add(tabCountText);
                    fenceHeaderBorder.Child = fenceHeaderPanel;
                    targetsPanel.Children.Add(fenceHeaderBorder);

                    // Option to move to fence main area
                    CreateMoveTargetButton(targetsPanel, "📋 Main Items", "↳", fence, null, item, sourceFence, moveWindow, true);

                    // Individual tabs
                    var tabs = fence.Tabs as JArray ?? new JArray();
                    for (int i = 0; i < tabs.Count; i++)
                    {
                        var tab = tabs[i] as JObject;
                        if (tab != null)
                        {
                            string tabName = tab["TabName"]?.ToString() ?? $"Tab {i}";
                            CreateMoveTargetButton(targetsPanel, $"📄 {tabName}", "↳", fence, i, item, sourceFence, moveWindow, true);
                        }
                    }

                    // Add separator after tabbed fence group
                    Border separator = new Border
                    {
                        Height = 1,
                        Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(230, 230, 230)),
                        Margin = new Thickness(0, 8, 0, 8)
                    };
                    targetsPanel.Children.Add(separator);
                }
            }
        }

        /// <summary>
        /// Gets the number of tabs for a fence (for display purposes)
        /// </summary>
        private static int GetTabCount(dynamic fence)
        {
            try
            {
                var tabs = fence.Tabs as JArray;
                return tabs?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Creates a clickable move target button for fence or tab
        /// </summary>
        private static void CreateMoveTargetButton(StackPanel parent, string displayText, string prefix,
            dynamic targetFence, int? targetTabIndex, dynamic item, dynamic sourceFence, Window moveWindow, bool isIndented)
        {
            Button targetButton = new Button
            {
                HorizontalContentAlignment = HorizontalAlignment.Left,
                Background = System.Windows.Media.Brushes.White,
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
           //     CornerRadius = new CornerRadius(6),
                Margin = new Thickness(isIndented ? 20 : 0, 2, 0, 2),
                Padding = new Thickness(12, 10, 12, 10),
                Cursor = Cursors.Hand,
                MinHeight = 40
            };

            StackPanel buttonContent = new StackPanel
            {
                Orientation = Orientation.Horizontal
            };

            if (isIndented)
            {
                TextBlock indentPrefix = new TextBlock
                {
                    Text = prefix,
                    FontSize = 12,
                    Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(120, 120, 120)),
                    Margin = new Thickness(0, 0, 8, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };
                buttonContent.Children.Add(indentPrefix);
            }

            TextBlock targetText = new TextBlock
            {
                Text = displayText,
                FontSize = 13,
                FontWeight = isIndented ? FontWeights.Normal : FontWeights.Medium,
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 33, 36)),
                VerticalAlignment = VerticalAlignment.Center
            };

            buttonContent.Children.Add(targetText);
            targetButton.Content = buttonContent;

            // Add hover effects
            targetButton.MouseEnter += (s, e) =>
            {
                targetButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(243, 244, 246));
                targetButton.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(66, 133, 244));
            };

            targetButton.MouseLeave += (s, e) =>
            {
                targetButton.Background = System.Windows.Media.Brushes.White;
                targetButton.BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(218, 220, 224));
            };

            // Handle click - perform the actual move
            targetButton.Click += (s, e) => HandleMoveToTarget(item, sourceFence, targetFence, targetTabIndex, moveWindow);

            parent.Children.Add(targetButton);
        }

        /// <summary>
        /// Handles the actual move operation to fence or tab
        /// </summary>
        private static void HandleMoveToTarget(dynamic item, dynamic sourceFence, dynamic targetFence,
            int? targetTabIndex, Window moveWindow)
        {
            try
            {
                IDictionary<string, object> itemDict = item is IDictionary<string, object> dict ?
                    dict : ((JObject)item).ToObject<IDictionary<string, object>>();
                string filename = itemDict.ContainsKey("Filename") ? itemDict["Filename"].ToString() : "Unknown";

                // Determine source location (main Items or tab)
                JArray sourceItems = null;
                bool sourceIsTabbed = sourceFence.TabsEnabled?.ToString().ToLower() == "true";

                if (sourceIsTabbed)
                {
                    var sourceTabs = sourceFence.Tabs as JArray ?? new JArray();
                    int sourceCurrentTab = Convert.ToInt32(sourceFence.CurrentTab?.ToString() ?? "0");
                    if (sourceCurrentTab >= 0 && sourceCurrentTab < sourceTabs.Count)
                    {
                        var sourceActiveTab = sourceTabs[sourceCurrentTab] as JObject;
                        sourceItems = sourceActiveTab?["Items"] as JArray ?? new JArray();
                    }
                }
                else
                {
                    sourceItems = sourceFence.Items as JArray ?? new JArray();
                }

                // Find item in source
                var itemToMove = sourceItems?.FirstOrDefault(i => i["Filename"]?.ToString() == filename);
                if (itemToMove == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                        $"Item '{filename}' not found in source location");
                    return;
                }

                // Determine destination location
                JArray destItems = null;
                string destinationDescription = "";

                if (targetTabIndex.HasValue)
                {
                    // Moving to specific tab
                    var targetTabs = targetFence.Tabs as JArray ?? new JArray();
                    if (targetTabIndex.Value >= 0 && targetTabIndex.Value < targetTabs.Count)
                    {
                        var targetTab = targetTabs[targetTabIndex.Value] as JObject;
                        destItems = targetTab?["Items"] as JArray ?? new JArray();
                        string tabName = targetTab?["TabName"]?.ToString() ?? $"Tab {targetTabIndex.Value}";
                        destinationDescription = $"tab '{tabName}' in fence '{targetFence.Title}'";
                    }
                }
                else
                {
                    // Moving to fence main Items
                    destItems = targetFence.Items as JArray ?? new JArray();
                    destinationDescription = $"main area of fence '{targetFence.Title}'";
                }

                if (destItems == null)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                        "Could not determine destination location");
                    return;
                }

                // Perform the move
                sourceItems.Remove(itemToMove);
                destItems.Add(itemToMove);

                // Save changes
                FenceDataManager.SaveFenceData();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.IconHandling,
                    $"Successfully moved item '{filename}' from '{sourceFence.Title}' to {destinationDescription}");

                moveWindow.Close();

                // Show success feedback and refresh UI
                // Show success feedback and refresh UI
                ShowMoveSuccessAndRefresh(filename, destinationDescription, sourceFence, targetFence);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.IconHandling,
                    $"Error during move operation: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Move failed: {ex.Message}", "Error");
            }
        }

        /// <summary>
        /// Shows success message and refreshes UI after move
        /// </summary>
        /// <summary>
        /// Shows success message and refreshes UI after move - only affects source and target fences
        /// </summary>
        private static void ShowMoveSuccessAndRefresh(string itemName, string destination, dynamic sourceFence, dynamic targetFence)
        {
            //// Show brief success notification
            //MessageBoxesManager.ShowOKOnlyMessageBoxForm(
            //    $"'{itemName}' moved successfully to {destination}.",
            //    "Move Complete");

            // Reload only the source and target fences to prevent duplicates
            Application.Current.Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    FenceManager.ReloadSpecificFences(sourceFence, targetFence);
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                        $"Error refreshing specific fences after move: {ex.Message}");
                }
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        #endregion
    }
}
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace Desktop_Fences
{
    /// <summary>
    /// Manages Note Fence specific functionality including content management, context menus, and formatting
    /// Dedicated class for Note Fence operations to maintain separation from Data/Portal fence logic
    /// Used by: FenceManager.CreateFence() for Note type fences
    /// Category: Note Fence Management
    /// </summary>
    public static class NoteFenceManager
    {
        #region Note Content Creation - Used by: FenceManager.CreateFence()

        /// <summary>
        /// Creates the note content area (TextBox) for Note fences
        /// Used by: FenceManager.CreateFence() when ItemsType == "Note"
        /// Category: UI Creation
        /// </summary>
        /// <param name="fence">The fence data object</param>
        /// <param name="dockPanel">The parent DockPanel to add content to</param>
        /// <returns>The created TextBox for note content</returns>
        public static TextBox CreateNoteContent(dynamic fence, DockPanel dockPanel)
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                    $"Creating note content for fence '{fence.Title}'");

                // Safely get note content with null checks
                string noteContent = "";
                try
                {
                    noteContent = fence.NoteContent?.ToString() ?? "";
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // NoteContent property doesn't exist yet - use empty string
                    noteContent = "";
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        "NoteContent property not found - using empty content for new Note fence");
                }

                // Safely get other Note properties with fallbacks
                string wordWrap = GetSafeNoteProperty(fence, "WordWrap", "true");
                string spellCheck = GetSafeNoteProperty(fence, "SpellCheck", "true");
                string noteFontSize = GetSafeNoteProperty(fence, "NoteFontSize", "Medium");
                string noteFontFamily = GetSafeNoteProperty(fence, "NoteFontFamily", "Segoe UI");

                // Debug logging to see what values we're getting
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Loading Note fence settings - FontSize: '{noteFontSize}', FontFamily: '{noteFontFamily}'");
               

                                // Create the main TextBox for note content
                                TextBox noteTextBox = new TextBox
                {
                    // Content and behavior
                    Text = noteContent,
                    AcceptsReturn = true,
                    AcceptsTab = true,
                    TextWrapping = GetTextWrapping(wordWrap),
                    SpellCheck = { IsEnabled = GetSpellCheck(spellCheck) },

                                    // Appearance - Start transparent, will be updated by ApplyNoteColorScheme
                                    //  Background = Brushes.Transparent,
                                    Background = GetTextBackgroundBrush(fence),
                                    Foreground = GetNoteForeground(fence),
                                 
                                    FontSize = GetNoteFontSize(noteFontSize),
                    FontFamily = GetNoteFontFamily(noteFontFamily),
                    BorderThickness = GetTextBoxBorder().thickness,
                    BorderBrush = GetTextBoxBorder().brush,

                                    //// Layout
                                    //HorizontalAlignment = HorizontalAlignment.Stretch,
                                    //VerticalAlignment = VerticalAlignment.Stretch,
                                    //Margin = new Thickness(18),
                                    //Padding = new Thickness(4, 6, 8, 6),
                                    // Layout - TWEAKABLE VALUES
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Stretch,
                    Margin = GetTextBoxMargin(), // Custom spacing from fence edges
                    Padding = GetTextBoxPadding(), // Text spacing inside TextBox

                                    // Scrolling
                    VerticalScrollBarVisibility = SettingsManager.DisableFenceScrollbars ?
                        ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto,
                    HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden,

                    // Cursor
                    Cursor = Cursors.IBeam,

//                    // Add drop shadow effect for better text visibility
//Effect = new DropShadowEffect
//{
//    Color = Colors.Black,
//    Direction = 315,
//    ShadowDepth = 2,
//    BlurRadius = 3,
//    Opacity = 0.8
//}

                                };

                // Apply exact fence background color
                ApplyNoteColorScheme(noteTextBox, fence);

                // Add visual state management and auto-save functionality
                SetupNoteEditingBehavior(noteTextBox, fence);

                // Add to DockPanel (fills remaining space after title)
                dockPanel.Children.Add(noteTextBox);

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.FenceCreation,
                    $"Successfully created note content for fence '{fence.Title}' with {noteContent.Length} characters");

                return noteTextBox;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error creating note content: {ex.Message}");

                // Return a basic TextBox as fallback - but make it functional
                var fallbackTextBox = new TextBox
                {
                    Text = "",  // Empty instead of error message
                    AcceptsReturn = true,
                    AcceptsTab = true,
                    Background = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Stretch,
                    VerticalAlignment = VerticalAlignment.Stretch,
                    Margin = new Thickness(2),
                    Padding = new Thickness(8, 6, 8, 6)
                };

                // Add to DockPanel
                dockPanel.Children.Add(fallbackTextBox);

                // Even for fallback, add basic functionality
                SetupNoteEditingBehavior(fallbackTextBox, fence);

                return fallbackTextBox;
            }
        }

        /// <summary>
        /// Safely gets Note fence properties with fallback values
        /// </summary>
        /// <summary>
        /// Safely gets Note fence properties with fallback values
        /// </summary>
        public static string GetSafeNoteProperty(dynamic fence, string propertyName, string fallbackValue)
        {
            try
            {
                // Method 1: Direct property access
                try
                {
                    var value = fence.GetType().GetProperty(propertyName)?.GetValue(fence);
                    if (value != null) return value.ToString();
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }

                // Method 2: Dictionary access  
                try
                {
                    var fenceDict = fence as IDictionary<string, object>;
                    if (fenceDict != null && fenceDict.ContainsKey(propertyName))
                    {
                        var value = fenceDict[propertyName]?.ToString();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                            $"GetSafeNoteProperty: Found {propertyName} = '{value}' in dictionary");
                        return value ?? fallbackValue;
                    }
                }
                catch { }

                // Method 3: JObject access (for JSON loaded fences)
                try
                {
                    var jObject = fence as Newtonsoft.Json.Linq.JObject;
                    if (jObject != null && jObject.ContainsKey(propertyName))
                    {
                        var value = jObject[propertyName]?.ToString();
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                            $"GetSafeNoteProperty: Found {propertyName} = '{value}' in JObject");
                        return value ?? fallbackValue;
                    }
                }
                catch { }

                return fallbackValue;
            }
            catch
            {
                return fallbackValue;
            }
        }
        #endregion

        #region Note Editing Behavior - Used by: CreateNoteContent()

    

        /// <summary>
        /// Sets up complete editing behavior with visual feedback and auto-save
        /// Used by: CreateNoteContent() during TextBox setup
        /// Category: Content Management
        /// </summary>
        private static void SetupNoteEditingBehavior(TextBox noteTextBox, dynamic fence)
        {
            try
            {
                // Capture the ID immediately so we can always look up the fresh object later
                string fenceId = fence.Id?.ToString();

                // Auto-save timer to prevent excessive saves during typing (can be disabled via settings)
                System.Windows.Threading.DispatcherTimer autoSaveTimer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(2) // Save 2 seconds after last change
                };

                autoSaveTimer.Tick += (s, e) =>
                {
                    autoSaveTimer.Stop();
                    if (!SettingsManager.DisableNoteAutoSave)
                    {
                        // FIX: Fetch fresh data for saving to ensure we don't overwrite other properties with stale data
                        var freshFence = FenceManager.GetFenceData().FirstOrDefault(f => f.Id?.ToString() == fenceId) ?? fence;
                        SaveNoteContent(freshFence, noteTextBox.Text);
                    }
                };

                // CRITICAL: Store original layout properties to maintain anchoring during editing
                Thickness originalMargin = noteTextBox.Margin;
                HorizontalAlignment originalHAlign = noteTextBox.HorizontalAlignment;
                VerticalAlignment originalVAlign = noteTextBox.VerticalAlignment;
                Brush originalTextColor = noteTextBox.Foreground; // Store purely for fallback

                // Mouse enter - very subtle indication it's clickable
                noteTextBox.MouseEnter += (s, e) =>
                {
                    if (!noteTextBox.IsFocused)
                    {
                        noteTextBox.BorderBrush = new SolidColorBrush(Color.FromArgb(80, 70, 130, 180));
                        noteTextBox.Cursor = Cursors.IBeam;
                    }
                };

                // Mouse leave - remove hover effects if not focused
                noteTextBox.MouseLeave += (s, e) =>
                {
                    if (!noteTextBox.IsFocused)
                    {
                        noteTextBox.BorderThickness = new Thickness(0);
                        noteTextBox.BorderBrush = null;
                    }
                };

                Button doneButton = null;

                // Handle focus - show editing state and optional Done button
                noteTextBox.GotFocus += (s, e) =>
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        "TextBox got focus - entering edit mode");

                    // --- FIX START: Get Fresh Data ---
                    // Retrieve the latest fence data from the global list to ensure we use the NEW color
                    var freshFence = FenceManager.GetFenceData().FirstOrDefault(f => f.Id?.ToString() == fenceId) ?? fence;
                    // --- FIX END ---

                    // CRITICAL: Maintain anchoring properties during editing
                    noteTextBox.Margin = originalMargin;
                    noteTextBox.HorizontalAlignment = originalHAlign;
                    noteTextBox.VerticalAlignment = originalVAlign;

                    // Use FRESH fence color settings
                    string fenceColor = freshFence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                    Color baseColor = Utility.GetColorFromName(fenceColor ?? "Gray");

                    // Apply same blending to ALL fences
                    Color highlightedColor = Color.FromRgb(
                        (byte)(baseColor.R * 0.75 + 255 * 0.25),
                        (byte)(baseColor.G * 0.75 + 255 * 0.25),
                        (byte)(baseColor.B * 0.75 + 255 * 0.25));

                    noteTextBox.Background = new SolidColorBrush(highlightedColor) { Opacity = 0.9 };

                    // During edit mode, stick to a standard high-contrast color (Dark Blue) 
                    // or calculate a high-contrast color against the highlight
                    noteTextBox.Foreground = Brushes.DarkBlue;

                    noteTextBox.BorderBrush = new SolidColorBrush(Color.FromRgb(70, 130, 180)); // Steel blue
                    noteTextBox.BorderThickness = new Thickness(1);

                    // Create "Done" button if needed
                    if (doneButton == null)
                    {
                        var buttonLayout = GetDoneButtonLayout();
                        doneButton = new Button
                        {
                            Content = "✓",
                            Width = buttonLayout.width,
                            Height = buttonLayout.height,
                            HorizontalAlignment = HorizontalAlignment.Right,
                            VerticalAlignment = VerticalAlignment.Bottom,
                            Margin = buttonLayout.margin,
                            Background = new SolidColorBrush(Color.FromArgb(200, 64, 169, 64)),
                            Foreground = Brushes.White,
                            BorderThickness = new Thickness(0),
                            FontSize = 14,
                            FontWeight = FontWeights.Bold,
                            Cursor = Cursors.Hand,
                            ToolTip = "Click to finish editing",
                            Visibility = Visibility.Collapsed,
                        };

                        // Add click handler for Done button
                        doneButton.Click += (ds, de) => {
                            // Use fresh fence for saving here too
                            var currentFence = FenceManager.GetFenceData().FirstOrDefault(f => f.Id?.ToString() == fenceId) ?? fence;
                            SaveNoteContent(currentFence, noteTextBox.Text);

                            var parentWindow = FindParentWindow(noteTextBox);
                            if (parentWindow != null) parentWindow.Focus();

                            doneButton.Visibility = Visibility.Collapsed;
                        };

                        // Add button logic (same as before)
                        var buttonParentWindow = FindParentWindow(noteTextBox);
                        if (buttonParentWindow != null)
                        {
                            var border = buttonParentWindow.Content as Border;
                            if (border != null)
                            {
                                Grid overlayGrid = new Grid();
                                var dockPanel = border.Child as DockPanel;
                                if (dockPanel != null)
                                {
                                    border.Child = overlayGrid;
                                    overlayGrid.Children.Add(dockPanel);
                                }
                                overlayGrid.Children.Add(doneButton);
                                Canvas.SetZIndex(doneButton, 1000);
                            }
                        }
                    }

                    doneButton.Visibility = Visibility.Visible;

                    var pWin = FindParentWindow(noteTextBox);
                    if (pWin is NonActivatingWindow naw)
                    {
                        naw.EnableFocusPrevention(false);
                    }
                };

                // Handle focus loss - restore fence background
                noteTextBox.LostFocus += (s, e) =>
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                        "TextBox focus lost - restoring background");

                    // --- FIX START: Get Fresh Data ---
                    var freshFence = FenceManager.GetFenceData().FirstOrDefault(f => f.Id?.ToString() == fenceId) ?? fence;
                    // --- FIX END ---

               
                    // Use FRESH fence color for the Highlight calculation
                    string fenceColor = freshFence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;
                    Color baseColor = Utility.GetColorFromName(fenceColor ?? "Gray");

                    // Calculate Highlight based on the NEW color
                    Color highlightedColor = Color.FromRgb(
                        (byte)(baseColor.R * 0.75 + 255 * 0.25),
                        (byte)(baseColor.G * 0.75 + 255 * 0.25),
                        (byte)(baseColor.B * 0.75 + 255 * 0.25));

                    noteTextBox.Background = new SolidColorBrush(highlightedColor) { Opacity = 0.9 };


                    autoSaveTimer.Stop();
                    if (!SettingsManager.DisableNoteAutoSave)
                    {
                        SaveNoteContent(freshFence, noteTextBox.Text);
                    }

                    var pWin = FindParentWindow(noteTextBox);
                    if (pWin is NonActivatingWindow naw)
                    {
                        naw.EnableFocusPrevention(true);
                    }

                    System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(
                        System.Windows.Threading.DispatcherPriority.Background,
                        new Action(() => {
                            try
                            {
                                noteTextBox.Margin = originalMargin;
                                noteTextBox.HorizontalAlignment = originalHAlign;
                                noteTextBox.VerticalAlignment = originalVAlign;

                                noteTextBox.BorderThickness = new Thickness(0);
                                noteTextBox.BorderBrush = null;

                                // Apply color scheme using the FRESH fence object
                                ApplyNoteColorScheme(noteTextBox, freshFence);
                            }
                            catch (Exception ex)
                            {
                                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                                    $"Error restoring fence background: {ex.Message}");
                            }
                        })
                    );
                };

                noteTextBox.TextChanged += (s, e) =>
                {
                    if (noteTextBox.IsFocused && !SettingsManager.DisableNoteAutoSave)
                    {
                        autoSaveTimer.Stop();
                        autoSaveTimer.Start();
                    }
                };

                noteTextBox.ToolTip = "Click to edit note content";
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error setting up note editing behavior: {ex.Message}");
            }
        }










        /// <summary>
        /// Helper method to find parent NonActivatingWindow
        /// </summary>
        private static Window FindParentWindow(DependencyObject child)
        {
            DependencyObject parentObject = VisualTreeHelper.GetParent(child);

            if (parentObject == null)
                return null;

            if (parentObject is Window window)
                return window;

            return FindParentWindow(parentObject);
        }

        /// <summary>
        /// Saves note content to JSON data
        /// Used by: Auto-save timer and focus lost events
        /// Category: Data Persistence
        /// </summary>
        private static void SaveNoteContent(dynamic fence, string content)
        {
            try
            {
                string fenceId = fence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId)) return;

                // Find fence in data and update content
                int fenceIndex = FenceDataManager.FenceData.FindIndex(f => f.Id?.ToString() == fenceId);
                if (fenceIndex >= 0)
                {
                    dynamic actualFence = FenceDataManager.FenceData[fenceIndex];
                    IDictionary<string, object> fenceDict = actualFence as IDictionary<string, object> ??
                        ((JObject)actualFence).ToObject<IDictionary<string, object>>();

                    fenceDict["NoteContent"] = content;
                    FenceDataManager.FenceData[fenceIndex] = JObject.FromObject(fenceDict);
                    FenceDataManager.SaveFenceData();

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceUpdate,
                        $"Auto-saved note content for fence '{fence.Title}' ({content.Length} characters)");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceUpdate,
                    $"Error saving note content: {ex.Message}");
            }
        }
        #endregion

        #region Context Menu Creation - Used by: FenceManager context menu logic

        /// <summary>
        /// Creates Note fence specific context menu items
        /// Used by: FenceManager context menu creation for Note fences
        /// Category: UI Context Menu
        /// </summary>
        public static void AddNoteContextMenuItems(ContextMenu menu, dynamic fence, TextBox noteTextBox)
        {
            try
            {
                
            

                //        // Text formatting submenu
                //    //    MenuItem formatMenu = new MenuItem { Header = "Text Format" };

                //        // Font Size submenu
                //        MenuItem fontSizeMenu = new MenuItem { Header = "Font Size" };
                //        foreach (string size in new[] { "Small", "Medium", "Large", "Extra Large" })
                //        {
                //            MenuItem sizeItem = new MenuItem { Header = size };
                //            sizeItem.Click += (s, e) => ChangeNoteFontSize(fence, noteTextBox, size);
                //            fontSizeMenu.Items.Add(sizeItem);
                //        }
                // //       formatMenu.Items.Add(fontSizeMenu);

                //        // Font Family submenu
                //        MenuItem fontFamilyMenu = new MenuItem { Header = "Font Family" };
                //        foreach (string family in new[] { "Segoe UI", "Consolas", "Times New Roman", "Arial", "Courier New" })
                //        {
                //            MenuItem familyItem = new MenuItem { Header = family };
                //            familyItem.Click += (s, e) => ChangeNoteFontFamily(fence, noteTextBox, family);
                //            fontFamilyMenu.Items.Add(familyItem);
                //        }
                // //       formatMenu.Items.Add(fontFamilyMenu);

                //        // Text Color submenu (uses same colors as fence icons)
                //        MenuItem textColorMenu = new MenuItem { Header = "Text Color" };
                //        foreach (string color in new[] { "Default", "Red", "Green", "Blue", "White", "Black", "Gray" })
                //        {
                //            MenuItem colorItem = new MenuItem { Header = color };
                //            colorItem.Click += (s, e) => ChangeNoteTextColor(fence, noteTextBox, color == "Default" ? null : color);
                //            textColorMenu.Items.Add(colorItem);
                //        }
                ////        formatMenu.Items.Add(textColorMenu);

                // //       menu.Items.Add(formatMenu);

                // Text operations
                //   MenuItem textOpsMenu = new MenuItem { Header = "Text Operations" };


                // Text Format form (new unified approach)
                MenuItem textFormatFormItem = new MenuItem { Header = "Text Format..." };
                textFormatFormItem.Click += (s, e) =>
                {
                    try
                    {
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                            $"Opening Text Format form for fence '{fence.Title}'");
                        TextFormatFormManager.ShowTextFormatForm(fence, noteTextBox);
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                            $"Error opening Text Format form: {ex.Message}");
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error opening Text Format form: {ex.Message}", "Form Error");
                    }
                };
                menu.Items.Add(textFormatFormItem);

                // A seperator to commnds to note-specific commands
                menu.Items.Add(new Separator());

                MenuItem copyAllItem = new MenuItem { Header = "Copy All Text" };
                copyAllItem.Click += (s, e) => CopyAllNoteText(noteTextBox);
                menu.Items.Add(copyAllItem);

                MenuItem clearAllItem = new MenuItem { Header = "Clear All Text" };
                clearAllItem.Click += (s, e) => ClearAllNoteText(fence, noteTextBox);
                menu.Items.Add(clearAllItem);

           //     textOpsMenu.Items.Add(new Separator());

                MenuItem wordWrapItem = new MenuItem
                {
                    Header = GetWordWrapMenuText(fence.WordWrap?.ToString())
                };
                wordWrapItem.Click += (s, e) => ToggleWordWrap(fence, noteTextBox);
         //       textOpsMenu.Items.Add(wordWrapItem);

                MenuItem spellCheckItem = new MenuItem
                {
                    Header = GetSpellCheckMenuText(fence.SpellCheck?.ToString())
                };
                spellCheckItem.Click += (s, e) => ToggleSpellCheck(fence, noteTextBox);
             //   textOpsMenu.Items.Add(spellCheckItem);

             //   menu.Items.Add(textOpsMenu);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    "Added Note fence context menu items");

                // Keep existing individual submenu for backwards compatibility



            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error creating note context menu: {ex.Message}");
            }
        }
        #endregion

        /// <summary>
        /// Creates a subtle background specifically for text area to improve visibility
        /// </summary>
        private static Brush GetTextBackgroundBrush(dynamic fence)
        {
            try
            {
                // Get the fence color name safely
                string fenceColorName = null;
                try { fenceColorName = fence.CustomColor?.ToString(); } catch { }

                if (string.IsNullOrEmpty(fenceColorName))
                    fenceColorName = SettingsManager.SelectedColor ?? "Gray";

                // Use the shared helper to determine the actual visual background color
                Color actualBg = GetActualFenceBackgroundColor(fenceColorName);
                double luminance = GetRelativeLuminance(actualBg);

                if (luminance > 0.5) // Light background
                {
                    // Dark semi-transparent background for light fences
                    return new SolidColorBrush(Color.FromArgb(30, 0, 0, 0));
                }
                else // Dark background
                {
                    // Light semi-transparent background for dark fences
                    return new SolidColorBrush(Color.FromArgb(30, 255, 255, 255));
                }
            }
            catch
            {
                return Brushes.Transparent;
            }
        }





        #region Text Formatting Methods - Used by: Context menu actions

        /// <summary>
        /// Changes the font size of the note content
        /// </summary>
        private static void ChangeNoteFontSize(dynamic fence, TextBox noteTextBox, string size)
        {
            try
            {
                double fontSize = GetNoteFontSizeValue(size);
                noteTextBox.FontSize = fontSize;

                // Save to fence data
                string fenceId = fence.Id?.ToString();
                if (!string.IsNullOrEmpty(fenceId))
                {
                    FenceManager.UpdateFenceProperty(fence, "NoteFontSize", size,
                        $"Changed font size to {size}");
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Changed note font size to {size} for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error changing font size: {ex.Message}");
            }
        }

        /// <summary>
        /// Changes the font family of the note content
        /// </summary>
        private static void ChangeNoteFontFamily(dynamic fence, TextBox noteTextBox, string family)
        {
            try
            {
                noteTextBox.FontFamily = new FontFamily(family);

                // Save to fence data
                string fenceId = fence.Id?.ToString();
                if (!string.IsNullOrEmpty(fenceId))
                {
                    FenceManager.UpdateFenceProperty(fence, "NoteFontFamily", family,
                        $"Changed font family to {family}");
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Changed note font family to {family} for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error changing font family: {ex.Message}");
            }
        }

        /// <summary>
        /// Changes the text color of the note content (uses fence TextColor property)
        /// </summary>
        private static void ChangeNoteTextColor(dynamic fence, TextBox noteTextBox, string colorName)
        {
            try
            {
                // Update the fence TextColor property (same as icons use)
                string fenceId = fence.Id?.ToString();
                if (!string.IsNullOrEmpty(fenceId))
                {
                    FenceManager.UpdateFenceProperty(fence, "TextColor", colorName,
                        $"Changed text color to {colorName}");
                }

                // Apply the color to the TextBox immediately
                if (!string.IsNullOrEmpty(colorName) && colorName != "Default")
                {
                    var textColor = Utility.GetColorFromName(colorName);
                    noteTextBox.Foreground = new SolidColorBrush(textColor);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        $"Applied text color {colorName} to TextBox immediately");
                }
                else
                {
                    noteTextBox.Foreground = Brushes.White; // Default
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        "Applied default white text color to TextBox");
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Changed note text color to {colorName} for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error changing text color: {ex.Message}");
            }
        }

        /// <summary>
        /// Copies all text from the note to clipboard
        /// </summary>
        private static void CopyAllNoteText(TextBox noteTextBox)
        {
            try
            {
                if (!string.IsNullOrEmpty(noteTextBox.Text))
                {
                    Clipboard.SetText(noteTextBox.Text);
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        "Copied all note text to clipboard");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error copying note text: {ex.Message}");
            }
        }

        /// <summary>
        /// Clears all text from the note
        /// </summary>
        private static void ClearAllNoteText(dynamic fence, TextBox noteTextBox)
        {
            try
            {
                if (MessageBox.Show("Are you sure you want to clear all text from this note?",
                    "Clear Note", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    noteTextBox.Text = "";
                    SaveNoteContent(fence, "");

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                        $"Cleared all text from note fence '{fence.Title}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error clearing note text: {ex.Message}");
            }
        }

        /// <summary>
        /// Toggles word wrap setting for the note
        /// </summary>
        private static void ToggleWordWrap(dynamic fence, TextBox noteTextBox)
        {
            try
            {
                bool currentWrap = noteTextBox.TextWrapping == TextWrapping.Wrap;
                bool newWrap = !currentWrap;

                noteTextBox.TextWrapping = newWrap ? TextWrapping.Wrap : TextWrapping.NoWrap;

                // Save to fence data
                string fenceId = fence.Id?.ToString();
                if (!string.IsNullOrEmpty(fenceId))
                {
                    FenceManager.UpdateFenceProperty(fence, "WordWrap", newWrap.ToString().ToLower(),
                        $"Toggled word wrap to {newWrap}");
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Toggled word wrap to {newWrap} for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error toggling word wrap: {ex.Message}");
            }
        }

        /// <summary>
        /// Toggles spell check setting for the note
        /// </summary>
        private static void ToggleSpellCheck(dynamic fence, TextBox noteTextBox)
        {
            try
            {
                bool currentSpellCheck = noteTextBox.SpellCheck.IsEnabled;
                bool newSpellCheck = !currentSpellCheck;

                noteTextBox.SpellCheck.IsEnabled = newSpellCheck;

                // Save to fence data
                string fenceId = fence.Id?.ToString();
                if (!string.IsNullOrEmpty(fenceId))
                {
                    FenceManager.UpdateFenceProperty(fence, "SpellCheck", newSpellCheck.ToString().ToLower(),
                        $"Toggled spell check to {newSpellCheck}");
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Toggled spell check to {newSpellCheck} for fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error toggling spell check: {ex.Message}");
            }
        }
        #endregion

        #region Helper Methods - Internal formatting and utility functions

        /// <summary>
        /// Gets TextWrapping enum from string value
        /// </summary>
        private static TextWrapping GetTextWrapping(string wordWrap)
        {
            return wordWrap?.ToLower() == "false" ? TextWrapping.NoWrap : TextWrapping.Wrap;
        }

        /// <summary>
        /// Gets spell check boolean from string value
        /// </summary>
        private static bool GetSpellCheck(string spellCheck)
        {
            return spellCheck?.ToLower() != "false"; // Default to true
        }

        


        public static Brush GetNoteForeground(dynamic fence)
        {
            try
            {
                // Try multiple ways to get the TextColor property
                string textColorName = null;
                // Method 1: Direct property access
                try
                {
                    textColorName = fence.TextColor?.ToString();
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // TextColor property doesn't exist - will use default
                }
                // Method 2: Dictionary access if fence is a dictionary
                if (string.IsNullOrEmpty(textColorName))
                {
                    try
                    {
                        var fenceDict = fence as IDictionary<string, object>;
                        if (fenceDict != null && fenceDict.ContainsKey("TextColor"))
                        {
                            textColorName = fenceDict["TextColor"]?.ToString();
                        }
                    }
                    catch { }
                }

                // Get fence background color for contrast checking
                string fenceColorName = null;
                try
                {
                    fenceColorName = fence.CustomColor?.ToString();
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // CustomColor property doesn't exist
                }
                // Try dictionary access for fence color too
                if (string.IsNullOrEmpty(fenceColorName))
                {
                    try
                    {
                        var fenceDict = fence as IDictionary<string, object>;
                        if (fenceDict != null && fenceDict.ContainsKey("CustomColor"))
                        {
                            fenceColorName = fenceDict["CustomColor"]?.ToString();
                        }
                    }
                    catch { }
                }
                // Fallback to SettingsManager if no fence-specific color
                if (string.IsNullOrEmpty(fenceColorName))
                {
                    fenceColorName = SettingsManager.SelectedColor;
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Note fence TextColor: '{textColorName ?? "null"}', FenceColor: '{fenceColorName ?? "null"}'");

                // Get the actual tinted fence background color for contrast calculation
                Color actualFenceBackground = GetActualFenceBackgroundColor(fenceColorName);

                // If text color is specified, check if it has good contrast with the tinted fence background
                if (!string.IsNullOrEmpty(textColorName) && textColorName != "null")
                {
                    Color originalTextColor = Utility.GetColorFromName(textColorName);

                    // Calculate contrast ratio between text and actual fence background
                    double contrastRatio = CalculateContrastRatio(originalTextColor, actualFenceBackground);

                    // ALWAYS use the user's chosen color - never change it
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        $"Applied user's chosen text color: {textColorName} (contrast ratio: {contrastRatio:F2})");
                    //return new SolidColorBrush(originalTextColor);
                    // Use bright/vibrant version of user's chosen color for text
                    Color brightTextColor = GetBrightTextVersion(textColorName);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                        $"Applied bright text version: {textColorName} -> bright variant (contrast ratio: {contrastRatio:F2})");
                    return new SolidColorBrush(brightTextColor);

                }

                // No text color specified - use smart default based on fence background
                string smartDefault = GetSmartDefaultTextColor(actualFenceBackground);
                var defaultColor = Utility.GetColorFromName(smartDefault);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Using smart default text color: {smartDefault} for fence background");

                return new SolidColorBrush(defaultColor);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Error getting note text color, using white: {ex.Message}");
                return Brushes.White; // Safe fallback
            }
        }


        /// <summary>
        /// Converts fence background colors to bright/vibrant text versions
        /// </summary>


        private static Color GetBrightTextVersion(string colorName)
        {
            // High-contrast bright text versions
            return colorName switch
            {
                "Red" => Color.FromRgb(255, 10, 10),        // Bright red with slight warmth
                "Green" => Color.FromRgb(10, 255, 10),      // Bright green with slight warmth
                "Blue" => Color.FromRgb(50, 80, 255),     // Bright blue with better readability
                "Purple" => Color.FromRgb(191, 0, 191),   // Bright purple with better contrast
                "Orange" => Color.FromRgb(230, 136, 50),    // Bright orange
                "Yellow" => Color.FromRgb(255, 255, 2),   // Bright yellow with slight depth
                "Fuchsia" => Color.FromRgb(248, 67, 250),   // Hot pink
                "Teal" => Color.FromRgb(25, 255, 194),      // Bright teal
                "White" => Color.FromRgb(237, 237, 237),    // Pure white
                "Black" => Color.FromRgb(23, 26, 25),          // Pure black
                "Gray" => Color.FromRgb(170, 170, 170),     // Light gray
                "Beige" => Color.FromRgb(210, 144, 14),    // Bright beige
                "Bismark" => Color.FromRgb(0, 135, 224),  // Bright blue-gray
                _ => Utility.GetColorFromName(colorName)
            };
        }

        /// <summary>
        /// Gets the actual tinted fence background color as it appears visually
        /// </summary>
        private static Color GetActualFenceBackgroundColor(string fenceColorName)
        {
            if (string.IsNullOrEmpty(fenceColorName))
                return Colors.Transparent;

            Color baseColor = Utility.GetColorFromName(fenceColorName);

            // Apply the same tint logic as the fence uses
            if (SettingsManager.TintValue > 0)
            {
                // Simulate the visual effect of tinted color over transparent background
                double opacity = SettingsManager.TintValue / 100.0;
                // Blend with white background (common desktop color)
                return Color.FromRgb(
                    (byte)(baseColor.R * opacity + 255 * (1 - opacity)),
                    (byte)(baseColor.G * opacity + 255 * (1 - opacity)),
                    (byte)(baseColor.B * opacity + 255 * (1 - opacity))
                );
            }

            // No tint - assume white/transparent background
            return Colors.White;
        }

        /// <summary>
        /// Calculates contrast ratio between two colors (WCAG standard)
        /// </summary>
        private static double CalculateContrastRatio(Color foreground, Color background)
        {
            double fgLuminance = GetRelativeLuminance(foreground);
            double bgLuminance = GetRelativeLuminance(background);

            double lighter = Math.Max(fgLuminance, bgLuminance);
            double darker = Math.Min(fgLuminance, bgLuminance);

            return (lighter + 0.05) / (darker + 0.05);
        }

        /// <summary>
        /// Gets relative luminance of a color (WCAG standard)
        /// </summary>
        private static double GetRelativeLuminance(Color color)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            r = r <= 0.03928 ? r / 12.92 : Math.Pow((r + 0.055) / 1.055, 2.4);
            g = g <= 0.03928 ? g / 12.92 : Math.Pow((g + 0.055) / 1.055, 2.4);
            b = b <= 0.03928 ? b / 12.92 : Math.Pow((b + 0.055) / 1.055, 2.4);

            return 0.2126 * r + 0.7152 * g + 0.0722 * b;
        }

 
     
        /// <summary>
        /// Gets smart default text color based on fence background
        /// </summary>
        private static string GetSmartDefaultTextColor(Color fenceBackground)
        {
            double luminance = GetRelativeLuminance(fenceBackground);
            return luminance > 0.5 ? "Black" : "White";
        }



        /// <summary>
        /// Gets font size value from string
        /// </summary>
        private static double GetNoteFontSize(string fontSize)
        {
            double result;
            switch (fontSize?.ToLower())
            {
                case "small": return 11;
                case "large": return 16;
                case "extra large": return 20;
                default: return 14; // Medium

                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
        $"GetNoteFontSize: '{fontSize}' -> {result}");
                    return result;

            }
        }

        /// <summary>
        /// Gets font size value for setting changes
        /// </summary>
        public static double GetNoteFontSizeValue(string size)
        {
            return GetNoteFontSize(size);
        }

        /// <summary>
        /// Gets FontFamily from string
        /// </summary>
        private static FontFamily GetNoteFontFamily(string fontFamily)
        {
            try
            {
                return new FontFamily(fontFamily ?? "Segoe UI");
            }
            catch
            {
                return new FontFamily("Segoe UI"); // Fallback
            }
        }

        /// <summary>
        /// Applies color scheme to note background based on fence colors
        /// </summary>
        /// <summary>
        /// Applies color scheme to note background AND FOREGROUND based on fence colors
        /// </summary>
        private static void ApplyNoteColorScheme(TextBox noteTextBox, dynamic fence)
        {
            try
            {
                // 1. BACKGROUND LOGIC
                string fenceColor = fence.CustomColor?.ToString() ?? SettingsManager.SelectedColor;

                if (fenceColor != null && fenceColor != "Default")
                {
                    Color baseColor = Utility.GetColorFromName(fenceColor);
                    if (SettingsManager.TintValue > 0)
                    {
                        var tintedBrush = new SolidColorBrush(baseColor) { Opacity = SettingsManager.TintValue / 100.0 };
                        noteTextBox.Background = tintedBrush;
                    }
                    else
                    {
                        noteTextBox.Background = Brushes.Transparent;
                    }
                }
                else
                {
                    noteTextBox.Background = Brushes.Transparent;
                }

                // 2. BORDER LOGIC
                noteTextBox.BorderThickness = new Thickness(0);
                noteTextBox.BorderBrush = null;

                // 3. FOREGROUND LOGIC (THE CRITICAL FIX)
                // Force text color recalculation based on the new background
                noteTextBox.Foreground = GetNoteForeground(fence);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Applied note scheme: BG={fenceColor}, FG={noteTextBox.Foreground}");
            }
            catch (Exception ex)
            {
                // Safe defaults
                noteTextBox.Background = Brushes.Transparent;
                noteTextBox.Foreground = Brushes.White;
            }
        }


        /// <summary>
        /// Publicly accessible method to force a visual refresh of the note
        /// Call this from FenceManager.UpdateFenceProperty when CustomColor changes
        /// </summary>
        public static void RefreshNoteVisuals(dynamic fence, TextBox noteTextBox)
        {
            if (noteTextBox == null) return;

            // Re-apply the full color scheme (Background + Foreground)
            ApplyNoteColorScheme(noteTextBox, fence);
        }

        /// <summary>
        /// Gets word wrap menu text based on current state
        /// </summary>
        private static string GetWordWrapMenuText(string currentState)
        {
            bool isEnabled = currentState?.ToLower() != "false";
            return isEnabled ? "✓ Word Wrap" : "Word Wrap";
        }

        /// <summary>
        /// Gets spell check menu text based on current state
        /// </summary>
        private static string GetSpellCheckMenuText(string currentState)
        {
            bool isEnabled = currentState?.ToLower() != "false";
            return isEnabled ? "✓ Spell Check" : "Spell Check";
        }
        #endregion

        #region Note Fence Validation - Used by: Data validation

        /// <summary>
        /// Validates if a fence is a Note type fence
        /// Used by: Various fence operations to check fence type
        /// Category: Validation
        /// </summary>
        public static bool IsNoteFence(dynamic fence)
        {
            return fence?.ItemsType?.ToString() == "Note";
        }

        /// <summary>
        /// Gets default Note fence properties for new fence creation
        /// Used by: FenceDataManager.ApplyFenceDefaults() for Note fences
        /// Category: Default Values
        /// </summary>
        public static void ApplyNoteDefaults(IDictionary<string, object> fenceDict)
        {
            try
            {
                fenceDict["NoteContent"] = "";
                fenceDict["NoteFontSize"] = "Medium";
                fenceDict["NoteFontFamily"] = "Segoe UI";
                fenceDict["WordWrap"] = "true";
                fenceDict["SpellCheck"] = "true";

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.FenceCreation,
                    "Applied default properties for Note fence");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.FenceCreation,
                    $"Error applying note defaults: {ex.Message}");
            }
        }
        #endregion

        #region Tweakable Layout Configuration - Adjust these values for perfect positioning

        
        /// <summary>
        /// TWEAK THESE VALUES: Text spacing inside the TextBox
        /// </summary>
        private static Thickness GetTextBoxPadding()
        {
            double leftPadding = 6;   // TWEAK: Text distance from TextBox left edge
            double topPadding = 4;    // TWEAK: Text distance from TextBox top edge  
            double rightPadding = 6;  // TWEAK: Text distance from TextBox right edge
            double bottomPadding = 4; // TWEAK: Text distance from TextBox bottom edge

            return new Thickness(leftPadding, topPadding, rightPadding, bottomPadding);
        }
        private static Thickness GetTextBoxMargin()
        {
            double leftMargin = 8;   // TWEAK: Equal distance from all fence edges
            double rightMargin = 24;
            double topMargin = 4;     // TWEAK: Distance from top (below title)
            double bottomMargin = 28; // TWEAK: Space for Done button below textbox

            return new Thickness(leftMargin, topMargin, rightMargin, bottomMargin);
        }
        /// <summary>
        /// TWEAK THESE VALUES: TextBox border appearance
        /// </summary>
        private static (Thickness thickness, Brush brush) GetTextBoxBorder()
        {
            double borderWidth = 1;   // TWEAK: Border thickness (0=no border, 1=thin, 2=thick)

            // TWEAK: Border color - adjust Alpha for transparency
            Color borderColor = Color.FromArgb(
                80,    // TWEAK: Alpha (0=invisible, 255=solid)
                100,   // TWEAK: Red component
                100,   // TWEAK: Green component  
                100    // TWEAK: Blue component
            );

            return (new Thickness(borderWidth), new SolidColorBrush(borderColor));
        }

        /// <summary>
        /// TWEAK THESE VALUES: Done button positioning
        /// </summary>
        private static (double width, double height, Thickness margin) GetDoneButtonLayout()
        {
            double buttonWidth = 20;   // TWEAK: Done button width
            double buttonHeight = 20;  // TWEAK: Done button height

            double rightMargin = 24;    // TWEAK: Distance from right fence edge (matches textbox margin)
            double bottomMargin = 4;   // TWEAK: Distance from bottom fence edge

            return (buttonWidth, buttonHeight, new Thickness(0, 0, rightMargin, bottomMargin));
        }
        #endregion

    }
}
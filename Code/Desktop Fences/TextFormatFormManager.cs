using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using Newtonsoft.Json.Linq;

namespace Desktop_Fences
{
    /// <summary>
    /// Modern WPF form for customizing Note fence text formatting properties
    /// Follows exact same design pattern as CustomizeFenceFormManager
    /// DPI-aware with monitor positioning support
    /// </summary>
    public class TextFormatFormManager : Window
    {
        #region Private Fields
        private dynamic _fence;
        private TextBox _noteTextBox;
        private bool _result = false;

        // Form controls
        private ComboBox _cmbFontSize;
        private ComboBox _cmbFontFamily;
        private ComboBox _cmbTextColor;
        private CheckBox _chkWordWrap;
        private CheckBox _chkSpellCheck;

        // Available values from existing NoteFenceManager code
        private readonly string[] _validFontSizes = { "Small", "Medium", "Large", "Extra Large" };
        private readonly string[] _validFontFamilies = { "Segoe UI", "Arial", "Consolas", "Times New Roman", "Courier New" };
        private readonly string[] _validTextColors = { "Red", "Green", "Teal", "Blue", "Bismark", "White", "Beige", "Gray", "Black", "Purple", "Fuchsia", "Yellow", "Orange" };

        private Color _userAccentColor;

        // Store original values for cancel functionality
        private string _originalFontSize;
        private string _originalFontFamily;
        private string _originalTextColor;
        private bool _originalWordWrap;
        private bool _originalSpellCheck;
        #endregion

        #region Constructor
        /// <summary>
        /// Initialize the Text Format form with current note fence data
        /// </summary>
        /// <param name="fence">The note fence object to format</param>
        /// <param name="noteTextBox">The TextBox control to apply changes to</param>
        public TextFormatFormManager(dynamic fence, TextBox noteTextBox)
        {
            try
            {
                // Get the most current fence data from FenceManager to avoid stale references
                _fence = GetCurrentFenceData(fence);
                _noteTextBox = noteTextBox;

                // Store original values for cancel functionality
                StoreOriginalValues();

                InitializeComponent();
                LoadCurrentValues();

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Text Format form initialized for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error initializing Text Format form: {ex.Message}");
                Close();
            }
        }
        #endregion

        #region Public Properties
        /// <summary>
        /// Gets whether the user clicked Save (true) or Cancel (false)
        /// </summary>
        public new bool DialogResult => _result;
        #endregion

        #region Form Initialization
        private void InitializeComponent()
        {
            try
            {
                // Get user's accent color for modern design elements (same as customize form)
                string selectedColorName = SettingsManager.SelectedColor;
                _userAccentColor = Utility.GetColorFromName(selectedColorName);

                // Modern WPF window setup with DPI awareness (same as customize form)
                this.Title = "Text Format";
                this.Width = 400;
                this.Height = 400;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;

                // Set icon from executable (same as customize form)
                try
                {
                    this.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { } // Ignore icon loading errors

                // Main container with modern card design (same as customize form)
                Border mainCard = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                    BorderThickness = new Thickness(1),
                    Margin = new Thickness(8),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 270,
                        ShadowDepth = 2,
                        BlurRadius = 10,
                        Opacity = 0.1
                    }
                };

                // Main grid layout (same as customize form)
                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) }); // Header
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) }); // Footer

                CreateHeader(mainGrid);
                CreateContent(mainGrid);
                CreateFooter(mainGrid);

                mainCard.Child = mainGrid;
                this.Content = mainCard;

                // Position form on the screen where mouse is currently located (same as customize form)
                PositionFormOnMouseScreen();

                // Add keyboard support for Enter/Escape keys (same as customize form)
                this.KeyDown += TextFormatForm_KeyDown;
                this.Focusable = true;
                this.Focus();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Initialized TextFormatForm for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error initializing TextFormatForm: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error initializing form: {ex.Message}", "Form Error");
            }
        }

        /// <summary>
        /// Handles keyboard input for TextFormatForm
        /// Enter = Save, Escape = Cancel
        /// </summary>
        private void TextFormatForm_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Enter:
                    SaveButton_Click(this, new RoutedEventArgs());
                    e.Handled = true;
                    break;
                case Key.Escape:
                    CancelButton_Click(this, new RoutedEventArgs());
                    e.Handled = true;
                    break;
            }
        }

        private void CreateHeader(Grid parent)
        {
            // Header border with accent color background (same as customize form)
            Border headerBorder = new Border
            {
                Background = new SolidColorBrush(_userAccentColor),
                Height = 50
            };
            Grid.SetRow(headerBorder, 0);

            // Header grid for layout (same as customize form)
            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });

            // Title label (same style as customize form)
            TextBlock titleBlock = new TextBlock
            {
                Text = "Text Format",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };
            Grid.SetColumn(titleBlock, 0);

            // Close button (✕) (same style as customize form)
            Button closeButton = new Button
            {
                Content = "✕",
                Width = 32,
                Height = 32,
                FontSize = 16,
                FontFamily = new FontFamily("Segoe UI"),
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            closeButton.Click += CloseButton_Click;
            closeButton.MouseEnter += (s, e) => closeButton.Background = new SolidColorBrush(Color.FromArgb(50, 255, 255, 255));
            closeButton.MouseLeave += (s, e) => closeButton.Background = Brushes.Transparent;
            Grid.SetColumn(closeButton, 1);

            headerGrid.Children.Add(titleBlock);
            headerGrid.Children.Add(closeButton);
            headerBorder.Child = headerGrid;
            parent.Children.Add(headerBorder);
        }

        private void CreateContent(Grid parent)
        {
            // Content scroll viewer (same as customize form)
            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Background = Brushes.White,
                Padding = new Thickness(16)
            };
            Grid.SetRow(scrollViewer, 1);

            // Main content stack panel (same as customize form)
            StackPanel contentStack = new StackPanel
            {
                Orientation = Orientation.Vertical
            };

            CreateTextAppearanceSection(contentStack);
            CreateTextBehaviorSection(contentStack);

            scrollViewer.Content = contentStack;
            parent.Children.Add(scrollViewer);
        }

        private void CreateFooter(Grid parent)
        {
            // Footer border (same style as customize form)
            Border footerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(16, 8, 16, 8)
            };
            Grid.SetRow(footerBorder, 2);

            // Button panel (same as customize form)
            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            // Apply button with green color (same as Default button in customize form)
            Button applyButton = new Button
            {
                Content = "Apply",
                Width = 100,
                Height = 34,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(Color.FromRgb(34, 139, 34)), // Green like Default button
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 10, 0)
            };
            applyButton.Click += ApplyButton_Click;

            // Cancel button (same style as customize form)
            Button cancelButton = new Button
            {
                Content = "Cancel",
                Width = 100,
                Height = 33,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 10, 0)
            };
            cancelButton.Click += CancelButton_Click;

            // Save button with accent color (same as customize form)
            Button saveButton = new Button
            {
                Content = "Save",
                Width = 100,
                Height = 34,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Background = new SolidColorBrush(_userAccentColor),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            saveButton.Click += SaveButton_Click;

            buttonPanel.Children.Add(applyButton);
            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(saveButton);

            footerBorder.Child = buttonPanel;
            parent.Children.Add(footerBorder);
        }

        private void CreateTextAppearanceSection(StackPanel parent)
        {
            // GroupBox with same style as customize form
            GroupBox textAppearanceGroupBox = new GroupBox
            {
                Header = "Text Appearance",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(_userAccentColor),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(8)
            };

            StackPanel textAppearanceStack = new StackPanel { Orientation = Orientation.Vertical };

            CreateDropdownField(textAppearanceStack, "Font Size:", _validFontSizes, out _cmbFontSize);
            CreateDropdownField(textAppearanceStack, "Font Family:", _validFontFamilies, out _cmbFontFamily);
            CreateDropdownField(textAppearanceStack, "Text Color:", _validTextColors, out _cmbTextColor);

            textAppearanceGroupBox.Content = textAppearanceStack;
            parent.Children.Add(textAppearanceGroupBox);
        }

        private void CreateTextBehaviorSection(StackPanel parent)
        {
            // GroupBox with same style as customize form
            GroupBox textBehaviorGroupBox = new GroupBox
            {
                Header = "Text Behavior",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(_userAccentColor),
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(8)
            };

            StackPanel textBehaviorStack = new StackPanel { Orientation = Orientation.Vertical };

            CreateCheckboxField(textBehaviorStack, "Word Wrap", out _chkWordWrap);
            CreateCheckboxField(textBehaviorStack, "Spell Check", out _chkSpellCheck);

            textBehaviorGroupBox.Content = textBehaviorStack;
            parent.Children.Add(textBehaviorGroupBox);
        }

        #region Helper Methods for Control Creation (copied exactly from CustomizeFenceFormManager)
        private void CreateDropdownField(StackPanel parent, string labelText, string[] items, out ComboBox comboBox)
        {
            Grid fieldGrid = new Grid
            {
                Margin = new Thickness(0, 5, 0, 5)
            };
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(180) });
            fieldGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock label = new TextBlock
            {
                Text = labelText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };
            Grid.SetColumn(label, 0);

            comboBox = new ComboBox
            {
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Width = 180,
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Center
            };

            comboBox.Items.Add("Default");
            foreach (var item in items)
            {
                comboBox.Items.Add(item);
            }
            comboBox.SelectedIndex = 0;

            Grid.SetColumn(comboBox, 1);

            fieldGrid.Children.Add(label);
            fieldGrid.Children.Add(comboBox);
            parent.Children.Add(fieldGrid);
        }

        private void CreateCheckboxField(StackPanel parent, string labelText, out CheckBox checkBox)
        {
            checkBox = new CheckBox
            {
                Content = labelText,
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 12,
                FontWeight = FontWeights.Normal,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                Margin = new Thickness(16, 5, 0, 5),
                VerticalAlignment = VerticalAlignment.Center
            };

            parent.Children.Add(checkBox);
        }
        #endregion

        private void PositionFormOnMouseScreen()
        {
            try
            {
                // Get current cursor position (same logic as customize form)
                var cursorPosition = System.Windows.Forms.Cursor.Position;

                // Convert to WPF coordinates and center form around cursor
                this.Left = cursorPosition.X - (this.Width / 2);
                this.Top = cursorPosition.Y - (this.Height / 2);

                // Ensure form stays on screen
                if (this.Left < 0) this.Left = 0;
                if (this.Top < 0) this.Top = 0;
                if (this.Left + this.Width > SystemParameters.PrimaryScreenWidth)
                    this.Left = SystemParameters.PrimaryScreenWidth - this.Width;
                if (this.Top + this.Height > SystemParameters.PrimaryScreenHeight)
                    this.Top = SystemParameters.PrimaryScreenHeight - this.Height;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Positioned TextFormatForm at ({this.Left}, {this.Top})");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error positioning form: {ex.Message}");
                this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }
        #endregion

        #region Current Fence Data Retrieval (same as customize form)
        private dynamic GetCurrentFenceData(dynamic originalFence)
        {
            try
            {
                string fenceId = originalFence.Id?.ToString();
                if (string.IsNullOrEmpty(fenceId))
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Original fence '{originalFence.Title}' has no Id, using original reference");
                    return originalFence;
                }

                var fenceData = FenceManager.GetFenceData();
                dynamic currentFence = fenceData.FirstOrDefault(f => f.Id?.ToString() == fenceId);

                if (currentFence != null)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Retrieved current fence data for '{currentFence.Title}' (Id: {fenceId})");
                    return currentFence;
                }
                else
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not find current fence data for Id '{fenceId}', using original reference");
                    return originalFence;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error retrieving current fence data: {ex.Message}, using original reference");
                return originalFence;
            }
        }
        #endregion

        #region Data Loading and Storage
        private void StoreOriginalValues()
        {
            try
            {
                _originalFontSize = GetSafeNoteProperty(_fence, "NoteFontSize", "Medium");
                _originalFontFamily = GetSafeNoteProperty(_fence, "NoteFontFamily", "Segoe UI");
                _originalTextColor = GetSafeNoteProperty(_fence, "TextColor", "Default");
                _originalWordWrap = GetSafeNoteProperty(_fence, "WordWrap", "true").ToLower() == "true";
                _originalSpellCheck = GetSafeNoteProperty(_fence, "SpellCheck", "true").ToLower() == "true";

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Stored original values for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error storing original values: {ex.Message}");
            }
        }

        private void LoadCurrentValues()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Loading current values for fence '{_fence.Title}'");

                // Load Font Size (same pattern as customize form)
                LoadDropdownValue(_cmbFontSize, _fence.NoteFontSize?.ToString(), "NoteFontSize");

                // Load Font Family
                LoadDropdownValue(_cmbFontFamily, _fence.NoteFontFamily?.ToString(), "NoteFontFamily");

                // Load Text Color
                LoadDropdownValue(_cmbTextColor, _fence.TextColor?.ToString(), "TextColor");

                // Load Word Wrap
                LoadCheckboxValue(_chkWordWrap, _fence.WordWrap?.ToString(), "WordWrap");

                // Load Spell Check
                LoadCheckboxValue(_chkSpellCheck, _fence.SpellCheck?.ToString(), "SpellCheck");

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Successfully loaded all current values for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error loading current values for fence '{_fence.Title}': {ex.Message}");
            }
        }

        // Helper methods copied exactly from CustomizeFenceFormManager
        private void LoadDropdownValue(ComboBox comboBox, string currentValue, string propertyName)
        {
            try
            {
                if (string.IsNullOrEmpty(currentValue))
                {
                    comboBox.SelectedIndex = 0; // Default
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to Default (null/empty value)");
                }
                else
                {
                    // Find matching item in ComboBox
                    for (int i = 0; i < comboBox.Items.Count; i++)
                    {
                        if (comboBox.Items[i].ToString().Equals(currentValue, StringComparison.OrdinalIgnoreCase))
                        {
                            comboBox.SelectedIndex = i;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to '{currentValue}' (index {i})");
                            return;
                        }
                    }
                    // If not found, set to Default
                    comboBox.SelectedIndex = 0;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Value '{currentValue}' not found for {propertyName}, set to Default");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading {propertyName}: {ex.Message}");
                comboBox.SelectedIndex = 0;
            }
        }

        private void LoadCheckboxValue(CheckBox checkBox, string currentValue, string propertyName)
        {
            try
            {
                if (string.IsNullOrEmpty(currentValue))
                {
                    checkBox.IsChecked = propertyName == "WordWrap" || propertyName == "SpellCheck"; // Default true for these
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to default (null/empty value)");
                }
                else
                {
                    bool parsedValue = currentValue.Equals("true", StringComparison.OrdinalIgnoreCase);
                    checkBox.IsChecked = parsedValue;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Set {propertyName} to '{parsedValue}'");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading {propertyName}: {ex.Message}");
                checkBox.IsChecked = false;
            }
        }

        private string GetDropdownValue(ComboBox comboBox)
        {
            string selectedValue = comboBox.SelectedItem?.ToString() ?? "Default";
            return selectedValue == "Default" ? null : selectedValue;
        }

        // Safe property getter (local implementation)
        private string GetSafeNoteProperty(dynamic fence, string propertyName, string fallbackValue)
        {
            try
            {
                // Method 1: Dictionary access  
                try
                {
                    var fenceDict = fence as IDictionary<string, object>;
                    if (fenceDict != null && fenceDict.ContainsKey(propertyName))
                    {
                        return fenceDict[propertyName]?.ToString() ?? fallbackValue;
                    }
                }
                catch { }

                // Method 2: JObject access (for JSON loaded fences)
                try
                {
                    var jObject = fence as Newtonsoft.Json.Linq.JObject;
                    if (jObject != null && jObject.ContainsKey(propertyName))
                    {
                        return jObject[propertyName]?.ToString() ?? fallbackValue;
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

        // Font size conversion (local implementation)
        private double GetNoteFontSizeValue(string size)
        {
            switch (size?.ToLower())
            {
                case "small": return 11;
                case "large": return 16;
                case "extra large": return 20;
                default: return 14; // Medium
            }
        }
        #endregion

        #region Button Event Handlers
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            _result = false;
            this.Close();
        }

        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ApplyChanges();
                SaveChangesToJson();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Applied text format changes for fence '{_fence.Title}' (form remains open)");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error applying changes: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error applying changes: {ex.Message}", "Apply Error");
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ApplyChanges();
                SaveChangesToJson();
                _result = true;
                Close();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Saved text format changes for fence '{_fence.Title}' and closed form");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error saving changes: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error saving changes: {ex.Message}", "Save Error");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
               // RestoreOriginalValues();
                _result = false;
                Close();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Cancelled text format changes for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error cancelling: {ex.Message}");
                Close();
            }
        }
        #endregion

        #region Apply and Save Logic
        private void ApplyChanges()
        {
            try
            {
                //// Apply Font Size
                string selectedFontSize = _cmbFontSize.SelectedItem?.ToString() ?? "Medium";
                //if (selectedFontSize != "Default")
                //{
                //    double fontSize = GetNoteFontSizeValue(selectedFontSize);
                //    _noteTextBox.FontSize = fontSize;
                //}

                //// Apply Font Family
                string selectedFontFamily = _cmbFontFamily.SelectedItem?.ToString() ?? "Segoe UI";
                //if (selectedFontFamily != "Default")
                //{
                //    _noteTextBox.FontFamily = new FontFamily(selectedFontFamily);
                //}

                // Apply Font Size - ALWAYS apply something
                string fontSizeToApply = selectedFontSize == "Default" ? "Medium" : selectedFontSize;
                double fontSize = GetNoteFontSizeValue(fontSizeToApply);
                _noteTextBox.FontSize = fontSize;

                // Apply Font Family - ALWAYS apply something  
                string fontFamilyToApply = selectedFontFamily == "Default" ? "Segoe UI" : selectedFontFamily;
                _noteTextBox.FontFamily = new FontFamily(fontFamilyToApply);



                // Apply Text Color
                string selectedTextColor = GetDropdownValue(_cmbTextColor);
                if (!string.IsNullOrEmpty(selectedTextColor))
                {
                    var textColor = Utility.GetColorFromName(selectedTextColor);
                    _noteTextBox.Foreground = new SolidColorBrush(textColor);
                }
                else
                {
                    // Use default white color
                    _noteTextBox.Foreground = Brushes.White;
                }

                // Apply Word Wrap
                bool wordWrapEnabled = _chkWordWrap.IsChecked ?? true;
                _noteTextBox.TextWrapping = wordWrapEnabled ? TextWrapping.Wrap : TextWrapping.NoWrap;

                // Apply Spell Check
                bool spellCheckEnabled = _chkSpellCheck.IsChecked ?? true;
                _noteTextBox.SpellCheck.IsEnabled = spellCheckEnabled;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Applied all text format changes to TextBox for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error applying changes to TextBox: {ex.Message}");
                throw;
            }
        }

        private void SaveChangesToJson()
        {
            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Saving text format properties to JSON for fence '{_fence.Title}'");

                // Get values from all controls (same pattern as customize form)
                string fontSizeValue = GetDropdownValue(_cmbFontSize);
                string fontFamilyValue = GetDropdownValue(_cmbFontFamily);
                string textColorValue = GetDropdownValue(_cmbTextColor);
                string wordWrapValue = (_chkWordWrap.IsChecked ?? true).ToString().ToLower();
                string spellCheckValue = (_chkSpellCheck.IsChecked ?? true).ToString().ToLower();

                // Save all properties using existing UpdateFenceProperty method (same as customize form)
                FenceManager.UpdateFenceProperty(_fence, "NoteFontSize", fontSizeValue ?? "Medium", $"NoteFontSize updated to '{fontSizeValue ?? "Medium"}'");
                FenceManager.UpdateFenceProperty(_fence, "NoteFontFamily", fontFamilyValue ?? "Segoe UI", $"NoteFontFamily updated to '{fontFamilyValue ?? "Segoe UI"}'");
                FenceManager.UpdateFenceProperty(_fence, "TextColor", textColorValue, $"TextColor updated to '{textColorValue}'");
                FenceManager.UpdateFenceProperty(_fence, "WordWrap", wordWrapValue, $"WordWrap updated to '{wordWrapValue}'");
                FenceManager.UpdateFenceProperty(_fence, "SpellCheck", spellCheckValue, $"SpellCheck updated to '{spellCheckValue}'");

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"Saved all text format settings for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error saving changes to fence data: {ex.Message}");
                throw;
            }
        }

        private void RestoreOriginalValues()
        {
            try
            {
                // Restore Font Size
                if (_originalFontSize != "Default")
                {
                    double fontSize = GetNoteFontSizeValue(_originalFontSize);
                    _noteTextBox.FontSize = fontSize;
                }

                // Restore Font Family
                if (_originalFontFamily != "Default")
                {
                    _noteTextBox.FontFamily = new FontFamily(_originalFontFamily);
                }

                // Restore Text Color
                if (!string.IsNullOrEmpty(_originalTextColor) && _originalTextColor != "Default")
                {
                    var textColor = Utility.GetColorFromName(_originalTextColor);
                    _noteTextBox.Foreground = new SolidColorBrush(textColor);
                }
                else
                {
                    _noteTextBox.Foreground = Brushes.White;
                }

                // Restore Word Wrap
                _noteTextBox.TextWrapping = _originalWordWrap ? TextWrapping.Wrap : TextWrapping.NoWrap;

                // Restore Spell Check
                _noteTextBox.SpellCheck.IsEnabled = _originalSpellCheck;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"Restored original values for fence '{_fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error restoring original values: {ex.Message}");
            }
        }
        #endregion

        #region Public Methods
        /// <summary>
        /// Shows the Text Format form and returns the result
        /// </summary>
        /// <returns>True if changes were saved, false if cancelled</returns>
        public new bool ShowDialog()
        {
            try
            {
                base.ShowDialog();
                return _result;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error showing Text Format dialog: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Static method to show Text Format form for a Note fence
        /// </summary>
        /// <param name="fence">The note fence to format</param>
        /// <param name="noteTextBox">The TextBox control to apply changes to</param>
        /// <returns>True if changes were saved, false if cancelled</returns>
        public static bool ShowTextFormatForm(dynamic fence, TextBox noteTextBox)
        {
            try
            {
                if (fence?.ItemsType?.ToString() != "Note")
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                        "Text Format form can only be used with Note fences");
                    return false;
                }

                var form = new TextFormatFormManager(fence, noteTextBox);
                return form.ShowDialog();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error showing Text Format form: {ex.Message}");
                return false;
            }
        }
        #endregion

        #region Window Drag Support (same as EditShortcutWindow)
        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
        #endregion
    }
}
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using Desktop_Fences;
using IWshRuntimeLibrary;

public partial class EditShortcutWindow : Window
{
    private string shortcutPath;
    public string NewDisplayName { get; private set; }
    private TextBox nameBox;
    private TextBox iconPathBox;
    private Image iconPreview; // For future icon preview functionality
    private TextBox targetPathBox;
    private TextBox argumentsBox;
    private Button _saveButton;

    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);

    public EditShortcutWindow(string shortcutPath, string currentDisplayName)
    {
        this.shortcutPath = shortcutPath;
        NewDisplayName = currentDisplayName;
        InitializeModernComponent();
    }

    private void InitializeModernComponent()
    {
        Title = "Edit Shortcut";
        Width = 540;
        Height = 550;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.NoResize;
        Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
        WindowStyle = WindowStyle.None;
        AllowsTransparency = false;

        //Effect = new DropShadowEffect
        //{
        //    Color = Colors.Black,
        //    Opacity = 0.15,
        //    BlurRadius = 12,
        //    ShadowDepth = 4,
        //    Direction = 270
        //};

        // Create main container like CustomizeFenceForm
        Grid mainContainer = new Grid
        {
            Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
            Margin = new Thickness(8)
        };

        // Create main card border (like CustomizeFenceForm mainCard)
        Border mainBorder = new Border
        {
            Background = Brushes.White,
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(0) // No rounded corners like CustomizeFenceForm
        };

        Grid rootGrid = new Grid();
        rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
        rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
        rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Footer

        CreateModernHeader(rootGrid);
        CreateModernContent(rootGrid);
        CreateModernFooter(rootGrid);

        mainBorder.Child = rootGrid;
        mainContainer.Children.Add(mainBorder);
        Content = mainContainer;

        // Initialize with current values
        LoadCurrentValues();

        // Add keyboard support for Enter/Escape keys
        this.KeyDown += EditShortcutWindow_KeyDown;
        this.Focusable = true;
        this.Focus();
    }

    /// <summary>
    /// Handles keyboard input for EditShortcutWindow
    /// Enter = Save, Escape = Cancel
    /// </summary>
    private void EditShortcutWindow_KeyDown(object sender, KeyEventArgs e)
    {
        switch (e.Key)
        {
            case Key.Enter:
                // Trigger Save button logic
                Save_Click(this, new RoutedEventArgs());
                e.Handled = true;
                break;
            case Key.Escape:
                // Trigger Cancel button logic
                DialogResult = false;
                Close();
                e.Handled = true;
                break;
        }
    }

    private void CreateModernHeader(Grid rootGrid)
    {
        // Header with accent color background (same as CustomizeFenceForm)
        Border headerBorder = new Border
        {
            Height = 50,
            Background = GetAccentColorBrush(), // Use user's selected theme color
            CornerRadius = new CornerRadius(0) // No rounded corners like CustomizeFenceForm
        };

        Grid headerGrid = new Grid();
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        // Title label with white text (same as CustomizeFenceForm)
        TextBlock titleText = new TextBlock
        {
            Text = "Edit Shortcut",
            FontSize = 14,
            FontWeight = FontWeights.Bold, // Bold like CustomizeFenceForm
            Foreground = Brushes.White, // White text on colored background
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(16, 0, 0, 0)
        };

        // Close button with hover effects (same as CustomizeFenceForm)
        Button closeButton = new Button
        {
            Content = "✕",
            Width = 32,
            Height = 32,
            FontSize = 12,
            FontWeight = FontWeights.Bold,
            Foreground = Brushes.White,
            Background = Brushes.Transparent,
            BorderBrush = Brushes.Transparent,
            Cursor = Cursors.Hand,
            Margin = new Thickness(0, 0, 9, 0),
            VerticalAlignment = VerticalAlignment.Center
        };

        // Add hover effect like CustomizeFenceForm
        closeButton.MouseEnter += (s, e) => closeButton.Background = new SolidColorBrush(Color.FromArgb(50, 255, 255, 255));
        closeButton.MouseLeave += (s, e) => closeButton.Background = Brushes.Transparent;
        closeButton.Click += (s, e) => { DialogResult = false; Close(); };

        headerGrid.Children.Add(titleText);
        headerGrid.Children.Add(closeButton);
        Grid.SetColumn(closeButton, 1);

        headerBorder.Child = headerGrid;
        Grid.SetRow(headerBorder, 0);
        rootGrid.Children.Add(headerBorder);
    }

    private void CreateModernContent(Grid rootGrid)
    {
        Border contentBorder = new Border
        {
            Background = Brushes.White,
            Padding = new Thickness(20, 10, 20, 10)
        };

        StackPanel contentPanel = new StackPanel
        {
            Orientation = Orientation.Vertical
        };

        // Display Name Section
        CreateFieldSection(contentPanel, "Display Name:", out nameBox, NewDisplayName);

        // Target Path Section
        CreateFieldWithButtonSection(contentPanel, "Target Path:", out targetPathBox, GetCurrentTargetPath(), "Browse...", BrowseTarget_Click);

        // Arguments Section
        CreateFieldSection(contentPanel, "Arguments:", out argumentsBox, GetCurrentArguments());

        // Icon Section with Preview Space
        CreateIconSection(contentPanel);

        contentBorder.Child = contentPanel;
        Grid.SetRow(contentBorder, 1);
        rootGrid.Children.Add(contentBorder);
    }

    private void CreateFieldSection(StackPanel parent, string labelText, out TextBox textBox, string initialValue)
    {
        Border fieldBorder = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            CornerRadius = new CornerRadius(6),
            Margin = new Thickness(0, 0, 0, 12),
            Padding = new Thickness(12, 12, 12, 12)
        };

        StackPanel fieldPanel = new StackPanel();

        TextBlock label = new TextBlock
        {
            Text = labelText,
            FontSize = 12,
            FontWeight = FontWeights.Medium,
            Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
            Margin = new Thickness(0, 0, 0, 8)
        };

        textBox = new TextBox
        {
            Text = initialValue ?? "",
            FontSize = 13,
            Padding = new Thickness(8, 6, 8, 6),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            Background = Brushes.White,
            Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36))
        };

        fieldPanel.Children.Add(label);
        fieldPanel.Children.Add(textBox);
        fieldBorder.Child = fieldPanel;
        parent.Children.Add(fieldBorder);
    }

    private void CreateFieldWithButtonSection(StackPanel parent, string labelText, out TextBox textBox, string initialValue, string buttonText, RoutedEventHandler buttonClick)
    {
        Border fieldBorder = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            CornerRadius = new CornerRadius(6),
            Margin = new Thickness(0, 0, 0, 12),
            Padding = new Thickness(12, 12, 12, 12)
        };

        StackPanel fieldPanel = new StackPanel();

        TextBlock label = new TextBlock
        {
            Text = labelText,
            FontSize = 12,
            FontWeight = FontWeights.Medium,
            Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
            Margin = new Thickness(0, 0, 0, 8)
        };

        Grid inputGrid = new Grid();
        inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        textBox = new TextBox
        {
            Text = initialValue ?? "",
            FontSize = 13,
            Padding = new Thickness(10, 8, 10, 8),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            Background = Brushes.White,
            Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
            Margin = new Thickness(0, 0, 10, 0)
        };

        Button browseButton = CreateModernButton(buttonText, false);
        browseButton.Click += buttonClick;

        inputGrid.Children.Add(textBox);
        inputGrid.Children.Add(browseButton);
        Grid.SetColumn(browseButton, 1);

        fieldPanel.Children.Add(label);
        fieldPanel.Children.Add(inputGrid);
        fieldBorder.Child = fieldPanel;
        parent.Children.Add(fieldBorder);
    }

    private void CreateIconSection(StackPanel parent)
    {
        Border iconBorder = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            CornerRadius = new CornerRadius(6),
            Margin = new Thickness(0, 0, 0, 12),
            Padding = new Thickness(12, 12, 12, 12)
        };

        StackPanel iconPanel = new StackPanel();

        TextBlock iconLabel = new TextBlock
        {
            Text = "Icon:",
            FontSize = 12,
            FontWeight = FontWeights.Medium,
            Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
            Margin = new Thickness(0, 0, 0, 8)
        };

        Grid iconGrid = new Grid();
        iconGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        iconGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        iconGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        iconPathBox = new TextBox
        {
            Text = GetCurrentIconLocation(),
            FontSize = 13,
            Padding = new Thickness(8, 6, 8, 6),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            Background = Brushes.White,
            Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
            Margin = new Thickness(0, 0, 10, 0),
            IsReadOnly = true
        };

        // Icon Preview Container (prepared for future functionality)
        Border iconPreviewBorder = new Border
        {
            Width = 36,
            Height = 36,
            Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            CornerRadius = new CornerRadius(4),
            Margin = new Thickness(0, 0, 10, 0)
        };

        iconPreview = new Image
        {
            Width = 28,
            Height = 28,
            Stretch = Stretch.Uniform,
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center
        };

        iconPreviewBorder.Child = iconPreview;

        Button browseIconButton = CreateModernButton("Browse...", false);
        browseIconButton.Click += BrowseIcon_Click;

        iconGrid.Children.Add(iconPathBox);
        iconGrid.Children.Add(iconPreviewBorder);
        iconGrid.Children.Add(browseIconButton);
        Grid.SetColumn(iconPreviewBorder, 1);
        Grid.SetColumn(browseIconButton, 2);

        iconPanel.Children.Add(iconLabel);
        iconPanel.Children.Add(iconGrid);
        iconBorder.Child = iconPanel;
        parent.Children.Add(iconBorder);
    }

    private void CreateModernFooter(Grid rootGrid)
    {
        Border footerBorder = new Border
        {
            Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(0, 1, 0, 0), // Only top border
            Padding = new Thickness(20, 16, 20, 16),
            CornerRadius = new CornerRadius(0) // No rounded corners
        };

        StackPanel buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };

        // Green default button (same as CustomizeFenceForm)
        Button defaultButton = new Button
        {
            Content = "Default",
            Height = 36,
            MinWidth = 80,
            FontSize = 13,
            FontWeight = FontWeights.Bold,
            Cursor = Cursors.Hand,
            BorderThickness = new Thickness(0),
            Padding = new Thickness(16, 0, 16, 0),
            Background = new SolidColorBrush(Color.FromArgb(255,34, 139, 34)), // Green like CustomizeFenceForm
            Foreground = Brushes.White,
            Margin = new Thickness(0, 0, 10, 0)
        };
        defaultButton.Click += Default_Click;

        Button cancelButton = CreateModernButton("Cancel", false);
        cancelButton.Margin = new Thickness(0, 0, 10, 0);
        cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

        // Save button with accent color (same as CustomizeFenceForm)
        _saveButton = new Button
        {
            Content = "Save",
            Height = 36,
            MinWidth = 80,
            FontSize = 13,
            FontWeight = FontWeights.Bold,
            Cursor = Cursors.Hand,
            BorderThickness = new Thickness(0),
            Padding = new Thickness(16, 0, 16, 0),
            Background = GetAccentColorBrush(), // Use user's selected theme color
            Foreground = Brushes.White
        };
        _saveButton.Click += Save_Click;

        buttonPanel.Children.Add(defaultButton);
        buttonPanel.Children.Add(cancelButton);
        buttonPanel.Children.Add(_saveButton);

        footerBorder.Child = buttonPanel;

        Grid.SetRow(footerBorder, 2);
        rootGrid.Children.Add(footerBorder);
    }

    private Button CreateModernButton(string text, bool isPrimary)
    {
        Button button = new Button
        {
            Content = text,
            Height = 36,
            MinWidth = 80,
            FontSize = 13,
            FontWeight = FontWeights.Medium,
            Cursor = Cursors.Hand,
            BorderThickness = new Thickness(1, 1, 1, 1),
            Padding = new Thickness(16, 0, 16, 0)
        };

        if (isPrimary)
        {
            button.Background = new SolidColorBrush(Color.FromRgb(66, 133, 244));
            button.Foreground = Brushes.White;
            button.BorderBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244));
        }
        else
        {
            button.Background = Brushes.White;
            button.Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36));
            button.BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224));
        }

        return button;
    }

    private void LoadCurrentValues()
    {
        // Load current values - existing functionality preserved
        try
        {
            nameBox.Text = NewDisplayName ?? "";
            targetPathBox.Text = GetCurrentTargetPath();
            argumentsBox.Text = GetCurrentArguments();
            iconPathBox.Text = GetCurrentIconLocation();

            // Load the actual icon being used on the fence
            LoadActualFenceIcon();
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading current values: {ex.Message}");
        }
    }

    // Existing methods preserved - just keeping method signatures for compatibility
    private string GetCurrentTargetPath()
    {
        try
        {
            // FIX: Manual parsing for .url files
            // WshShell often fails to read the "Target" of a .url file if the header is non-standard.
            if (System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url")
            {
                if (System.IO.File.Exists(shortcutPath))
                {
                    // Read file as text to find URL=...
                    string[] lines = System.IO.File.ReadAllLines(shortcutPath);
                    foreach (string line in lines)
                    {
                        if (line.Trim().StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                        {
                            return line.Trim().Substring(4);
                        }
                    }
                }
            }

            // Fallback to WshShell for .lnk files
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            return shortcut.TargetPath ?? "";
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error getting target path: {ex.Message}");
            return "";
        }
    }

    private string GetCurrentArguments()
    {
        try
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            return shortcut.Arguments ?? "";
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error getting arguments: {ex.Message}");
            return "";
        }
    }



    private string GetCurrentIconLocation()
    {
        try
        {
            // FIX: Robust manual parsing for .url files
            if (System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url")
            {
                if (System.IO.File.Exists(shortcutPath))
                {
                    var lines = System.IO.File.ReadAllLines(shortcutPath);
                    string iconFile = null;
                    string iconIndex = null;

                    foreach (var line in lines)
                    {
                        string trimmed = line.Trim();
                        if (trimmed.StartsWith("IconFile=", StringComparison.OrdinalIgnoreCase))
                            iconFile = trimmed.Substring(9).Trim();
                        else if (trimmed.StartsWith("IconIndex=", StringComparison.OrdinalIgnoreCase))
                            iconIndex = trimmed.Substring(10).Trim();
                    }

                    if (!string.IsNullOrEmpty(iconFile))
                    {
                        // Return in format "path,index" or just "path"
                        return string.IsNullOrEmpty(iconIndex) ? iconFile : $"{iconFile},{iconIndex}";
                    }
                }
                return "Default";
            }

            // Standard .lnk handling
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            return string.IsNullOrEmpty(shortcut.IconLocation) ? "Default" : shortcut.IconLocation;
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error getting icon location: {ex.Message}");
            return "Default";
        }
    }



    private void BrowseTarget_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var fileDialog = new System.Windows.Forms.OpenFileDialog
            {
                Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*",
                Title = "Select Target File",
                CheckFileExists = true,
                CheckPathExists = true
            };

            string currentTarget = targetPathBox?.Text ?? "";
            if (!string.IsNullOrEmpty(currentTarget) && System.IO.File.Exists(currentTarget))
            {
                fileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(currentTarget);
                fileDialog.FileName = System.IO.Path.GetFileName(currentTarget);
            }

            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                targetPathBox.Text = fileDialog.FileName;
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Target path updated to: {fileDialog.FileName}");
            }
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error in target browse dialog: {ex.Message}");
            MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error opening file browser: {ex.Message}", "Browse Error");
        }
    }

    private void BrowseIcon_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.OpenFileDialog
        {
            Filter = "Icon Files (*.ico)|*.ico|Executable Files (*.exe)|*.exe|DLL Files (*.dll)|*.dll|All Files (*.*)|*.*",
            Title = "Select an Icon"
        };

        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            string selectedPath = dialog.FileName;
            string extension = System.IO.Path.GetExtension(selectedPath).ToLower();

            if (extension == ".dll" || extension == ".exe")
            {
                var picker = new IconPickerDialog(selectedPath);
                if (picker.ShowDialog() == true && picker.SelectedIndex >= 0)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"User selected icon index: {picker.SelectedIndex} from {selectedPath}");
                    iconPathBox.Text = $"{selectedPath},{picker.SelectedIndex}";
                    LoadIconPreview(iconPathBox.Text);
                }
            }
            else
            {
                iconPathBox.Text = selectedPath;
                LoadIconPreview(iconPathBox.Text);
            }
        }
    }

    private void Default_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Restoring default values for shortcut: {shortcutPath}");

            string originalTarget = "";
            string originalName = System.IO.Path.GetFileNameWithoutExtension(shortcutPath) ?? "";
            bool isUrl = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url";

            if (isUrl)
            {
                // FIX: Manual parsing for .url files to avoid COM errors
                if (System.IO.File.Exists(shortcutPath))
                {
                    var lines = System.IO.File.ReadAllLines(shortcutPath);
                    foreach (var line in lines)
                    {
                        if (line.Trim().StartsWith("URL=", StringComparison.OrdinalIgnoreCase))
                        {
                            originalTarget = line.Trim().Substring(4);
                            break;
                        }
                    }
                }
            }
            else
            {
                // Standard .lnk handling
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                originalTarget = shortcut.TargetPath ?? "";
            }

            // Reset UI Fields
            nameBox.Text = originalName;
            targetPathBox.Text = originalTarget;
            argumentsBox.Text = "";

            // Setting this to "Default" triggers the removal logic in Save_Click
            iconPathBox.Text = "Default";

            // Force the preview to update immediately
            LoadIconPreview("Default");

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Default values restored successfully");
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error restoring defaults: {ex.Message}");
            MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error restoring default values: {ex.Message}", "Default Error");
        }
    }


    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Starting save process for shortcut: {shortcutPath}");

            string newDisplayName = nameBox?.Text?.Trim() ?? "";
            string newTargetPath = targetPathBox?.Text?.Trim() ?? "";
            string newIconPath = iconPathBox?.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(newDisplayName) || string.IsNullOrWhiteSpace(newTargetPath))
            {
                MessageBoxesManager.ShowOKOnlyMessageBoxForm("Display name and Target path cannot be empty.", "Validation Error");
                return;
            }

            bool isUrl = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url";

            if (isUrl)
            {
                // --- FIX: ROBUST INI WRITING FOR .URL ---
                var lines = System.IO.File.ReadAllLines(shortcutPath).ToList();

                void UpdateIniKey(string key, string value)
                {
                    // Find existing key properly ignoring case and whitespace
                    int idx = lines.FindIndex(l => l.Trim().StartsWith(key + "=", StringComparison.OrdinalIgnoreCase));

                    if (idx != -1)
                    {
                        if (value == null) lines.RemoveAt(idx);
                        else lines[idx] = $"{key}={value}";
                    }
                    else if (value != null)
                    {
                        // Key missing: Find header to insert after
                        int headerIdx = lines.FindIndex(l => l.Trim().Equals("[InternetShortcut]", StringComparison.OrdinalIgnoreCase));
                        if (headerIdx == -1)
                        {
                            // Header missing? Add it at the top (safest fallback)
                            lines.Insert(0, "[InternetShortcut]");
                            headerIdx = 0;
                        }
                        lines.Insert(headerIdx + 1, $"{key}={value}");
                    }
                }

                UpdateIniKey("URL", newTargetPath);

                // Handle Icon
                if (!string.IsNullOrEmpty(newIconPath) && newIconPath != "Default")
                {
                    string iconFile = newIconPath;
                    string iconIndex = "0";
                    if (newIconPath.Contains(","))
                    {
                        var parts = newIconPath.Split(',');
                        iconFile = parts[0];
                        if (parts.Length > 1) iconIndex = parts[1];
                    }
                    UpdateIniKey("IconFile", iconFile);
                    UpdateIniKey("IconIndex", iconIndex);
                }
                else
                {
                    // Reset: Remove the keys entirely
                    UpdateIniKey("IconFile", null);
                    UpdateIniKey("IconIndex", null);
                }

                System.IO.File.WriteAllLines(shortcutPath, lines);
            }
            else
            {
                // Standard .lnk Saving (Keep existing working logic)
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = newTargetPath;
                string newArguments = argumentsBox?.Text?.Trim() ?? "";
                if (!string.IsNullOrEmpty(newArguments)) shortcut.Arguments = newArguments;

                // Handle Working Directory
                try
                {
                    if (System.IO.File.Exists(newTargetPath))
                    {
                        string wd = System.IO.Path.GetDirectoryName(newTargetPath);
                        if (!string.IsNullOrEmpty(wd)) shortcut.WorkingDirectory = wd;
                    }
                }
                catch { }

                // Handle Icon
                if (!string.IsNullOrEmpty(newIconPath) && newIconPath != "Default")
                    shortcut.IconLocation = newIconPath;
                else
                {
                    if (System.IO.File.Exists(newTargetPath)) shortcut.IconLocation = $"{newTargetPath},0";
                    else shortcut.IconLocation = ",0";
                }
                shortcut.Save();
            }

            NewDisplayName = newDisplayName;
            System.Windows.Application.Current.Dispatcher.Invoke(() => {
                FenceManager.RefreshIconClickHandlers(shortcutPath, newDisplayName);
            });

            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error saving shortcut: {ex.Message}");
            MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to save shortcut: {ex.Message}", "Save Error");
        }
    }








    /// <summary>
    /// Loads and displays the icon in the preview Image control
    /// </summary>
    private void LoadIconPreview(string iconPath)
    {
        try
        {
            ImageSource iconSource = null;
            bool isUrl = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url";

            // Handle "Default" case
            if (string.IsNullOrEmpty(iconPath) || iconPath == "Default")
            {
                if (isUrl)
                {
                    // FIX: Show the correct App Theme icon for Web Links
                    iconSource = new BitmapImage(new Uri("pack://application:,,,/Resources/link-White.png"));
                }
                else
                {
                    // Standard .lnk logic (Extract from target)
                    string targetPath = GetCurrentTargetPath();

                    if (!string.IsNullOrEmpty(targetPath) && System.IO.File.Exists(targetPath))
                    {
                        iconSource = IconManager.ExtractIconFromFile(targetPath, 0);
                    }
                    else if (System.IO.Directory.Exists(targetPath))
                    {
                        iconSource = new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                    }
                }
            }
            else if (iconPath.Contains(","))
            {
                // Handle format: "path,index"
                string[] parts = iconPath.Split(',');
                string filePath = parts[0];
                int iconIndex = 0;

                if (parts.Length == 2 && int.TryParse(parts[1], out int parsedIndex))
                    iconIndex = parsedIndex;

                if (System.IO.File.Exists(filePath))
                    iconSource = IconManager.ExtractIconFromFile(filePath, iconIndex);
            }
            else
            {
                // Handle direct file path
                if (System.IO.File.Exists(iconPath))
                {
                    string extension = System.IO.Path.GetExtension(iconPath).ToLower();
                    if (extension == ".ico")
                        iconSource = new BitmapImage(new Uri(iconPath));
                    else if (extension == ".exe" || extension == ".dll")
                        iconSource = IconManager.ExtractIconFromFile(iconPath, 0);
                }
            }

            // Apply to UI
            if (iconSource != null)
            {
                iconPreview.Source = iconSource;
            }
            else
            {
                // Fallback for unknowns
                iconPreview.Source = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
            }
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading icon preview: {ex.Message}");
            iconPreview.Source = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
        }
    }

    /// <summary>
    /// Loads the actual icon being used on the fence (same logic as fence display)
    /// </summary>
    private void LoadActualFenceIcon()
    {
        try
        {
            string targetPath = GetCurrentTargetPath();
            bool isShortcut = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".lnk";
            bool isFolder = !isShortcut && System.IO.Directory.Exists(targetPath);
            bool isLink = System.IO.Path.GetExtension(shortcutPath).ToLower() == ".url";

            // Use the same icon extraction logic as the fence
            ImageSource fenceIcon = IconManager.GetIconForFile(targetPath, shortcutPath, isFolder, isLink, isShortcut, null);

            if (fenceIcon != null)
            {
                iconPreview.Source = fenceIcon;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Loaded actual fence icon for: {shortcutPath}");
            }
            else
            {
                // Fallback to LoadIconPreview method
                LoadIconPreview(iconPathBox.Text);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Using LoadIconPreview fallback");
            }
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Error loading fence icon: {ex.Message}, using fallback");
            // Fallback to the original method
            LoadIconPreview(iconPathBox.Text);
        }
    }


    /// <summary>
    /// Gets the user's selected accent color (same method as CustomizeFenceForm)
    /// </summary>
    private SolidColorBrush GetAccentColorBrush()
    {
        try
        {
            string selectedColorName = SettingsManager.SelectedColor;
            var mediaColor = Utility.GetColorFromName(selectedColorName);
            return new SolidColorBrush(mediaColor);
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error getting accent color: {ex.Message}");
            // Fallback to blue
            return new SolidColorBrush(Color.FromRgb(66, 133, 244));
        }
    }

    protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
    {
        base.OnMouseLeftButtonDown(e);
        if (e.ButtonState == MouseButtonState.Pressed)
        {
            DragMove();
        }
    }
}
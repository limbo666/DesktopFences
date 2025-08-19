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

        Border mainBorder = new Border
        {
            Background = Brushes.White,
            CornerRadius = new CornerRadius(8),
            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
            BorderThickness = new Thickness(1, 1, 1, 1),
            Margin = new Thickness(0, 0, 0, 0)  // Remove margin when using transparency
        };

        Grid rootGrid = new Grid();
        rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
        rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
        rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Footer

        CreateModernHeader(rootGrid);
        CreateModernContent(rootGrid);
        CreateModernFooter(rootGrid);

        mainBorder.Child = rootGrid;
        Content = mainBorder;

        // Initialize with current values
        LoadCurrentValues();
    }

    private void CreateModernHeader(Grid rootGrid)
    {
        Grid headerGrid = new Grid
        {
            Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
            Height = 50
        };
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        TextBlock titleText = new TextBlock
        {
            Text = "Edit Shortcut",
            FontSize = 16,
            FontWeight = FontWeights.SemiBold,
            Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(20, 0, 0, 0)
        };

        Button closeButton = new Button
        {
            Width = 36,
            Height = 36,
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0, 0, 0, 0),
            Margin = new Thickness(0, 0, 10, 0),
            Cursor = Cursors.Hand,
            VerticalAlignment = VerticalAlignment.Center,
            Content = new TextBlock
            {
                Text = "✕",
                FontSize = 16,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            }
        };
        closeButton.Click += (s, e) => { DialogResult = false; Close(); };

        headerGrid.Children.Add(titleText);
        headerGrid.Children.Add(closeButton);
        Grid.SetColumn(closeButton, 1);

        Grid.SetRow(headerGrid, 0);
        rootGrid.Children.Add(headerGrid);
    }

    private void CreateModernContent(Grid rootGrid)
    {
        Border contentBorder = new Border
        {
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
            BorderThickness = new Thickness(0, 1, 0, 0),
            Padding = new Thickness(20, 16, 20, 16),
            CornerRadius = new CornerRadius(0, 0, 8, 8)
        };

        StackPanel buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };

        Button defaultButton = CreateModernButton("Default", false);
        defaultButton.Margin = new Thickness(0, 0, 10, 0);
        defaultButton.Click += Default_Click;

        Button cancelButton = CreateModernButton("Cancel", false);
        cancelButton.Margin = new Thickness(0, 0, 10, 0);
        cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

        _saveButton = CreateModernButton("Save", true);
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
            // Icon preview will be implemented later
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
            TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error opening file browser: {ex.Message}", "Browse Error");
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
                }
            }
            else
            {
                iconPathBox.Text = selectedPath;
            }
        }
    }

    private void Default_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Restoring default values for shortcut: {shortcutPath}");

            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

            string originalTarget = shortcut.TargetPath ?? "";
            string originalName = System.IO.Path.GetFileNameWithoutExtension(shortcutPath) ?? "";

            nameBox.Text = originalName;
            targetPathBox.Text = originalTarget;
            argumentsBox.Text = "";
            iconPathBox.Text = "Default";

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Default values restored successfully");
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error restoring defaults: {ex.Message}");
            TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Error restoring default values: {ex.Message}", "Default Error");
        }
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Starting save process for shortcut: {shortcutPath}");

            string newDisplayName = nameBox?.Text?.Trim() ?? "";
            string newTargetPath = targetPathBox?.Text?.Trim() ?? "";
            string newArguments = argumentsBox?.Text?.Trim() ?? "";
            string newIconPath = iconPathBox?.Text?.Trim() ?? "";

            if (string.IsNullOrWhiteSpace(newDisplayName))
            {
                TrayManager.Instance.ShowOKOnlyMessageBoxForm("Display name cannot be empty.", "Validation Error");
                return;
            }

            if (string.IsNullOrWhiteSpace(newTargetPath))
            {
                TrayManager.Instance.ShowOKOnlyMessageBoxForm("Target path cannot be empty.", "Validation Error");
                return;
            }

            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

            shortcut.TargetPath = newTargetPath;
            shortcut.Arguments = newArguments;

            if (!string.IsNullOrEmpty(newIconPath) && newIconPath != "Default")
            {
                shortcut.IconLocation = newIconPath;
            }
            else
            {
                shortcut.IconLocation = "";
            }

            shortcut.Save();

            NewDisplayName = newDisplayName;

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                FenceManager.RefreshIconClickHandlers(shortcutPath, newDisplayName);
            });

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "Shortcut saved successfully");
            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error saving shortcut: {ex.Message}");
            TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to save shortcut: {ex.Message}", "Save Error");
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
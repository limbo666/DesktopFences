using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Desktop_Fences;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using IWshRuntimeLibrary;
using System.Runtime.InteropServices;

public partial class EditShortcutWindow : Window
{
    private string shortcutPath;
    public string NewDisplayName { get; private set; }
    private TextBox nameBox; // Field to store reference
    private TextBox iconPathBox; // Field to store reference

    private Image iconPreview;

    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern uint ExtractIconEx(string szFileName, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

    [DllImport("user32.dll")]
    private static extern bool DestroyIcon(IntPtr hIcon);


    public EditShortcutWindow(string shortcutPath, string currentDisplayName)
    {
        this.shortcutPath = shortcutPath;
        NewDisplayName = currentDisplayName;
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.Title = "Edit Shortcut";
        this.Width = 400;
        this.Height = 200;
        this.WindowStartupLocation = WindowStartupLocation.CenterScreen;

        // Root Grid
        Grid mainGrid = new Grid();
        mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        // GroupBox
        GroupBox groupBox = new GroupBox
        {
            Header = "Shortcut Properties",
            Margin = new Thickness(10),
        };
        Grid.SetRow(groupBox, 0);

        // Inner Grid
        Grid innerGrid = new Grid();
        innerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        innerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        innerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        innerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        innerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        // Display Name Label
        Label nameLabel = new Label { Content = "Display Name:", Margin = new Thickness(5, 5, 0, 0) };
        Grid.SetRow(nameLabel, 0);
        Grid.SetColumn(nameLabel, 0);

        // Display Name TextBox
        nameBox = new TextBox
        {
            Text = NewDisplayName,
            Margin = new Thickness(5, 5, 5, 0),
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Stretch
        };
        Grid.SetRow(nameBox, 0);
        Grid.SetColumn(nameBox, 1);

        // Icon Label
        Label iconLabel = new Label { Content = "Icon:", Margin = new Thickness(5, 5, 0, 0) };
        Grid.SetRow(iconLabel, 1);
        Grid.SetColumn(iconLabel, 0);

        // Icon StackPanel
        StackPanel iconStack = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetRow(iconStack, 1);
        Grid.SetColumn(iconStack, 1);

        iconPathBox = new TextBox
        {
            Text = GetCurrentIconLocation(),
            IsReadOnly = true,
            Margin = new Thickness(0, 5, 5, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Width = 200
        };

        iconPreview = new Image
        {
            Width = 32,
            Height = 32,
            VerticalAlignment = VerticalAlignment.Center,
            Stretch = Stretch.Uniform
        };

        iconPreview.Source = null;

        var iconContainer = new Border
        {
            Width = 32,
            Height = 32,
            Margin = new Thickness(5, 0, 0, 0),
            Background = Brushes.DarkRed,
            BorderBrush = Brushes.Gray,           // visible outline
            BorderThickness = new Thickness(3),
            Padding = new Thickness(2),       // so the Image inside doesn’t fill the whole thing
            Child = iconPreview
        };

        iconStack.Children.Add(iconPathBox);
        iconStack.Children.Add(iconContainer);

        // Browse Button
        Button browseIconButton = new Button
        {
            Content = "Browse...",
            Margin = new Thickness(0, 5, 5, 0),
            Width = 80,
            Height = 25
        };
        browseIconButton.Click += BrowseIcon_Click;
        Grid.SetRow(browseIconButton, 1);
        Grid.SetColumn(browseIconButton, 2);

        // Add all controls
        innerGrid.Children.Add(nameLabel);
        innerGrid.Children.Add(nameBox);
        innerGrid.Children.Add(iconLabel);
        innerGrid.Children.Add(iconStack);
        innerGrid.Children.Add(browseIconButton);

        // Safe call to icon preview method
        try
        {
            UpdateIconPreview();
        }
        catch (Exception ex)
        {
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Error setting initial icon preview: {ex.Message}");
        }

        groupBox.Content = innerGrid;

        // Bottom Buttons
        StackPanel buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 0, 10, 10)
        };
        Grid.SetRow(buttonPanel, 1);

        Button defaultButton = new Button { Content = "Default", Margin = new Thickness(0, 0, 10, 0), Width = 80, Height = 25 };
        defaultButton.Click += Default_Click;

        Button saveButton = new Button { Content = "Save", Margin = new Thickness(0, 0, 10, 0), Width = 80, Height = 25 };
        saveButton.Click += Save_Click;

        Button cancelButton = new Button { Content = "Cancel", Width = 80, Height = 25 };
        cancelButton.Click += (s, e) => this.DialogResult = false;

        buttonPanel.Children.Add(defaultButton);
        buttonPanel.Children.Add(saveButton);
        buttonPanel.Children.Add(cancelButton);

        // Final layout assembly
        mainGrid.Children.Add(groupBox);
        mainGrid.Children.Add(buttonPanel);

        this.Content = mainGrid;
    }




    private void UpdateIconPreview()
    {
        var path = iconPathBox.Text;
        if (System.IO.File.Exists(path))
            iconPreview.Source = new BitmapImage(new Uri(path));

        //if (iconPathBox != null && iconPreview != null)
        //{
        //    iconPreview.Source = GetIconSource(iconPathBox.Text, shortcutPath);

        //}
    }

    private void Default_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Entering Default_Click for {shortcutPath}");
            if (nameBox == null || iconPathBox == null)
            {
                throw new InvalidOperationException("TextBox controls not initialized.");
            }
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, "TextBox controls confirmed initialized");

            string defaultName = System.IO.Path.GetFileNameWithoutExtension(shortcutPath);
            nameBox.Text = defaultName;
            NewDisplayName = defaultName;
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Set default name to {defaultName}");

            WshShell shell = new WshShell();
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, "WshShell created");
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, "IWshShortcut created");

            string targetPath = shortcut.TargetPath;
            if (string.IsNullOrEmpty(targetPath) || !System.IO.File.Exists(targetPath))
            {
                FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Target path invalid or missing: {targetPath}, resetting without icon change");
                // Optionally handle missing targets differently if needed
            }
            else
            {
                shortcut.IconLocation = targetPath; // Use target's default icon
                FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Set IconLocation to target: {targetPath}");
            }
            shortcut.Save();
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, "Shortcut saved");

            iconPathBox.Text = "Default";
            UpdateIconPreview(); // Update the icon preview
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Restored default icon and name for {shortcutPath}");
        }
        catch (Exception ex)
        {
            FenceManager.Log(FenceManager.LogLevel.Error, FenceManager.LogCategory.General, $"Failed to restore defaults for {shortcutPath}: {ex.Message}\nStackTrace: {ex.StackTrace}");
            //      MessageBox.Show($"Failed to restore defaults: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to restore defaults: {ex.Message}", "Error");
        }
        finally
        {
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, "Exiting Default_Click");
        }
    }

    private ImageSource GetIconSource(string iconLocation, string shortcutPath)
    {
        try
        {
            if (string.IsNullOrEmpty(iconLocation) || iconLocation == "Default")
            {
                string targetPath = GetTargetPath(shortcutPath);
                if (System.IO.Directory.Exists(targetPath))
                {
                    return new BitmapImage(new Uri("pack://application:,,,/Resources/folder-White.png"));
                }
                else if (System.IO.File.Exists(targetPath))
                {
                    try
                    {
                        return System.Drawing.Icon.ExtractAssociatedIcon(targetPath).ToImageSource();
                    }
                    catch
                    {
                        return new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));
                    }
                }
                return null;
            }

            string[] parts = iconLocation.Split(',');
            string iconPath = parts[0];
            int iconIndex = 0;

            if (parts.Length > 1 && int.TryParse(parts[1], out int index))
            {
                iconIndex = index;
            }

            if (!System.IO.File.Exists(iconPath))
            {
                return null;
            }

            if (iconPath.ToLower().EndsWith(".ico"))
            {
                // Load .ico file directly
                using (var stream = new System.IO.FileStream(iconPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.StreamSource = stream;
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    bitmap.Freeze(); // Ensures it's usable across threads
                    return bitmap;
                }
            }
            else
            {
                // Use ExtractIconEx for .exe/.dll icons
                IntPtr[] hIcon = new IntPtr[1];
                uint result = ExtractIconEx(iconPath, iconIndex, hIcon, null, 1);
                if (result > 0 && hIcon[0] != IntPtr.Zero)
                {
                    try
                    {
                        return Imaging.CreateBitmapSourceFromHIcon(
                            hIcon[0],
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions()
                        );
                    }
                    finally
                    {
                        DestroyIcon(hIcon[0]);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Error extracting icon from {iconLocation}: {ex.Message}");
        }

        return null;
    }

    private string GetTargetPath(string shortcutPath)
    {
        WshShell shell = new WshShell();
        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
        return shortcut.TargetPath;
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
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Error getting current icon location: {ex.Message}");
            return "Default";
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
                    FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"User selected icon index: {picker.SelectedIndex} from {selectedPath}");
                    iconPathBox.Text = $"{selectedPath},{picker.SelectedIndex}";
                    UpdateIconPreview(); // Update the icon preview
                }
            }
            else
            {
                iconPathBox.Text = selectedPath;
                UpdateIconPreview(); // Update the icon preview



            }
        }
    }

    private ImageSource ExtractIconForPreview(string iconPath, int iconIndex)
    {
        if (!System.IO.File.Exists(iconPath))
        {
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Icon file not found for preview: {iconPath}");
            return null;
        }

        try
        {
            IntPtr[] hIcon = new IntPtr[1];
            uint result = ExtractIconEx(iconPath, iconIndex, hIcon, null, 1);
            if (result > 0 && hIcon[0] != IntPtr.Zero)
            {
                try
                {
                    ImageSource source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        hIcon[0],
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                    return source;
                }
                finally
                {
                    DestroyIcon(hIcon[0]);
                }
            }
            else
            {
                FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"ExtractIconEx failed for {iconPath} at index {iconIndex}. Result: {result}");
                return null;
            }
        }
        catch (Exception ex)
        {
            FenceManager.Log(FenceManager.LogLevel.Error, FenceManager.LogCategory.General, $"Error extracting icon for preview from {iconPath} at index {iconIndex}: {ex.Message}");
            return null;
        }
    }


    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, "Save_Click started");

            NewDisplayName = string.IsNullOrWhiteSpace(nameBox.Text)
                ? System.IO.Path.GetFileNameWithoutExtension(shortcutPath)
                : nameBox.Text;

            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"New display name: {NewDisplayName}");

            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

            string iconPath = iconPathBox?.Text;
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Icon path from textbox: {iconPath}");

            if (!string.IsNullOrEmpty(iconPath) && iconPath != "Default")
            {
                // Don't modify the iconPath value - use it exactly as stored
                // It should already be in the correct format (path,index)
                shortcut.IconLocation = iconPath;
                FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, $"Set shortcut.IconLocation to: {shortcut.IconLocation}");
            }

            shortcut.Save();
            FenceManager.Log(FenceManager.LogLevel.Debug, FenceManager.LogCategory.General, "Shortcut saved successfully");

            this.DialogResult = true;
        }
        catch (Exception ex)
        {
            FenceManager.Log(FenceManager.LogLevel.Error, FenceManager.LogCategory.General, $"Failed to save shortcut: {ex.Message}\nStack trace: {ex.StackTrace}");
            //  MessageBox.Show($"Failed to save shortcut: {ex.Message}", "Error",
            //                  MessageBoxButton.OK, MessageBoxImage.Error);
            TrayManager.Instance.ShowOKOnlyMessageBoxForm($"Failed to save shortcut: {ex.Message}", "Error");
        }
    }

    //private void Log(string message)
    //{
    //    bool isLogEnabled = true; // Adjust based on your _options
    //    if (isLogEnabled)
    //    {
    //        string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), "desktop_fences.log");
    //        System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
    //    }
    //}
}
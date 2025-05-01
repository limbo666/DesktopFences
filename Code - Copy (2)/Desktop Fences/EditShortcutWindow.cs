//using IWshRuntimeLibrary;
//using System.Windows.Controls;
//using System.Windows;
//using System;
//using System.Linq;

//public partial class EditShortcutWindow : Window
//{
//    private string shortcutPath;
//    public string NewDisplayName { get; private set; }

//    public EditShortcutWindow(string shortcutPath, string currentDisplayName)
//    {
//        this.shortcutPath = shortcutPath;
//        NewDisplayName = currentDisplayName;
//        InitializeComponent();
//    }

//    private void InitializeComponent()
//    {
//        this.Title = "Edit Shortcut";
//        this.Width = 400;
//        this.Height = 200; // Reduced from 300 to 2/3 (200)
//        this.WindowStartupLocation = WindowStartupLocation.CenterScreen;

//        Grid mainGrid = new Grid();
//        mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
//        mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

//        // GroupBox for controls
//        GroupBox groupBox = new GroupBox
//        {
//            Header = "Shortcut Properties",
//            Margin = new Thickness(10),
//        };
//        Grid.SetRow(groupBox, 0);

//        Grid innerGrid = new Grid();
//        innerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
//        innerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
//        innerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
//        innerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

//        // Display Name
//        Label nameLabel = new Label { Content = "Display Name:", Margin = new Thickness(5, 5, 0, 0) };
//        Grid.SetRow(nameLabel, 0);
//        Grid.SetColumn(nameLabel, 0);

//        TextBox nameBox = new TextBox
//        {
//            Text = NewDisplayName,
//            Margin = new Thickness(100, 5, 5, 0),
//            VerticalAlignment = VerticalAlignment.Center
//        };
//        Grid.SetRow(nameBox, 0);
//        Grid.SetColumn(nameBox, 0);
//        Grid.SetColumnSpan(nameBox, 2);

//        // Icon Selection
//        Label iconLabel = new Label { Content = "Icon:", Margin = new Thickness(5, 5, 0, 0) };
//        Grid.SetRow(iconLabel, 1);
//        Grid.SetColumn(iconLabel, 0);

//        TextBox iconPathBox = new TextBox
//        {
//            Text = GetCurrentIconLocation(),
//            IsReadOnly = true,
//            Margin = new Thickness(100, 5, 5, 0),
//            VerticalAlignment = VerticalAlignment.Center
//        };
//        Grid.SetRow(iconPathBox, 1);
//        Grid.SetColumn(iconPathBox, 0);

//        Button browseIconButton = new Button
//        {
//            Content = "Browse...",
//            Margin = new Thickness(0, 5, 5, 0),
//            Width = 80,
//            Height = 25
//        };
//        browseIconButton.Click += BrowseIcon_Click;
//        Grid.SetRow(browseIconButton, 1);
//        Grid.SetColumn(browseIconButton, 1);

//        innerGrid.Children.Add(nameLabel);
//        innerGrid.Children.Add(nameBox);
//        innerGrid.Children.Add(iconLabel);
//        innerGrid.Children.Add(iconPathBox);
//        innerGrid.Children.Add(browseIconButton);
//        groupBox.Content = innerGrid;

//        // Buttons
//        StackPanel buttonPanel = new StackPanel
//        {
//            Orientation = Orientation.Horizontal,
//            HorizontalAlignment = HorizontalAlignment.Right,
//            Margin = new Thickness(0, 0, 10, 10)
//        };
//        Grid.SetRow(buttonPanel, 1);

//        Button saveButton = new Button
//        {
//            Content = "Save",
//            Margin = new Thickness(0, 0, 10, 0),
//            Width = 80,
//            Height = 25
//        };
//        saveButton.Click += Save_Click;

//        Button cancelButton = new Button
//        {
//            Content = "Cancel",
//            Width = 80,
//            Height = 25
//        };
//        cancelButton.Click += (s, e) => this.DialogResult = false;

//        buttonPanel.Children.Add(saveButton);
//        buttonPanel.Children.Add(cancelButton);

//        mainGrid.Children.Add(groupBox);
//        mainGrid.Children.Add(buttonPanel);

//        this.Content = mainGrid;
//    }

//    private string GetCurrentIconLocation()
//    {
//        WshShell shell = new WshShell();
//        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
//        return string.IsNullOrEmpty(shortcut.IconLocation) ? "Default" : shortcut.IconLocation.Split(',')[0];
//    }

//    private void BrowseIcon_Click(object sender, RoutedEventArgs e)
//    {
//        var dialog = new System.Windows.Forms.OpenFileDialog
//        {
//            Filter = "Icon Files (*.ico)|*.ico|Executable Files (*.exe)|*.exe|All Files (*.*)|*.*",
//            Title = "Select an Icon"
//        };
//        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
//        {
//            var iconPathBox = ((this.Content as Grid)?.Children.OfType<GroupBox>().FirstOrDefault()?.Content as Grid)?
//                .Children.OfType<TextBox>().FirstOrDefault(t => Grid.GetRow(t) == 1);
//            if (iconPathBox != null)
//            {
//                iconPathBox.Text = dialog.FileName;
//            }
//        }
//    }

//    private void Save_Click(object sender, RoutedEventArgs e)
//    {
//        try
//        {
//            var nameBox = ((this.Content as Grid)?.Children.OfType<GroupBox>().FirstOrDefault()?.Content as Grid)?
//                .Children.OfType<TextBox>().FirstOrDefault(t => Grid.GetRow(t) == 0);
//            var iconPathBox = ((this.Content as Grid)?.Children.OfType<GroupBox>().FirstOrDefault()?.Content as Grid)?
//                .Children.OfType<TextBox>().FirstOrDefault(t => Grid.GetRow(t) == 1);

//            if (nameBox == null)
//            {
//                throw new InvalidOperationException("Display Name TextBox not found.");
//            }

//            NewDisplayName = string.IsNullOrWhiteSpace(nameBox.Text) ? System.IO.Path.GetFileNameWithoutExtension(shortcutPath) : nameBox.Text;

//            WshShell shell = new WshShell();
//            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

//            string iconPath = iconPathBox?.Text;
//            if (!string.IsNullOrEmpty(iconPath) && iconPath != "Default" && System.IO.File.Exists(iconPath))
//            {
//                shortcut.IconLocation = iconPath;
//                Log($"Set custom icon for {shortcutPath} to {iconPath}");
//            }
//            else
//            {
//                Log($"No new icon selected for {shortcutPath}, keeping existing: {shortcut.IconLocation}");
//            }

//            shortcut.Save();
//            this.DialogResult = true;
//            Log($"Saved changes to shortcut {shortcutPath}");
//        }
//        catch (Exception ex)
//        {
//            Log($"Failed to save shortcut {shortcutPath}: {ex.Message}");
//            MessageBox.Show($"Failed to save changes: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
//        }
//    }

//    private void Log(string message)
//    {
//        bool isLogEnabled = true; // Adjust based on your _options
//        if (isLogEnabled)
//        {
//            string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), "desktop_fences.log");
//            System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
//        }
//    }
//}
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using IWshRuntimeLibrary;

public partial class EditShortcutWindow : Window
{
    private string shortcutPath;
    public string NewDisplayName { get; private set; }
    private TextBox nameBox; // Field to store reference
    private TextBox iconPathBox; // Field to store reference

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

        Grid mainGrid = new Grid();
        mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        GroupBox groupBox = new GroupBox
        {
            Header = "Shortcut Properties",
            Margin = new Thickness(10),
        };
        Grid.SetRow(groupBox, 0);

        Grid innerGrid = new Grid();
        innerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        innerGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        innerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        innerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        Label nameLabel = new Label { Content = "Display Name:", Margin = new Thickness(5, 5, 0, 0) };
        Grid.SetRow(nameLabel, 0);
        Grid.SetColumn(nameLabel, 0);

        nameBox = new TextBox // Assign to field
        {
            Text = NewDisplayName,
            Margin = new Thickness(100, 5, 5, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetRow(nameBox, 0);
        Grid.SetColumn(nameBox, 0);
        Grid.SetColumnSpan(nameBox, 2);

        Label iconLabel = new Label { Content = "Icon:", Margin = new Thickness(5, 5, 0, 0) };
        Grid.SetRow(iconLabel, 1);
        Grid.SetColumn(iconLabel, 0);

        iconPathBox = new TextBox // Assign to field
        {
            Text = GetCurrentIconLocation(),
            IsReadOnly = true,
            Margin = new Thickness(100, 5, 5, 0),
            VerticalAlignment = VerticalAlignment.Center
        };
        Grid.SetRow(iconPathBox, 1);
        Grid.SetColumn(iconPathBox, 0);

        Button browseIconButton = new Button
        {
            Content = "Browse...",
            Margin = new Thickness(0, 5, 5, 0),
            Width = 80,
            Height = 25
        };
        browseIconButton.Click += BrowseIcon_Click;
        Grid.SetRow(browseIconButton, 1);
        Grid.SetColumn(browseIconButton, 1);

        innerGrid.Children.Add(nameLabel);
        innerGrid.Children.Add(nameBox);
        innerGrid.Children.Add(iconLabel);
        innerGrid.Children.Add(iconPathBox);
        innerGrid.Children.Add(browseIconButton);
        groupBox.Content = innerGrid;

        StackPanel buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 0, 10, 10)
        };
        Grid.SetRow(buttonPanel, 1);

        Button defaultButton = new Button
        {
            Content = "Default",
            Margin = new Thickness(0, 0, 10, 0),
            Width = 80,
            Height = 25
        };
        defaultButton.Click += Default_Click;

        Button saveButton = new Button
        {
            Content = "Save",
            Margin = new Thickness(0, 0, 10, 0),
            Width = 80,
            Height = 25
        };
        saveButton.Click += Save_Click;

        Button cancelButton = new Button
        {
            Content = "Cancel",
            Width = 80,
            Height = 25
        };
        cancelButton.Click += (s, e) => this.DialogResult = false;

        buttonPanel.Children.Add(defaultButton);
        buttonPanel.Children.Add(saveButton);
        buttonPanel.Children.Add(cancelButton);

        mainGrid.Children.Add(groupBox);
        mainGrid.Children.Add(buttonPanel);

        this.Content = mainGrid;
    }

    private void Default_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Log($"Entering Default_Click for {shortcutPath}");
            if (nameBox == null || iconPathBox == null)
            {
                throw new InvalidOperationException("TextBox controls not initialized.");
            }
            Log("TextBox controls confirmed initialized");

            string defaultName = System.IO.Path.GetFileNameWithoutExtension(shortcutPath);
            nameBox.Text = defaultName;
            NewDisplayName = defaultName;
            Log($"Set default name to {defaultName}");

            WshShell shell = new WshShell();
            Log("WshShell created");
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            Log("IWshShortcut created");

            string targetPath = shortcut.TargetPath;
            if (string.IsNullOrEmpty(targetPath) || !System.IO.File.Exists(targetPath))
            {
                Log($"Target path invalid or missing: {targetPath}, resetting without icon change");
                // Optionally handle missing targets differently if needed
            }
            else
            {
                shortcut.IconLocation = targetPath; // Use target's default icon
                Log($"Set IconLocation to target: {targetPath}");
            }
            shortcut.Save();
            Log("Shortcut saved");

            iconPathBox.Text = "Default";
            Log($"Restored default icon and name for {shortcutPath}");
        }
        catch (Exception ex)
        {
            Log($"Failed to restore defaults for {shortcutPath}: {ex.Message}\nStackTrace: {ex.StackTrace}");
            MessageBox.Show($"Failed to restore defaults: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    //private void Default_Click(object sender, RoutedEventArgs e)
    //{
    //    Log($"nameBox is {(nameBox == null ? "null" : "not null")}, iconPathBox is {(iconPathBox == null ? "null" : "not null")}");
    //    try
    //    {
    //        if (nameBox == null || iconPathBox == null)
    //        {
    //            throw new InvalidOperationException("TextBox controls not initialized.");
    //        }

    //        string defaultName = System.IO.Path.GetFileNameWithoutExtension(shortcutPath);
    //        nameBox.Text = defaultName;
    //        NewDisplayName = defaultName;

    //        WshShell shell = new WshShell();
    //        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
    //        shortcut.IconLocation = ""; // Clear custom icon
    //        shortcut.Save();

    //        iconPathBox.Text = "Default";
    //        Log($"Restored default icon and name for {shortcutPath}");
    //    }
    //    catch (Exception ex)
    //    {
    //        Log($"Failed to restore defaults for {shortcutPath}: {ex.Message}");
    //        MessageBox.Show($"Failed to restore defaults: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
    //    }
    //}

    private string GetCurrentIconLocation()
    {
        WshShell shell = new WshShell();
        IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
        return string.IsNullOrEmpty(shortcut.IconLocation) ? "Default" : shortcut.IconLocation.Split(',')[0];
    }

    private void BrowseIcon_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.OpenFileDialog
        {
            Filter = "Icon Files (*.ico)|*.ico|Executable Files (*.exe)|*.exe|All Files (*.*)|*.*",
            Title = "Select an Icon"
        };
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            if (iconPathBox != null)
            {
                iconPathBox.Text = dialog.FileName;
            }
        }
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (nameBox == null)
            {
                throw new InvalidOperationException("Display Name TextBox not found.");
            }

            NewDisplayName = string.IsNullOrWhiteSpace(nameBox.Text) ? System.IO.Path.GetFileNameWithoutExtension(shortcutPath) : nameBox.Text;

            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

            string iconPath = iconPathBox?.Text;
            if (!string.IsNullOrEmpty(iconPath) && iconPath != "Default" && System.IO.File.Exists(iconPath))
            {
                shortcut.IconLocation = iconPath;
                Log($"Set custom icon for {shortcutPath} to {iconPath}");
            }
            else
            {
                Log($"No new icon selected or reset to default for {shortcutPath}, keeping existing: {shortcut.IconLocation}");
            }

            shortcut.Save();
            this.DialogResult = true;
            Log($"Saved changes to shortcut {shortcutPath}");
        }
        catch (Exception ex)
        {
            Log($"Failed to save shortcut {shortcutPath}: {ex.Message}");
            MessageBox.Show($"Failed to save changes: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Log(string message)
    {
        bool isLogEnabled = true; // Adjust based on your _options
        if (isLogEnabled)
        {
            string logPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location), "desktop_fences.log");
            System.IO.File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Shapes;

namespace Desktop_Fences
{
    public class ProfileManagerForm : Window
    {
        private StackPanel _listPanel;

        public ProfileManagerForm()
        {
            // Window Setup
            Title = "Profile Manager";
            Width = 480;
            Height = 720; // Height increased to prevent footer cutoff
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;

            // Main Container (Card) - Updated to match CustomizeFenceForm style
            Border mainBorder = new Border
            {
                Background = Brushes.White,
                // Added distinct border definition
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                // Squared corners to match reference
                CornerRadius = new CornerRadius(0),
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

            Grid rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 0: Header
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 1: List
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 2: Add New
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 3: Footer

            // --- 1. HEADER ---
            Border header = new Border
            {
                Background = GetAccentBrush(),
                // Squared corners
                CornerRadius = new CornerRadius(0),
                Padding = new Thickness(15),
                Height = 50 // Fixed height for consistency
            };

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            StackPanel titleStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            titleStack.Children.Add(new TextBlock { Text = "Profile Manager", FontSize = 18, FontWeight = FontWeights.Bold, Foreground = Brushes.White });
            // Subtitle removed or kept small to fit clean header style

            Button closeBtn = new Button { Content = "✕", Background = Brushes.Transparent, BorderThickness = new Thickness(0), Foreground = Brushes.White, FontSize = 16, Cursor = System.Windows.Input.Cursors.Hand, VerticalAlignment = VerticalAlignment.Center };
            closeBtn.Click += (s, e) => Close();

            headerGrid.Children.Add(titleStack);
            headerGrid.Children.Add(closeBtn); Grid.SetColumn(closeBtn, 1);
            header.Child = headerGrid;
            header.MouseLeftButtonDown += (s, e) => DragMove();

            // --- 2. LIST AREA ---
            ScrollViewer scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Margin = new Thickness(0, 10, 0, 10) };
            _listPanel = new StackPanel { Margin = new Thickness(16, 0, 16, 0) };
            scroll.Content = _listPanel;

            // --- 3. ADD NEW SECTION ---
            Border addBorder = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(16),
                Background = new SolidColorBrush(Color.FromRgb(250, 250, 250))
            };
            Grid addGrid = new Grid();
            addGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            addGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            TextBox txtNewName = new TextBox { Height = 34, VerticalContentAlignment = VerticalAlignment.Center, Padding = new Thickness(5, 0, 5, 0), Tag = "New Profile Name..." };

            Button btnAdd = new Button
            {
                Content = "Create",
                Height = 34,
                Width = 100,
                Margin = new Thickness(10, 0, 0, 0),
                Background = GetAccentBrush(),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                FontWeight = FontWeights.Bold,
                Cursor = System.Windows.Input.Cursors.Hand
            };
            btnAdd.Click += (s, e) =>
            {
                if (!string.IsNullOrWhiteSpace(txtNewName.Text))
                {
                    if (ProfileManager.CreateProfile(txtNewName.Text))
                    {
                        txtNewName.Text = "";
                        RefreshList();
                        TrayManager.Instance?.UpdateProfilesMenu();
                    }
                    else MessageBoxesManager.ShowOKOnlyMessageBoxForm("Invalid name or profile already exists.", "Error");
                }
            };

            addGrid.Children.Add(txtNewName);
            addGrid.Children.Add(btnAdd); Grid.SetColumn(btnAdd, 1);
            addBorder.Child = addGrid;

            // --- 4. FOOTER ---
            Border footer = new Border
            {
                Padding = new Thickness(16),
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                // Squared corners
                CornerRadius = new CornerRadius(0),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0)
            };
            Button btnClose = new Button
            {
                Content = "Close",
                Width = 100,
                Height = 34,
                HorizontalAlignment = HorizontalAlignment.Right,
                Cursor = System.Windows.Input.Cursors.Hand,
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1)
            };
            btnClose.Click += (s, e) => Close();
            footer.Child = btnClose;

            // Assembly
            rootGrid.Children.Add(header);
            rootGrid.Children.Add(scroll); Grid.SetRow(scroll, 1);
            rootGrid.Children.Add(addBorder); Grid.SetRow(addBorder, 2);
            rootGrid.Children.Add(footer); Grid.SetRow(footer, 3);

            mainBorder.Child = rootGrid;
            Content = mainBorder;

            RefreshList();
        }

        private void RefreshList()
        {
            _listPanel.Children.Clear();
            var profiles = ProfileManager.GetProfiles();
            string current = ProfileManager.CurrentProfileName;

            for (int i = 0; i < profiles.Count; i++)
            {
                var p = profiles[i];
                bool isActive = p.Name.Equals(current, StringComparison.OrdinalIgnoreCase);

                Border card = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(0),
                    Margin = new Thickness(0, 0, 0, 8),
                    Padding = new Thickness(10),
                    // Show Hand cursor if actionable (inactive), otherwise Arrow
                    Cursor = isActive ? System.Windows.Input.Cursors.Arrow : System.Windows.Input.Cursors.Hand
                };

                // CLICK TO SWITCH EVENT
                if (!isActive)
                {
                    card.MouseLeftButtonUp += (s, e) =>
                    {
                        // Sync automation home to this manual choice
                        ProfileManager.SetManualBaseProfile(p.Name);
                        ProfileManager.SwitchToProfile(p.Name);

                        // Refresh UI elements
                        RefreshList();
                        TrayManager.Instance?.UpdateProfilesMenu();
                        TrayManager.Instance?.UpdateTrayIcon();
                    };
                }

                Grid row = new Grid();
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) }); // ID
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Name
                row.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Controls

                // ID
                row.Children.Add(new TextBlock { Text = p.Id.ToString(), FontWeight = FontWeights.Bold, Foreground = Brushes.Gray, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center });

                // Name
                StackPanel nameStack = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                TextBlock txtName = new TextBlock { Text = p.Name, FontWeight = FontWeights.SemiBold, FontSize = 14 };
                nameStack.Children.Add(txtName);

                if (isActive)
                {
                    Border badge = new Border { Background = new SolidColorBrush(Color.FromRgb(220, 255, 220)), CornerRadius = new CornerRadius(0), Margin = new Thickness(10, 0, 0, 0), Padding = new Thickness(5, 1, 5, 1) };
                    badge.Child = new TextBlock { Text = "Active", FontSize = 10, Foreground = Brushes.Green };
                    nameStack.Children.Add(badge);
                }
                Grid.SetColumn(nameStack, 1);
                row.Children.Add(nameStack);

                // Controls
                StackPanel controls = new StackPanel { Orientation = Orientation.Horizontal };

                // Reorder Buttons
                if (i > 0)
                {
                    string neighborName = profiles[i - 1].Name;
                    Button btnUp = CreateIconButton("▲", "Move Up");
                    btnUp.Click += (s, e) =>
                    {
                        e.Handled = true; // Prevent card click
                        ProfileManager.SwapProfileIds(p.Name, neighborName);
                        RefreshList();
                    };
                    controls.Children.Add(btnUp);
                }

                if (i < profiles.Count - 1)
                {
                    string neighborName = profiles[i + 1].Name;
                    Button btnDown = CreateIconButton("▼", "Move Down");
                    btnDown.Click += (s, e) =>
                    {
                        e.Handled = true; // Prevent card click
                        ProfileManager.SwapProfileIds(p.Name, neighborName);
                        RefreshList();
                    };
                    controls.Children.Add(btnDown);
                }

                // Rename
                Button btnRename = CreateIconButton("✎", "Rename");
                btnRename.IsEnabled = !isActive;

                if (btnRename.IsEnabled)
                {
                    btnRename.Click += (s, e) =>
                    {
                        e.Handled = true; // Prevent card click
                        string newName = Microsoft.VisualBasic.Interaction.InputBox("Rename profile:", "Rename", p.Name);
                        if (!string.IsNullOrWhiteSpace(newName) && newName != p.Name)
                        {
                            if (ProfileManager.RenameProfile(p.Name, newName))
                            {
                                RefreshList();
                                TrayManager.Instance?.UpdateProfilesMenu();
                            }
                            else MessageBoxesManager.ShowOKOnlyMessageBoxForm("Rename failed. Name may be in use or invalid.", "Error");
                        }
                    };
                }
                else btnRename.Opacity = 0.3;
                controls.Children.Add(btnRename);

                // Delete
                Button btnDelete = CreateIconButton("🗑", "Delete");
                btnDelete.IsEnabled = !isActive && profiles.Count > 1;

                if (btnDelete.IsEnabled)
                {
                    btnDelete.Foreground = Brushes.Red;
                    btnDelete.Click += (s, e) =>
                    {
                        e.Handled = true; // Prevent card click
                        if (MessageBoxesManager.ShowCustomYesNoMessageBox($"Are you sure you want to delete profile '{p.Name}'?\nThis cannot be undone.", "Delete Profile"))
                        {
                            if (ProfileManager.DeleteProfile(p.Name))
                            {
                                RefreshList();
                                TrayManager.Instance?.UpdateProfilesMenu();
                            }
                        }
                    };
                }
                else btnDelete.Opacity = 0.3;
                controls.Children.Add(btnDelete);

                Grid.SetColumn(controls, 2);
                row.Children.Add(controls);

                card.Child = row;
                _listPanel.Children.Add(card);
            }
        }
        private Button CreateIconButton(string icon, string tooltip)
        {
            return new Button
            {
                Content = icon,
                Width = 30,
                Height = 30,
                Margin = new Thickness(2, 0, 2, 0),
                Background = Brushes.Transparent,
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                ToolTip = tooltip,
                Cursor = System.Windows.Input.Cursors.Hand
            };
        }

        private SolidColorBrush GetAccentBrush()
        {
            try { return new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor)); }
            catch { return new SolidColorBrush(Color.FromRgb(66, 133, 244)); }
        }
    }
}
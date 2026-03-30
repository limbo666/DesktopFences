using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;

namespace Desktop_Fences
{
    public class AutomationRulesForm : Window
    {
        private ListBox _rulesList;
        private ComboBox _profileCombo;
        private ComboBox _processSearchCombo;
        private CheckBox _persistedChk;
        private Slider _delaySlider;
        private TextBlock _delayValText;

        // Colors
        private Color _colorPurple = Color.FromRgb(128, 0, 128);
        private Color _colorGreen = Color.FromRgb(34, 139, 34);
        private Color _colorRed = Color.FromRgb(220, 53, 69);
        private Color _userAccentColor;

        // Win32 Imports
        [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(System.Drawing.Point p);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);
        private const uint GA_ROOT = 2;

        public AutomationRulesForm()
        {
            InitializeComponent();
            LoadRules();
            RefreshProcessList();
        }

        private void InitializeComponent()
        {
            try
            {
                _userAccentColor = Utility.GetColorFromName(SettingsManager.SelectedColor);

                this.Title = "Automation Rules";
                this.Width = 550;
                this.Height = 750;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;

                // Main Card
                Border mainCard = new Border
                {
                    Background = Brushes.White,
                    BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                    BorderThickness = new Thickness(1),
                    Margin = new Thickness(8),
                    Effect = new DropShadowEffect { Color = Colors.Black, Direction = 270, ShadowDepth = 2, BlurRadius = 10, Opacity = 0.1 }
                };

                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) }); // Header
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(60) }); // Footer

                CreateHeader(mainGrid);
                CreateContent(mainGrid);
                CreateFooter(mainGrid);

                mainCard.Child = mainGrid;
                this.Content = mainCard;

                PositionFormOnMouseScreen();
                this.KeyDown += (s, e) => { if (e.Key == Key.Escape) Close(); };
            }
            catch (Exception ex)
            {
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error initializing: {ex.Message}", "Error");
            }
        }

        private void CreateHeader(Grid parent)
        {
            Border headerBorder = new Border { Background = new SolidColorBrush(_userAccentColor), Height = 50 };
            Grid.SetRow(headerBorder, 0);

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });

            StackPanel titleStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(16, 0, 0, 0) };
            titleStack.Children.Add(new TextBlock { Text = "Manage Automation Rules", FontFamily = new FontFamily("Segoe UI"), FontSize = 18, FontWeight = FontWeights.Bold, Foreground = Brushes.White });

            Button closeButton = new Button { Content = "✕", Width = 32, Height = 32, FontSize = 16, Foreground = Brushes.White, Background = Brushes.Transparent, BorderThickness = new Thickness(0), Cursor = Cursors.Hand };
            closeButton.Click += (s, e) => Close();

            Grid.SetColumn(titleStack, 0); headerGrid.Children.Add(titleStack);
            Grid.SetColumn(closeButton, 1); headerGrid.Children.Add(closeButton);
            headerBorder.Child = headerGrid;
            headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) DragMove(); };

            parent.Children.Add(headerBorder);
        }

        private void CreateContent(Grid parent)
        {
            ScrollViewer scroll = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, Padding = new Thickness(16) };
            Grid.SetRow(scroll, 1);

            StackPanel contentStack = new StackPanel();

            // --- 1. EXISTING RULES ---
            GroupBox listGroup = new GroupBox
            {
                Header = "Existing Rules",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(_userAccentColor),
                Margin = new Thickness(0, 0, 0, 20),
                Padding = new Thickness(8)
            };

            StackPanel listPanel = new StackPanel();
            _rulesList = new ListBox { Height = 150, Margin = new Thickness(0, 0, 0, 10), BorderBrush = new SolidColorBrush(Color.FromRgb(200, 200, 200)) };
            _rulesList.MouseDoubleClick += (s, e) => LoadSelectedRuleForEdit();

            // RED Delete Button (Width 120)
            Button btnDelete = new Button
            {
                Content = "Delete Selected",
                Width = 120,
                Height = 34,
                HorizontalAlignment = HorizontalAlignment.Right,
                Background = new SolidColorBrush(_colorRed),
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btnDelete.Click += (s, e) => DeleteRule();

            listPanel.Children.Add(new TextBlock { Text = "Double-click to edit rule", FontSize = 11, Foreground = Brushes.Gray, Margin = new Thickness(0, 0, 0, 5) });
            listPanel.Children.Add(_rulesList);
            listPanel.Children.Add(btnDelete);
            listGroup.Content = listPanel;
            contentStack.Children.Add(listGroup);

            // --- 2. RULE DEFINITION ---
            GroupBox ruleGroup = new GroupBox
            {
                Header = "Rule Definition",
                FontFamily = new FontFamily("Segoe UI"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(_userAccentColor),
                Padding = new Thickness(8)
            };

            StackPanel ruleStack = new StackPanel();

            // Process Name + Pick Button
            CreateLabel(ruleStack, "Process Name (Pick or Type):");
            Grid procGrid = new Grid();
            procGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            procGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            _processSearchCombo = new ComboBox { IsEditable = true, Height = 34, Margin = new Thickness(0, 0, 10, 10), VerticalContentAlignment = VerticalAlignment.Center };

            // PURPLE Pick Button (Width 120 to match Delete)
            Button btnPick = new Button
            {
                Content = "🎯 Pick Window",
                Width = 120,
                Height = 34,
                Margin = new Thickness(0, 0, 0, 10),
                Background = new SolidColorBrush(_colorPurple),
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btnPick.Click += async (s, e) => await PickWindowInteractively();

            procGrid.Children.Add(_processSearchCombo);
            procGrid.Children.Add(btnPick); Grid.SetColumn(btnPick, 1);
            ruleStack.Children.Add(procGrid);

            // Target Profile
            CreateLabel(ruleStack, "Target Profile:");
            _profileCombo = new ComboBox { Height = 34, Margin = new Thickness(0, 0, 0, 10), VerticalContentAlignment = VerticalAlignment.Center };
            foreach (var p in ProfileManager.GetProfiles()) _profileCombo.Items.Add(p.Name);
            ruleStack.Children.Add(_profileCombo);

            // Delay Slider
            CreateLabel(ruleStack, "Activation Delay:");
            Grid delayGrid = new Grid { Margin = new Thickness(0, 0, 0, 10) };
            delayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            delayGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            _delaySlider = new Slider { Minimum = 0, Maximum = 10, TickFrequency = 1, IsSnapToTickEnabled = true, Value = 1 };
            _delayValText = new TextBlock { Text = "1s", FontWeight = FontWeights.Bold, Margin = new Thickness(10, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center };
            _delaySlider.ValueChanged += (s, e) => { if (_delayValText != null) _delayValText.Text = $"{(int)e.NewValue}s"; };

            delayGrid.Children.Add(_delaySlider);
            delayGrid.Children.Add(_delayValText); Grid.SetColumn(_delayValText, 1);
            ruleStack.Children.Add(delayGrid);

            // Persistence
            _persistedChk = new CheckBox { Content = "Persisted Mode (Stay on profile after close)", Margin = new Thickness(0, 5, 0, 0) };
            ruleStack.Children.Add(_persistedChk);

            ruleGroup.Content = ruleStack;
            contentStack.Children.Add(ruleGroup);

            scroll.Content = contentStack;
            parent.Children.Add(scroll);
        }

        private void CreateFooter(Grid parent)
        {
            Border footerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(16, 8, 16, 8)
            };
            Grid.SetRow(footerBorder, 2);

            StackPanel buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            // Close
            Button btnCancel = new Button
            {
                Content = "Close",
                Width = 100,
                Height = 34,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Margin = new Thickness(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnCancel.Click += (s, e) => Close();

            // GREEN Apply
            Button btnApply = new Button
            {
                Content = "Apply",
                Width = 100,
                Height = 34,
                Background = new SolidColorBrush(_colorGreen),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 10, 0),
                Cursor = Cursors.Hand
            };
            btnApply.Click += (s, e) => ApplyRule(closeAfter: false);

            // THEME Save
            Button btnSave = new Button
            {
                Content = "Save",
                Width = 100,
                Height = 34,
                Background = new SolidColorBrush(_userAccentColor),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                FontWeight = FontWeights.Bold,
                Cursor = Cursors.Hand
            };
            btnSave.Click += (s, e) => ApplyRule(closeAfter: true);

            buttonPanel.Children.Add(btnCancel);
            buttonPanel.Children.Add(btnApply);
            buttonPanel.Children.Add(btnSave);
            footerBorder.Child = buttonPanel;
            parent.Children.Add(footerBorder);
        }

        private void CreateLabel(StackPanel p, string text)
        {
            p.Children.Add(new TextBlock { Text = text, FontFamily = new FontFamily("Segoe UI"), FontSize = 12, Foreground = Brushes.Gray, Margin = new Thickness(0, 0, 0, 5) });
        }

        // --- Logic ---

        private void LoadRules()
        {
            _rulesList.Items.Clear();
            foreach (var rule in ProfileManager.AutomationRules)
            {
                string mode = rule.IsPersisted ? "[P]" : "[D]";
                _rulesList.Items.Add($"{mode} {rule.ProcessName} → {rule.TargetProfile} ({rule.DelaySeconds}s)");
            }
        }

        private void ApplyRule(bool closeAfter)
        {
            // Silent Validation
            if (string.IsNullOrWhiteSpace(_processSearchCombo.Text) || _profileCombo.SelectedItem == null)
            {
                // If the user clicked "Save" (closeAfter=true) but fields are empty,
                // treat it as a "Cancel/Close" action instead of doing nothing.
                if (closeAfter) Close();
                return;
            }

            var existing = ProfileManager.AutomationRules.FirstOrDefault(r => r.ProcessName.Equals(_processSearchCombo.Text, StringComparison.OrdinalIgnoreCase));
            if (existing != null) ProfileManager.AutomationRules.Remove(existing);

            ProfileManager.AutomationRules.Add(new AutomationRule
            {
                ProcessName = _processSearchCombo.Text.Trim(),
                TargetProfile = _profileCombo.SelectedItem.ToString(),
                IsPersisted = _persistedChk.IsChecked == true,
                DelaySeconds = (int)_delaySlider.Value
            });

            ProfileManager.SaveConfigInternal();
            LoadRules();

            if (closeAfter) Close();
        }

        private void LoadSelectedRuleForEdit()
        {
            if (_rulesList.SelectedIndex < 0) return;
            var rule = ProfileManager.AutomationRules[_rulesList.SelectedIndex];
            _processSearchCombo.Text = rule.ProcessName;
            _profileCombo.SelectedItem = rule.TargetProfile;
            _delaySlider.Value = rule.DelaySeconds;
            _persistedChk.IsChecked = rule.IsPersisted;
        }

        private void DeleteRule()
        {
            if (_rulesList.SelectedIndex < 0) return;
            ProfileManager.AutomationRules.RemoveAt(_rulesList.SelectedIndex);
            ProfileManager.SaveConfigInternal();
            LoadRules();
        }

        private void RefreshProcessList()
        {
            var currentText = _processSearchCombo.Text;
            _processSearchCombo.Items.Clear();
            var processes = Process.GetProcesses().Select(p => p.ProcessName).Distinct().OrderBy(n => n);
            foreach (var n in processes) _processSearchCombo.Items.Add(n);
            _processSearchCombo.Text = currentText;
        }

        private async Task PickWindowInteractively()
        {
            this.Visibility = Visibility.Hidden;
            await Task.Delay(150);

            Window overlay = new Window
            {
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = new SolidColorBrush(Color.FromArgb(1, 0, 0, 0)),
                Topmost = true,
                ShowInTaskbar = false,
                Cursor = Cursors.Cross,
                ForceCursor = true,
                Left = SystemParameters.VirtualScreenLeft,
                Top = SystemParameters.VirtualScreenTop,
                Width = SystemParameters.VirtualScreenWidth,
                Height = SystemParameters.VirtualScreenHeight
            };

            Grid overlayGrid = new Grid();

            // --- UNIFIED UI REPLACEMENT ---
            // Replaced manual Border/TextBlock with the Centralized Factory
            var msgBorder = MessageBoxesManager.CreateUnifiedMessage("CLICK ON THE TARGET WINDOW");

            // Important: Center the message in the grid
            msgBorder.HorizontalAlignment = HorizontalAlignment.Center;
            msgBorder.VerticalAlignment = VerticalAlignment.Center;

            overlayGrid.Children.Add(msgBorder);
            // -----------------------------

            overlay.Content = overlayGrid;

            bool picked = false;
            overlay.PreviewMouseDown += (s, e) =>
            {
                picked = true;
                try
                {
                    System.Drawing.Point p = System.Windows.Forms.Cursor.Position;
                    overlay.Close();

                    IntPtr hWnd = WindowFromPoint(p);
                    if (hWnd != IntPtr.Zero)
                    {
                        IntPtr root = GetAncestor(hWnd, GA_ROOT);
                        if (root != IntPtr.Zero) hWnd = root;

                        GetWindowThreadProcessId(hWnd, out uint pid);
                        Process proc = Process.GetProcessById((int)pid);
                        _processSearchCombo.Text = proc.ProcessName;
                    }
                }
                catch { }
            };

            overlay.Show();
            overlay.KeyDown += (s, e) => { if (e.Key == Key.Escape) overlay.Close(); };

            while (overlay.IsVisible)
            {
                await Task.Delay(100);
            }

            this.Visibility = Visibility.Visible;
            this.Activate();
        }

        // --- MULTI-SCREEN PICKER LOGIC ---
        //private async Task PickWindowInteractively()
        //{
        //    this.Visibility = Visibility.Hidden;
        //    await Task.Delay(150);

        //    Window overlay = new Window
        //    {
        //        WindowStyle = WindowStyle.None,
        //        AllowsTransparency = true,
        //        Background = new SolidColorBrush(Color.FromArgb(1, 0, 0, 0)), // Nearly invisible hit-test layer
        //        Topmost = true,
        //        ShowInTaskbar = false,
        //        Cursor = Cursors.Cross, // FORCE CROSSHAIR CURSOR
        //        ForceCursor = true, // Ensure child elements (if any) don't override it
        //        // Span all screens manually
        //        Left = SystemParameters.VirtualScreenLeft,
        //        Top = SystemParameters.VirtualScreenTop,
        //        Width = SystemParameters.VirtualScreenWidth,
        //        Height = SystemParameters.VirtualScreenHeight
        //    };

        //    Grid overlayGrid = new Grid();
        //    Border msgBorder = new Border
        //    {
        //        Background = new SolidColorBrush(Color.FromArgb(200, 0, 0, 0)),
        //        CornerRadius = new CornerRadius(10),
        //        HorizontalAlignment = HorizontalAlignment.Center,
        //        VerticalAlignment = VerticalAlignment.Center,
        //        Padding = new Thickness(20)
        //    };
        //    TextBlock msg = new TextBlock
        //    {
        //        Text = "CLICK ON THE TARGET WINDOW",
        //        Foreground = Brushes.White,
        //        FontSize = 24,
        //        FontWeight = FontWeights.Bold
        //    };
        //    msgBorder.Child = msg;
        //    overlayGrid.Children.Add(msgBorder);
        //    overlay.Content = overlayGrid;

        //    bool picked = false;
        //    overlay.PreviewMouseDown += (s, e) =>
        //    {
        //        picked = true;
        //        try
        //        {
        //            System.Drawing.Point p = System.Windows.Forms.Cursor.Position;
        //            overlay.Close();

        //            IntPtr hWnd = WindowFromPoint(p);
        //            if (hWnd != IntPtr.Zero)
        //            {
        //                IntPtr root = GetAncestor(hWnd, GA_ROOT);
        //                if (root != IntPtr.Zero) hWnd = root;

        //                GetWindowThreadProcessId(hWnd, out uint pid);
        //                Process proc = Process.GetProcessById((int)pid);
        //                _processSearchCombo.Text = proc.ProcessName;
        //            }
        //        }
        //        catch { }
        //    };

        //    overlay.Show();
        //    overlay.KeyDown += (s, e) => { if (e.Key == Key.Escape) overlay.Close(); };

        //    while (overlay.IsVisible)
        //    {
        //        await Task.Delay(100);
        //    }

        //    this.Visibility = Visibility.Visible;
        //    this.Activate();
        //}

        private void PositionFormOnMouseScreen()
        {
            try
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                var mouseScreen = System.Windows.Forms.Screen.FromPoint(mousePosition);
                double dpiScale = GetFormDpiScaleFactor();
                double centerX = (mouseScreen.Bounds.Left / dpiScale) + ((mouseScreen.Bounds.Width / dpiScale) - this.Width) / 2;
                double centerY = (mouseScreen.Bounds.Top / dpiScale) + ((mouseScreen.Bounds.Height / dpiScale) - this.Height) / 2;
                this.Left = centerX;
                this.Top = centerY;
            }
            catch { this.WindowStartupLocation = WindowStartupLocation.CenterScreen; }
        }

        private double GetFormDpiScaleFactor()
        {
            try { using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero)) return graphics.DpiX / 96.0; }
            catch { return 1.0; }
        }
    }
}
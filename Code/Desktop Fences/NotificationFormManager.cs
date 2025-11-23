using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Threading;
// No "using System.Windows.Input;" to avoid potential collisions with WinForms types.

namespace Desktop_Fences
{
    public class NotificationFormManager : Window
    {
        private RemoteAnnouncement _msg;
        private DispatcherTimer _autoCloseTimer;
        private ProgressBar _timerBar;
        private CheckBox _chkDismiss;

        public static void Show(RemoteAnnouncement msg)
        {
            var win = new NotificationFormManager(msg);
            win.Show();
        }

        public NotificationFormManager(RemoteAnnouncement msg)
        {
            _msg = msg;
            InitializeComponent();

            // Increment display count immediately
            RegistryHelper.IncrementMessageCount(_msg.Id);

            if (_msg.AutoCloseSeconds > 0)
                StartAutoCloseTimer();
        }

        private void InitializeComponent()
        {
            // Window Props
            this.Title = "Desktop Fences + Notification"; // Taskbar/System title
            this.Width = 360;
            this.Height = 200;
            this.SizeToContent = SizeToContent.Height;
            this.WindowStyle = WindowStyle.None;
            this.AllowsTransparency = true;
            this.Background = Brushes.Transparent;
            this.ResizeMode = ResizeMode.NoResize;
            this.Topmost = true;
            this.ShowInTaskbar = false;

            // Position: Bottom Right of Work Area (Tray area)
            double screenRight = SystemParameters.WorkArea.Right;
            double screenBottom = SystemParameters.WorkArea.Bottom;
            this.Left = screenRight - this.Width - 10;
            this.Top = screenBottom; // Start off-screen

            // Animation on Load
            this.Loaded += (s, e) =>
            {
                DoubleAnimation slideUp = new DoubleAnimation(screenBottom, screenBottom - this.ActualHeight - 10, TimeSpan.FromMilliseconds(300));
                slideUp.EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut };
                this.BeginAnimation(TopProperty, slideUp);
            };

            // Main Card
            Border card = new Border
            {
                Background = Brushes.White,
                CornerRadius = new CornerRadius(8),
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 220)),
                Effect = new DropShadowEffect { BlurRadius = 15, ShadowDepth = 5, Opacity = 0.3 }
            };

            // Colors based on Type
            Color headerColor = Utility.GetColorFromName(SettingsManager.SelectedColor); // Default
            string iconSymbol = "ℹ";

            if (_msg.Type == "Warning") { headerColor = Colors.Orange; iconSymbol = "⚠️"; }
            if (_msg.Type == "Alert") { headerColor = Colors.Red; iconSymbol = "🛑"; }

            Grid grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Body
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Footer
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Timer

            // 1. Header
            Border header = new Border
            {
                Background = new SolidColorBrush(headerColor),
                CornerRadius = new CornerRadius(8, 8, 0, 0),
                Padding = new Thickness(10)
            };
            DockPanel headerDock = new DockPanel();

            // Icon
            TextBlock txtIcon = new TextBlock { Text = iconSymbol, Foreground = Brushes.White, FontSize = 16, Margin = new Thickness(0, 0, 10, 0) };

            // Title - Explicit Branding
            // Format: "Desktop Fences + | Message Title"
            TextBlock txtTitle = new TextBlock
            {
                Text = $"Desktop Fences + | {_msg.Title}",
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                FontSize = 14
            };

            // Close X
            Button btnClose = new Button
            {
                Content = "✕",
                Foreground = Brushes.White,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Cursor = System.Windows.Input.Cursors.Hand,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            btnClose.Click += (s, e) => this.Close();

            headerDock.Children.Add(txtIcon);
            headerDock.Children.Add(btnClose);
            DockPanel.SetDock(btnClose, Dock.Right);
            headerDock.Children.Add(txtTitle);
            header.Child = headerDock;
            Grid.SetRow(header, 0);

            // 2. Body
            TextBlock txtBody = new TextBlock
            {
                Text = _msg.Body,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(15, 15, 15, 5),
                FontSize = 13,
                Foreground = new SolidColorBrush(Color.FromRgb(50, 50, 50))
            };
            Grid.SetRow(txtBody, 1);

            // 3. Footer (Actions)
            StackPanel footer = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(15, 10, 15, 15) };

            // Link Button
            if (!string.IsNullOrEmpty(_msg.Link))
            {
                Button btnLink = new Button
                {
                    Content = _msg.Type == "Info" ? "Read More..." : "Update Now",
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Padding = new Thickness(10, 5, 10, 5),
                    Margin = new Thickness(0, 0, 0, 10),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                btnLink.Click += (s, e) => { try { Process.Start(new ProcessStartInfo(_msg.Link) { UseShellExecute = true }); } catch { } };
                footer.Children.Add(btnLink);
            }

            // "Don't Show Again"
            if (_msg.CanUserDismiss)
            {
                _chkDismiss = new CheckBox
                {
                    Content = "Don't show this again",
                    FontSize = 11,
                    Foreground = Brushes.Gray,
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                _chkDismiss.Checked += (s, e) => RegistryHelper.SetMessageDismissed(_msg.Id);
                footer.Children.Add(_chkDismiss);
            }

            Grid.SetRow(footer, 2);

            // 4. Auto-Close Timer Bar
            if (_msg.AutoCloseSeconds > 0)
            {
                _timerBar = new ProgressBar
                {
                    Height = 4,
                    Minimum = 0,
                    Maximum = _msg.AutoCloseSeconds * 100,
                    Value = _msg.AutoCloseSeconds * 100,
                    BorderThickness = new Thickness(0),
                    Foreground = new SolidColorBrush(headerColor),
                    Background = Brushes.Transparent,
                    VerticalAlignment = VerticalAlignment.Bottom
                };
                Grid.SetRow(_timerBar, 3);
                grid.Children.Add(_timerBar);
            }

            grid.Children.Add(header);
            grid.Children.Add(txtBody);
            grid.Children.Add(footer);
            card.Child = grid;
            this.Content = card;
        }

        private void StartAutoCloseTimer()
        {
            _autoCloseTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(10) };
            _autoCloseTimer.Tick += (s, e) =>
            {
                if (_timerBar.Value > 0)
                {
                    _timerBar.Value -= 1;
                }
                else
                {
                    _autoCloseTimer.Stop();
                    this.Close();
                }
            };
            _autoCloseTimer.Start();
        }
    }
}
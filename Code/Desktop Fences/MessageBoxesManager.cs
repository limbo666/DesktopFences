using System;
using System.IO;
using System.Media;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;

namespace Desktop_Fences
{
    /// <summary>
    /// Manager class for modern WPF message boxes and dialogs
    /// DPI-aware with monitor positioning support and consistent styling
    /// </summary>
    public static class MessageBoxesManager
    {
        #region ShowCustomMessageBoxForm - Delete Confirmation Dialog
        /// <summary>
        /// Shows a modern WPF confirmation dialog for fence deletion
        /// Converted from Windows Forms with identical functionality and layout
        /// </summary>
        /// <returns>True if user clicked Yes, False if user clicked No</returns>
        public static bool ShowCustomMessageBoxForm()
        {
            bool result = false;

            try
            {
                var messageBox = new CustomMessageBoxWindow();
                messageBox.ShowDialog();
                result = messageBox.DialogResult;

                return result;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing CustomMessageBox form: {ex.Message}");
                return false;
            }
        }
        #endregion

        #region ShowCustomYesNoMessageBox - Generic Confirmation Dialog
        /// <summary>
        /// Shows a modern WPF Yes/No confirmation dialog with custom message and title
        /// </summary>
        /// <param name="message">The question to ask the user</param>
        /// <param name="title">The dialog title</param>
        /// <returns>True if Yes was clicked, False if No was clicked</returns>
        public static bool ShowCustomYesNoMessageBox(string message, string title)
        {
            try
            {
                var messageBox = new CustomYesNoMessageBoxWindow(message, title);
                messageBox.ShowDialog();
                return messageBox.DialogResult;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing Yes/No MessageBox: {ex.Message}");
                return false;
            }
        }
        #endregion



        #region CustomYesNoMessageBoxWindow - Internal WPF Window Class
        /// <summary>
        /// Internal WPF window for generic Yes/No confirmation dialogs
        /// </summary>
        private class CustomYesNoMessageBoxWindow : Window
        {
            private bool _result = false;
            private Color _userAccentColor;
            private string _message;
            private string _title;

            public new bool DialogResult => _result;

            public CustomYesNoMessageBoxWindow(string message, string title)
            {
                _message = message;
                _title = title;
                InitializeComponent();
                PositionWindowOnMouseScreen(this);
                PlayDingSound();
            }

            private void InitializeComponent()
            {
                // Get user's accent color
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                _userAccentColor = mediaColor;

                // Modern WPF window setup
                this.Title = _title;
                this.Width = 450;
                this.Height = 220;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;
                this.Topmost = true;

                // Set icon
                try
                {
                    this.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { }

                CreateContent();
            }

            private void CreateContent()
            {
                // Main white card
                Border mainCard = new Border
                {
                    Background = Brushes.White,
                    Margin = new Thickness(8, 8, 8, 1),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 270,
                        ShadowDepth = 2,
                        BlurRadius = 10,
                        Opacity = 0.1
                    }
                };

                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) });
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

                // Accent header
                Border accentHeader = new Border
                {
                    Background = new SolidColorBrush(_userAccentColor),
                    Height = 8
                };
                Grid.SetRow(accentHeader, 0);

                // Content area
                Grid contentGrid = new Grid { Margin = new Thickness(24) };
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(64) });
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) });

                Grid.SetRow(contentGrid, 1);

                // Warning icon
                Border iconContainer = new Border
                {
                    Width = 48,
                    Height = 48,
                    Background = new SolidColorBrush(Color.FromArgb(15, _userAccentColor.R, _userAccentColor.G, _userAccentColor.B)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(0, 4, 0, 0)
                };
                TextBlock warningIcon = new TextBlock
                {
                    Text = "⚠",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 32,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(_userAccentColor),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };
                iconContainer.Child = warningIcon;
                Grid.SetColumn(iconContainer, 0); Grid.SetRow(iconContainer, 0);

                // Message area
                StackPanel messageArea = new StackPanel { Margin = new Thickness(8, 4, 0, 0) };
                TextBlock titleLabel = new TextBlock
                {
                    Text = _title,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 18,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                    Margin = new Thickness(0, 0, 0, 4)
                };
                TextBlock messageLabel = new TextBlock
                {
                    Text = _message,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 300
                };
                messageArea.Children.Add(titleLabel); messageArea.Children.Add(messageLabel);
                Grid.SetColumn(messageArea, 1); Grid.SetRow(messageArea, 0);

                // Button area
                StackPanel buttonArea = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(8, 8, 0, 0)
                };

                // No Button
                Button btnNo = new Button
                {
                    Content = "No",
                    Width = 80,
                    Height = 32,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = new SolidColorBrush(_userAccentColor),
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    Margin = new Thickness(0, 0, 4, 0)
                };
                btnNo.MouseEnter += (s, e) => btnNo.Background = new SolidColorBrush(Color.FromRgb((byte)Math.Max(0, _userAccentColor.R - 25), (byte)Math.Max(0, _userAccentColor.G - 25), (byte)Math.Max(0, _userAccentColor.B - 25)));
                btnNo.MouseLeave += (s, e) => btnNo.Background = new SolidColorBrush(_userAccentColor);
                btnNo.Click += (s, e) => { _result = false; this.Close(); };

                // Yes Button (Red for Reset/Danger actions)
                Button btnYes = new Button
                {
                    Content = "Yes",
                    Width = 80,
                    Height = 32,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = new SolidColorBrush(Color.FromRgb(234, 67, 53)),
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };
                btnYes.MouseEnter += (s, e) => btnYes.Background = new SolidColorBrush(Color.FromRgb(219, 50, 36));
                btnYes.MouseLeave += (s, e) => btnYes.Background = new SolidColorBrush(Color.FromRgb(234, 67, 53));
                btnYes.Click += (s, e) => { _result = true; this.Close(); };

                buttonArea.Children.Add(btnNo); buttonArea.Children.Add(btnYes);
                Grid.SetColumn(buttonArea, 1); Grid.SetRow(buttonArea, 1);

                contentGrid.Children.Add(iconContainer); contentGrid.Children.Add(messageArea); contentGrid.Children.Add(buttonArea);
                mainGrid.Children.Add(accentHeader); mainGrid.Children.Add(contentGrid);
                mainCard.Child = mainGrid;
                this.Content = mainCard;

                this.KeyDown += (s, e) => { if (e.Key == Key.Enter) { _result = true; this.Close(); } else if (e.Key == Key.Escape) { _result = false; this.Close(); } };
                this.Focusable = true; this.Focus();
            }
        }
        #endregion

        public static void ShowOKOnlyMessageBoxFormStatic(string msgboxMessage, string msgboxTitle)
        {
            MessageBoxesManager.ShowOKOnlyMessageBoxForm(msgboxMessage, msgboxTitle);
        }




        #region AutoClosingMessageBoxWindow - Internal WPF Window Class
        /// <summary>
        /// Internal WPF window for auto-closing notification dialogs
        /// </summary>
        private class AutoClosingMessageBoxWindow : Window
        {
            private Color _userAccentColor;
            private string _message;
            private string _title;
            private int _autoCloseTimeMs;
            private System.Windows.Threading.DispatcherTimer _autoCloseTimer;
            private TextBlock _autoCloseLabel;
            private int _remainingSeconds;

            public AutoClosingMessageBoxWindow(string message, string title, int autoCloseTimeMs)
            {
                _message = message;
                _title = title;
                _autoCloseTimeMs = autoCloseTimeMs;
                _remainingSeconds = autoCloseTimeMs / 1000;
                InitializeComponent();
                PositionWindowOnMouseScreen(this);
                PlayNotificationSound();
                StartAutoCloseTimer();
            }

            private void InitializeComponent()
            {
                // Get user's accent color
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                _userAccentColor = mediaColor;

                // Modern WPF window setup with dynamic sizing
                this.Title = _title;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;
                this.Topmost = true;

                // Dynamic sizing based on message length (slightly smaller for notifications)
                var textSize = MeasureText(_message);
                int dynamicHeight = Math.Max(200, (int)(textSize.Height + 180));
                this.Width = 420;
                this.Height = dynamicHeight;

                // Set icon
                try
                {
                    this.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { }

                CreateAutoClosingContent();
            }

            private Size MeasureText(string text)
            {
                var typeface = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
                var formattedText = new FormattedText(text, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeface, 14, Brushes.Black, 96);
                formattedText.MaxTextWidth = 380;
                return new Size(formattedText.Width, formattedText.Height);
            }

            private void CreateAutoClosingContent()
            {
                // Main white card
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

                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(45) }); // Header (slightly smaller)
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content

                // Header with accent color and auto-close indicator
                Border headerBorder = new Border
                {
                    Background = new SolidColorBrush(_userAccentColor),
                    Height = 45
                };
                Grid.SetRow(headerBorder, 0);

                Grid headerGrid = new Grid();
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(130) });
                headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(30) });

                // Title
                TextBlock titleLabel = new TextBlock
                {
                    Text = _title,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 16,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(15, 0, 0, 0)
                };
                Grid.SetColumn(titleLabel, 0);

                // Auto-close indicator
                _autoCloseLabel = new TextBlock
                {
                    Text = $"Auto-closing in {_remainingSeconds}s",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Color.FromArgb(220, 255, 255, 255)),
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 0, 5, 0)
                };
                Grid.SetColumn(_autoCloseLabel, 1);

                // Manual close button
                Button manualCloseButton = new Button
                {
                    Content = "✕",
                    Width = 24,
                    Height = 20,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 10,
                    FontWeight = FontWeights.Bold,
                    Foreground = Brushes.White,
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                manualCloseButton.MouseEnter += (s, e) => manualCloseButton.Background = new SolidColorBrush(Color.FromArgb(30, 255, 255, 255));
                manualCloseButton.MouseLeave += (s, e) => manualCloseButton.Background = Brushes.Transparent;
                manualCloseButton.Click += (s, e) => this.Close();
                Grid.SetColumn(manualCloseButton, 2);

                headerGrid.Children.Add(titleLabel);
                headerGrid.Children.Add(_autoCloseLabel);
                headerGrid.Children.Add(manualCloseButton);
                headerBorder.Child = headerGrid;

                // Content area with centered message
                Border contentBorder = new Border
                {
                    Background = Brushes.White,
                    Padding = new Thickness(20)
                };
                Grid.SetRow(contentBorder, 1);

                TextBlock messageLabel = new TextBlock
                {
                    Text = _message,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(60, 64, 67)),
                    TextWrapping = TextWrapping.Wrap,
                    TextAlignment = TextAlignment.Center, // Center for notifications
                    VerticalAlignment = VerticalAlignment.Center,
                    HorizontalAlignment = HorizontalAlignment.Center
                };

                contentBorder.Child = messageLabel;

                mainGrid.Children.Add(headerBorder);
                mainGrid.Children.Add(contentBorder);

                mainCard.Child = mainGrid;
                this.Content = mainCard;

                // Add keyboard support for Enter/Escape keys
                this.KeyDown += AutoClosingMessageBox_KeyDown;
                this.Focusable = true;
                this.Focus();
            }

            /// <summary>
            /// Handles keyboard input for auto-closing notification message box
            /// Enter/Escape = Close immediately (stop auto-close timer)
            /// </summary>
            private void AutoClosingMessageBox_KeyDown(object sender, KeyEventArgs e)
            {
                switch (e.Key)
                {
                    case Key.Enter:
                    case Key.Escape:
                        _autoCloseTimer?.Stop();
                        this.Close();
                        e.Handled = true;
                        break;
                }
            }

            private void StartAutoCloseTimer()
            {
                _autoCloseTimer = new System.Windows.Threading.DispatcherTimer();
                _autoCloseTimer.Interval = TimeSpan.FromSeconds(1);
                _autoCloseTimer.Tick += AutoCloseTimer_Tick;
                _autoCloseTimer.Start();
            }

            private void AutoCloseTimer_Tick(object sender, EventArgs e)
            {
                _remainingSeconds--;
                if (_remainingSeconds <= 0)
                {
                    _autoCloseTimer.Stop();
                    this.Close();
                }
                else
                {
                    _autoCloseLabel.Text = $"Auto-closing in {_remainingSeconds}s";
                }
            }

            private void PlayNotificationSound()
            {
                try
                {
                    if (SettingsManager.IsLogEnabled)
                    {
                        System.Media.SystemSounds.Asterisk.Play();
                    }
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error playing sound: {ex.Message}");
                }
            }

            protected override void OnClosed(EventArgs e)
            {
                _autoCloseTimer?.Stop();
                _autoCloseTimer = null;
                base.OnClosed(e);
            }
        }
        #endregion

        #region WaitWindow - Internal WPF Window Class
        /// <summary>
        /// Internal WPF window for wait/loading operations
        /// </summary>
        private class WaitWindow : Window
        {
            public WaitWindow(string title, string message)
            {
                InitializeComponent(title, message);
            }

            private void InitializeComponent(string title, string message)
            {
                this.Title = title;
                this.Width = 350;
                this.Height = 180;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.AllowsTransparency = true;
                this.Topmost = true;
                this.ResizeMode = ResizeMode.NoResize;

                CreateWaitContent(title, message);
            }

            private void CreateWaitContent(string title, string message)
            {
                // Main container with modern styling
                Border mainBorder = new Border
                {
                    Background = Brushes.White,
                    Margin = new Thickness(8),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 270,
                        ShadowDepth = 4,
                        Opacity = 0.15,
                        BlurRadius = 8
                    }
                };

                StackPanel waitStack = new StackPanel
                {
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    Orientation = Orientation.Vertical
                };

                // App title
                TextBlock titleText = new TextBlock
                {
                    Text = title,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 16,
                    FontWeight = FontWeights.Medium,
                    Foreground = new SolidColorBrush(Color.FromRgb(66, 133, 244)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                waitStack.Children.Add(titleText);

                // Logo with fallback
                System.Windows.Controls.Image logoImage = new System.Windows.Controls.Image
                {
                    Width = 32,
                    Height = 32,
                    Margin = new Thickness(0, 0, 0, 10),
                    HorizontalAlignment = HorizontalAlignment.Center
                };

                try
                {
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var resourceStream = assembly.GetManifestResourceStream("Desktop_Fences.Resources.logo1.png");
                    if (resourceStream != null)
                    {
                        var bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.StreamSource = resourceStream;
                        bitmap.EndInit();
                        logoImage.Source = bitmap;
                    }
                    else
                    {
                        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
                        logoImage.Source = Utility.ToImageSource(System.Drawing.Icon.ExtractAssociatedIcon(exePath));
                    }
                }
                catch
                {
                    try
                    {
                        string exePath = System.Reflection.Assembly.GetEntryAssembly().Location;
                        logoImage.Source = Utility.ToImageSource(System.Drawing.Icon.ExtractAssociatedIcon(exePath));
                    }
                    catch
                    {
                        // Hide logo if all attempts fail
                        logoImage.Visibility = Visibility.Collapsed;
                    }
                }
                waitStack.Children.Add(logoImage);

                // Wait message
                TextBlock waitText = new TextBlock
                {
                    Text = message,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 12,
                    Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                    HorizontalAlignment = HorizontalAlignment.Center
                };
                waitStack.Children.Add(waitText);

                mainBorder.Child = waitStack;
                this.Content = mainBorder;
            }
        }

        #region ShowAutoClosingMessageBoxForm - Auto-Closing Notification Dialog
        /// <summary>
        /// Shows a modern WPF auto-closing notification dialog
        /// Converted from Windows Forms with identical functionality and auto-close timer
        /// </summary>
        /// <param name="message">The message to display</param>
        /// <param name="title">The dialog title</param>
        /// <param name="autoCloseTimeMs">Time in milliseconds before auto-close (default: 2000)</param>
        public static void ShowAutoClosingMessageBoxForm(string message, string title, int autoCloseTimeMs = 2000)
        {
            try
            {
                var messageBox = new AutoClosingMessageBoxWindow(message, title, autoCloseTimeMs);
                messageBox.Show(); // Use Show() instead of ShowDialog() for non-blocking notifications
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing auto-closing MessageBox: {ex.Message}");
                // Fallback to regular message box
                System.Windows.MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        #endregion

        #region CreateWaitWindow - Standardized Wait Window
        /// <summary>
        /// Creates and shows a standardized wait window for long operations
        /// Converted from WPF with improved styling and DPI awareness
        /// </summary>
        /// <param name="title">The title of the wait window</param>
        /// <param name="message">The message to display</param>
        /// <returns>The wait window instance (caller should close it when done)</returns>
        public static Window CreateWaitWindow(string title = "Desktop Fences +", string message = "Please wait...")
        {
            try
            {
                var waitWindow = new WaitWindow(title, message);
                PositionWindowOnMouseScreen(waitWindow);
                return waitWindow;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error creating wait window: {ex.Message}");
                // Return a basic fallback window
                return new Window
                {
                    Title = title,
                    Width = 350,
                    Height = 180,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    Content = new TextBlock { Text = message, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center }
                };
            }
        }

        #region OKOnlyMessageBoxWindow - Internal WPF Window Class
        /// <summary>
        /// Internal WPF window for OK-only information/error dialogs with dynamic sizing
        /// </summary>
        private class OKOnlyMessageBoxWindow : Window
        {
            private Color _userAccentColor;
            private string _message;
            private string _title;

            public OKOnlyMessageBoxWindow(string message, string title)
            {
                _message = message;
                _title = title;
                InitializeComponent();
                PositionWindowOnMouseScreen(this);
                PlayDingSound();
            }

            private void InitializeComponent()
            {
                // Get user's accent color
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                _userAccentColor = mediaColor;

                // Modern WPF window setup with dynamic sizing
                this.Title = _title;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;
                this.Topmost = true;

                // Dynamic sizing based on message length
                var textSize = MeasureText(_message);
                int dynamicHeight = Math.Max(220, (int)(textSize.Height + 200));
                this.Width = 450;
                this.Height = dynamicHeight;

                // Set icon
                try
                {
                    this.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { }

                CreateOKOnlyContent();
            }

            private Size MeasureText(string text)
            {
                var typeface = new Typeface(new FontFamily("Segoe UI"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
                var formattedText = new FormattedText(text, System.Globalization.CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeface, 14, Brushes.Black, 96);
                formattedText.MaxTextWidth = 380;
                return new Size(formattedText.Width, formattedText.Height);
            }

            private void CreateOKOnlyContent()
            {
                // Main white card
                Border mainCard = new Border
                {
                    Background = Brushes.White,
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

                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(6) }); // Accent header
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content

                // Accent header
                Border accentHeader = new Border
                {
                    Background = new SolidColorBrush(_userAccentColor),
                    Height = 6
                };
                Grid.SetRow(accentHeader, 0);

                // Content area
                Grid contentGrid = new Grid
                {
                    Margin = new Thickness(24)
                };
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(64) });
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) });

                Grid.SetRow(contentGrid, 1);

                // Info icon
                Border iconContainer = new Border
                {
                    Width = 48,
                    Height = 48,
                    Background = new SolidColorBrush(Color.FromArgb(15, _userAccentColor.R, _userAccentColor.G, _userAccentColor.B)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(0, 4, 0, 0)
                };

                TextBlock infoIcon = new TextBlock
                {
                    Text = "ℹ",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 42,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(_userAccentColor),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };

                iconContainer.Child = infoIcon;
                Grid.SetColumn(iconContainer, 0);
                Grid.SetRow(iconContainer, 0);

                // Message area
                StackPanel messageArea = new StackPanel
                {
                    Margin = new Thickness(8, 4, 0, 0)
                };

                TextBlock titleLabel = new TextBlock
                {
                    Text = _title,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 18,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                    Margin = new Thickness(0, 0, 0, 8)
                };

                TextBlock messageLabel = new TextBlock
                {
                    Text = _message,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 320
                };

                messageArea.Children.Add(titleLabel);
                messageArea.Children.Add(messageLabel);
                Grid.SetColumn(messageArea, 1);
                Grid.SetRow(messageArea, 0);

                // Button area
                StackPanel buttonArea = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(8, 8, 0, 0)
                };

                // OK button with user's accent color
                Button btnOK = new Button
                {
                    Content = "OK",
                    Width = 90,
                    Height = 36,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = new SolidColorBrush(_userAccentColor),
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };

                Color okButtonHover = Color.FromRgb(
                    (byte)Math.Max(0, _userAccentColor.R - 25),
                    (byte)Math.Max(0, _userAccentColor.G - 25),
                    (byte)Math.Max(0, _userAccentColor.B - 25)
                );

                btnOK.MouseEnter += (s, e) => btnOK.Background = new SolidColorBrush(okButtonHover);
                btnOK.MouseLeave += (s, e) => btnOK.Background = new SolidColorBrush(_userAccentColor);
                btnOK.Click += (s, e) => this.Close();

                buttonArea.Children.Add(btnOK);
                Grid.SetColumn(buttonArea, 1);
                Grid.SetRow(buttonArea, 1);

                contentGrid.Children.Add(iconContainer);
                contentGrid.Children.Add(messageArea);
                contentGrid.Children.Add(buttonArea);

                mainGrid.Children.Add(accentHeader);
                mainGrid.Children.Add(contentGrid);

                mainCard.Child = mainGrid;
                this.Content = mainCard;

                // Add keyboard support for Enter/Escape keys
                this.KeyDown += OKOnlyMessageBox_KeyDown;
                this.Focusable = true;
                this.Focus();
            }

            /// <summary>
            /// Handles keyboard input for OK-only message box
            /// Enter = OK/Close, Escape = Close
            /// </summary>
            private void OKOnlyMessageBox_KeyDown(object sender, KeyEventArgs e)
            {
                switch (e.Key)
                {
                    case Key.Enter:
                    case Key.Escape:
                        this.Close();
                        e.Handled = true;
                        break;
                }
            }
        }
        #endregion

        #region TabDeleteConfirmationWindow - Internal WPF Window Class
        /// <summary>
        /// Internal WPF window for tab deletion confirmation dialog
        /// </summary>
        private class TabDeleteConfirmationWindow : Window
        {
            private bool _result = false;
            private Color _userAccentColor;
            private string _tabName;
            private int _itemCount;

            public new bool DialogResult => _result;

            public TabDeleteConfirmationWindow(string tabName, int itemCount)
            {
                _tabName = tabName;
                _itemCount = itemCount;
                InitializeComponent();
                PositionWindowOnMouseScreen(this);
                PlayDingSound();
            }

            private void InitializeComponent()
            {
                // Get user's accent color
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                _userAccentColor = mediaColor;

                // Modern WPF window setup
                this.Title = "Delete Tab";
                this.Width = 420;
                this.Height = 220;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;
                this.Topmost = true;

                // Set icon
                try
                {
                    this.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { }

                CreateTabDeleteContent();
            }

            private void CreateTabDeleteContent()
            {
                // Main white card
                Border mainCard = new Border
                {
                    Background = Brushes.White,
                    Margin = new Thickness(8, 8, 8, 1),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 270,
                        ShadowDepth = 2,
                        BlurRadius = 10,
                        Opacity = 0.1
                    }
                };

                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) });
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

                // Accent header
                Border accentHeader = new Border
                {
                    Background = new SolidColorBrush(_userAccentColor),
                    Height = 8
                };
                Grid.SetRow(accentHeader, 0);

                // Content area
                Grid contentGrid = new Grid
                {
                    Margin = new Thickness(24)
                };
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(64) });
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) });

                Grid.SetRow(contentGrid, 1);

                // Warning icon
                Border iconContainer = new Border
                {
                    Width = 48,
                    Height = 48,
                    Background = new SolidColorBrush(Color.FromArgb(15, _userAccentColor.R, _userAccentColor.G, _userAccentColor.B)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(0, 4, 0, 0)
                };

                TextBlock warningIcon = new TextBlock
                {
                    Text = "🗂",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 42,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(_userAccentColor),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };

                iconContainer.Child = warningIcon;
                Grid.SetColumn(iconContainer, 0);
                Grid.SetRow(iconContainer, 0);

                // Message area
                StackPanel messageArea = new StackPanel
                {
                    Margin = new Thickness(8, 4, 0, 0)
                };

                TextBlock titleLabel = new TextBlock
                {
                    Text = "Delete Tab",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 18,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                    Margin = new Thickness(0, 0, 0, 4)
                };

                string itemText = _itemCount == 1 ? "item" : "items";
                TextBlock messageLabel = new TextBlock
                {
                    Text = $"Are you sure you want to delete tab\n '{_tabName}'?\nThis tab contains {_itemCount} {itemText} that will be permanently removed.",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 280
                };

                messageArea.Children.Add(titleLabel);
                messageArea.Children.Add(messageLabel);
                Grid.SetColumn(messageArea, 1);
                Grid.SetRow(messageArea, 0);

                // Button area
                StackPanel buttonArea = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(8, 8, 0, 0)
                };

                // No button (safe action)
                Button btnNo = new Button
                {
                    Content = "No",
                    Width = 80,
                    Height = 32,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = new SolidColorBrush(_userAccentColor),
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    Margin = new Thickness(0, 0, 4, 0)
                };

                Color noButtonHover = Color.FromRgb(
                    (byte)Math.Max(0, _userAccentColor.R - 25),
                    (byte)Math.Max(0, _userAccentColor.G - 25),
                    (byte)Math.Max(0, _userAccentColor.B - 25)
                );

                btnNo.MouseEnter += (s, e) => btnNo.Background = new SolidColorBrush(noButtonHover);
                btnNo.MouseLeave += (s, e) => btnNo.Background = new SolidColorBrush(_userAccentColor);
                btnNo.Click += (s, e) => { _result = false; this.Close(); };

                // Yes button (danger action)
                Button btnYes = new Button
                {
                    Content = "Yes",
                    Width = 80,
                    Height = 32,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = new SolidColorBrush(Color.FromRgb(234, 67, 53)),
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };

                btnYes.MouseEnter += (s, e) => btnYes.Background = new SolidColorBrush(Color.FromRgb(219, 50, 36));
                btnYes.MouseLeave += (s, e) => btnYes.Background = new SolidColorBrush(Color.FromRgb(234, 67, 53));
                btnYes.Click += (s, e) => { _result = true; this.Close(); };

                buttonArea.Children.Add(btnNo);
                buttonArea.Children.Add(btnYes);
                Grid.SetColumn(buttonArea, 1);
                Grid.SetRow(buttonArea, 1);

                contentGrid.Children.Add(iconContainer);
                contentGrid.Children.Add(messageArea);
                contentGrid.Children.Add(buttonArea);

                mainGrid.Children.Add(accentHeader);
                mainGrid.Children.Add(contentGrid);

                mainCard.Child = mainGrid;
                this.Content = mainCard;

                // Add keyboard support for Enter/Escape keys
                this.KeyDown += TabDeleteConfirmation_KeyDown;
                this.Focusable = true;
                this.Focus();
            }

            /// <summary>
            /// Handles keyboard input for tab deletion confirmation message box
            /// Enter = Yes (delete tab), Escape = No (cancel)
            /// </summary>
            private void TabDeleteConfirmation_KeyDown(object sender, KeyEventArgs e)
            {
                switch (e.Key)
                {
                    case Key.Enter:
                        _result = true;
                        this.Close();
                        e.Handled = true;
                        break;
                    case Key.Escape:
                        _result = false;
                        this.Close();
                        e.Handled = true;
                        break;
                }
            }
        }
        #endregion

        #region Helper Methods for Future Message Boxes
        /// <summary>
        /// Converts WPF Color to System.Drawing.Color for compatibility
        /// </summary>
        private static System.Drawing.Color ConvertToDrawingColor(Color mediaColor)
        {
            return System.Drawing.Color.FromArgb(mediaColor.A, mediaColor.R, mediaColor.G, mediaColor.B);
        }

        /// <summary>
        /// Positions window on the screen where the mouse is currently located
        /// </summary>
        /// <summary>
        /// Positions window on the screen where the mouse is currently located with DPI scaling support
        /// </summary>
        private static void PositionWindowOnMouseScreen(Window window)
        {
            try
            {
                var mousePosition = System.Windows.Forms.Cursor.Position;
                var mouseScreen = System.Windows.Forms.Screen.FromPoint(mousePosition);

                // Get DPI scale factor for proper WPF positioning
                double dpiScale = GetMessageBoxDpiScaleFactor();
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Mouse position: X={mousePosition.X}, Y={mousePosition.Y}, DPI scale: {dpiScale}");

                // Convert physical pixels to device-independent units (DIUs)
                double screenLeftDiu = mouseScreen.WorkingArea.Left / dpiScale;
                double screenTopDiu = mouseScreen.WorkingArea.Top / dpiScale;
                double screenWidthDiu = mouseScreen.WorkingArea.Width / dpiScale;
                double screenHeightDiu = mouseScreen.WorkingArea.Height / dpiScale;

                // Calculate center position in DIUs
                double centerX = screenLeftDiu + (screenWidthDiu - window.Width) / 2;
                double centerY = screenTopDiu + (screenHeightDiu - window.Height) / 2;

                window.Left = centerX;
                window.Top = centerY;

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Positioned message box at X={centerX}, Y={centerY} on screen '{mouseScreen.DeviceName}' (DPI scale: {dpiScale})");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error positioning message box: {ex.Message}");
                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Falling back to CenterScreen positioning");
            }
        }

        /// <summary>
        /// Gets the DPI scale factor for proper WPF message box positioning
        /// </summary>
        private static double GetMessageBoxDpiScaleFactor()
        {
            try
            {
                // Use Graphics to get the screen's DPI
                using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    float dpiX = graphics.DpiX; // Horizontal DPI
                    return dpiX / 96.0; // Standard DPI is 96, so scale factor = dpiX / 96
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"Could not get DPI scale factor for message box: {ex.Message}. Using default scale of 1.0");
                return 1.0; // Default to no scaling if DPI detection fails
            }
        }

        /// <summary>
        /// Plays the ding.wav sound from embedded resources
        /// </summary>
        private static void PlayDingSound()
        {
           // System.Windows.MessageBox.Show($"Enable sound is: {SettingsManager.EnableSounds}", "Info", MessageBoxButton.OK);
            if (SettingsManager.EnableSounds == false)


            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    "MessageBoxes: Sound is muted, skipping ding sound playback");
                return;
            }


            try
            {
                using (Stream soundStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Desktop_Fences.Resources.ui-8-warning-sound-effect-336254.wav"))
                {
                    if (soundStream != null)
                    {
                        using (SoundPlayer player = new SoundPlayer(soundStream))
                        {
                            player.Play();
                        }
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, "Sound resource 'ding.wav' not found.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error playing sound: {ex.Message}");
            }
        }
        #endregion

        #region CustomMessageBoxWindow - Internal WPF Window Class
        /// <summary>
        /// Internal WPF window for the delete confirmation dialog
        /// </summary>
        private class CustomMessageBoxWindow : Window
        {
            private bool _result = false;
            private Color _userAccentColor;

            public new bool DialogResult => _result;

            public CustomMessageBoxWindow()
            {
                InitializeComponent();
                PositionWindowOnMouseScreen(this);
                PlayDingSound();
            }

            private void InitializeComponent()
            {
                // Get user's accent color
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                _userAccentColor = mediaColor;

                // Modern WPF window setup
                this.Title = "Confirm delete";
                this.Width = 420;
                this.Height = 200;
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.WindowStyle = WindowStyle.None;
                this.AllowsTransparency = true;
                this.Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
                this.ResizeMode = ResizeMode.NoResize;
                this.Topmost = true;

                // Set icon from executable
                try
                {
                    this.Icon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(
                        System.Drawing.Icon.ExtractAssociatedIcon(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName).Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );
                }
                catch { }

                CreateContent();
            }

            private void CreateContent()
            {
                // Main white card with shadow effect
                Border mainCard = new Border
                {
                    Background = Brushes.White,
                    Margin = new Thickness(8, 8, 8, 1),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 270,
                        ShadowDepth = 2,
                        BlurRadius = 10,
                        Opacity = 0.1
                    }
                };

                // Main grid layout
                Grid mainGrid = new Grid();
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(8) }); // Accent header
                mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content

                // Creative header - thin accent bar using user's color
                Border accentHeader = new Border
                {
                    Background = new SolidColorBrush(_userAccentColor),
                    Height = 8
                };
                Grid.SetRow(accentHeader, 0);

                // Content area
                Grid contentGrid = new Grid
                {
                    Margin = new Thickness(24)
                };
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(64) }); // Icon column
                contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Content column
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Message area
                contentGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(50) }); // Button area

                Grid.SetRow(contentGrid, 1);

                // Warning icon with user's accent color
                Border iconContainer = new Border
                {
                    Width = 48,
                    Height = 48,
                    Background = new SolidColorBrush(Color.FromArgb(15, _userAccentColor.R, _userAccentColor.G, _userAccentColor.B)),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Top,
                    Margin = new Thickness(0, 4, 0, 0)
                };

                TextBlock warningIcon = new TextBlock
                {
                    Text = "⚠",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 32,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(_userAccentColor),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };

                iconContainer.Child = warningIcon;
                Grid.SetColumn(iconContainer, 0);
                Grid.SetRow(iconContainer, 0);

                // Message area
                StackPanel messageArea = new StackPanel
                {
                    Margin = new Thickness(8, 4, 0, 0)
                };

                TextBlock titleLabel = new TextBlock
                {
                    Text = "Delete Fence",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 18,
                    FontWeight = FontWeights.Bold,
                    Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                    Margin = new Thickness(0, 0, 0, 4)
                };

                TextBlock messageLabel = new TextBlock
                {
                    Text = "Are you sure you want to delete this fence?",
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 280
                };

                messageArea.Children.Add(titleLabel);
                messageArea.Children.Add(messageLabel);
                Grid.SetColumn(messageArea, 1);
                Grid.SetRow(messageArea, 0);

                // Button area
                StackPanel buttonArea = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(8, 8, 0, 0)
                };

                // No button with user's accent color (safe action)
                Button btnNo = new Button
                {
                    Content = "No",
                    Width = 80,
                    Height = 32,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = new SolidColorBrush(_userAccentColor),
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand,
                    Margin = new Thickness(0, 0, 4, 0)
                };

                // Dynamic hover color based on user's accent
                Color noButtonHover = Color.FromRgb(
                    (byte)Math.Max(0, _userAccentColor.R - 25),
                    (byte)Math.Max(0, _userAccentColor.G - 25),
                    (byte)Math.Max(0, _userAccentColor.B - 25)
                );

                btnNo.MouseEnter += (s, e) => btnNo.Background = new SolidColorBrush(noButtonHover);
                btnNo.MouseLeave += (s, e) => btnNo.Background = new SolidColorBrush(_userAccentColor);
                btnNo.Click += (s, e) => { _result = false; this.Close(); };

                // Yes button with danger styling
                Button btnYes = new Button
                {
                    Content = "Yes",
                    Width = 80,
                    Height = 32,
                    FontFamily = new FontFamily("Segoe UI"),
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Background = new SolidColorBrush(Color.FromRgb(234, 67, 53)), // Material Red
                    Foreground = Brushes.White,
                    BorderThickness = new Thickness(0),
                    Cursor = Cursors.Hand
                };

                btnYes.MouseEnter += (s, e) => btnYes.Background = new SolidColorBrush(Color.FromRgb(219, 50, 36));
                btnYes.MouseLeave += (s, e) => btnYes.Background = new SolidColorBrush(Color.FromRgb(234, 67, 53));
                btnYes.Click += (s, e) => { _result = true; this.Close(); };

                buttonArea.Children.Add(btnNo);
                buttonArea.Children.Add(btnYes);
                Grid.SetColumn(buttonArea, 1);
                Grid.SetRow(buttonArea, 1);

                contentGrid.Children.Add(iconContainer);
                contentGrid.Children.Add(messageArea);
                contentGrid.Children.Add(buttonArea);

                mainGrid.Children.Add(accentHeader);
                mainGrid.Children.Add(contentGrid);

                mainCard.Child = mainGrid;
                this.Content = mainCard;

                // Add keyboard support for Enter/Escape keys
                this.KeyDown += CustomMessageBox_KeyDown;
                this.Focusable = true;
                this.Focus();
            }

            /// <summary>
            /// Handles keyboard input for delete confirmation message box
            /// Enter = Yes (delete), Escape = No (cancel)
            /// </summary>
            private void CustomMessageBox_KeyDown(object sender, KeyEventArgs e)
            {
                switch (e.Key)
                {
                    case Key.Enter:
                        _result = true;
                        this.Close();
                        e.Handled = true;
                        break;
                    case Key.Escape:
                        _result = false;
                        this.Close();
                        e.Handled = true;
                        break;
                }
            }
        }
        #endregion

        #region ShowOKOnlyMessageBoxForm - Information/Error Dialog
        /// <summary>
        /// Shows a modern WPF OK-only dialog with dynamic sizing for variable message lengths
        /// Converted from TrayManager.ShowOKOnlyMessageBoxForm with identical functionality
        /// </summary>
        /// <param name="message">The message to display</param>
        /// <param name="title">The dialog title</param>
        public static void ShowOKOnlyMessageBoxForm(string message, string title)
        {
            try
            {
                var messageBox = new OKOnlyMessageBoxWindow(message, title);
                messageBox.ShowDialog();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing OK-only MessageBox: {ex.Message}");
                // Fallback to system message box
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region ShowTabDeleteConfirmationForm - Tab Deletion Dialog
        /// <summary>
        /// Shows a modern WPF tab deletion confirmation dialog
        /// Converted from TrayManager.ShowTabDeleteConfirmationForm with identical functionality
        /// </summary>
        /// <param name="tabName">The name of the tab to delete</param>
        /// <param name="itemCount">The number of items in the tab</param>
        /// <returns>True if Yes was clicked, False if No was clicked</returns>
        public static bool ShowTabDeleteConfirmationForm(string tabName, int itemCount)
        {
            try
            {
                var messageBox = new TabDeleteConfirmationWindow(tabName, itemCount);
                messageBox.ShowDialog();
                return messageBox.DialogResult;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error showing tab delete confirmation: {ex.Message}");
                return false;
            }
        }
        #endregion
    }
#endregion
}


#endregion


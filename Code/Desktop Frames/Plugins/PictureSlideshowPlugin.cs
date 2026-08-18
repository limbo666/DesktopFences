using Desktop_Frames.Localization;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace Desktop_Frames.Plugins
{
    public class PictureSlideshowPlugin : IFramePlugin
    {
        public string PluginId => "PictureSlideshow";
        public string DisplayName => "Photo Frame";

        public int DevelopmentState => 1; // Set to 1, 2, or 3 based on your testing phase

        // --- DUAL IMAGE ENGINE FOR TRANSITIONS ---
        private Grid _containerGrid;
        private Image _img1;
        private Image _img2;
        private bool _useImg1 = true;

        private DispatcherTimer _timer;
        private List<string> _imagePaths = new List<string>();
        private int _currentIndex = -1;
        private bool _isPaused = false;
        private string _currentTransition = "Fade";

        // --- NEW: Rescan Tracking ---
        private bool _autoRescan = false;
        private string _currentPath = "";

        public FrameworkElement CreateVisualElement()
        {
            _containerGrid = new Grid
            {
                // YOUR CUSTOM FRAME MARGINS
                Margin = new Thickness(0, 8, 16, 16)
            };

            _img1 = new Image { Stretch = Stretch.UniformToFill, Opacity = 1.0, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };
            _img2 = new Image { Stretch = Stretch.UniformToFill, Opacity = 0.0, HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center };

            _containerGrid.Children.Add(_img1);
            _containerGrid.Children.Add(_img2);

            return _containerGrid;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            _containerGrid = visual as Grid;
            if (_containerGrid == null) return;

            _timer = new DispatcherTimer();
            _timer.Tick += Timer_Tick;

            ApplySettings(settings);
        }

        public void Pause()
        {
            _isPaused = true;
            _timer?.Stop();
        }

        public void Resume()
        {
            _isPaused = false;
            if (_imagePaths.Count > 1) _timer?.Start();
        }

        public void Cleanup()
        {
            _timer?.Stop();
            if (_img1 != null) _img1.Source = null;
            if (_img2 != null) _img2.Source = null;
            _imagePaths.Clear();
        }

        private void ApplySettings(Dictionary<string, object> settings)
        {
            if (settings == null) return;

            string path = settings.ContainsKey("Path") ? settings["Path"]?.ToString() : string.Empty;
            int intervalSeconds = 30;

            if (settings.ContainsKey("IntervalSeconds") && int.TryParse(settings["IntervalSeconds"]?.ToString(), out int parsed))
            {
                intervalSeconds = Math.Max(1, parsed);
            }

            _timer.Interval = TimeSpan.FromSeconds(intervalSeconds);

            // Apply Stretch Mode
            string stretchMode = settings.ContainsKey("StretchMode") ? settings["StretchMode"]?.ToString() : "UniformToFill";
            if (Enum.TryParse(stretchMode, out Stretch stretch))
            {
                _img1.Stretch = stretch;
                _img2.Stretch = stretch;
            }

            // Apply Transition
            _currentTransition = settings.ContainsKey("Transition") ? settings["Transition"]?.ToString() : "Fade";

            // Apply Rescan Setting
            _autoRescan = settings.ContainsKey("AutoRescan") && settings["AutoRescan"]?.ToString().ToLower() == "true";

            LoadImagesFromPath(path);
        }

        public void ShowSettingsWindow(Window ownerWindow, dynamic frameData)
        {
            Dictionary<string, object> settings = new Dictionary<string, object>();
            try
            {
                if (frameData is JObject jFrame && jFrame["PluginSettings"] is JObject settingsObj)
                    settings = settingsObj.ToObject<Dictionary<string, object>>();
                else if (frameData.PluginSettings != null)
                {
                    var dict = frameData.PluginSettings as IDictionary<string, object>;
                    if (dict != null) settings = new Dictionary<string, object>(dict);
                }
            }
            catch { }

            string currentPath = settings.ContainsKey("Path") ? settings["Path"]?.ToString() : "";
            string currentInterval = settings.ContainsKey("IntervalSeconds") ? settings["IntervalSeconds"]?.ToString() : "30";
            string currentStretch = settings.ContainsKey("StretchMode") ? settings["StretchMode"]?.ToString() : "UniformToFill";
            string currentTransition = settings.ContainsKey("Transition") ? settings["Transition"]?.ToString() : "Fade";
            bool currentRescan = settings.ContainsKey("AutoRescan") && settings["AutoRescan"]?.ToString().ToLower() == "true";

            // --- 1. Get Dynamic Accent Color ---
            SolidColorBrush accentBrush;
            try
            {
                var mediaColor = Utility.GetColorFromName(SettingsManager.SelectedColor);
                accentBrush = new SolidColorBrush(mediaColor);
            }
            catch
            {
                accentBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244)); // Fallback Blue
            }

            // --- 2. Build Modern Window Shell ---
            Window win = new Window
            {
                Title = Strings.PhotoSettings,
                Width = 540,
                SizeToContent = SizeToContent.Height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = ownerWindow,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.None, // Custom Titlebar
                AllowsTransparency = false,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250))
            };

            // Enable Dragging
            win.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) win.DragMove(); };

            Grid mainContainer = new Grid { Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)), Margin = new Thickness(8) };
            Border mainBorder = new Border { Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(0) };

            Grid rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Footer

            // --- 3. Build Header ---
            Border headerBorder = new Border { Height = 50, Background = accentBrush, CornerRadius = new CornerRadius(0) };
            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            headerGrid.Children.Add(new TextBlock { Text = Strings.PhotoSettings, FontSize = 14, FontWeight = FontWeights.Bold, Foreground = Brushes.White, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(16, 0, 0, 0) });

            Button closeBtn = new Button { Content = "✕", Width = 32, Height = 32, FontSize = 12, FontWeight = FontWeights.Bold, Foreground = Brushes.White, Background = Brushes.Transparent, BorderBrush = Brushes.Transparent, Cursor = Cursors.Hand, Margin = new Thickness(0, 0, 9, 0), VerticalAlignment = VerticalAlignment.Center };
            closeBtn.MouseEnter += (s, e) => closeBtn.Background = new SolidColorBrush(Color.FromArgb(50, 255, 255, 255));
            closeBtn.MouseLeave += (s, e) => closeBtn.Background = Brushes.Transparent;
            closeBtn.Click += (s, e) => win.Close();
            Grid.SetColumn(closeBtn, 1);
            headerGrid.Children.Add(closeBtn);

            headerBorder.Child = headerGrid;
            Grid.SetRow(headerBorder, 0);
            rootGrid.Children.Add(headerBorder);

            // --- 4. Build Content (Grouped Fields) ---
            Border contentBorder = new Border { Background = Brushes.White, Padding = new Thickness(20, 10, 20, 10) };
            StackPanel contentPanel = new StackPanel { Orientation = Orientation.Vertical };

            // Helper function to create grouped fields
            Border CreateGroupField(string labelText, UIElement control)
            {
                Border fieldBorder = new Border { Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Margin = new Thickness(0, 0, 0, 12), Padding = new Thickness(12) };
                StackPanel sp = new StackPanel();
                sp.Children.Add(new TextBlock { Text = labelText, FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 0, 0, 8) });
                sp.Children.Add(control);
                fieldBorder.Child = sp;
                return fieldBorder;
            }

            // Path Group
            Grid pathGrid = new Grid();
            pathGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            pathGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            TextBox txtPath = new TextBox { Text = currentPath, FontSize = 13, Padding = new Thickness(8, 6, 8, 6), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Margin = new Thickness(0, 0, 10, 0) };
            Button btnBrowse = new Button { Content = Strings.BtnBrowse, Height = 32, MinWidth = 80, FontSize = 13, Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Cursor = Cursors.Hand };
            btnBrowse.Click += (s, e) => { using (var dialog = new System.Windows.Forms.FolderBrowserDialog()) { if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) txtPath.Text = dialog.SelectedPath; } };
            Grid.SetColumn(txtPath, 0); Grid.SetColumn(btnBrowse, 1);
            pathGrid.Children.Add(txtPath); pathGrid.Children.Add(btnBrowse);
            contentPanel.Children.Add(CreateGroupField("Picture Folder Path:", pathGrid));

            // Interval Group
            TextBox txtInterval = new TextBox { Text = currentInterval, FontSize = 13, Padding = new Thickness(8, 6, 8, 6), Width = 100, HorizontalAlignment = HorizontalAlignment.Left, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)) };
            contentPanel.Children.Add(CreateGroupField("Slideshow Interval (Seconds):", txtInterval));

            // Fit & Transition Group
            Grid comboGrid = new Grid();
            comboGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            comboGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) }); // Spacer
            comboGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            StackPanel stretchPanel = new StackPanel();
            stretchPanel.Children.Add(new TextBlock { Text = Strings.PhotoFitMode, FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 0, 0, 8) });
            ComboBox cmbStretch = new ComboBox { FontSize = 13, Height = 30, Padding = new Thickness(8, 4, 8, 4) };
            cmbStretch.Items.Add(new ComboBoxItem { Content = Strings.PhotoCropToFill, Tag = "UniformToFill" });
            cmbStretch.Items.Add(new ComboBoxItem { Content = Strings.PhotoFitInside, Tag = "Uniform" });
            cmbStretch.Items.Add(new ComboBoxItem { Content = Strings.PhotoStretch, Tag = "Fill" });
            cmbStretch.Items.Add(new ComboBoxItem { Content = Strings.PhotoOriginalSize, Tag = "None" });
            foreach (ComboBoxItem item in cmbStretch.Items) { if (item.Tag.ToString() == currentStretch) item.IsSelected = true; }
            stretchPanel.Children.Add(cmbStretch);
            Grid.SetColumn(stretchPanel, 0);

            StackPanel transPanel = new StackPanel();
            transPanel.Children.Add(new TextBlock { Text = Strings.PhotoTransition, FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 0, 0, 8) });
            ComboBox cmbTransition = new ComboBox { FontSize = 13, Height = 30, Padding = new Thickness(8, 4, 8, 4) };
            cmbTransition.Items.Add(new ComboBoxItem { Content = Strings.PhotoCrossfade, Tag = "Fade" });
            cmbTransition.Items.Add(new ComboBoxItem { Content = Strings.PhotoBlurCrossfade, Tag = "Blur" });
            cmbTransition.Items.Add(new ComboBoxItem { Content = Strings.PhotoVerticalWipe, Tag = "Wipe" });
            cmbTransition.Items.Add(new ComboBoxItem { Content = Strings.PhotoSubtleTwist, Tag = "Twist" });
            cmbTransition.Items.Add(new ComboBoxItem { Content = Strings.PhotoNoTransition, Tag = "None" });
            foreach (ComboBoxItem item in cmbTransition.Items) { if (item.Tag.ToString() == currentTransition) item.IsSelected = true; }
            transPanel.Children.Add(cmbTransition);
            Grid.SetColumn(transPanel, 2);

            comboGrid.Children.Add(stretchPanel);
            comboGrid.Children.Add(transPanel);

            Border displayBorder = new Border { Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Margin = new Thickness(0, 0, 0, 12), Padding = new Thickness(12) };
            displayBorder.Child = comboGrid;
            contentPanel.Children.Add(displayBorder);

            // Rescan Checkbox
            CheckBox chkRescan = new CheckBox { Content = Strings.PhotoLiveRescan, IsChecked = currentRescan, FontSize = 13, Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)), Margin = new Thickness(2, 5, 0, 5) };
            contentPanel.Children.Add(chkRescan);

            contentBorder.Child = contentPanel;
            Grid.SetRow(contentBorder, 1);
            rootGrid.Children.Add(contentBorder);

            // --- 5. Build Footer ---
            Border footerBorder = new Border { Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(0, 1, 0, 0), Padding = new Thickness(20, 16, 20, 16) };
            StackPanel buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            Button cancelBtn = new Button { Content = Strings.BtnCancel, Height = 36, MinWidth = 80, FontSize = 13, FontWeight = FontWeights.Medium, Background = Brushes.White, BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Cursor = Cursors.Hand, Margin = new Thickness(0, 0, 10, 0) };
            cancelBtn.Click += (s, e) => win.Close();

            Button saveBtn = new Button { Content = Strings.BtnSave, Height = 36, MinWidth = 80, FontSize = 13, FontWeight = FontWeights.Bold, Background = accentBrush, Foreground = Brushes.White, BorderThickness = new Thickness(0), Cursor = Cursors.Hand, Padding = new Thickness(16, 0, 16, 0) };
            saveBtn.Click += (s, e) =>
            {
                settings["Path"] = txtPath.Text;
                settings["IntervalSeconds"] = txtInterval.Text;
                if (cmbStretch.SelectedItem is ComboBoxItem selectedStretch) settings["StretchMode"] = selectedStretch.Tag.ToString();
                if (cmbTransition.SelectedItem is ComboBoxItem selectedTrans) settings["Transition"] = selectedTrans.Tag.ToString();
                settings["AutoRescan"] = chkRescan.IsChecked == true ? "true" : "false";

                if (frameData is JObject jFrame) jFrame["PluginSettings"] = JObject.FromObject(settings);
                else { var frameDict = (IDictionary<string, object>)frameData; frameDict["PluginSettings"] = settings; }

                FrameDataManager.SaveFrameData();
                ApplySettings(settings);
                win.Close();
            };

            buttonPanel.Children.Add(cancelBtn);
            buttonPanel.Children.Add(saveBtn);
            footerBorder.Child = buttonPanel;
            Grid.SetRow(footerBorder, 2);
            rootGrid.Children.Add(footerBorder);

            // Assemble Window
            mainBorder.Child = rootGrid;
            mainContainer.Children.Add(mainBorder);
            win.Content = mainContainer;

            // Handle Enter/Escape
            win.KeyDown += (s, e) => {
                if (e.Key == Key.Escape) win.Close();
                else if (e.Key == Key.Enter) saveBtn.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            };

            win.ShowDialog();
        }

        private void LoadImagesFromPath(string path)
        {
            _timer.Stop();
            _imagePaths.Clear();
            _currentIndex = -1;
            _img1.Source = null;
            _img2.Source = null;
            _currentPath = path;

            if (string.IsNullOrEmpty(path)) return;

            try
            {
                if (File.Exists(path) && IsImageFile(path)) _imagePaths.Add(path);
                else if (Directory.Exists(path))
                {
                    _imagePaths.AddRange(Directory.EnumerateFiles(path, "*.*", SearchOption.TopDirectoryOnly).Where(s => IsImageFile(s)));
                }

                if (_imagePaths.Count > 0)
                {
                    AdvanceImage();
                    if (_imagePaths.Count > 1 && !_isPaused) _timer.Start();
                }
            }
            catch { }
        }

        private void Timer_Tick(object sender, EventArgs e) => AdvanceImage();

        private void AdvanceImage()
        {
            // --- NEW: Live Rescan Logic ---
            if (_autoRescan && !string.IsNullOrEmpty(_currentPath) && Directory.Exists(_currentPath))
            {
                try
                {
                    var freshPaths = Directory.EnumerateFiles(_currentPath, "*.*", SearchOption.TopDirectoryOnly)
                                              .Where(s => IsImageFile(s)).ToList();
                    _imagePaths = freshPaths;

                    // Failsafe in case pictures were deleted and our index is now out of bounds
                    if (_currentIndex >= _imagePaths.Count) _currentIndex = -1;
                }
                catch { } // Fail silently if folder is locked or disconnected
            }

            if (_imagePaths.Count == 0) return;
            _currentIndex = (_currentIndex + 1) % _imagePaths.Count;
            DisplayImage(_imagePaths[_currentIndex]);
        }

        private void DisplayImage(string filePath)
        {
            try
            {
                BitmapImage bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.UriSource = new Uri(filePath, UriKind.Absolute);
                bitmap.EndInit();
                bitmap.Freeze();

                Image activeImg = _useImg1 ? _img1 : _img2;
                Image inactiveImg = _useImg1 ? _img2 : _img1;

                // --- CRITICAL FIX: Z-Index and State Reset ---
                // Forces the new image to the top
                Panel.SetZIndex(activeImg, 1);
                Panel.SetZIndex(inactiveImg, 0);

                // Completely wipe all previous animations to prevent WPF state locking
                activeImg.BeginAnimation(UIElement.OpacityProperty, null);
                inactiveImg.BeginAnimation(UIElement.OpacityProperty, null);

                // Reset New Image (activeImg)
                activeImg.RenderTransform = new TransformGroup();
                activeImg.RenderTransformOrigin = new Point(0.5, 0.5);
                activeImg.Effect = null;
                activeImg.OpacityMask = null;
                activeImg.Opacity = 1.0;
                activeImg.Source = bitmap;

                // Reset Old Image (inactiveImg)
                inactiveImg.RenderTransform = new TransformGroup();
                inactiveImg.RenderTransformOrigin = new Point(0.5, 0.5);
                inactiveImg.Effect = null;
                inactiveImg.OpacityMask = null;
                inactiveImg.Opacity = 1.0; // MUST be fully visible before transition starts

                TimeSpan duration = TimeSpan.FromSeconds(0.8);
                var ease = new QuadraticEase { EasingMode = EasingMode.EaseInOut };

                switch (_currentTransition)
                {
                    case "Fade":
                        activeImg.Opacity = 0.0;
                        activeImg.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(0.0, 1.0, duration));
                        inactiveImg.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(1.0, 0.0, duration));
                        break;

                    case "Blur":
                        var blur = new System.Windows.Media.Effects.BlurEffect { Radius = 0 };
                        inactiveImg.Effect = blur;

                        activeImg.Opacity = 0.0;
                        activeImg.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(0.0, 1.0, duration));

                        blur.BeginAnimation(System.Windows.Media.Effects.BlurEffect.RadiusProperty, new DoubleAnimation(0, 30, duration) { EasingFunction = ease });
                        inactiveImg.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(1.0, 0.0, duration));
                        break;

                    case "Wipe":
                        // Create a perfect 2-point soft feathered gradient mask
                        var mask = new LinearGradientBrush { StartPoint = new Point(0, 0), EndPoint = new Point(0, 1) };
                        var stop1 = new GradientStop(Colors.Black, -0.2); // Starts above the image
                        var stop2 = new GradientStop(Colors.Transparent, 0.0);
                        mask.GradientStops.Add(stop1);
                        mask.GradientStops.Add(stop2);

                        activeImg.OpacityMask = mask;

                        // Sweep the mask down
                        stop1.BeginAnimation(GradientStop.OffsetProperty, new DoubleAnimation(-0.2, 1.0, duration));
                        stop2.BeginAnimation(GradientStop.OffsetProperty, new DoubleAnimation(0.0, 1.2, duration));

                        // Delayed fade-out of the old image to prevent aspect-ratio residues sticking out
                        var delayedFade = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(0.3)) { BeginTime = TimeSpan.FromSeconds(0.5) };
                        inactiveImg.BeginAnimation(UIElement.OpacityProperty, delayedFade);
                        break;

                    case "Twist":
                        var group = new TransformGroup();
                        var scale = new ScaleTransform(1.03, 1.03); // Slight zoom hides edges during tilt
                        var rotate = new RotateTransform(-2.0);
                        group.Children.Add(scale);
                        group.Children.Add(rotate);

                        activeImg.RenderTransform = group;
                        activeImg.Opacity = 0.0;

                        activeImg.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(0.0, 1.0, duration));
                        rotate.BeginAnimation(RotateTransform.AngleProperty, new DoubleAnimation(-2.0, 0.0, duration) { EasingFunction = ease });
                        scale.BeginAnimation(ScaleTransform.ScaleXProperty, new DoubleAnimation(1.03, 1.0, duration) { EasingFunction = ease });
                        scale.BeginAnimation(ScaleTransform.ScaleYProperty, new DoubleAnimation(1.03, 1.0, duration) { EasingFunction = ease });

                        inactiveImg.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(1.0, 0.0, duration));
                        break;

                    default: // None
                        activeImg.Opacity = 1.0;
                        inactiveImg.Opacity = 0.0;
                        break;
                }

                _useImg1 = !_useImg1; // Swap for next tick
            }
            catch { }
        }

        private bool IsImageFile(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            return ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp" || ext == ".webp";
        }
    }
}
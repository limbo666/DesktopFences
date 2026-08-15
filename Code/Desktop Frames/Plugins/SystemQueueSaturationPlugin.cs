using Desktop_Frames.Localization;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Desktop_Frames.Plugins
{
    public class SystemQueueSaturationPlugin : IFramePlugin
    {
        public string PluginId => "SystemQueueSaturation";
        public string DisplayName => "System Queue Saturation";
        public int DevelopmentState => 1;

        // UI Components
        private Grid _rootGrid;
        private Border _mainCard;
        private TextBlock _titleText;
        private TextBlock _coresText;
        private TextBlock _cpuValuesText;
        private TextBlock _diskValuesText;

        // Custom Progress Bars
        private ColumnDefinition _cpuFillColumn;
        private ColumnDefinition _cpuEmptyColumn;
        private Border _cpuFillBorder;

        private ColumnDefinition _diskFillColumn;
        private ColumnDefinition _diskEmptyColumn;
        private Border _diskFillBorder;

        // Hardware & Math Variables
        private PerformanceCounter _cpuCounter;
        private PerformanceCounter _diskCounter;
        private DispatcherTimer _timer;
        private int _logicalCores;
        private double _cpuThreshold;
        private double _diskThreshold = 5.0;
        private bool _countersReady = false;

        // Settings State
        private bool _showTitle = true;
        private bool _showNumerics = true;
        private bool _showCores = true;
        private bool _useDynamicColors = true;
        private string _staticColor = "DeepSkyBlue";
        private int _refreshRateMs = 1000;

        // Brushes
        private readonly SolidColorBrush _colorGreen = new SolidColorBrush(Color.FromRgb(50, 205, 50));

        private SolidColorBrush GetStaticBrush()
        {
            try { return new SolidColorBrush((Color)ColorConverter.ConvertFromString(_staticColor)); }
            catch { return new SolidColorBrush(Color.FromRgb(66, 133, 244)); }
        }
        private readonly SolidColorBrush _colorYellow = new SolidColorBrush(Color.FromRgb(255, 215, 0));
        private readonly SolidColorBrush _colorRed = new SolidColorBrush(Color.FromRgb(255, 69, 0));
        private readonly SolidColorBrush _colorGray = new SolidColorBrush(Color.FromRgb(60, 60, 60));

        public FrameworkElement CreateVisualElement()
        {
            _logicalCores = Environment.ProcessorCount;
            _cpuThreshold = _logicalCores * 2;

            _rootGrid = new Grid
            {
                // Master UI Spacing Blueprint
                Margin = new Thickness(5, 5, 25, 10),
                MinWidth = 240
            };

            _mainCard = new Border
            {
                Background = Brushes.Transparent, // Matches the frame's background
                BorderThickness = new Thickness(0), // Removed the inner border
                // Internal Card Blueprint
                Margin = new Thickness(10, 5, 10, 15),
                Padding = new Thickness(10, 5, 10, 5)
            };

            StackPanel panel = new StackPanel();

            // Header Section
            _titleText = new TextBlock
            {
                Text = Strings.QueueTitle,
                Foreground = Brushes.White,
                FontSize = 11,
                FontWeight = FontWeights.Bold,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 0, 0, 10)
            };
            panel.Children.Add(_titleText);

            // CPU Block (Now includes the tiny logical cores subtitle)
            panel.Children.Add(CreateDiagnosticRow("CPU QUEUE", $"Logical cores: {_logicalCores}", out _cpuValuesText, out _coresText, out _cpuFillColumn, out _cpuEmptyColumn, out _cpuFillBorder));

            // Spacer
            panel.Children.Add(new Border { Height = 1, Background = new SolidColorBrush(Color.FromArgb(30, 255, 255, 255)), Margin = new Thickness(0, 10, 0, 10) });

            // Disk Block
            panel.Children.Add(CreateDiagnosticRow("DISK I/O", null, out _diskValuesText, out _, out _diskFillColumn, out _diskEmptyColumn, out _diskFillBorder));
            _mainCard.Child = panel;
            _rootGrid.Children.Add(_mainCard);

            // Set initial wait indicator while the background thread loads
            if (_cpuValuesText != null) _cpuValuesText.Text = Strings.QueueInitializing;
            if (_diskValuesText != null) _diskValuesText.Text = Strings.QueueInitializing;

            _timer = new DispatcherTimer();
            _timer.Tick += Timer_Tick;

            return _rootGrid;
        }

        private Grid CreateDiagnosticRow(string label, string subLabel, out TextBlock valuesText, out TextBlock subLabelText, out ColumnDefinition fillCol, out ColumnDefinition emptyCol, out Border fillBorder)
        {
            Grid row = new Grid();
            row.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            row.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Label & Numerics
            Grid textGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };

            StackPanel labelsPanel = new StackPanel { HorizontalAlignment = HorizontalAlignment.Left };
            labelsPanel.Children.Add(new TextBlock { Text = label, Foreground = Brushes.LightGray, FontSize = 12, FontWeight = FontWeights.SemiBold });

            subLabelText = null;
            if (!string.IsNullOrEmpty(subLabel))
            {
                subLabelText = new TextBlock { Text = subLabel, Foreground = Brushes.White, FontSize = 9, Margin = new Thickness(0, 2, 0, 0) };
                labelsPanel.Children.Add(subLabelText);
            }
            textGrid.Children.Add(labelsPanel);

            valuesText = new TextBlock
            {
                Text = "0.00",
                Foreground = Brushes.White,
                FontSize = 11,
                FontFamily = new FontFamily("Lucida Console"),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top
            };
            textGrid.Children.Add(valuesText);

            Grid.SetRow(textGrid, 0);
            row.Children.Add(textGrid);

            // Custom Flat Progress Bar (Thicker)
            Border barContainer = new Border
            {
                Background = _colorGray,
                Height = 14, // Made the bar significantly thicker
                CornerRadius = new CornerRadius(3)
            };

            Grid barGrid = new Grid();
            fillCol = new ColumnDefinition { Width = new GridLength(0, GridUnitType.Star) };
            emptyCol = new ColumnDefinition { Width = new GridLength(100, GridUnitType.Star) };
            barGrid.ColumnDefinitions.Add(fillCol);
            barGrid.ColumnDefinitions.Add(emptyCol);

            fillBorder = new Border
            {
                Background = _colorGreen,
                CornerRadius = new CornerRadius(3, 0, 0, 3)
            };
            Grid.SetColumn(fillBorder, 0);
            barGrid.Children.Add(fillBorder);

            barContainer.Child = barGrid;
            Grid.SetRow(barContainer, 1);
            row.Children.Add(barContainer);

            return row;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            ApplySettingsData(settings);
            ApplyLayoutSettings();

            _timer.Interval = TimeSpan.FromMilliseconds(_refreshRateMs);
            _timer.Start();

            // Offload the heavy WMI registry queries to a background thread
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _cpuCounter = new PerformanceCounter("System", "Processor Queue Length");
                    _diskCounter = new PerformanceCounter("PhysicalDisk", "Avg. Disk Queue Length", "_Total");

                    // Dummy read to wake up the sensor (the first read is natively always 0)
                    _cpuCounter.NextValue();
                    _diskCounter.NextValue();

                    _countersReady = true;
                }
                catch
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (_cpuValuesText != null) _cpuValuesText.Text = Strings.QueueWmiError;
                        if (_diskValuesText != null) _diskValuesText.Text = Strings.QueueWmiError;
                    });
                }
            });
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            // Skip updates until the background thread finishes loading the counters
            if (!_countersReady || _cpuCounter == null || _diskCounter == null) return;

            try
            {
                float cpuData = _cpuCounter.NextValue();
                float diskData = _diskCounter.NextValue();

                // CPU Math
                double cpuRatio = cpuData / _logicalCores;
                double cpuPercent = Math.Min((cpuData / _cpuThreshold) * 100, 100);

                // Disk Math
                double diskPercent = Math.Min((diskData / _diskThreshold) * 100, 100);

                // CPU Color Logic
                SolidColorBrush cpuBrush = GetStaticBrush();
                if (_useDynamicColors)
                {
                    cpuBrush = _colorRed;
                    if (cpuRatio < 1) cpuBrush = _colorGreen;
                    else if (cpuRatio < 2) cpuBrush = _colorYellow;
                }

                // Disk Color Logic
                SolidColorBrush diskBrush = GetStaticBrush();
                if (_useDynamicColors)
                {
                    diskBrush = _colorRed;
                    if (diskData < 1) diskBrush = _colorGreen;
                    else if (diskData < 3) diskBrush = _colorYellow;
                }

                // Animate Data
                UpdateBar(_cpuFillColumn, _cpuEmptyColumn, _cpuFillBorder, cpuPercent, cpuBrush);
                UpdateBar(_diskFillColumn, _diskEmptyColumn, _diskFillBorder, diskPercent, diskBrush);

                if (_showNumerics)
                {
                    _cpuValuesText.Text = $"[{cpuRatio:0.00} Threads/Core]";
                    _diskValuesText.Text = $"[{diskData:0.00} Pending I/O]";

                    _cpuValuesText.Foreground = cpuBrush;
                    _diskValuesText.Foreground = diskBrush;
                }
            }
            catch
            {
                // Handle potential transient WMI/PerfCounter drops silently
            }
        }

        private void UpdateBar(ColumnDefinition fill, ColumnDefinition empty, Border fillBorder, double percent, SolidColorBrush targetBrush)
        {
            if (double.IsNaN(percent) || percent < 0) percent = 0;
            if (percent > 100) percent = 100;

            fill.Width = new GridLength(percent, GridUnitType.Star);
            empty.Width = new GridLength(100 - percent, GridUnitType.Star);

            // Adjust corner radius dynamically if full
            fillBorder.CornerRadius = percent >= 99.9 ? new CornerRadius(2) : new CornerRadius(2, 0, 0, 2);
            fillBorder.Background = targetBrush;
        }

        private void ApplySettingsData(Dictionary<string, object> settings)
        {
            if (settings == null) return;
            if (settings.ContainsKey("ShowTitle") && bool.TryParse(settings["ShowTitle"].ToString(), out bool st)) _showTitle = st;
            if (settings.ContainsKey("ShowNumerics") && bool.TryParse(settings["ShowNumerics"].ToString(), out bool sn)) _showNumerics = sn;
            if (settings.ContainsKey("ShowCores") && bool.TryParse(settings["ShowCores"].ToString(), out bool sc)) _showCores = sc;
            if (settings.ContainsKey("UseDynamicColors") && bool.TryParse(settings["UseDynamicColors"].ToString(), out bool udc)) _useDynamicColors = udc;
            if (settings.ContainsKey("StaticColor") && settings["StaticColor"] != null) _staticColor = settings["StaticColor"].ToString();
            if (settings.ContainsKey("RefreshRateMs") && int.TryParse(settings["RefreshRateMs"].ToString(), out int rr)) _refreshRateMs = rr;
        }

        private void ApplyLayoutSettings()
        {
            if (_titleText != null) _titleText.Visibility = _showTitle ? Visibility.Visible : Visibility.Collapsed;
            if (_cpuValuesText != null) _cpuValuesText.Visibility = _showNumerics ? Visibility.Visible : Visibility.Collapsed;
            if (_diskValuesText != null) _diskValuesText.Visibility = _showNumerics ? Visibility.Visible : Visibility.Collapsed;
            if (_coresText != null) _coresText.Visibility = _showCores ? Visibility.Visible : Visibility.Collapsed;
        }

        public void Pause() => _timer?.Stop();

        public void Resume() => _timer?.Start();

        public void Cleanup()
        {
            _timer?.Stop();
            _cpuCounter?.Dispose();
            _diskCounter?.Dispose();
        }

        public void ShowSettingsWindow(Window ownerWindow, dynamic frameData)
        {
            SolidColorBrush accentBrush;
            try { accentBrush = new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor)); }
            catch { accentBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244)); }

            Window win = new Window
            {
                Title = Strings.QueueConfig,
                Owner = ownerWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowStyle = WindowStyle.None,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                SizeToContent = SizeToContent.Height,
                Width = 400
            };

            Border headerBorder = new Border { Height = 50, Background = accentBrush };
            headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) win.DragMove(); };
            Grid headerGrid = new Grid();
            headerGrid.Children.Add(new TextBlock { Text = Strings.QueueDiagnosticSettings, Foreground = Brushes.White, FontWeight = FontWeights.Bold, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(15, 0, 0, 0) });
            Button btnClose = new Button { Content = "X", Foreground = Brushes.White, Background = Brushes.Transparent, BorderThickness = new Thickness(0), Width = 40, HorizontalAlignment = HorizontalAlignment.Right };
            btnClose.Click += (s, e) => win.Close();
            headerGrid.Children.Add(btnClose);
            headerBorder.Child = headerGrid;

            Border contentBorder = new Border { Background = Brushes.White, Padding = new Thickness(20) };
            StackPanel contentPanel = new StackPanel();

            CheckBox chkTitle = new CheckBox { Content = Strings.QueueShowTitle, IsChecked = _showTitle, Margin = new Thickness(0, 0, 0, 12), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkTitle);

            CheckBox chkNumerics = new CheckBox { Content = Strings.QueueShowValues, IsChecked = _showNumerics, Margin = new Thickness(0, 0, 0, 12), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkNumerics);

            CheckBox chkCores = new CheckBox { Content = Strings.QueueShowCores, IsChecked = _showCores, Margin = new Thickness(0, 0, 0, 12), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkCores);

            CheckBox chkDynamic = new CheckBox { Content = Strings.QueueDynamicColors, IsChecked = _useDynamicColors, Margin = new Thickness(0, 0, 0, 10), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkDynamic);

            StackPanel colorSp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 15) };
            colorSp.Children.Add(new TextBlock { Text = Strings.PerfStaticBarColor, Width = 150, VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeights.SemiBold });
            ComboBox cmbStaticColor = new ComboBox { Width = 120 };
            string[] presetColors = { "DeepSkyBlue", "Cyan", "LimeGreen", "Gold", "Orange", "Magenta", "White", "LightGray" };
            foreach (var c in presetColors) cmbStaticColor.Items.Add(c);
            cmbStaticColor.SelectedItem = presetColors.Contains(_staticColor) ? _staticColor : "DeepSkyBlue";
            colorSp.Children.Add(cmbStaticColor);
            contentPanel.Children.Add(colorSp);

            StackPanel rateSp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 10) };
            rateSp.Children.Add(new TextBlock { Text = Strings.QueueRefreshRate, Width = 150, VerticalAlignment = VerticalAlignment.Center, FontWeight = FontWeights.SemiBold });
            ComboBox cmbRate = new ComboBox { Width = 120 };
            cmbRate.Items.Add("500 ms");
            cmbRate.Items.Add("1000 ms");
            cmbRate.Items.Add("2000 ms");
            cmbRate.Items.Add("5000 ms");
            cmbRate.SelectedItem = $"{_refreshRateMs} ms";
            if (cmbRate.SelectedItem == null) cmbRate.SelectedItem = "1000 ms";
            rateSp.Children.Add(cmbRate);
            contentPanel.Children.Add(rateSp);

            contentBorder.Child = contentPanel;

            Border footerBorder = new Border { Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)), BorderThickness = new Thickness(0, 1, 0, 0), BorderBrush = Brushes.LightGray, Padding = new Thickness(15) };
            StackPanel footerSp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            Button btnCancel = new Button { Content = Strings.BtnCancel, Background = Brushes.White, BorderBrush = Brushes.Gray, Width = 80, Height = 30, Margin = new Thickness(0, 0, 10, 0) };
            btnCancel.Click += (s, e) => win.Close();
            Button btnSave = new Button { Content = Strings.BtnSave, Background = accentBrush, Foreground = Brushes.White, FontWeight = FontWeights.Bold, BorderThickness = new Thickness(0), Width = 80, Height = 30 };
            btnSave.Click += (s, e) =>
            {
                int newRate = 1000;
                if (cmbRate.SelectedItem != null)
                {
                    string rStr = cmbRate.SelectedItem.ToString().Replace(" ms", "");
                    int.TryParse(rStr, out newRate);
                }

                Dictionary<string, object> newSettings = new Dictionary<string, object> {
                    { "ShowTitle", chkTitle.IsChecked == true },
                    { "ShowNumerics", chkNumerics.IsChecked == true },
                    { "ShowCores", chkCores.IsChecked == true },
                    { "UseDynamicColors", chkDynamic.IsChecked == true },
                    { "StaticColor", cmbStaticColor.SelectedItem?.ToString() ?? "DeepSkyBlue" },
                    { "RefreshRateMs", newRate }
                };

                // Secure Host Save Pipeline Execution
                if (frameData is Newtonsoft.Json.Linq.JObject jFrame) jFrame["PluginSettings"] = Newtonsoft.Json.Linq.JObject.FromObject(newSettings);
                else ((IDictionary<string, object>)frameData)["PluginSettings"] = newSettings;
                try { FrameDataManager.SaveFrameData(); } catch { }

                ApplySettingsData(newSettings);
                ApplyLayoutSettings();

                if (_timer != null)
                {
                    _timer.Interval = TimeSpan.FromMilliseconds(_refreshRateMs);
                }

                win.Close();
            };
            footerSp.Children.Add(btnCancel); footerSp.Children.Add(btnSave); footerBorder.Child = footerSp;

            DockPanel rootPanel = new DockPanel();
            DockPanel.SetDock(headerBorder, Dock.Top); DockPanel.SetDock(footerBorder, Dock.Bottom);
            rootPanel.Children.Add(headerBorder); rootPanel.Children.Add(footerBorder); rootPanel.Children.Add(contentBorder);

            win.Content = rootPanel;
            win.ShowDialog();
        }
    }
}
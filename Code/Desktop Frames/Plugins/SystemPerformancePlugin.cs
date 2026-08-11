using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Desktop_Frames.Plugins
{
    public class SystemPerformancePlugin : IFramePlugin
    {
        public string PluginId => "SysPerfGauge";
        public string DisplayName => "System Performance Gauges";

        public int DevelopmentState => 1; // Set to 1, 2, or 3 based on your testing phase

        private DispatcherTimer _timer;
        private PerformanceCounter _cpuCounter;

        private Grid _layoutGrid;

        private RotateTransform _cpuNeedleTransform;
        private RotateTransform _ramNeedleTransform;
        private TextBlock _cpuValueText;
        private TextBlock _ramValueText;

        private string _layoutMode = "Both";
        private int _refreshRateMs = 1000;

        private double _cpuCurrentAngle = -90;
        private double _ramCurrentAngle = -90;

        private bool _countersReady = false;

        // Bar Theme Settings & Variables
        private string _visualTheme = "Gauges"; // Options: "Gauges" or "Bars"
        private ColumnDefinition _cpuFillColumn, _cpuEmptyColumn, _ramFillColumn, _ramEmptyColumn;
        private Border _cpuFillBorder, _ramFillBorder;
        private TextBlock _cpuBarValueText, _ramBarValueText;

        // Shared Diagnostic Theme Brushes
        private bool _useDynamicColors = true;
        private string _staticColor = "DeepSkyBlue";

        private readonly SolidColorBrush _colorGreen = new SolidColorBrush(Color.FromRgb(50, 205, 50));
        private readonly SolidColorBrush _colorYellow = new SolidColorBrush(Color.FromRgb(255, 215, 0));
        private readonly SolidColorBrush _colorRed = new SolidColorBrush(Color.FromRgb(255, 69, 0));
        private readonly SolidColorBrush _colorGray = new SolidColorBrush(Color.FromRgb(60, 60, 60));

        private SolidColorBrush GetStaticBrush()
        {
            try { return new SolidColorBrush((Color)ColorConverter.ConvertFromString(_staticColor)); }
            catch { return new SolidColorBrush(Color.FromRgb(66, 133, 244)); }
        }
        // Native API for accurate Task Manager RAM Matching
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private class MEMORYSTATUSEX
        {
            public uint dwLength;
            public uint dwMemoryLoad; // This precisely matches Task Manager's Physical RAM %
            public ulong ullTotalPhys;
            public ulong ullAvailPhys;
            public ulong ullTotalPageFile;
            public ulong ullAvailPageFile;
            public ulong ullTotalVirtual;
            public ulong ullAvailVirtual;
            public ulong ullAvailExtendedVirtual;
            public MEMORYSTATUSEX() { this.dwLength = (uint)Marshal.SizeOf(typeof(MEMORYSTATUSEX)); }
        }

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool GlobalMemoryStatusEx([In, Out] MEMORYSTATUSEX lpBuffer);

        public FrameworkElement CreateVisualElement()
        {
            _layoutGrid = new Grid
            {
                Margin = new Thickness(5, 5, 25, 10), // Master UI Spacing Blueprint
                MinWidth = 240 // Matches SystemQueueSaturation limits
            };

            return _layoutGrid;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            ApplySettingsData(settings);
            BuildLayout();

            if (_cpuValueText != null) _cpuValueText.Text = "..."; // Visual wait indicator for CPU

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(_refreshRateMs) };
            _timer.Tick += Timer_Tick;
            _timer.Start();

            // Offload WMI registry query to prevent UI freeze on load
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    // Modern Windows uses Processor Utility for Task Manager parity
                    _cpuCounter = new PerformanceCounter("Processor Information", "% Processor Utility", "_Total");
                    _cpuCounter.NextValue(); // Wake up the sensor
                }
                catch
                {
                    try
                    {
                        // Fallback for older systems
                        _cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
                        _cpuCounter.NextValue();
                    }
                    catch { }
                }

                _countersReady = true;
            });
        }

        private void ApplySettingsData(Dictionary<string, object> settings)
        {
            if (settings != null)
            {
                if (settings.ContainsKey("VisualTheme")) _visualTheme = settings["VisualTheme"].ToString();
                if (settings.ContainsKey("LayoutMode")) _layoutMode = settings["LayoutMode"].ToString();
                if (settings.ContainsKey("UseDynamicColors") && bool.TryParse(settings["UseDynamicColors"].ToString(), out bool udc)) _useDynamicColors = udc;
                if (settings.ContainsKey("StaticColor") && settings["StaticColor"] != null) _staticColor = settings["StaticColor"].ToString();
                if (settings.ContainsKey("RefreshRateMs") && int.TryParse(settings["RefreshRateMs"].ToString(), out int rate)) _refreshRateMs = rate;
            }
        }

        private void BuildLayout()
        {
            _layoutGrid.Children.Clear();
            _layoutGrid.ColumnDefinitions.Clear();
            _layoutGrid.RowDefinitions.Clear();

            _cpuCurrentAngle = -90;
            _ramCurrentAngle = -90;

            bool showCpu = _layoutMode == "CPU" || _layoutMode == "Both";
            bool showRam = _layoutMode == "RAM" || _layoutMode == "Both";

            if (_visualTheme == "Bars")
            {
                // Match the exact container structure of SystemQueueSaturationPlugin
                Border mainCard = new Border
                {
                    Background = Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Margin = new Thickness(10, 5, 10, 15),
                    Padding = new Thickness(10, 5, 10, 5)
                };

                StackPanel panel = new StackPanel();

                if (showCpu)
                {
                    panel.Children.Add(CreateDiagnosticRow("CPU UTILIZATION", out _cpuBarValueText, out _cpuFillColumn, out _cpuEmptyColumn, out _cpuFillBorder));
                }

                if (showCpu && showRam)
                {
                    panel.Children.Add(new Border { Height = 1, Background = new SolidColorBrush(Color.FromArgb(30, 255, 255, 255)), Margin = new Thickness(0, 10, 0, 10) });
                }

                if (showRam)
                {
                    panel.Children.Add(CreateDiagnosticRow("RAM UTILIZATION", out _ramBarValueText, out _ramFillColumn, out _ramEmptyColumn, out _ramFillBorder));
                }

                mainCard.Child = panel;
                _layoutGrid.Children.Add(mainCard);
            }
            else // Default classic Gauges
            {
                // Dynamically inject the Viewbox wrapper only for the Gauges theme
                Viewbox gaugeViewbox = new Viewbox
                {
                    Stretch = Stretch.Uniform,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center
                };

                Grid gaugeContainer = new Grid();

                if (showCpu && showRam)
                {
                    gaugeContainer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    gaugeContainer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    var cpuGauge = CreateGauge("CPU %", out _cpuNeedleTransform, out _cpuValueText);
                    var ramGauge = CreateGauge("RAM %", out _ramNeedleTransform, out _ramValueText);

                    Grid.SetColumn(cpuGauge, 0);
                    Grid.SetColumn(ramGauge, 1);

                    cpuGauge.Margin = new Thickness(0, 0, 10, 0);
                    ramGauge.Margin = new Thickness(10, 0, 0, 0);

                    gaugeContainer.Children.Add(cpuGauge);
                    gaugeContainer.Children.Add(ramGauge);
                }
                else if (showCpu)
                {
                    gaugeContainer.Children.Add(CreateGauge("CPU %", out _cpuNeedleTransform, out _cpuValueText));
                }
                else if (showRam)
                {
                    gaugeContainer.Children.Add(CreateGauge("RAM %", out _ramNeedleTransform, out _ramValueText));
                }

                gaugeViewbox.Child = gaugeContainer;
                _layoutGrid.Children.Add(gaugeViewbox);
            }
        }

        private Canvas CreateGauge(string title, out RotateTransform needleTransform, out TextBlock valueText)
        {
            Canvas canvas = new Canvas { Width = 120, Height = 105 };

            valueText = new TextBlock
            {
                Text = "0",
                Foreground = Brushes.Cyan,
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Width = 40,
                TextAlignment = TextAlignment.Center
            };
            Canvas.SetTop(valueText, 40);
            Canvas.SetLeft(valueText, 40);
            canvas.Children.Add(valueText);

            Path dial = new Path
            {
                Stroke = Brushes.DarkGray,
                StrokeThickness = 4,
                Data = Geometry.Parse("M 10,70 A 50,50 0 0,1 110,70")
            };
            canvas.Children.Add(dial);

            canvas.Children.Add(CreateLabel("0", 5, 75));
            canvas.Children.Add(CreateLabel("50", 52, 0));
            canvas.Children.Add(CreateLabel("100", 95, 75));

            needleTransform = new RotateTransform(-90, 60, 70);
            Line needle = new Line
            {
                X1 = 60,
                Y1 = 70,
                X2 = 60,
                Y2 = 25,
                Stroke = Brushes.Red,
                StrokeThickness = 2,
                RenderTransform = needleTransform
            };
            canvas.Children.Add(needle);

            Ellipse pin = new Ellipse { Width = 8, Height = 8, Fill = Brushes.White };
            Canvas.SetTop(pin, 66);
            Canvas.SetLeft(pin, 56);
            canvas.Children.Add(pin);

            TextBlock titleText = new TextBlock
            {
                Text = title,
                Foreground = Brushes.White,
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Width = 120,
                TextAlignment = TextAlignment.Center
            };
            Canvas.SetTop(titleText, 85);
            Canvas.SetLeft(titleText, 0);
            canvas.Children.Add(titleText);

            return canvas;
        }

        private TextBlock CreateLabel(string text, double left, double top)
        {
            var tb = new TextBlock { Text = text, Foreground = Brushes.LightGray, FontSize = 10 };
            Canvas.SetLeft(tb, left);
            Canvas.SetTop(tb, top);
            return tb;
        }


        private Grid CreateDiagnosticRow(string label, out TextBlock valuesText, out ColumnDefinition fillCol, out ColumnDefinition emptyCol, out Border fillBorder)
        {
            Grid row = new Grid();
            row.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            row.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            Grid textGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };

            textGrid.Children.Add(new TextBlock { Text = label, Foreground = Brushes.LightGray, FontSize = 12, FontWeight = FontWeights.SemiBold, HorizontalAlignment = HorizontalAlignment.Left });

            valuesText = new TextBlock
            {
                Text = "0%",
                Foreground = Brushes.White,
                FontSize = 11,
                FontFamily = new FontFamily("Lucida Console"),
                HorizontalAlignment = HorizontalAlignment.Right
            };
            textGrid.Children.Add(valuesText);

            Grid.SetRow(textGrid, 0);
            row.Children.Add(textGrid);

            // Custom Flat Progress Bar
            Border barContainer = new Border
            {
                Background = _colorGray,
                Height = 14,
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

        private void UpdateDiagnosticBar(ColumnDefinition fill, ColumnDefinition empty, Border fillBorder, double percent)
        {
            if (fill == null || empty == null || fillBorder == null) return;

            if (double.IsNaN(percent) || percent < 0) percent = 0;
            if (percent > 100) percent = 100;

            // Establish color severity thresholds for standard system loads
            SolidColorBrush targetBrush = GetStaticBrush();
            if (_useDynamicColors)
            {
                targetBrush = _colorGreen;
                if (percent >= 85) targetBrush = _colorRed;
                else if (percent >= 60) targetBrush = _colorYellow;
            }

            fill.Width = new GridLength(percent, GridUnitType.Star);
            empty.Width = new GridLength(100 - percent, GridUnitType.Star);

            // Ensures full rounding on the right side if the bar hits 100% capacity
            fillBorder.CornerRadius = percent >= 99.9 ? new CornerRadius(3) : new CornerRadius(3, 0, 0, 3);
            fillBorder.Background = targetBrush;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if ((_layoutMode == "CPU" || _layoutMode == "Both") && _countersReady)
            {
                float cpuVal = _cpuCounter?.NextValue() ?? 0f;
                if (cpuVal > 100f) cpuVal = 100f; // Processor Utility can occasionally spike above 100 on turbo boost, cap it.

                if (_visualTheme == "Bars")
                {
                    UpdateDiagnosticBar(_cpuFillColumn, _cpuEmptyColumn, _cpuFillBorder, cpuVal);
                    if (_cpuBarValueText != null)
                    {
                        _cpuBarValueText.Text = $"{cpuVal:0.0}%";
                        _cpuBarValueText.Foreground = _cpuFillBorder.Background; // Syncs text color dynamically
                    }
                }
                else
                {
                    AnimateNeedle(_cpuNeedleTransform, cpuVal, ref _cpuCurrentAngle);
                    if (_cpuValueText != null) _cpuValueText.Text = ((int)cpuVal).ToString();
                }
            }

            if (_layoutMode == "RAM" || _layoutMode == "Both")
            {
                float ramVal = 0f;
                MEMORYSTATUSEX memStatus = new MEMORYSTATUSEX();
                if (GlobalMemoryStatusEx(memStatus))
                {
                    ramVal = memStatus.dwMemoryLoad;
                }

                if (_visualTheme == "Bars")
                {
                    UpdateDiagnosticBar(_ramFillColumn, _ramEmptyColumn, _ramFillBorder, ramVal);
                    if (_ramBarValueText != null)
                    {
                        _ramBarValueText.Text = $"{ramVal:0.0}%";
                        _ramBarValueText.Foreground = _ramFillBorder.Background;
                    }
                }
                else
                {
                    AnimateNeedle(_ramNeedleTransform, ramVal, ref _ramCurrentAngle);
                    if (_ramValueText != null) _ramValueText.Text = ((int)ramVal).ToString();
                }
            }
        }

        private void AnimateNeedle(RotateTransform transform, float percentage, ref double currentAngle)
        {
            if (transform == null) return;

            double targetAngle = -90 + (percentage * 1.8);

            DoubleAnimation anim = new DoubleAnimation
            {
                From = currentAngle,
                To = targetAngle,
                Duration = TimeSpan.FromMilliseconds(_refreshRateMs * 0.8),
                EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut }
            };

            transform.BeginAnimation(RotateTransform.AngleProperty, null);
            transform.Angle = currentAngle;
            transform.BeginAnimation(RotateTransform.AngleProperty, anim);

            currentAngle = targetAngle;
        }

        public void Pause()
        {
            _timer?.Stop();
        }

        public void Resume()
        {
            _timer?.Start();
        }

        public void Cleanup()
        {
            _timer?.Stop();
            _cpuCounter?.Dispose();
        }

        public void ShowSettingsWindow(Window ownerWindow, dynamic frameData)
        {
            SolidColorBrush accentBrush;
            try
            {
                accentBrush = new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor));
            }
            catch
            {
                accentBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244));
            }

            Window win = new Window
            {
                Owner = ownerWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = false,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                SizeToContent = SizeToContent.Height,
                Width = 450
            };

            Border headerBorder = new Border { Height = 50, Background = accentBrush };
            headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) win.DragMove(); };

            Grid headerGrid = new Grid();
            headerGrid.Children.Add(new TextBlock
            {
                Text = "System Performance Settings",
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(15, 0, 0, 0)
            });

            Button btnClose = new Button
            {
                Content = "X",
                Foreground = Brushes.White,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Width = 40,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            btnClose.Click += (s, e) => win.Close();
            headerGrid.Children.Add(btnClose);
            headerBorder.Child = headerGrid;

            Border contentBorder = new Border
            {
                Background = Brushes.White,
                Padding = new Thickness(20, 10, 20, 10)
            };
            StackPanel contentPanel = new StackPanel();

            Border groupBox = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(0, 0, 0, 12),
                Padding = new Thickness(12)
            };

            StackPanel groupSp = new StackPanel();
            groupSp.Children.Add(new TextBlock { Text = "Visual Theme & Layout", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 12) });

            // 1. Visual Theme (Determines gauge logic vs bar logic)
            groupSp.Children.Add(new TextBlock { Text = "Visual Style:", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 5) });
            ComboBox cmbTheme = new ComboBox { Margin = new Thickness(0, 0, 0, 15) };
            cmbTheme.Items.Add("Gauges");
            cmbTheme.Items.Add("Bars");
            cmbTheme.SelectedItem = _visualTheme;
            groupSp.Children.Add(cmbTheme);

            CheckBox chkDynamic = new CheckBox { Content = "Use Dynamic Colors for Bars", IsChecked = _useDynamicColors, Margin = new Thickness(0, 0, 0, 10), FontWeight = FontWeights.SemiBold };
            groupSp.Children.Add(chkDynamic);

            StackPanel colorSp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 15) };
            colorSp.Children.Add(new TextBlock { Text = "Static Bar Color:", Width = 150, FontWeight = FontWeights.SemiBold, VerticalAlignment = VerticalAlignment.Center });
            ComboBox cmbStaticColor = new ComboBox { Width = 120 };
            string[] presetColors = { "DeepSkyBlue", "Cyan", "LimeGreen", "Gold", "Orange", "Magenta", "White", "LightGray" };
            foreach (var c in presetColors) cmbStaticColor.Items.Add(c);
            cmbStaticColor.SelectedItem = presetColors.Contains(_staticColor) ? _staticColor : "DeepSkyBlue";
            colorSp.Children.Add(cmbStaticColor);
            groupSp.Children.Add(colorSp);

            // 2. Hardware Toggles
            groupSp.Children.Add(new TextBlock { Text = "Sensors to Display:", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 5) });
            ComboBox cmbLayout = new ComboBox { Margin = new Thickness(0, 0, 0, 15) };
            cmbLayout.Items.Add("CPU");
            cmbLayout.Items.Add("RAM");
            cmbLayout.Items.Add("Both");
            cmbLayout.SelectedItem = _layoutMode;
            groupSp.Children.Add(cmbLayout);

            // 3. Refresh Rate handling via predefined combo rather than loose textbox
            groupSp.Children.Add(new TextBlock { Text = "Sensor Refresh Rate (ms):", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 5) });
            ComboBox cmbRate = new ComboBox { Margin = new Thickness(0, 0, 0, 5) };
            cmbRate.Items.Add("500");
            cmbRate.Items.Add("1000");
            cmbRate.Items.Add("2000");
            cmbRate.Items.Add("5000");
            cmbRate.SelectedItem = _refreshRateMs.ToString();
            if (cmbRate.SelectedItem == null) cmbRate.SelectedItem = "1000";
            groupSp.Children.Add(cmbRate);

            groupBox.Child = groupSp;
            contentPanel.Children.Add(groupBox);
            contentBorder.Child = contentPanel;

            Border footerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(15)
            };

            StackPanel footerSp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            Button btnCancel = new Button
            {
                Content = "Cancel",
                Background = Brushes.White,
                BorderBrush = Brushes.Gray,
                Foreground = Brushes.Black,
                Width = 80,
                Height = 30,
                Margin = new Thickness(0, 0, 10, 0)
            };
            btnCancel.Click += (s, e) => win.Close();

            Button btnSave = new Button
            {
                Content = "Save",
                Background = accentBrush,
                Foreground = Brushes.White,
                FontWeight = FontWeights.Bold,
                BorderThickness = new Thickness(0),
                Width = 80,
                Height = 30
            };
            btnSave.Click += (s, e) =>
            {
                int newRate = 1000;
                if (cmbRate.SelectedItem != null) int.TryParse(cmbRate.SelectedItem.ToString(), out newRate);

                Dictionary<string, object> newSettings = new Dictionary<string, object>
                {
                    { "VisualTheme", cmbTheme.SelectedItem.ToString() },
                    { "UseDynamicColors", chkDynamic.IsChecked == true },
                    { "StaticColor", cmbStaticColor.SelectedItem?.ToString() ?? "DeepSkyBlue" },
                    { "LayoutMode", cmbLayout.SelectedItem.ToString() },
                    { "RefreshRateMs", newRate }
                };

                if (frameData is Newtonsoft.Json.Linq.JObject jFrame)
                    jFrame["PluginSettings"] = Newtonsoft.Json.Linq.JObject.FromObject(newSettings);
                else
                    ((IDictionary<string, object>)frameData)["PluginSettings"] = newSettings;

                try { FrameDataManager.SaveFrameData(); } catch { }
                ApplySettingsData(newSettings);

                if (_timer != null)
                {
                    _timer.Interval = TimeSpan.FromMilliseconds(_refreshRateMs);
                }
                BuildLayout();

                win.Close();
            };

            footerSp.Children.Add(btnCancel);
            footerSp.Children.Add(btnSave);
            footerBorder.Child = footerSp;

            DockPanel rootPanel = new DockPanel();
            DockPanel.SetDock(headerBorder, Dock.Top);
            DockPanel.SetDock(footerBorder, Dock.Bottom);

            rootPanel.Children.Add(headerBorder);
            rootPanel.Children.Add(footerBorder);
            rootPanel.Children.Add(contentBorder);

            win.Content = rootPanel;
            win.ShowDialog();
        }
    }
}
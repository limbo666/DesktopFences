using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        public int DevelopmentState => 2; // Set to 1, 2, or 3 based on your testing phase

        private DispatcherTimer _timer;
        private PerformanceCounter _cpuCounter;

        private Viewbox _rootViewbox;
        private Grid _layoutGrid;

        private RotateTransform _cpuNeedleTransform;
        private RotateTransform _ramNeedleTransform;
        private TextBlock _cpuValueText;
        private TextBlock _ramValueText;

        private string _layoutMode = "Both";
        private int _refreshRateMs = 1000;

        private double _cpuCurrentAngle = -90;
        private double _ramCurrentAngle = -90;

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
            _rootViewbox = new Viewbox
            {
                Stretch = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(5, 5, 25, 10) // Master UI Spacing Blueprint
            };

            _layoutGrid = new Grid();
            _rootViewbox.Child = _layoutGrid;

            return _rootViewbox;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            ApplySettingsData(settings);

            try
            {
                // Modern Windows uses Processor Utility for Task Manager parity
                _cpuCounter = new PerformanceCounter("Processor Information", "% Processor Utility", "_Total");
            }
            catch
            {
                // Fallback for older systems
                _cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
            }

            BuildLayout();

            _timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(_refreshRateMs) };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void ApplySettingsData(Dictionary<string, object> settings)
        {
            if (settings != null)
            {
                if (settings.ContainsKey("LayoutMode")) _layoutMode = settings["LayoutMode"].ToString();
                if (settings.ContainsKey("RefreshRateMs") && int.TryParse(settings["RefreshRateMs"].ToString(), out int rate)) _refreshRateMs = rate;
            }
        }

        private void BuildLayout()
        {
            _layoutGrid.Children.Clear();
            _layoutGrid.ColumnDefinitions.Clear();

            _cpuCurrentAngle = -90;
            _ramCurrentAngle = -90;

            bool showCpu = _layoutMode == "CPU" || _layoutMode == "Both";
            bool showRam = _layoutMode == "RAM" || _layoutMode == "Both";

            if (showCpu && showRam)
            {
                _layoutGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                _layoutGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var cpuGauge = CreateGauge("CPU %", out _cpuNeedleTransform, out _cpuValueText);
                var ramGauge = CreateGauge("RAM %", out _ramNeedleTransform, out _ramValueText);

                Grid.SetColumn(cpuGauge, 0);
                Grid.SetColumn(ramGauge, 1);

                cpuGauge.Margin = new Thickness(0, 0, 10, 0);
                ramGauge.Margin = new Thickness(10, 0, 0, 0);

                _layoutGrid.Children.Add(cpuGauge);
                _layoutGrid.Children.Add(ramGauge);
            }
            else if (showCpu)
            {
                _layoutGrid.Children.Add(CreateGauge("CPU %", out _cpuNeedleTransform, out _cpuValueText));
            }
            else if (showRam)
            {
                _layoutGrid.Children.Add(CreateGauge("RAM %", out _ramNeedleTransform, out _ramValueText));
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

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (_layoutMode == "CPU" || _layoutMode == "Both")
            {
                float cpuVal = _cpuCounter?.NextValue() ?? 0f;
                // Processor Utility can occasionally spike above 100 on turbo boost, cap it.
                if (cpuVal > 100f) cpuVal = 100f;

                AnimateNeedle(_cpuNeedleTransform, cpuVal, ref _cpuCurrentAngle);
                if (_cpuValueText != null) _cpuValueText.Text = ((int)cpuVal).ToString();
            }

            if (_layoutMode == "RAM" || _layoutMode == "Both")
            {
                float ramVal = 0f;
                MEMORYSTATUSEX memStatus = new MEMORYSTATUSEX();
                if (GlobalMemoryStatusEx(memStatus))
                {
                    ramVal = memStatus.dwMemoryLoad; // Fetch exact Physical RAM %
                }

                AnimateNeedle(_ramNeedleTransform, ramVal, ref _ramCurrentAngle);
                if (_ramValueText != null) _ramValueText.Text = ((int)ramVal).ToString();
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
            groupSp.Children.Add(new TextBlock { Text = "Display Options", FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10) });

            groupSp.Children.Add(new TextBlock { Text = "Layout Mode:", Margin = new Thickness(0, 0, 0, 5) });
            ComboBox cmbLayout = new ComboBox { Margin = new Thickness(0, 0, 0, 15) };
            cmbLayout.Items.Add("CPU");
            cmbLayout.Items.Add("RAM");
            cmbLayout.Items.Add("Both");
            cmbLayout.SelectedItem = _layoutMode;
            groupSp.Children.Add(cmbLayout);

            groupSp.Children.Add(new TextBlock { Text = "Refresh Rate (ms):", Margin = new Thickness(0, 0, 0, 5) });
            TextBox txtRefresh = new TextBox { Text = _refreshRateMs.ToString() };
            groupSp.Children.Add(txtRefresh);

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
                Dictionary<string, object> newSettings = new Dictionary<string, object>
                {
                    { "LayoutMode", cmbLayout.SelectedItem.ToString() },
                    { "RefreshRateMs", txtRefresh.Text }
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
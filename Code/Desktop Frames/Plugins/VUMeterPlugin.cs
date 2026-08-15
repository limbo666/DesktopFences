using Desktop_Frames.Localization;
using System;
using System.Collections.Generic;
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
    public class VUMeterPlugin : IFramePlugin
    {
        public string PluginId => "AnalogVUMeter";
        public string DisplayName => "Audio VU Meter";

        public int DevelopmentState => 1; // Set to 1, 2, or 3 based on your testing phase

        private DispatcherTimer _timer;
        private Viewbox _rootViewbox;
        private Grid _layoutGrid;

        private RotateTransform _leftNeedleTransform;
        private RotateTransform _rightNeedleTransform;

        // Plugin Settings
        private string _layoutMode = "Stereo";
        private double _gain = 1.2;
        private int _attackMs = 25;
        private int _decayMs = 200;
        private int _refreshRateMs = 30;

        // True Analog State Tracking (Mathematical Smoothing)
        private float _leftSmoothedValue = 0f;
        private float _rightSmoothedValue = 0f;

        // COM Objects for Audio Metering
        private IMMDeviceEnumerator _deviceEnumerator;
        private IMMDevice _device;
        private IAudioMeterInformation _audioMeter;
        private AudioNotificationClient _notificationClient;
        private bool _deviceChanged = false;

        public FrameworkElement CreateVisualElement()
        {
            _rootViewbox = new Viewbox
            {
                Stretch = Stretch.Uniform,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(8)
            };

            _layoutGrid = new Grid();
            _rootViewbox.Child = _layoutGrid;

            return _rootViewbox;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            ApplySettingsData(settings);
            InitAudioMeter();
            BuildLayout();

            _timer = new DispatcherTimer(DispatcherPriority.Render) { Interval = TimeSpan.FromMilliseconds(_refreshRateMs) };
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void InitAudioMeter()
        {
            try
            {
                if (_deviceEnumerator == null)
                {
                    _deviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();

                    _notificationClient = new AudioNotificationClient(() => _deviceChanged = true);
                    _deviceEnumerator.RegisterEndpointNotificationCallback(_notificationClient);
                }

                if (_audioMeter != null) { Marshal.ReleaseComObject(_audioMeter); _audioMeter = null; }
                if (_device != null) { Marshal.ReleaseComObject(_device); _device = null; }

                _deviceEnumerator.GetDefaultAudioEndpoint(0, 1, out _device);

                if (_device != null)
                {
                    Guid iid = typeof(IAudioMeterInformation).GUID;
                    _device.Activate(ref iid, 2 /*CLSCTX_ALL*/, IntPtr.Zero, out object meterObj);
                    _audioMeter = (IAudioMeterInformation)meterObj;
                }
            }
            catch { /* Fail silently */ }
        }

        private void ApplySettingsData(Dictionary<string, object> settings)
        {
            if (settings != null)
            {
                if (settings.ContainsKey("LayoutMode")) _layoutMode = settings["LayoutMode"].ToString();
                if (settings.ContainsKey("Gain") && double.TryParse(settings["Gain"].ToString(), out double g)) _gain = Math.Max(0.1, Math.Min(10.0, g));
                if (settings.ContainsKey("AttackMs") && int.TryParse(settings["AttackMs"].ToString(), out int a)) _attackMs = Math.Max(1, Math.Min(1000, a));
                if (settings.ContainsKey("DecayMs") && int.TryParse(settings["DecayMs"].ToString(), out int d)) _decayMs = Math.Max(10, Math.Min(3000, d));
            }
        }

        private void BuildLayout()
        {
            _layoutGrid.Children.Clear();
            _layoutGrid.ColumnDefinitions.Clear();

            _leftSmoothedValue = 0f;
            _rightSmoothedValue = 0f;

            if (_layoutMode == "Stereo")
            {
                _layoutGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                _layoutGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var leftGauge = CreateGauge("LEFT (L)", out _leftNeedleTransform);
                var rightGauge = CreateGauge("RIGHT (R)", out _rightNeedleTransform);

                Grid.SetColumn(leftGauge, 0);
                Grid.SetColumn(rightGauge, 1);

                leftGauge.Margin = new Thickness(0, 0, 10, 0);
                rightGauge.Margin = new Thickness(10, 0, 0, 0);

                _layoutGrid.Children.Add(leftGauge);
                _layoutGrid.Children.Add(rightGauge);
            }
            else if (_layoutMode == "Combined")
            {
                _layoutGrid.Children.Add(CreateGauge("MASTER VU", out _leftNeedleTransform));
            }
            else if (_layoutMode == "Left")
            {
                _layoutGrid.Children.Add(CreateGauge("LEFT VU", out _leftNeedleTransform));
            }
            else if (_layoutMode == "Right")
            {
                _layoutGrid.Children.Add(CreateGauge("RIGHT VU", out _rightNeedleTransform));
            }
        }

        private Canvas CreateGauge(string title, out RotateTransform needleTransform)
        {
            Canvas canvas = new Canvas { Width = 140, Height = 120 };

            //Path dialBase = new Path
            //{
            //    Stroke = Brushes.DarkGray,
            //    StrokeThickness = 6,
            //    // Fake logarithmic dash scale: { Dash, Gap, Dash, Gap... }
            //    // Multipliers of StrokeThickness (6px) creating progressively wider dashes and tighter gaps
            //    StrokeDashArray = new DoubleCollection { 0.2, 1.8, 0.4, 1.5, 0.6, 1.2, 0.8, 1.0, 1.2, 0.8, 2.0, 0.6, 3.0, 0.5, 5.0, 0.5 },
            //    Data = Geometry.Parse("M 20,80 A 50,50 0 0,1 120,80")
            //};

            Path dialBase = new Path
            {
                Stroke = Brushes.DarkGray,
                StrokeThickness = 6,
                // Reversed fake logarithmic dash scale: Thick dashes on the left, thinning out toward the right
                StrokeDashArray = new DoubleCollection { 5.0, 0.5, 3.0, 0.5, 2.0, 0.6, 1.2, 0.8, 0.8, 1.0, 0.6, 1.2, 0.4, 1.5, 0.2, 1.8 },
                Data = Geometry.Parse("M 20,80 A 50,50 0 0,1 120,80")
            };

            canvas.Children.Add(dialBase);

            Path dialPeak = new Path
            {
                Stroke = Brushes.Red,
                StrokeThickness = 6,
                Data = Geometry.Parse("M 95,36.7 A 50,50 0 0,1 120,80")
            };
            canvas.Children.Add(dialPeak);

            // Shifted left (from 5 to -5) to completely clear the start of the curve at X=20
            // Precision VU Numbering (Surgically aligned to the dash gaps)
            canvas.Children.Add(CreateLabel("-20", -1, 72));
            canvas.Children.Add(CreateLabel("-10", 12, 35)); // 1st gap
            canvas.Children.Add(CreateLabel("-7", 34, 19));  // 2nd gap
            canvas.Children.Add(CreateLabel("-5", 50, 13));   // 3rd gap
            canvas.Children.Add(CreateLabel("-3", 65, 12));   // 4th gap
            canvas.Children.Add(CreateLabel("-1", 79, 13));  // 5th gap
            canvas.Children.Add(CreateLabel("0", 94, 17));  // Start of Red Peak
            canvas.Children.Add(CreateLabel("+3", 125, 72));

            needleTransform = new RotateTransform(-90, 70, 80);
            Line needle = new Line
            {
                X1 = 70,
                Y1 = 80,
                X2 = 70,
                Y2 = 35,
                Stroke = Brushes.White,
                StrokeThickness = 2,
                RenderTransform = needleTransform
            };
            canvas.Children.Add(needle);

            Ellipse pin = new Ellipse { Width = 10, Height = 10, Fill = Brushes.Silver };
            Canvas.SetTop(pin, 75);
            Canvas.SetLeft(pin, 65);
            canvas.Children.Add(pin);

            TextBlock titleText = new TextBlock
            {
                Text = title,
                Foreground = Brushes.White,
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Width = 140,
                TextAlignment = TextAlignment.Center
            };
            Canvas.SetTop(titleText, 100);
            Canvas.SetLeft(titleText, 0);
            canvas.Children.Add(titleText);

            return canvas;
        }

        private TextBlock CreateLabel(string text, double left, double top)
        {
            var tb = new TextBlock { Text = text, Foreground = Brushes.LightGray, FontSize = 10, FontWeight = FontWeights.SemiBold };
            Canvas.SetLeft(tb, left);
            Canvas.SetTop(tb, top);
            return tb;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (_deviceChanged)
            {
                _deviceChanged = false;
                InitAudioMeter();
            }

            if (_audioMeter == null) return;

            try
            {
                // Calculate EMA Integration factors (0.0 to 1.0) based on UI settings
                float attackFactor = Math.Min(1.0f, (float)_refreshRateMs / _attackMs);
                float decayFactor = Math.Min(1.0f, (float)_refreshRateMs / _decayMs);

                float rawLeft = 0f;
                float rawRight = 0f;

                if (_layoutMode == "Combined")
                {
                    _audioMeter.GetPeakValue(out float masterPeak);
                    rawLeft = Math.Min(masterPeak * (float)_gain * 100f, 100f);
                }
                else
                {
                    _audioMeter.GetMeteringChannelCount(out int channelCount);
                    if (channelCount > 0)
                    {
                        float[] peaks = new float[channelCount];
                        _audioMeter.GetChannelsPeakValues(channelCount, peaks);

                        rawLeft = Math.Min(peaks[0] * (float)_gain * 100f, 100f);
                        rawRight = Math.Min(((channelCount > 1) ? peaks[1] : peaks[0]) * (float)_gain * 100f, 100f);
                    }
                }

                // Envelope Follower Math: Smooths the erratic jumps BEFORE animating
                _leftSmoothedValue += (rawLeft - _leftSmoothedValue) * (rawLeft > _leftSmoothedValue ? attackFactor : decayFactor);
                _rightSmoothedValue += (rawRight - _rightSmoothedValue) * (rawRight > _rightSmoothedValue ? attackFactor : decayFactor);

                if (_layoutMode == "Stereo")
                {
                    AnimateNeedle(_leftNeedleTransform, _leftSmoothedValue);
                    AnimateNeedle(_rightNeedleTransform, _rightSmoothedValue);
                }
                else if (_layoutMode == "Left" || _layoutMode == "Combined")
                {
                    AnimateNeedle(_leftNeedleTransform, _leftSmoothedValue);
                }
                else if (_layoutMode == "Right")
                {
                    AnimateNeedle(_rightNeedleTransform, _rightSmoothedValue);
                }
            }
            catch
            {
                InitAudioMeter();
            }
        }

        private void AnimateNeedle(RotateTransform transform, float smoothedPercentage)
        {
            if (transform == null) return;

            double targetAngle = -90 + (smoothedPercentage * 1.8);

            // Since the math handles the analog smoothing, WPF only needs to draw a linear bridge to the next frame
            DoubleAnimation anim = new DoubleAnimation
            {
                To = targetAngle,
                Duration = TimeSpan.FromMilliseconds(_refreshRateMs)
            };

            transform.BeginAnimation(RotateTransform.AngleProperty, anim);
        }

        public void Pause() => _timer?.Stop();

        public void Resume()
        {
            InitAudioMeter();
            _timer?.Start();
        }

        public void Cleanup()
        {
            _timer?.Stop();

            try
            {
                if (_deviceEnumerator != null && _notificationClient != null)
                {
                    _deviceEnumerator.UnregisterEndpointNotificationCallback(_notificationClient);
                }
            }
            catch { }

            if (_audioMeter != null) Marshal.ReleaseComObject(_audioMeter);
            if (_device != null) Marshal.ReleaseComObject(_device);
            if (_deviceEnumerator != null) Marshal.ReleaseComObject(_deviceEnumerator);
        }

        // ==========================================
        // UI BOILERPLATE FOR DESKTOP FRAMES +
        // ==========================================
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
                Text = Strings.VuSettings,
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

            // Group 1: Display Options
            Border groupBox1 = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(0, 0, 0, 12),
                Padding = new Thickness(12)
            };
            StackPanel groupSp1 = new StackPanel();
            groupSp1.Children.Add(new TextBlock { Text = Strings.VuDisplayOptions, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10) });

            groupSp1.Children.Add(new TextBlock { Text = Strings.VuLayoutMode, Margin = new Thickness(0, 0, 0, 5) });
            ComboBox cmbLayout = new ComboBox { Margin = new Thickness(0, 0, 0, 15) };
            cmbLayout.Items.Add("Stereo");
            cmbLayout.Items.Add("Combined");
            cmbLayout.Items.Add("Left");
            cmbLayout.Items.Add("Right");
            cmbLayout.SelectedItem = _layoutMode;
            groupSp1.Children.Add(cmbLayout);
            groupBox1.Child = groupSp1;
            contentPanel.Children.Add(groupBox1);

            // Group 2: Meter Ballistics
            Border groupBox2 = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(0, 0, 0, 12),
                Padding = new Thickness(12)
            };
            StackPanel groupSp2 = new StackPanel();
            groupSp2.Children.Add(new TextBlock { Text = Strings.VuBallistics, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10) });

            groupSp2.Children.Add(new TextBlock { Text = Strings.VuSignalGain, Margin = new Thickness(0, 0, 0, 2) });
            groupSp2.Children.Add(new TextBlock { Text = Strings.VuGainHint, Foreground = Brushes.Gray, FontSize = 10, Margin = new Thickness(0, 0, 0, 5) });
            TextBox txtGain = new TextBox { Text = _gain.ToString(), Margin = new Thickness(0, 0, 0, 15) };
            groupSp2.Children.Add(txtGain);

            groupSp2.Children.Add(new TextBlock { Text = Strings.VuAttack, Margin = new Thickness(0, 0, 0, 2) });
            groupSp2.Children.Add(new TextBlock { Text = Strings.VuAttackHint, Foreground = Brushes.Gray, FontSize = 10, Margin = new Thickness(0, 0, 0, 5) });
            TextBox txtAttack = new TextBox { Text = _attackMs.ToString(), Margin = new Thickness(0, 0, 0, 15) };
            groupSp2.Children.Add(txtAttack);

            groupSp2.Children.Add(new TextBlock { Text = Strings.VuDecay, Margin = new Thickness(0, 0, 0, 2) });
            groupSp2.Children.Add(new TextBlock { Text = Strings.VuDecayHint, Foreground = Brushes.Gray, FontSize = 10, Margin = new Thickness(0, 0, 0, 5) });
            TextBox txtDecay = new TextBox { Text = _decayMs.ToString(), Margin = new Thickness(0, 0, 0, 10) };
            groupSp2.Children.Add(txtDecay);

            groupBox2.Child = groupSp2;
            contentPanel.Children.Add(groupBox2);
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
                if (double.TryParse(txtGain.Text, out double parsedGain)) _gain = Math.Max(0.1, Math.Min(10.0, parsedGain));
                if (int.TryParse(txtAttack.Text, out int parsedAttack)) _attackMs = Math.Max(1, Math.Min(1000, parsedAttack));
                if (int.TryParse(txtDecay.Text, out int parsedDecay)) _decayMs = Math.Max(10, Math.Min(3000, parsedDecay));

                Dictionary<string, object> newSettings = new Dictionary<string, object>
                {
                    { "LayoutMode", cmbLayout.SelectedItem.ToString() },
                    { "Gain", _gain },
                    { "AttackMs", _attackMs },
                    { "DecayMs", _decayMs }
                };

                if (frameData is Newtonsoft.Json.Linq.JObject jFrame)
                    jFrame["PluginSettings"] = Newtonsoft.Json.Linq.JObject.FromObject(newSettings);
                else
                    ((IDictionary<string, object>)frameData)["PluginSettings"] = newSettings;

                try { FrameDataManager.SaveFrameData(); } catch { }
                ApplySettingsData(newSettings);

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

        // ==========================================
        // WINDOWS CORE AUDIO API (WASAPI) COM INTERFACES
        // ==========================================
        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        internal class MMDeviceEnumerator { }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDeviceEnumerator
        {
            [PreserveSig]
            int EnumAudioEndpoints(int dataFlow, int stateMask, out IntPtr ppDevices);

            [PreserveSig]
            int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppEndpoint);

            [PreserveSig]
            int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);

            [PreserveSig]
            int RegisterEndpointNotificationCallback(IMMNotificationClient pClient);

            [PreserveSig]
            int UnregisterEndpointNotificationCallback(IMMNotificationClient pClient);
        }

        [Guid("7991EEC9-7E89-4D85-8390-6C703CEC60C0"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMNotificationClient
        {
            [PreserveSig] int OnDeviceStateChanged([MarshalAs(UnmanagedType.LPWStr)] string deviceId, uint dwNewState);
            [PreserveSig] int OnDeviceAdded([MarshalAs(UnmanagedType.LPWStr)] string deviceId);
            [PreserveSig] int OnDeviceRemoved([MarshalAs(UnmanagedType.LPWStr)] string deviceId);
            [PreserveSig] int OnDefaultDeviceChanged(int flow, int role, [MarshalAs(UnmanagedType.LPWStr)] string defaultDeviceId);
            [PreserveSig] int OnPropertyValueChanged([MarshalAs(UnmanagedType.LPWStr)] string deviceId, PROPERTYKEY key);
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct PROPERTYKEY
        {
            public Guid fmtid;
            public uint pid;
        }

        private class AudioNotificationClient : IMMNotificationClient
        {
            private Action _onDefaultDeviceChanged;
            public AudioNotificationClient(Action onDefaultDeviceChanged) { _onDefaultDeviceChanged = onDefaultDeviceChanged; }

            public int OnDeviceStateChanged(string deviceId, uint dwNewState) => 0;
            public int OnDeviceAdded(string deviceId) => 0;
            public int OnDeviceRemoved(string deviceId) => 0;

            public int OnDefaultDeviceChanged(int flow, int role, string defaultDeviceId)
            {
                if (flow == 0 && role == 1)
                {
                    _onDefaultDeviceChanged?.Invoke();
                }
                return 0;
            }

            public int OnPropertyValueChanged(string deviceId, PROPERTYKEY key) => 0;
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDevice
        {
            [PreserveSig]
            int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        }

        [Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioMeterInformation
        {
            [PreserveSig]
            int GetPeakValue(out float pfPeak);
            [PreserveSig]
            int GetMeteringChannelCount(out int pnChannelCount);
            [PreserveSig]
            int GetChannelsPeakValues(int u32ChannelCount, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] float[] afPeakValues);
        }
    }
}
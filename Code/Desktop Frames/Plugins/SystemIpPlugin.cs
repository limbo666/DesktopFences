using Desktop_Frames.Localization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.IO;
using System.Windows.Input;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Desktop_Frames.Plugins
{
    public class SystemIpPlugin : IFramePlugin
    {
        public string PluginId => "SystemIpMonitor";
        public string DisplayName => "Network IP Monitor";
        public int DevelopmentState => 2; // Set to 1, 2, or 3 based on your testing phase

        private DispatcherTimer _localIpTimer;
        private DispatcherTimer _publicIpTimer;
        private ScrollViewer _rootVisual;
        private StackPanel _cardsPanel;

        // HTTP Client for Public IP (Shared, robust, background-safe)
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };

        // Settings
        private bool _showDisconnected = false;
        private bool _showPublicIp = true;
        private List<string> _shownInterfaces = new List<string>();

        // Safety flag: If the user hasn't saved a whitelist yet, show all interfaces by default
        private bool _isWhitelistInitialized = false;

        // State tracking
        private string _lastRenderedStateHash = "";
        private string _currentPublicIp = "Checking...";

        // Memory for Flash Animations
        private string _previousPublicIp = "";
        private Dictionary<string, string> _previousIps = new Dictionary<string, string>();
        private bool _isFirstLoad = true;

        // Brush Caching (Prevents animations from being killed during UI redraws)
        private SolidColorBrush _publicIpBrush = new SolidColorBrush(Color.FromArgb(25, 100, 150, 255));
        private Dictionary<string, SolidColorBrush> _cardBrushes = new Dictionary<string, SolidColorBrush>();

        // Memory for Multiple IP Selections
        private Dictionary<string, int> _selectedIpIndices = new Dictionary<string, int>();

        // Settings Reference for Persistent Saving
        private Dictionary<string, object> _settingsRef;

        public FrameworkElement CreateVisualElement()
        {
            _rootVisual = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(10),
                Margin = new Thickness(5, 5, 5, 10) // Master UI Spacing Blueprint
            };

            _cardsPanel = new StackPanel { Orientation = Orientation.Vertical };
            _rootVisual.Content = _cardsPanel;

            return _rootVisual;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            _settingsRef = settings; // Store reference to save states silently
            ApplySettingsData(settings);

            // 1. Local IP Fast Polling Timer (Every 2 seconds)
            _localIpTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromSeconds(2)
            };
            _localIpTimer.Tick += (s, e) => UpdateInterfaceData();
            _localIpTimer.Start();

            // 2. Public IP Polling Timer (Every 30 seconds to respect API limits)
            _publicIpTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromSeconds(30)
            };
            _publicIpTimer.Tick += (s, e) => _ = FetchPublicIpAsync();
            _publicIpTimer.Start();

            // Initial UI Draw
            UpdateInterfaceData(forceRedraw: true);

            // Delayed background fetch for Public IP to guarantee zero UI blocking on startup
            Task.Run(async () =>
            {
                await Task.Delay(2000);
                await FetchPublicIpAsync();
            });
        }

        private async Task FetchPublicIpAsync()
        {
            if (!_showPublicIp) return;

            // Pre-flight check: Is the network card even online?
            if (!NetworkInterface.GetIsNetworkAvailable())
            {
                UpdatePublicIpState(Strings.NetOffline);
                return;
            }

            try
            {
                // Simple, fast, plain-text IP API
                string ip = await _httpClient.GetStringAsync("https://api.ipify.org");
                UpdatePublicIpState(ip.Trim());
            }
            catch
            {
                UpdatePublicIpState(Strings.NetUnreachable);
            }
        }

        private void UpdatePublicIpState(string newState)
        {
            if (_currentPublicIp != newState)
            {
                _currentPublicIp = newState;
                // Force a redraw on the UI thread when the public IP state changes
                Application.Current.Dispatcher.InvokeAsync(() => UpdateInterfaceData(forceRedraw: true));
            }
        }

        private void ApplySettingsData(Dictionary<string, object> settings)
        {
            if (settings != null)
            {
                if (settings.ContainsKey("ShowDisconnected") && bool.TryParse(settings["ShowDisconnected"].ToString(), out bool sd))
                    _showDisconnected = sd;

                if (settings.ContainsKey("ShowPublicIp") && bool.TryParse(settings["ShowPublicIp"].ToString(), out bool spi))
                    _showPublicIp = spi;

                if (settings.ContainsKey("ShownInterfaces"))
                {
                    string shownStr = settings["ShownInterfaces"].ToString();
                    _shownInterfaces = shownStr.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    _isWhitelistInitialized = true; // The user has a saved preference
                }

                // Restore previous carousel positions from local cache file
                try
                {
                    string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemIpState.dat");
                    if (File.Exists(path))
                    {
                        string indicesStr = File.ReadAllText(path);
                        var pairs = indicesStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var pair in pairs)
                        {
                            var parts = pair.Split(':');
                            if (parts.Length == 2 && int.TryParse(parts[1], out int index))
                            {
                                _selectedIpIndices[parts[0]] = index;
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private void SaveIpSelectionState()
        {
            try
            {
                string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SystemIpState.dat");
                string data = string.Join("|", _selectedIpIndices.Select(kv => $"{kv.Key}:{kv.Value}"));
                File.WriteAllText(path, data);
            }
            catch { } // Fail silently if folder permissions are strictly read-only
        }
        private void UpdateInterfaceData(bool forceRedraw = false)
        {
var interfaces = NetworkInterface.GetAllNetworkInterfaces()
                .Where(ni => ni.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                // If whitelist isn't initialized, show all. Otherwise, strictly follow the whitelist.
                .Where(ni => !_isWhitelistInitialized || _shownInterfaces.Contains(ni.Name))
                .Where(ni => _showDisconnected || ni.OperationalStatus == OperationalStatus.Up)
                .ToList();

            // Hash includes local interfaces AND the public IP state
            string currentHash = $"{_currentPublicIp}|" + string.Join("|", interfaces.Select(ni =>
                $"{ni.Name}-{ni.OperationalStatus}-{string.Join(",", GetIPv4List(ni))}"));

            if (currentHash == _lastRenderedStateHash && !forceRedraw)
                return; // Nothing changed, skip redraw

            _lastRenderedStateHash = currentHash;

            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                _cardsPanel.Children.Clear();

                SolidColorBrush accentBrush;
                try { accentBrush = new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor)); }
                catch { accentBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244)); }

                // --- 1. Draw Public IP Card at the very top ---
                if (_showPublicIp)
                {
                    bool publicIpChanged = !_isFirstLoad && _previousPublicIp != "" && _previousPublicIp != _currentPublicIp;
                    _previousPublicIp = _currentPublicIp;

                    _cardsPanel.Children.Add(CreatePublicIpCard(publicIpChanged));
                }

                // --- 2. Draw Local Interface Cards ---
                if (interfaces.Count == 0 && !_showPublicIp)
                {
                    _cardsPanel.Children.Add(new TextBlock
                    {
                        Text = Strings.NetNoInterfaces,
                        Foreground = Brushes.Gray,
                        FontStyle = FontStyles.Italic,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(0, 20, 0, 0)
                    });
                    _isFirstLoad = false;
                    return;
                }

                foreach (var ni in interfaces)
                {
                    List<string> ips = GetIPv4List(ni);

                    if (!_selectedIpIndices.ContainsKey(ni.Name))
                        _selectedIpIndices[ni.Name] = 0;

                    // Safety bounds check in case an IP is dropped from the interface
                    if (_selectedIpIndices[ni.Name] >= ips.Count)
                        _selectedIpIndices[ni.Name] = 0;

                    int selectedIndex = _selectedIpIndices[ni.Name];
                    string displayedIp = ips[selectedIndex];

                    // Check if the currently displayed IP changed (from network shift, not user scrolling)
                    bool localIpChanged = !_isFirstLoad && _previousIps.ContainsKey(ni.Name) && _previousIps[ni.Name] != displayedIp;
                    _previousIps[ni.Name] = displayedIp;

                    _cardsPanel.Children.Add(CreateInterfaceCard(ni, accentBrush, ips, selectedIndex, localIpChanged));
                }

                _isFirstLoad = false;
            });
        }

        private List<string> GetIPv4List(NetworkInterface ni)
        {
            var ipProps = ni.GetIPProperties();
            var ipv4s = ipProps.UnicastAddresses
                .Where(a => a.Address.AddressFamily == AddressFamily.InterNetwork)
                .Select(a => a.Address.ToString())
                .ToList();

            if (ipv4s.Count == 0) ipv4s.Add("No IPv4 Address");
            return ipv4s;
        }

        private Border CreatePublicIpCard(bool recentlyChanged)
        {
            bool isUp = _currentPublicIp != Strings.NetOffline && _currentPublicIp != Strings.NetUnreachable && _currentPublicIp != Strings.NetChecking;

            Border card = new Border
            {
                Background = _publicIpBrush, // Use cached brush so animations survive redraws
                BorderBrush = new SolidColorBrush(Color.FromArgb(40, 255, 255, 255)), // Faint white outline
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(10, 5, 30, 15), // Preserved your custom right margin (30) to balance the frame
                Padding = new Thickness(15, 12, 15, 12)
            };

            // Trigger smooth Gold pulse if the IP just changed
            if (recentlyChanged)
            {
                ColorAnimation flashAnim = new ColorAnimation
                {
                    From = Color.FromArgb(180, 255, 215, 0), // Bright Gold
                    To = Color.FromArgb(25, 100, 150, 255),  // Fade back to normal blue glass
                    Duration = TimeSpan.FromSeconds(2.5),
                    EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut }
                };
                _publicIpBrush.BeginAnimation(SolidColorBrush.ColorProperty, flashAnim);
            }

            StackPanel cardContent = new StackPanel { Orientation = Orientation.Vertical };

            // Header
            Grid headerPanel = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // Globe icon instead of standard dot
            TextBlock globeIcon = new TextBlock
            {
                Text = "🌐",
                Foreground = Brushes.White,
                FontSize = 14,
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(globeIcon, 0);

            TextBlock lblName = new TextBlock
            {
                Text = Strings.NetPublicWan,
                Foreground = Brushes.LightSkyBlue, // Distinctive title color
                FontWeight = FontWeights.Bold,
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(lblName, 1);

            headerPanel.Children.Add(globeIcon);
            headerPanel.Children.Add(lblName);

            // IP Data Row
            Grid ipGrid = new Grid();
            ipGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
            ipGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock lblIpPrefix = new TextBlock
            {
                Text = Strings.NetIpLabel,
                Foreground = Brushes.DarkGray,
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(lblIpPrefix, 0);

            TextBlock lblIpValue = new TextBlock
            {
                Text = _currentPublicIp,
                Foreground = isUp ? Brushes.White : Brushes.Orange, // Orange for checking/unreachable
                FontFamily = new FontFamily("Consolas"),
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(lblIpValue, 1);

            ipGrid.Children.Add(lblIpPrefix);
            ipGrid.Children.Add(lblIpValue);

            cardContent.Children.Add(headerPanel);
            cardContent.Children.Add(ipGrid);
            card.Child = cardContent;

            return card;
        }

        private Border CreateInterfaceCard(NetworkInterface ni, SolidColorBrush accentBrush, List<string> ips, int selectedIndex, bool recentlyChanged)
        {
            bool isUp = ni.OperationalStatus == OperationalStatus.Up;

            if (!_cardBrushes.ContainsKey(ni.Name))
            {
                _cardBrushes[ni.Name] = new SolidColorBrush(Color.FromArgb(15, 255, 255, 255));
            }
            SolidColorBrush bgBrush = _cardBrushes[ni.Name];

            Border card = new Border
            {
                Background = bgBrush,
                BorderThickness = new Thickness(0),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(10, 5, 30, 10),
                Padding = new Thickness(15, 12, 15, 12)
            };

            if (recentlyChanged)
            {
                ColorAnimation flashAnim = new ColorAnimation
                {
                    From = Color.FromArgb(150, 0, 255, 200),
                    To = Color.FromArgb(15, 255, 255, 255),
                    Duration = TimeSpan.FromSeconds(2.5),
                    EasingFunction = new QuarticEase { EasingMode = EasingMode.EaseOut }
                };
                bgBrush.BeginAnimation(SolidColorBrush.ColorProperty, flashAnim);
            }

            StackPanel cardContent = new StackPanel { Orientation = Orientation.Vertical };

            // Header Row
            Grid headerPanel = new Grid { Margin = new Thickness(0, 0, 0, 6) };
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            headerPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Ellipse statusDot = new Ellipse
            {
                Width = 8,
                Height = 8,
                Fill = isUp ? Brushes.LimeGreen : Brushes.Crimson,
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            };

            if (isUp) statusDot.Effect = new DropShadowEffect { Color = Colors.LimeGreen, BlurRadius = 5, ShadowDepth = 0 };
            Grid.SetColumn(statusDot, 0);

            TextBlock lblName = new TextBlock
            {
                Text = ni.Name,
                Foreground = Brushes.White,
                FontWeight = FontWeights.SemiBold,
                FontSize = 13,
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap
            };
            Grid.SetColumn(lblName, 1);

            headerPanel.Children.Add(statusDot);
            headerPanel.Children.Add(lblName);

            // IP Data Row
            Grid ipGrid = new Grid();
            ipGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(40) });
            ipGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            TextBlock lblIpPrefix = new TextBlock
            {
                Text = Strings.NetIpLabel,
                Foreground = Brushes.DarkGray,
                FontSize = 11,
                VerticalAlignment = VerticalAlignment.Center
            };
            Grid.SetColumn(lblIpPrefix, 0);

            // Carousel Layout Container
            StackPanel carouselPanel = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };

            // Left Arrow
            if (ips.Count > 1)
            {
                TextBlock leftArrow = new TextBlock { Text = "<", Foreground = Brushes.Gray, Margin = new Thickness(0, 0, 8, 0), Cursor = Cursors.Hand };
                leftArrow.MouseLeftButtonDown += (s, e) => {
                    _selectedIpIndices[ni.Name] = (selectedIndex - 1 + ips.Count) % ips.Count;
                    _previousIps[ni.Name] = ips[_selectedIpIndices[ni.Name]]; // Prevent false flash animation
                    SaveIpSelectionState(); // Persist to disk instantly
                    UpdateInterfaceData(forceRedraw: true);
                };
                carouselPanel.Children.Add(leftArrow);
            }

            // Actual IP
            TextBlock lblIpValue = new TextBlock
            {
                Text = ips[selectedIndex],
                Foreground = isUp ? Brushes.WhiteSmoke : Brushes.Gray,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 14,
                FontWeight = FontWeights.Bold
            };
            carouselPanel.Children.Add(lblIpValue);

            // Right Arrow & Counter
            if (ips.Count > 1)
            {
                TextBlock rightArrow = new TextBlock { Text = ">", Foreground = Brushes.Gray, Margin = new Thickness(8, 0, 0, 0), Cursor = Cursors.Hand };
                rightArrow.MouseLeftButtonDown += (s, e) => {
                    _selectedIpIndices[ni.Name] = (selectedIndex + 1) % ips.Count;
                    _previousIps[ni.Name] = ips[_selectedIpIndices[ni.Name]]; // Prevent false flash animation
                    SaveIpSelectionState(); // Persist to disk instantly
                    UpdateInterfaceData(forceRedraw: true);
                };
                carouselPanel.Children.Add(rightArrow);

                TextBlock indicator = new TextBlock
                {
                    Text = $" ({selectedIndex + 1}/{ips.Count})",
                    Foreground = Brushes.DimGray,
                    FontSize = 11,
                    Margin = new Thickness(5, 0, 0, 0),
                    VerticalAlignment = VerticalAlignment.Center
                };
                carouselPanel.Children.Add(indicator);
            }

            Grid.SetColumn(carouselPanel, 1);
            ipGrid.Children.Add(lblIpPrefix);
            ipGrid.Children.Add(carouselPanel);

            cardContent.Children.Add(headerPanel);
            cardContent.Children.Add(ipGrid);
            card.Child = cardContent;

            return card;
        }
        public void Pause()
        {
            _localIpTimer?.Stop();
            _publicIpTimer?.Stop();
        }

        public void Resume()
        {
            UpdateInterfaceData(forceRedraw: true);
            _localIpTimer?.Start();
            _publicIpTimer?.Start();
            _ = FetchPublicIpAsync(); // Instant check on resume
        }

        public void Cleanup()
        {
            _localIpTimer?.Stop();
            _publicIpTimer?.Stop();
        }

        // ==========================================
        // UI SETTINGS WINDOW
        // ==========================================
        public void ShowSettingsWindow(Window ownerWindow, dynamic frameData)
        {
            SolidColorBrush accentBrush;
            try { accentBrush = new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor)); }
            catch { accentBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244)); }

            Window win = new Window
            {
                Owner = ownerWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowStyle = WindowStyle.None,
                AllowsTransparency = false,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                SizeToContent = SizeToContent.Height,
                Width = 420
            };

            Border headerBorder = new Border { Height = 50, Background = accentBrush };
            headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) win.DragMove(); };
            Grid headerGrid = new Grid();
            headerGrid.Children.Add(new TextBlock { Text = Strings.NetSettings, Foreground = Brushes.White, FontWeight = FontWeights.Bold, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(15, 0, 0, 0) });
            Button btnClose = new Button { Content = "X", Foreground = Brushes.White, Background = Brushes.Transparent, BorderThickness = new Thickness(0), Width = 40, HorizontalAlignment = HorizontalAlignment.Right };
            btnClose.Click += (s, e) => win.Close();
            headerGrid.Children.Add(btnClose);
            headerBorder.Child = headerGrid;

            Border contentBorder = new Border { Background = Brushes.White, Padding = new Thickness(20) };
            StackPanel contentPanel = new StackPanel();

            // --- EXTERNAL NETWORK SETTINGS ---
            contentPanel.Children.Add(new TextBlock { Text = Strings.NetExternal, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10), Foreground = accentBrush });

            CheckBox chkShowPublicIp = new CheckBox
            {
                Content = Strings.NetShowPublicWan,
                IsChecked = _showPublicIp,
                Margin = new Thickness(0, 0, 0, 20),
                FontWeight = FontWeights.SemiBold
            };
            contentPanel.Children.Add(chkShowPublicIp);

            // --- LOCAL NETWORK SETTINGS ---
            contentPanel.Children.Add(new TextBlock { Text = Strings.NetLocalAdapters, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10), Foreground = accentBrush });

            CheckBox chkShowDisconnected = new CheckBox
            {
                Content = Strings.NetShowDisconnected,
                IsChecked = _showDisconnected,
                Margin = new Thickness(0, 0, 0, 15),
                FontWeight = FontWeights.SemiBold
            };
            contentPanel.Children.Add(chkShowDisconnected);


            contentPanel.Children.Add(new TextBlock
            {
                Text = Strings.NetSelectInterfaces,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 5)
            });

            ScrollViewer interfacesScroll = new ScrollViewer { Height = 180, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
            StackPanel interfacesPanel = new StackPanel { Margin = new Thickness(5) };

            var allInterfaces = NetworkInterface.GetAllNetworkInterfaces()
                .Where(ni => ni.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                .OrderBy(ni => ni.Name)
                .ToList();

            List<CheckBox> interfaceCheckboxes = new List<CheckBox>();

            foreach (var ni in allInterfaces)
            {
                // If whitelist isn't initialized yet (first run), check them all by default
                bool shouldCheck = !_isWhitelistInitialized || _shownInterfaces.Contains(ni.Name);

                CheckBox chk = new CheckBox
                {
                    Content = $"{ni.Name} ({(ni.OperationalStatus == OperationalStatus.Up ? Strings.NetStatusUp : Strings.NetStatusDown)})",
                    Tag = ni.Name,
                    IsChecked = shouldCheck,
                    Margin = new Thickness(0, 5, 0, 5)
                };
                interfaceCheckboxes.Add(chk);
                interfacesPanel.Children.Add(chk);
            }
            interfacesScroll.Content = interfacesPanel;

            Border listBorder = new Border
            {
                BorderBrush = Brushes.LightGray,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Child = interfacesScroll
            };

            contentPanel.Children.Add(listBorder);
            contentBorder.Child = contentPanel;

            Border footerBorder = new Border { Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)), BorderThickness = new Thickness(0, 1, 0, 0), BorderBrush = Brushes.LightGray, Padding = new Thickness(15) };
            StackPanel footerSp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };

            Button btnCancel = new Button { Content = Strings.BtnCancel, Background = Brushes.White, BorderBrush = Brushes.Gray, Width = 80, Height = 30, Margin = new Thickness(0, 0, 10, 0) };
            btnCancel.Click += (s, e) => win.Close();

            Button btnSave = new Button { Content = Strings.BtnSave, Background = accentBrush, Foreground = Brushes.White, FontWeight = FontWeights.Bold, BorderThickness = new Thickness(0), Width = 80, Height = 30 };
            btnSave.Click += (s, e) =>
            {
                _shownInterfaces.Clear();
                foreach (var chk in interfaceCheckboxes)
                {
                    if (chk.IsChecked == true) _shownInterfaces.Add(chk.Tag.ToString());
                }

                _isWhitelistInitialized = true; // Mark as successfully initialized

                Dictionary<string, object> newSettings = new Dictionary<string, object>
                {
                    { "ShowDisconnected", chkShowDisconnected.IsChecked == true },
                    { "ShowPublicIp", chkShowPublicIp.IsChecked == true },
                    { "ShownInterfaces", string.Join(",", _shownInterfaces) }
                    // SelectedIpIndices removed because it's now handled by the local .dat file
                };

                if (frameData is Newtonsoft.Json.Linq.JObject jFrame)
                    jFrame["PluginSettings"] = Newtonsoft.Json.Linq.JObject.FromObject(newSettings);
                else
                    ((IDictionary<string, object>)frameData)["PluginSettings"] = newSettings;

                try { FrameDataManager.SaveFrameData(); } catch { }
                ApplySettingsData(newSettings);

                // If they just enabled the public IP, fetch it immediately
                if (chkShowPublicIp.IsChecked == true) _ = FetchPublicIpAsync();

                UpdateInterfaceData(forceRedraw: true);
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
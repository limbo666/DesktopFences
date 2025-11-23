using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation; // Needed for Animation
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;

namespace Desktop_Fences
{
    public class SearchFormManager : Window
    {
        #region Private Fields
        private TextBox _searchBox;
        private TextBlock _watermark; // "Search..." text
        private TextBlock _appTitle;  // "Desktop Fences +" text
        private WrapPanel _resultsPanel;
        private ScrollViewer _scrollViewer;
        private List<SearchResult> _allShortcuts;
        private static SearchFormManager _instance;
        private bool _isClosing = false;

        // Layout Constants
        private const double WINDOW_WIDTH = 600;
        private const double HEADER_HEIGHT = 50; // Height of search box area
        private const double TITLE_HEIGHT = 30;  // Height of "Desktop Fences +" label
        private const double ITEM_HEIGHT = 90;   // Height of one icon row
        private const double ITEM_WIDTH = 80;    // Width of one icon
        private const int ITEMS_PER_ROW = 7;     // 600 / 80 = 7.5
        private const double MAX_WINDOW_HEIGHT = 600;
        #endregion

        #region Data Structure
        public class SearchResult
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public bool IsFolder { get; set; }
            public string FenceId { get; set; }
            public string FenceTitle { get; set; }
        }
        #endregion

        #region Singleton / Static Access
        public static void ToggleSearch()
        {
            if (_instance == null || !_instance.IsLoaded || _instance._isClosing)
            {
                _instance = new SearchFormManager();
                _instance.Show();
                _instance.Activate();
                _instance.FocusSearch();
            }
            else
            {
                _instance.SafeClose();
            }
        }
        #endregion

        public SearchFormManager()
        {
            InitializeComponent();
            LoadAllShortcuts();
        }

        private void InitializeComponent()
        {
            // 1. Window Setup
            this.Title = "Desktop Fences Search";
            this.Width = WINDOW_WIDTH;
            this.Height = HEADER_HEIGHT + TITLE_HEIGHT; // Initial Compact Size
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.WindowStyle = WindowStyle.None;
            this.AllowsTransparency = true;
            this.Background = Brushes.Transparent;
            this.ResizeMode = ResizeMode.NoResize;
            this.Topmost = true;
            this.ShowInTaskbar = false;

            this.Deactivated += (s, e) => this.SafeClose();

            // 2. Main Card (Rounded, White, Shadow)
            Border mainBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(252, 252, 252)),
                CornerRadius = new CornerRadius(12), // Modern rounded corners
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.FromArgb(40, 0, 0, 0)),
                Effect = new DropShadowEffect { BlurRadius = 25, ShadowDepth = 10, Opacity = 0.4, Color = Colors.Black }
            };

            Grid mainGrid = new Grid();
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(HEADER_HEIGHT) }); // Row 0: Search Input
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Row 1: Dynamic Content

            // --- SEARCH INPUT AREA ---
            Grid searchContainer = new Grid { Margin = new Thickness(10, 0, 10, 0) };

            _watermark = new TextBlock
            {
                Text = "Type to search...",
                FontSize = 18,
                Foreground = Brushes.Gray,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 0, 0),
                IsHitTestVisible = false
            };

            _searchBox = new TextBox
            {
                FontSize = 18,
                FontFamily = new FontFamily("Segoe UI"),
                VerticalContentAlignment = VerticalAlignment.Center,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Foreground = Brushes.Black,
                CaretBrush = Brushes.Black,
                Tag = "Search..."
            };
            _searchBox.TextChanged += OnSearchTextChanged;
            _searchBox.PreviewKeyDown += OnSearchKeyDown;

            searchContainer.Children.Add(_watermark);
            searchContainer.Children.Add(_searchBox);
            Grid.SetRow(searchContainer, 0);

            // --- SEPARATOR LINE ---
            Border separator = new Border
            {
                BorderThickness = new Thickness(0, 0, 0, 1),
                BorderBrush = new SolidColorBrush(Color.FromArgb(30, 0, 0, 0)),
                VerticalAlignment = VerticalAlignment.Bottom
            };
            Grid.SetRow(separator, 0);

            // --- DYNAMIC CONTENT AREA (Row 1) ---
            // This grid holds both the "Title" (when empty) and the "Results" (when searching)
            Grid contentArea = new Grid();
            Grid.SetRow(contentArea, 1);

            // 1. The App Title (Visible initially)
            _appTitle = new TextBlock
            {
                Text = "Desktop Fences +",
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor)),
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 5, 0, 0),
                Opacity = 1.0
            };

            // 2. The Results Panel (Hidden initially)
            _resultsPanel = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                ItemWidth = ITEM_WIDTH,
                ItemHeight = ITEM_HEIGHT,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            _scrollViewer = new ScrollViewer
            {
                Content = _resultsPanel,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(0, 5, 0, 5),
                Opacity = 0.0, // Start hidden
                Visibility = Visibility.Collapsed
            };

            contentArea.Children.Add(_appTitle);
            contentArea.Children.Add(_scrollViewer);

            // Assemble
            mainGrid.Children.Add(searchContainer);
            mainGrid.Children.Add(separator);
            mainGrid.Children.Add(contentArea);

            mainBorder.Child = mainGrid;
            this.Content = mainBorder;
        }

        public void FocusSearch()
        {
            _searchBox.Focus();
        }

        private void SafeClose()
        {
            if (_isClosing) return;
            _isClosing = true;
            this.Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            _isClosing = true;
            if (_instance == this) _instance = null;
        }

        // --- LOGIC: Load Data ---
        private void LoadAllShortcuts()
        {
            _allShortcuts = new List<SearchResult>();
            var fences = FenceManager.GetFenceData();

            foreach (var fence in fences)
            {
                if (fence.ItemsType?.ToString() != "Data") continue;
                JArray items = fence.Items as JArray;
                bool tabsEnabled = fence.TabsEnabled?.ToString().ToLower() == "true";

                if (tabsEnabled)
                {
                    var tabs = fence.Tabs as JArray;
                    if (tabs != null)
                    {
                        foreach (var tab in tabs)
                        {
                            var tabItems = tab["Items"] as JArray;
                            if (tabItems != null) AddItemsToCache(tabItems, fence);
                        }
                    }
                }
                else if (items != null)
                {
                    AddItemsToCache(items, fence);
                }
            }
        }

        private void AddItemsToCache(JArray items, dynamic fence)
        {
            foreach (var item in items)
            {
                string name = item["DisplayName"]?.ToString() ?? System.IO.Path.GetFileNameWithoutExtension(item["Filename"]?.ToString());
                _allShortcuts.Add(new SearchResult
                {
                    Name = name,
                    Path = item["Filename"]?.ToString(),
                    IsFolder = item["IsFolder"]?.ToObject<bool>() ?? false,
                    FenceId = fence.Id?.ToString(),
                    FenceTitle = fence.Title?.ToString()
                });
            }
        }

        // --- LOGIC: Search & Animation ---
        private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
        {
            string query = _searchBox.Text.Trim().ToLower();

            // 1. Toggle Watermark
            _watermark.Visibility = string.IsNullOrEmpty(_searchBox.Text) ? Visibility.Visible : Visibility.Collapsed;

            _resultsPanel.Children.Clear();

            if (string.IsNullOrEmpty(query))
            {
                // STATE: Empty -> Show Title, Shrink Window
                AnimateTransition(showResults: false);
                return;
            }

            // 2. Filter Logic
            var matches = _allShortcuts
                .Where(x => x.Name.ToLower().Contains(query))
                .Take(21) // Limit (3 rows of 7)
                .ToList();

            if (matches.Count == 0)
            {
                // STATE: No Matches -> Show Title (or maybe "No Results"), Shrink
                AnimateTransition(showResults: false);
                return;
            }

            // 3. Populate Results
            foreach (var match in matches)
            {
                // Create Icon UI
                StackPanel sp = new StackPanel
                {
                    Width = ITEM_WIDTH - 10,
                    Height = ITEM_HEIGHT - 10,
                    Margin = new Thickness(5),
                    Cursor = Cursors.Hand,
                    Background = Brushes.Transparent
                };

                ImageSource iconSrc = Utility.GetShellIcon(match.Path, match.IsFolder);
                if (iconSrc == null) iconSrc = new BitmapImage(new Uri("pack://application:,,,/Resources/file-WhiteX.png"));

                Image img = new Image { Source = iconSrc, Width = 32, Height = 32, Margin = new Thickness(0, 5, 0, 2) };

                TextBlock txt = new TextBlock
                {
                    Text = match.Name,
                    TextAlignment = TextAlignment.Center,
                    TextWrapping = TextWrapping.Wrap,
                    FontSize = 11,
                    Foreground = Brushes.Black,
                    MaxHeight = 32,
                    TextTrimming = TextTrimming.CharacterEllipsis
                };

                sp.Children.Add(img);
                sp.Children.Add(txt);
                sp.ToolTip = $"Location: {match.FenceTitle}\nPath: {match.Path}";

                sp.MouseLeftButtonDown += (s, ev) => { if (ev.ClickCount == 1) LaunchResult(match); };
                _resultsPanel.Children.Add(sp);
            }

            // 4. STATE: Matches Found -> Show Results, Expand Window
            AnimateTransition(showResults: true, matchCount: matches.Count);
        }

        private void AnimateTransition(bool showResults, int matchCount = 0)
        {
            // Calculate Target Height
            double targetHeight;
            if (showResults)
            {
                int rows = (int)Math.Ceiling((double)matchCount / ITEMS_PER_ROW);
                // Header + (Rows * ItemHeight) + Padding
                targetHeight = HEADER_HEIGHT + (rows * ITEM_HEIGHT) + 20;
                if (targetHeight > MAX_WINDOW_HEIGHT) targetHeight = MAX_WINDOW_HEIGHT;
            }
            else
            {
                targetHeight = HEADER_HEIGHT + TITLE_HEIGHT;
            }

            // 1. Animate Height
            DoubleAnimation heightAnim = new DoubleAnimation(targetHeight, TimeSpan.FromMilliseconds(250))
            {
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };
            this.BeginAnimation(Window.HeightProperty, heightAnim);

            // 2. Cross-Fade Content
            if (showResults)
            {
                // Fade OUT Title
                _appTitle.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(150)));

                // Fade IN Results
                _scrollViewer.Visibility = Visibility.Visible;
                _scrollViewer.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(1, TimeSpan.FromMilliseconds(200)));
            }
            else
            {
                // Fade IN Title
                _appTitle.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(1, TimeSpan.FromMilliseconds(200)));

                // Fade OUT Results
                var fadeOut = new DoubleAnimation(0, TimeSpan.FromMilliseconds(150));
                fadeOut.Completed += (s, e) => _scrollViewer.Visibility = Visibility.Collapsed;
                _scrollViewer.BeginAnimation(UIElement.OpacityProperty, fadeOut);
            }
        }

        private void OnSearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape) this.SafeClose();

            if (e.Key == Key.Enter && _resultsPanel.Children.Count > 0)
            {
                var firstMatch = _allShortcuts.FirstOrDefault(x => x.Name.ToLower().Contains(_searchBox.Text.Trim().ToLower()));
                if (firstMatch != null) LaunchResult(firstMatch);
            }
        }

        private void LaunchResult(SearchResult result)
        {
            try
            {
                // Resolve Path logic (Same as before)
                string path = result.Path;
                if (!System.IO.Path.IsPathRooted(path))
                {
                    string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                    path = System.IO.Path.Combine(exeDir, path);
                }

                bool isStoreApp = Utility.IsStoreAppShortcut(path);
                bool isUrlFile = path.EndsWith(".url", StringComparison.OrdinalIgnoreCase);
                bool isWebString = result.Path.StartsWith("http", StringComparison.OrdinalIgnoreCase) ||
                                   result.Path.StartsWith("www", StringComparison.OrdinalIgnoreCase);

                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();
                psi.UseShellExecute = true;

                if (isWebString) psi.FileName = result.Path;
                else if (isStoreApp)
                {
                    psi.FileName = "explorer.exe";
                    psi.Arguments = $"\"{path}\"";
                    psi.WorkingDirectory = "";
                }
                else if (result.IsFolder)
                {
                    psi.FileName = "explorer.exe";
                    psi.Arguments = $"\"{path}\"";
                }
                else if (isUrlFile)
                {
                    psi.FileName = path;
                }
                else
                {
                    if (!System.IO.File.Exists(path))
                    {
                        MessageBoxesManager.ShowOKOnlyMessageBoxForm($"File not found: {path}", "Launch Error");
                        return;
                    }
                    psi.FileName = path;
                    try
                    {
                        string dir = System.IO.Path.GetDirectoryName(path);
                        if (!string.IsNullOrEmpty(dir)) psi.WorkingDirectory = dir;
                    }
                    catch { }
                }

                System.Diagnostics.Process.Start(psi);
                this.SafeClose();
            }
            catch (Exception ex)
            {
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Error launching: {ex.Message}", "Error");
            }
        }
    }
}
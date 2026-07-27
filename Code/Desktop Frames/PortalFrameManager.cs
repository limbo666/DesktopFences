using IWshRuntimeLibrary;
using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace Desktop_Frames
{
    public class PortalFramemanager
    {
        // New field for the active filter
        private string _currentFilter = null;
        private int _sortMode = 0; // 0=Name, 1=Date Modified, 2=Type, 3=Size
        private bool _sortAscending = true; // --- STEP 7: Track sort direction for List View ---
        private GridViewColumn _currentSortColumn = null; // --- STEP 7: Track active sorted column ---


        private readonly dynamic _frame;
        private readonly WrapPanel _wpcont;
        private readonly ScrollViewer _parentScrollViewer; // --- NEW: Track parent for dual-mode switching ---
        private ListView _detailsListView; // --- NEW: Tabular view control ---
        private bool _isDetailsView = false; // --- NEW: Active mode flag ---
        private readonly FileSystemWatcher _watcher;
        private string _targetFolderPath;
        private readonly Dispatcher _dispatcher;
        private readonly DispatcherTimer _debounceTimer;
        private int _navigationGeneration = 0; // Tracks active navigation to prevent thread collisions


        private Style GetThemedContextMenuStyle()
        {
            try
            {
                // Safely extract the ID
                string targetId = null;
                try
                {
                    if (_frame is Newtonsoft.Json.Linq.JObject jObj) targetId = jObj["Id"]?.ToString();
                    else targetId = _frame.Id?.ToString();
                }
                catch { }

                // Fetch live data safely to apply real-time color changes
                dynamic liveFrame = _frame;
                if (!string.IsNullOrEmpty(targetId))
                {
                    foreach (dynamic f in FrameDataManager.FrameData)
                    {
                        string fId = null;
                        try { fId = (f is Newtonsoft.Json.Linq.JObject jf) ? jf["Id"]?.ToString() : f.Id?.ToString(); } catch { }
                        if (fId == targetId) { liveFrame = f; break; }
                    }
                }

                // Determine Background Color
                string frameColorName = null;
                try
                {
                    if (liveFrame is Newtonsoft.Json.Linq.JObject jf) frameColorName = jf["CustomColor"]?.ToString();
                    else frameColorName = liveFrame.CustomColor?.ToString();
                }
                catch { }

                if (string.IsNullOrEmpty(frameColorName)) frameColorName = SettingsManager.SelectedColor;

                System.Windows.Media.Color bgColor = System.Windows.Media.Colors.Gray;
                try
                {
                    var drawingColor = Utility.GetColorFromName(frameColorName);
                    bgColor = System.Windows.Media.Color.FromArgb(255, drawingColor.R, drawingColor.G, drawingColor.B);
                }
                catch { }

                // Determine Foreground Color
                double luminance = (0.299 * bgColor.R + 0.587 * bgColor.G + 0.114 * bgColor.B) / 255.0;
                System.Windows.Media.Color fgColor = luminance > 0.5 ? System.Windows.Media.Colors.Black : System.Windows.Media.Colors.White;

                string bgHex = $"#{bgColor.A:X2}{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
                string fgHex = $"#{fgColor.A:X2}{fgColor.R:X2}{fgColor.G:X2}{fgColor.B:X2}";

                string xamlStyle = $@"
                <Style xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' 
                       xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml' 
                       TargetType='ContextMenu'>
                    <Style.Resources>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.MenuBrushKey}}' Color='{bgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.ControlBrushKey}}' Color='{bgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.WindowBrushKey}}' Color='{bgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.MenuTextBrushKey}}' Color='{fgHex}'/>
                        <SolidColorBrush x:Key='{{x:Static SystemColors.ControlTextBrushKey}}' Color='{fgHex}'/>
<Style TargetType='MenuItem'>
                            <Setter Property='Background' Value='{bgHex}'/>
                            <Setter Property='Foreground' Value='{fgHex}'/>
                            <Setter Property='BorderThickness' Value='0'/>
                            <Setter Property='BorderBrush' Value='Transparent'/>
                        </Style>
                    </Style.Resources>
                    <Setter Property='Background' Value='{bgHex}'/>
                    <Setter Property='Foreground' Value='{fgHex}'/>
                    <Setter Property='Template'>
                        <Setter.Value>
                            <ControlTemplate TargetType='ContextMenu'>
                                <Border Background='{{TemplateBinding Background}}' BorderBrush='#555555' BorderThickness='1' Padding='1'>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width='30'/>
                                            <ColumnDefinition Width='*'/>
                                        </Grid.ColumnDefinitions>
                                        <Border Grid.Column='0' Background='#151515'>
                                            <Image Source='pack://application:,,,/Resources/DesktopFramesVertical.png' Stretch='Uniform' VerticalAlignment='Top' Margin='0,5,0,0'/>
                                        </Border>
                                        <ScrollViewer Grid.Column='1' Margin='2,0,0,0' VerticalScrollBarVisibility='Hidden'>
                                            <ItemsPresenter KeyboardNavigation.DirectionalNavigation='Cycle'/>
                                        </ScrollViewer>
                                    </Grid>
                                </Border>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>";

                return (Style)System.Windows.Markup.XamlReader.Parse(xamlStyle);
            }
            catch { return null; }
        }



        // --- FILTERING ENGINE START ---

        /// <summary>
        /// Updates the current filter and refreshes the visibility of all items.
        /// Publicly called by Framemanager when the user types in the filter bar.
        /// </summary>
        /// 
        // --- COMPILER FIX: Native safe property resolver ---
        private string GetSafeProperty(dynamic obj, string propertyName)
        {
            try
            {
                if (obj is IDictionary<string, object> dict)
                    return dict.ContainsKey(propertyName) ? dict[propertyName]?.ToString() : null;
                if (obj is JObject jObj)
                    return jObj[propertyName]?.ToString();
                return obj?.GetType().GetProperty(propertyName)?.GetValue(obj)?.ToString();
            }
            catch { return null; }
        }
        public void ApplyFilter(string filterText)
        {
            _currentFilter = filterText;
            _dispatcher.Invoke(() =>
            {
                // --- STEP 5: Apply Filter to Details View ---
                if (_isDetailsView && _detailsListView != null)
                {
                    var view = System.Windows.Data.CollectionViewSource.GetDefaultView(_detailsListView.Items);
                    if (view != null)
                    {
                        view.Filter = item =>
                        {
                            if (item is PortalItemModel model && !string.IsNullOrEmpty(model.FullPath))
                            {
                                return ShouldShowItem(model.FullPath);
                            }
                            return true;
                        };
                        view.Refresh();
                    }
                    return;
                }
                // --------------------------------------------
                foreach (StackPanel sp in _wpcont.Children.OfType<StackPanel>())
                {
                    if (sp.Tag != null)
                    {
                        // Safely retrieve path from anonymous type or object
                        string path = sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)?.ToString();
                        if (!string.IsNullOrEmpty(path))
                        {
                            sp.Visibility = ShouldShowItem(path) ? Visibility.Visible : Visibility.Collapsed;
                        }
                    }
                }
            });
        }



        /// <summary>
        /// Determines if a file should be visible based on the current filter.
        /// Supports "Smart Match" if NoWildcardsOnPortalFilter is enabled.
        /// </summary>
        private bool ShouldShowItem(string filePath)
        {
            if (string.IsNullOrWhiteSpace(_currentFilter)) return true;

            string fileName = System.IO.Path.GetFileName(filePath);
            var terms = _currentFilter.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(t => t.Trim())
                                      .ToList();

            bool hasIncludeRules = terms.Any(t => !t.StartsWith(">"));
            bool matchesInclude = !hasIncludeRules;
            bool matchesExclude = false;

            foreach (var term in terms)
            {
                if (string.IsNullOrEmpty(term)) continue;

                string pattern = term;
                bool isExclude = false;

                // 1. Identify Exclusion
                if (pattern.StartsWith(">"))
                {
                    isExclude = true;
                    pattern = pattern.Substring(1); // Remove '>' prefix
                }

                // 2. Apply Smart Wildcards (Hidden Option)
                // Logic: If user wants "No Wildcards", we treat text as "Contains".
                // We only auto-wrap if the user hasn't typed wildcards themselves.
                if (SettingsManager.NoWildcardsOnPortalFilter)
                {
                    if (!pattern.Contains("*") && !pattern.Contains("?"))
                    {
                        pattern = "*" + pattern + "*";
                    }
                }

                // 3. Match
                if (isExclude)
                {
                    if (IsMatch(fileName, pattern))
                    {
                        matchesExclude = true;
                        break; // Hard fail
                    }
                }
                else
                {
                    if (IsMatch(fileName, pattern))
                    {
                        matchesInclude = true;
                    }
                }
            }

            return !matchesExclude && matchesInclude;
        }




        /// <summary>
        /// Simple glob matching (* and ?)
        /// </summary>
        private bool IsMatch(string text, string pattern)
        {
            // Use VB's Like operator or simple Regex. 
            // For a dependency-free C# solution, we convert glob to regex.
            try
            {
                string regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
                                      .Replace(@"\*", ".*")
                                      .Replace(@"\?", ".") + "$";
                return System.Text.RegularExpressions.Regex.IsMatch(text, regexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            }
            catch { return false; }
        }
        // --- FILTERING ENGINE END ---


        // --- SORTING ENGINE START ---
        public bool IsDetailsView => _isDetailsView; // --- STEP 6: Expose view mode cleanly ---

        public string CycleSortMode()
        {
            // --- STEP 6 FIX: Completely disable CTRL+Click sort cycling when in Details/List View ---
            if (_isDetailsView) return null;

            _sortMode++;
            if (_sortMode > 3) _sortMode = 0;

            // Save state using the existing updater
            Framemanager.UpdateFrameProperty(_frame, "SortMode", _sortMode.ToString(), "Updated portal sort mode");

            string[] modeNames = { "Name", "Date Modified", "Type", "Size" };
            string activeMode = modeNames[_sortMode];

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Portal frame sorted by: {activeMode}");

            SortContents();

            return activeMode;
        }

        // --- STEP 8: Column Persistence Engine (Width, Order, Visibility) ---
        private DispatcherTimer _columnSaveTimer = null;

        private void TriggerColumnSave()
        {
            if (_columnSaveTimer == null)
            {
                _columnSaveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                _columnSaveTimer.Tick += (s, e) => { _columnSaveTimer.Stop(); SaveColumnState(); };
            }
            _columnSaveTimer.Stop();
            _columnSaveTimer.Start();
        }

        private void SaveColumnState()
        {
            if (!_isDetailsView || _detailsListView?.View as GridView == null) return;
            var gridView = _detailsListView.View as GridView;

            var parts = new List<string>();
            foreach (var col in gridView.Columns)
            {
                string cleanName = col.Header?.ToString().Replace(" ▲", "").Replace(" ▼", "") ?? "";
                if (string.IsNullOrEmpty(cleanName)) continue;
                double w = col.Width;
                string wStr = double.IsNaN(w) ? "-1" : w.ToString("F0");
                parts.Add($"{cleanName}:{wStr}");
            }
            string savedString = string.Join(";", parts);
            Framemanager.UpdateFrameProperty(_frame, "DetailsViewColumns", savedString, "Saved Details View columns");
            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Saved portal column state: {savedString}");
        }

        private void LoadColumnState(GridView gridView, string overrideString = null)
        {
            if (gridView == null) return;

            string savedColumns = overrideString;
            if (savedColumns == null)
            {
                try
                {
                    IDictionary<string, object> frameDict = _frame is IDictionary<string, object> dict ? dict : ((Newtonsoft.Json.Linq.JObject)_frame).ToObject<IDictionary<string, object>>();
                    savedColumns = frameDict.ContainsKey("DetailsViewColumns") ? frameDict["DetailsViewColumns"]?.ToString() : null;
                }
                catch { }
            }

            if (string.IsNullOrEmpty(savedColumns)) return;

            try
            {
                var existingCols = gridView.Columns.ToList();
                gridView.Columns.Clear();

                var entries = savedColumns.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                var processedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var entry in entries)
                {
                    var parts = entry.Split(':');
                    if (parts.Length != 2) continue;
                    string colName = parts[0];
                    if (double.TryParse(parts[1], out double w))
                    {
                        var match = existingCols.FirstOrDefault(c => string.Equals(c.Header?.ToString().Replace(" ▲", "").Replace(" ▼", ""), colName, StringComparison.OrdinalIgnoreCase));
                        if (match != null)
                        {
                            match.Width = (w == -1) ? double.NaN : w;
                            gridView.Columns.Add(match);
                            processedNames.Add(colName);
                        }
                    }
                }

                // Append any fallback columns that weren't in the saved string
                foreach (var col in existingCols)
                {
                    string cleanName = col.Header?.ToString().Replace(" ▲", "").Replace(" ▼", "") ?? "";
                    if (!processedNames.Contains(cleanName))
                    {
                        gridView.Columns.Add(col);
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error loading column state: {ex.Message}");
            }
        }
        // --------------------------------------------------------------------



        // --- STEP 7: Handle Column Header Clicking & Arrow Indicators ---
        private void SortDetailsViewColumn(GridViewColumnHeader headerClicked)
        {
            GridView gridView = _detailsListView?.View as GridView;
            if (gridView == null || headerClicked?.Column == null) return;
            string headerText = headerClicked.Column.Header?.ToString();
            if (string.IsNullOrEmpty(headerText) || headerText.StartsWith("#")) return; // Ignore index column

            string cleanHeader = headerText.Replace(" ▲", "").Replace(" ▼", "");

            // Toggle direction if clicking the same column, otherwise default to Ascending
            if (_currentSortColumn == headerClicked.Column)
            {
                _sortAscending = !_sortAscending;
            }
            else
            {
                _currentSortColumn = headerClicked.Column;
                _sortAscending = true;
            }

            // Strip arrows from all columns
            foreach (var col in gridView.Columns)
            {
                if (col.Header != null)
                {
                    col.Header = col.Header.ToString().Replace(" ▲", "").Replace(" ▼", "");
                }
            }

            // Apply arrow to clicked column
            headerClicked.Column.Header = cleanHeader + (_sortAscending ? " ▲" : " ▼");

            // Map column name to _sortMode
            switch (cleanHeader)
            {
                case "Date modified": _sortMode = 1; break;
                case "Type": _sortMode = 2; break;
                case "Size": _sortMode = 3; break;
                default: _sortMode = 0; break; // Name
            }

            Framemanager.UpdateFrameProperty(_frame, "SortMode", _sortMode.ToString(), "Updated portal sort mode from header click");
            SortContents();
        }

        private void SortContents()
        {
            _dispatcher.Invoke(() =>
            {
                // --- STEP 7: Branch Sorting for Details View with Direction Support ---
                if (_isDetailsView && _detailsListView != null)
                {
                    var listItems = _detailsListView.Items.OfType<PortalItemModel>().ToList();
                    if (listItems.Count == 0) return;

                    IEnumerable<PortalItemModel> sortedList;
                    switch (_sortMode)
                    {
                        case 1:
                            sortedList = _sortAscending
                                ? listItems.OrderByDescending(i => i.IsFolder).ThenBy(i => i.RawDateModified)
                                : listItems.OrderByDescending(i => i.IsFolder).ThenByDescending(i => i.RawDateModified);
                            break;
                        case 2:
                            sortedList = _sortAscending
                                ? listItems.OrderByDescending(i => i.IsFolder).ThenBy(i => i.Type)
                                : listItems.OrderByDescending(i => i.IsFolder).ThenByDescending(i => i.Type);
                            break;
                        case 3:
                            sortedList = _sortAscending
                                ? listItems.OrderByDescending(i => i.IsFolder).ThenBy(i => i.RawSizeBytes)
                                : listItems.OrderByDescending(i => i.IsFolder).ThenByDescending(i => i.RawSizeBytes);
                            break;
                        default:
                            sortedList = _sortAscending
                                ? listItems.OrderByDescending(i => i.IsFolder).ThenBy(i => i.Name)
                                : listItems.OrderByDescending(i => i.IsFolder).ThenByDescending(i => i.Name);
                            break;
                    }

                    _detailsListView.Items.Clear();
                    foreach (var item in sortedList) _detailsListView.Items.Add(item);
                    return;
                }
                // -----------------------------------------------

                var children = _wpcont.Children.OfType<StackPanel>().ToList();
                if (children.Count == 0) return;

                _wpcont.Children.Clear();

                string GetPath(StackPanel sp)
                {
                    dynamic tag = sp.Tag;
                    return tag?.GetType().GetProperty("FilePath")?.GetValue(tag)?.ToString() ?? "";
                }

                bool IsFolder(StackPanel sp)
                {
                    dynamic tag = sp.Tag;
                    return tag != null && tag.GetType().GetProperty("IsFolder")?.GetValue(tag) as bool? == true;
                }

                IEnumerable<StackPanel> sorted;

                switch (_sortMode)
                {
                    case 1: // Date Modified (Newest first)
                        sorted = children.OrderByDescending(IsFolder)
                                         .ThenByDescending(sp => { try { return System.IO.File.GetLastWriteTime(GetPath(sp)); } catch { return DateTime.MinValue; } });
                        break;
                    case 2: // Type (A-Z)
                        sorted = children.OrderByDescending(IsFolder)
                                         .ThenBy(sp => System.IO.Path.GetExtension(GetPath(sp))?.ToLower() ?? "");
                        break;
                    case 3: // Size (Largest first)
                        sorted = children.OrderByDescending(IsFolder)
                                         .ThenByDescending(sp => { try { return IsFolder(sp) ? 0 : new System.IO.FileInfo(GetPath(sp)).Length; } catch { return 0; } });
                        break;
                    default: // 0 = Name (A-Z)
                        sorted = children.OrderByDescending(IsFolder)
                                         .ThenBy(sp => System.IO.Path.GetFileName(GetPath(sp))?.ToLower() ?? "");
                        break;
                }

                foreach (var sp in sorted)
                {
                    _wpcont.Children.Add(sp);
                }
            });
        }
        // --- SORTING ENGINE END ---


        public PortalFramemanager(dynamic frame, WrapPanel wpcont)
        {
            _frame = frame;
            _wpcont = wpcont;
            _dispatcher = _wpcont.Dispatcher;

            // Initialize debounce timer with longer interval for Excel temp files
            _debounceTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500) // Increased for better stability
            };
            _debounceTimer.Tick += ProcessPendingEvents;

            // Extract folder path and properties safely
            IDictionary<string, object> frameDict = frame is IDictionary<string, object> dict ? dict : ((JObject)frame).ToObject<IDictionary<string, object>>();
            _targetFolderPath = frameDict.ContainsKey("Path") ? frameDict["Path"]?.ToString() : null;

            if (frameDict.ContainsKey("FilterString"))
                _currentFilter = frameDict["FilterString"]?.ToString();

            if (frameDict.ContainsKey("SortMode"))
                _sortMode = Convert.ToInt32(frameDict["SortMode"]?.ToString() ?? "0");

            // --- COMPILER FIX: Single, clean ViewMode initialization ---
            _parentScrollViewer = _wpcont.Parent as ScrollViewer;
            string activeViewMode = frameDict.ContainsKey("ViewMode") ? frameDict["ViewMode"]?.ToString() : null;
            if (string.IsNullOrEmpty(activeViewMode)) activeViewMode = SettingsManager.DefaultPortalView;
            _isDetailsView = string.Equals(activeViewMode, "Details", StringComparison.OrdinalIgnoreCase);

            if (string.IsNullOrEmpty(_targetFolderPath))
            {
                throw new Exception("No folder path defined for Portal Frame. Please recreate the frame.");
            }

            if (!Directory.Exists(_targetFolderPath))
            {
                throw new Exception($"The folder '{_targetFolderPath}' does not exist. Please update the Portal Frame settings.");
            }

            _watcher = new FileSystemWatcher(_targetFolderPath)
            {
                EnableRaisingEvents = true,
                IncludeSubdirectories = false,
                InternalBufferSize = 65536 // --- BUG FIX: Maximize buffer to survive massive I/O operations ---
            };

            // --- BUG FIX: Simplified "State Reconciler" Trigger ---
            // The watcher is now just a "ping" to tell us something changed. 
            // We listen to the Error event to specifically catch buffer overflows!
            _watcher.Created += (s, e) => TriggerSync();
            _watcher.Deleted += (s, e) => TriggerSync();
            _watcher.Renamed += (s, e) => TriggerSync();
            _watcher.Error += (s, e) => TriggerSync();

            InitializeFrameContents();
            //  // --- TEST CODE START ---
            //  // Hardcode a filter to prove the engine works.
            //   // This simulates a user typing "*.txt" into the filter bar.
            //   ApplyFilter("*.txt");
            //  // --- TEST CODE END ---
        }

        private void TriggerSync(bool immediate = false)
        {
            _dispatcher.InvokeAsync(() =>
            {
                _debounceTimer.Stop();
                if (immediate)
                {
                    _ = RunReconcilerAsync();
                }
                else
                {
                    _debounceTimer.Start();
                }
            });
        }

        private void ProcessPendingEvents(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            _ = RunReconcilerAsync();
        }

        private async System.Threading.Tasks.Task RunReconcilerAsync()
        {
            int myGeneration = ++_navigationGeneration;
            string targetPath = _targetFolderPath;

            try
            {
                if (!Directory.Exists(targetPath)) return;

                // 1. Read Disk & UI State (Background Thread)
                var diff = await System.Threading.Tasks.Task.Run(() =>
                {
                    // --- PERFORMANCE FIX: Use EnumerateFileSystemInfos to avoid N+1 disk I/O calls ---
                    // This fetches attributes instantly during the directory scan instead of hitting the disk for every single file.
                    var currentDiskFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    try
                    {
                        var dirInfo = new DirectoryInfo(targetPath);
                        foreach (var fsi in dirInfo.EnumerateFileSystemInfos())
                        {
                            if (CoreUtilities.IsTemporaryFile(fsi.FullName)) continue;

                            // Attributes are pre-cached in 'fsi', ZERO extra disk I/O required!
                            if ((fsi.Attributes & FileAttributes.Hidden) == 0 &&
                                (fsi.Attributes & FileAttributes.System) == 0)
                            {
                                currentDiskFiles.Add(fsi.FullName);
                            }
                        }
                    }
                    catch { } // Handle access denied gracefully

                    List<string> currentUIFiles = new List<string>();
                    _dispatcher.Invoke(() =>
                    {
                        // --- STEP 5: Scan Details View if active ---
                        if (_isDetailsView && _detailsListView != null)
                        {
                            currentUIFiles = _detailsListView.Items.OfType<PortalItemModel>()
                                .Select(i => i.FullPath)
                                .Where(p => p != null).ToList();
                        }
                        else
                        {
                            currentUIFiles = _wpcont.Children.OfType<StackPanel>()
                                .Select(sp => sp.Tag?.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)?.ToString())
                                .Where(p => p != null).ToList();
                        }
                    });

                    var uiSet = currentUIFiles.ToHashSet(StringComparer.OrdinalIgnoreCase);
                    return new
                    {
                        ToRemove = currentUIFiles.Where(p => !currentDiskFiles.Contains(p)).ToList(),
                        ToAdd = currentDiskFiles.Where(p => !uiSet.Contains(p)).ToList()
                    };
                });

                // Abort if the user navigated away while we were scanning!
                if (myGeneration != _navigationGeneration) return;

                // 2. Remove old icons instantly
                if (diff.ToRemove.Count > 0)
                {
                    _dispatcher.Invoke(() =>
                    {
                        foreach (var path in diff.ToRemove) RemoveIcon(path);
                    });
                }

                if (myGeneration != _navigationGeneration) return;

                // 3. Add new icons (SMOOTH CHUNKING)
                // Instead of locking the UI thread to load 100 icons at once, we yield to the Background priority.
                // This keeps the app responsive during massive folder loads and avoids freezing.
                if (diff.ToAdd.Count > 0)
                {
                    foreach (var path in diff.ToAdd)
                    {
                        // Stop immediately if user navigated to another folder
                        if (myGeneration != _navigationGeneration) break;

                        await _dispatcher.InvokeAsync(() =>
                        {
                            AddIcon(path);
                        }, DispatcherPriority.Background);
                    }

                    if (myGeneration == _navigationGeneration)
                    {
                        _dispatcher.Invoke(() => SortContents());
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Portal Sync Error: {ex.Message}");
            }
        }

        private void AddIcon(string path)
        {



            // Enhanced filter during add to prevent duplicates
            FileAttributes attributes;
            bool isFolder = false;

            try
            {
                // --- PERFORMANCE FIX: 1 Disk Read instead of 4 ---
                // We grab attributes once. This immediately tells us if it exists, is a folder, and if it's hidden/system.
                attributes = System.IO.File.GetAttributes(path);
                isFolder = (attributes & FileAttributes.Directory) == FileAttributes.Directory;

                if (CoreUtilities.IsTemporaryFile(path)) return;
                if ((attributes & FileAttributes.Hidden) == FileAttributes.Hidden) return;
                if ((attributes & FileAttributes.System) == FileAttributes.System) return;

                // Check if icon already exists in UI (Safety Check)
                var existingPanel = _wpcont.Children.OfType<StackPanel>()
                    .FirstOrDefault(sp => sp.Tag != null &&
                                    sp.Tag.GetType().GetProperty("FilePath")?.GetValue(sp.Tag)?.ToString() == path);

                if (existingPanel != null) return;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Inaccessible or missing item {path}: {ex.Message}");
                return;
            }

            dynamic icon = new System.Dynamic.ExpandoObject();
            IDictionary<string, object> iconDict = icon;
            iconDict["Filename"] = path;
            iconDict["IsFolder"] = isFolder;





            // --- RESTORED: Network Path Detection ---
            iconDict["IsNetwork"] = Framemanager.IsNetworkPath(path);


            string displayName;

            try
            {
                // FIX: Handle Extensions based on Global Setting
                if (SettingsManager.ShowPortalExtensions && !isFolder)
                {
                    // Force display name WITH extension
                    displayName = Path.GetFileName(path);
                }
                else
                {
                    if (isFolder)
                    {
                        // Folders → keep full name even if they contain dots
                        displayName = Path.GetFileName(path);
                    }
                    else
                    {
                        // Files → strip extension (default behavior)
                        displayName = Path.GetFileNameWithoutExtension(path);
                    }
                }
            }
            catch
            {
                // Fallback: act like it's a file without extension
                displayName = Path.GetFileNameWithoutExtension(path);
            }

            iconDict["DisplayName"] = displayName;

            // --- STEP 4: Branch Rendering for Details View ---
            if (_isDetailsView)
            {
                BuildOrSwitchToDetailsView();

                // Safety check against duplicates
                if (_detailsListView.Items.OfType<PortalItemModel>().Any(i => string.Equals(i.FullPath, path, StringComparison.OrdinalIgnoreCase)))
                    return;

                var itemModel = new PortalItemModel
                {
                    Index = _detailsListView.Items.Count + 1,
                    Name = displayName,
                    FullPath = path,
                    IsFolder = isFolder,
                    Icon = Utility.GetShellIcon(path, isFolder),
                    Type = isFolder ? "File folder" : "Loading...",
                    Size = isFolder ? "" : "..."
                };

                // Build Item Context Menu for this row
                ContextMenu contextMenu = new ContextMenu();

                Style themedStyle = GetThemedContextMenuStyle();
                if (themedStyle != null) contextMenu.Style = themedStyle;

                MenuItem copyFileItem = new MenuItem { Header = "Copy Item" };
                copyFileItem.Click += (s, e) =>
                {
                    try
                    {
                        var paths = new System.Collections.Specialized.StringCollection { path };
                        Clipboard.SetFileDropList(paths);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Copied item to clipboard: {path}");
                    }
                    catch (Exception ex) { LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error copying item: {ex.Message}"); }
                };
                contextMenu.Items.Add(copyFileItem);

                MenuItem cutFileItem = new MenuItem { Header = "Cut Item" };
                cutFileItem.Click += (s, e) =>
                {
                    try
                    {
                        var paths = new System.Collections.Specialized.StringCollection { path };
                        DataObject data = new DataObject();
                        data.SetFileDropList(paths);
                        byte[] moveEffect = new byte[] { 2, 0, 0, 0 };
                        data.SetData("Preferred DropEffect", new System.IO.MemoryStream(moveEffect));
                        Clipboard.SetDataObject(data, true);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Cut item to clipboard: {path}");
                    }
                    catch (Exception ex) { LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error cutting item: {ex.Message}"); }
                };
                contextMenu.Items.Add(cutFileItem);

                MenuItem renameItem = new MenuItem { Header = "Rename item" };
                renameItem.Click += (s, e) => RenameItem(path, null);
                contextMenu.Items.Add(renameItem);

                MenuItem deleteItem = new MenuItem { Header = "Delete item" };
                deleteItem.Click += (s, e) =>
                {
                    DeleteItem(path, null);
                    if (_detailsListView.Items.Contains(itemModel)) _detailsListView.Items.Remove(itemModel);
                };
                contextMenu.Items.Add(deleteItem);

                if (!isFolder)
                {
                    MenuItem openWithItem = new MenuItem { Header = "Open with..." };
                    openWithItem.Click += (s, e) =>
                    {
                        try
                        {
                            // Modern approach: Native Windows shell verb
                            var psi = new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = path,
                                UseShellExecute = true,
                                Verb = "openas"
                            };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch
                        {
                            try
                            {
                                // Legacy fallback
                                System.Diagnostics.Process.Start("rundll32.exe", $"shell32.dll,OpenAs_RunDLL {path}");
                            }
                            catch (Exception ex) { LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error opening 'Open with' dialog: {ex.Message}"); }
                        }
                    };
                    contextMenu.Items.Add(openWithItem);
                }

                contextMenu.Items.Add(new Separator());

                MenuItem copyItemPathItem = new MenuItem { Header = "Copy item path" };
                copyItemPathItem.Click += (s, e) =>
                {
                    try
                    {
                        Clipboard.SetText(path);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Copied item path to clipboard: {path}");
                    }
                    catch (Exception ex) { LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error copying item path: {ex.Message}"); }
                };
                contextMenu.Items.Add(copyItemPathItem);
                MenuItem copyPathItem = new MenuItem { Header = "Copy folder path" };
                copyPathItem.Click += (s, e) => CopyPathOrTarget(path);
                contextMenu.Items.Add(copyPathItem);

                // --- LIVE THEME FIX ---
                contextMenu.Opened += (s, e) =>
                {
                    try
                    {
                        Style liveStyle = Framemanager.GetThemedContextMenuStyle(_frame);
                        if (liveStyle != null) contextMenu.Style = liveStyle;
                    }
                    catch { }
                };

                itemModel.RowMenu = contextMenu;
                _detailsListView.Items.Add(itemModel);

                // --- INSTANT ASYNC HYDRATION (Solves "Loading..." Hang) ---
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        if (!isFolder)
                        {
                            var fi = new System.IO.FileInfo(path);
                            long len = fi.Length;
                            string sizeStr = len > 1048576 ? $"{len / 1048576.0:F1} MB" : $"{len / 1024.0:F0} KB";
                            string typeStr = $"{fi.Extension.TrimStart('.').ToUpper()} File";
                            var time = System.IO.File.GetLastWriteTime(path);

                            _dispatcher.InvokeAsync(() =>
                            {
                                itemModel.RawSizeBytes = len;
                                itemModel.Size = sizeStr;
                                itemModel.Type = typeStr;
                                itemModel.RawDateModified = time;
                                itemModel.DateModified = time.ToString("yyyy-MM-dd HH:mm");
                            });
                        }
                        else
                        {
                            var time = System.IO.File.GetLastWriteTime(path);
                            _dispatcher.InvokeAsync(() =>
                            {
                                itemModel.RawDateModified = time;
                                itemModel.DateModified = time.ToString("yyyy-MM-dd HH:mm");
                            });
                        }
                    }
                    catch { }
                });

                return;
            }
            // -------------------------------------------------

            // --- FIX: ONE CALL ONLY ---
            // We use the new signature that passes '_frame' context.
            // This applies the custom settings (Size, Color, etc.) immediately.
            Framemanager.AddIcon(icon, _wpcont, _frame);

            // Now we grab the StackPanel that was just added to attach logic
            StackPanel sp = _wpcont.Children[_wpcont.Children.Count - 1] as StackPanel;
            if (sp != null)
            {
                // FIX: Apply filter immediately upon creation
                sp.Visibility = ShouldShowItem(path) ? Visibility.Visible : Visibility.Collapsed;

                Framemanager.ClickEventAdder(sp, path, Directory.Exists(path));


                // Create and attach context menu
                ContextMenu contextMenu = new ContextMenu();

                Style themedStyle = GetThemedContextMenuStyle();
                if (themedStyle != null) contextMenu.Style = themedStyle;

                // 1. Copy Item (File Object)
                MenuItem copyFileItem = new MenuItem { Header = "Copy Item" };
                copyFileItem.Click += (s, e) =>
                {
                    try
                    {
                        // Add file to clipboard as a FileDropList (Standard Windows Copy)
                        var paths = new System.Collections.Specialized.StringCollection();
                        paths.Add(path);
                        Clipboard.SetFileDropList(paths);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Copied item to clipboard: {path}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error copying item: {ex.Message}");
                    }
                };
                contextMenu.Items.Add(copyFileItem);

                // 2. Cut Item (File Object with Move Effect)
                MenuItem cutFileItem = new MenuItem { Header = "Cut Item" };
                cutFileItem.Click += (s, e) =>
                {
                    try
                    {
                        var paths = new System.Collections.Specialized.StringCollection();
                        paths.Add(path);

                        // Create a DataObject to hold both the file list and the "Move" flag
                        DataObject data = new DataObject();
                        data.SetFileDropList(paths);

                        // Set "Preferred DropEffect" to Move (Byte value 2)
                        // This tells Windows Explorer to perform a MOVe operation on Paste
                        byte[] moveEffect = new byte[] { 2, 0, 0, 0 };
                        System.IO.MemoryStream stream = new System.IO.MemoryStream(moveEffect);
                        data.SetData("Preferred DropEffect", stream);

                        Clipboard.SetDataObject(data, true);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Cut item to clipboard: {path}");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error cutting item: {ex.Message}");
                    }
                };
                contextMenu.Items.Add(cutFileItem);

                // 3. Rename item (Existing)
                MenuItem renameItem = new MenuItem { Header = "Rename item" };
                renameItem.Click += (s, e) => RenameItem(path, sp);
                contextMenu.Items.Add(renameItem);

                // 4. Delete item (Existing)
                MenuItem deleteItem = new MenuItem { Header = "Delete item" };
                deleteItem.Click += (s, e) => DeleteItem(path, sp);
                contextMenu.Items.Add(deleteItem);

                // 4.5 Open with (Files only)
                if (!isFolder)
                {
                    MenuItem openWithItem = new MenuItem { Header = "Open with..." };
                    openWithItem.Click += (s, e) =>
                    {
                        try
                        {
                            // Modern approach: Native Windows shell verb
                            var psi = new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = path,
                                UseShellExecute = true,
                                Verb = "openas"
                            };
                            System.Diagnostics.Process.Start(psi);
                        }
                        catch
                        {
                            try
                            {
                                // Legacy fallback
                                System.Diagnostics.Process.Start("rundll32.exe", $"shell32.dll,OpenAs_RunDLL {path}");
                            }
                            catch (Exception ex) { LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error opening 'Open with' dialog: {ex.Message}"); }
                        }
                    };
                    contextMenu.Items.Add(openWithItem);
                }

                // 5. Separator
                contextMenu.Items.Add(new Separator());

                // 6. Copy item path
                MenuItem copyItemPathItem = new MenuItem { Header = "Copy item path" };
                copyItemPathItem.Click += (s, e) =>
                {
                    try
                    {
                        Clipboard.SetText(path);
                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Copied item path to clipboard: {path}");
                    }
                    catch (Exception ex) { LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error copying item path: {ex.Message}"); }
                };
                contextMenu.Items.Add(copyItemPathItem);

                // 7. Copy folder path
                MenuItem copyPathItem = new MenuItem { Header = "Copy folder path" };
                copyPathItem.Click += (s, e) => CopyPathOrTarget(path);
                contextMenu.Items.Add(copyPathItem);

                // --- LIVE THEME FIX ---
                contextMenu.Opened += (s, e) =>
                {
                    try
                    {
                        Style liveStyle = Framemanager.GetThemedContextMenuStyle(_frame);
                        if (liveStyle != null) contextMenu.Style = liveStyle;
                    }
                    catch { }
                };

                sp.ContextMenu = contextMenu;



            }
        }

        private void RenameItem(string currentPath, StackPanel sp)
        {
            try
            {
                bool isFolder = Directory.Exists(currentPath);

                string currentName;
                string extension;

                if (isFolder)
                {
                    // Folders don't use extensions; dots are just part of the folder name
                    currentName = Path.GetFileName(currentPath);
                    extension = "";
                }
                else
                {
                    // Files protect their extensions during rename
                    currentName = Path.GetFileNameWithoutExtension(currentPath);
                    extension = Path.GetExtension(currentPath);
                }

                // Simple input dialog (you can replace with a proper dialog if you have one)
                string newName = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter new name:",
                    "Rename Item",
                    currentName);

                if (string.IsNullOrEmpty(newName) || newName == currentName)
                    return;

                string newPath = Path.Combine(Path.GetDirectoryName(currentPath), newName + extension);

                // Check if target name already exists
                if (System.IO.File.Exists(newPath) || Directory.Exists(newPath))
                {
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm("A file or folder with that name already exists.", "Rename Error");
                    return;
                }

                // Perform the rename
                if (Directory.Exists(currentPath))
                {
                    Directory.Move(currentPath, newPath);
                }
                else if (System.IO.File.Exists(currentPath))
                {
                    System.IO.File.Move(currentPath, newPath);
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Renamed {currentPath} to {newPath}");

                // The FileSystemWatcher will automatically handle UI updates
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to rename {currentPath}: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Failed to rename item: {ex.Message}", "Rename Error");
            }
        }

        private void InitializeFrameContents()
        {
            _dispatcher.Invoke(() =>
            {
                _wpcont.Children.Clear();
                if (_detailsListView != null) _detailsListView.Items.Clear(); // --- STEP 5: Clear Details table ---
            });

            // --- NAVIGATION LAG FIX ---
            // Pass 'true' to completely bypass the FileWatcher's 500ms debounce timer.
            // This guarantees the folder begins loading instantly upon click.
            TriggerSync(immediate: true);

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Requested immediate async initialization for {_targetFolderPath}");
        }

        private void CopyPathOrTarget(string path)
        {
            try
            {
                string pathToCopy;
                if (Path.GetExtension(path).ToLower() == ".lnk")
                {
                    // If it's a shortcut, get the target path
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(path);
                    pathToCopy = shortcut.TargetPath;
                }
                else
                {
                    // Otherwise, copy the folder path (portal frame path)
                    pathToCopy = Path.GetDirectoryName(path); // Gets the parent directory
                }

                // Copy to clipboard
                Clipboard.SetText(pathToCopy);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Copied path to clipboard: {pathToCopy}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Failed to copy path for {path}: {ex.Message}");
                MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Unable to copy path.", "Error");
            }
        }

        private void DeleteItem(string path, StackPanel sp)
        {
            bool UseRecycleBin = SettingsManager.UseRecycleBin;
            if (UseRecycleBin == true)
            {
                try
                {
                    // First, check if the item exists
                    if (!Directory.Exists(path) && !System.IO.File.Exists(path))
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Item not found for deletion: {path}");
                        return;
                    }

                    // Use SHFileOperation to move to recycle bin
                    SHFILEOPSTRUCT shf = new SHFILEOPSTRUCT();
                    shf.wFunc = FO_DELETE;
                    shf.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION;
                    shf.pFrom = path + '\0' + '\0'; // Double null-terminated string

                    int result = SHFileOperation(ref shf);

                    if (result != 0)
                    {
                        throw new Exception($"Failed to move to recycle bin (error code: {result})");
                    }

                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Moved to recycle bin: {path}");

                    // Remove the icon from the UI
                    _wpcont.Children.Remove(sp);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, ($"Removed icon for {path} from UI"));
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Failed to move item {path} to recycle bin: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Unable to move item to recycle bin.", "Error");
                }
            }
            else
            {
                try
                {
                    if (Directory.Exists(path))
                    {
                        // Delete folder
                        Directory.Delete(path, true); // true for recursive deletion
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Deleted folder: {path}");
                    }
                    else if (System.IO.File.Exists(path))
                    {
                        // Delete file
                        System.IO.File.Delete(path);
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Deleted file: {path}");
                    }
                    else
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Item not found for deletion: {path}");
                        return;
                    }

                    // Remove the icon from the UI
                    _wpcont.Children.Remove(sp);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Removed icon for {path} from UI");
                }
                catch (Exception ex)
                {
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Failed to delete item {path}: {ex.Message}");
                    MessageBoxesManager.ShowOKOnlyMessageBoxForm($"Unable to delete item.", "Error");
                }
            }
        }

        // Corrected Win32 API declarations
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            public uint wFunc;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pFrom;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string pTo;
            public ushort fFlags;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHFileOperation([In] ref SHFILEOPSTRUCT lpFileOp);

        const uint FO_DELETE = 0x0003;
        const ushort FOF_ALLOWUNDO = 0x0040;
        const ushort FOF_NOCONFIRMATION = 0x0010;

        private void RemoveIcon(string path)
        {
            // --- STEP 4 FIX: Remove row from Details View if active ---
            if (_isDetailsView && _detailsListView != null)
            {
                var item = _detailsListView.Items.OfType<PortalItemModel>()
                    .FirstOrDefault(i => string.Equals(i.FullPath, path, StringComparison.OrdinalIgnoreCase));
                if (item != null)
                {
                    _detailsListView.Items.Remove(item);
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Removed Details View row for {path}");
                }
                return;
            }
            // ----------------------------------------------------------

            var sp = _wpcont.Children.OfType<StackPanel>().FirstOrDefault(s =>
            {
                string p = s.Tag?.GetType().GetProperty("FilePath")?.GetValue(s.Tag)?.ToString();
                return string.Equals(p, path, StringComparison.OrdinalIgnoreCase);
            });

            if (sp != null)
            {
                _wpcont.Children.Remove(sp);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Successfully removed icon for {path}");
            }
            else
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General, $"Failed to find StackPanel for {path} in RemoveIcon");
            }
        }
        // TEST: Filter for only text files (REMOVE AFTER TEST)
        // ApplyFilter("*.txt");



 
        // --- DETAILS VIEW ENGINE START (Step 2 & Step 3) ---

        public class PortalItemModel : System.ComponentModel.INotifyPropertyChanged
        {
            public event System.ComponentModel.PropertyChangedEventHandler PropertyChanged;
            protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));

            public int Index { get; set; }
            public string Name { get; set; }
            public string FullPath { get; set; }
            public bool IsFolder { get; set; }
            public System.Windows.Media.ImageSource Icon { get; set; }
            public ContextMenu RowMenu { get; set; } // --- NEW: Binds item context menu ---

            private string _dateModified;
            public string DateModified { get => _dateModified; set { _dateModified = value; OnPropertyChanged(nameof(DateModified)); } }

            private string _type;
            public string Type { get => _type; set { _type = value; OnPropertyChanged(nameof(Type)); } }

            private string _size;
            public string Size { get => _size; set { _size = value; OnPropertyChanged(nameof(Size)); } }

            public long RawSizeBytes { get; set; }
            public DateTime RawDateModified { get; set; }
        }

        private void BuildOrSwitchToDetailsView()
        {
            if (_parentScrollViewer == null) return;

            if (_detailsListView == null)
            {
                _detailsListView = new ListView
                {
                    Background = System.Windows.Media.Brushes.Transparent,
                    BorderThickness = new Thickness(0),
                    Foreground = System.Windows.Media.Brushes.White,
                    FontFamily = new System.Windows.Media.FontFamily(SettingsManager.GlobalFontFamily),
                    FontSize = SettingsManager.DefaultItemFontSize,
                    SelectionMode = SelectionMode.Single,
                    Margin = new Thickness(5, 0, 24, 8) // --- VISUAL BALANCE: Prevents right/bottom edge collision ---
                };

                // --- THEME FIX: Strip Windows OS chrome and apply semi-transparent adaptive styling ---
                string themeXaml = @"
                <ResourceDictionary xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation' xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'>
             <Style x:Key='HeaderGripper' TargetType='Thumb'>
                        <Setter Property='Width' Value='8' />
                        <Setter Property='Background' Value='Transparent' />
                        <Setter Property='Template'>
                            <Setter.Value>
                                <ControlTemplate TargetType='Thumb'>
                                    <Border Background='Transparent' />
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                    <Style TargetType='GridViewColumnHeader'>
                        <Setter Property='OverridesDefaultStyle' Value='True' />
                        <Setter Property='Foreground' Value='White' />
                        <Setter Property='HorizontalContentAlignment' Value='Left' />
                        <Setter Property='Template'>
                            <Setter.Value>
                                <ControlTemplate TargetType='GridViewColumnHeader'>
                                    <Grid>
                                        <Border x:Name='bd' Background='#33000000' BorderBrush='#33FFFFFF' BorderThickness='0,0,1,1' Padding='6,4'>
                                            <ContentPresenter HorizontalAlignment='{TemplateBinding HorizontalContentAlignment}' VerticalAlignment='Center' />
                                        </Border>
                                        <Thumb x:Name='PART_HeaderGripper' HorizontalAlignment='Right' Style='{StaticResource HeaderGripper}' />
                                    </Grid>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property='IsMouseOver' Value='True'>
                                            <Setter TargetName='bd' Property='Background' Value='#25FFFFFF' />
                                        </Trigger>
                                        <Trigger Property='IsPressed' Value='True'>
                                            <Setter TargetName='bd' Property='Background' Value='#55000000' />
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                    
  <Style x:Key='GlassItemStyle' TargetType='ListViewItem'>
                        <!-- Nuclear option: Stops WPF from falling back to Windows Aero/Metro system styles -->
                        <Setter Property='OverridesDefaultStyle' Value='True' />
                        <Setter Property='Foreground' Value='White' />
                        <Setter Property='HorizontalContentAlignment' Value='Stretch' />
                        <Setter Property='Margin' Value='0,1' />
                        <Setter Property='ContextMenu' Value='{Binding RowMenu}' />
                        <Setter Property='Template'>
                            <Setter.Value>
                                <ControlTemplate TargetType='ListViewItem'>
                                    <Border x:Name='bd' Background='Transparent' BorderThickness='1' BorderBrush='Transparent' CornerRadius='3' Padding='4,2'>
                                        <!-- Clean tag prevents silent XAML parse exceptions -->
                                        <GridViewRowPresenter />
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property='IsMouseOver' Value='True'>
                                            <Setter TargetName='bd' Property='Background' Value='#15FFFFFF' />
                                            <Setter TargetName='bd' Property='BorderBrush' Value='#22FFFFFF' />
                                        </Trigger>
                                        <Trigger Property='IsSelected' Value='True'>
                                            <!-- Deep semi-transparent black glass makes the row pop while keeping text white -->
                                            <Setter TargetName='bd' Property='Background' Value='#55000000' />
                                            <Setter TargetName='bd' Property='BorderBrush' Value='#88FFFFFF' />
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </ResourceDictionary>";
                try
                {
                    _detailsListView.Resources = (ResourceDictionary)System.Windows.Markup.XamlReader.Parse(themeXaml);
                    _detailsListView.ItemContainerStyle = (Style)_detailsListView.Resources["GlassItemStyle"];

                    // --- STYLE ENFORCER HOOK: Intercepts row generation to defeat downstream style overrides ---
                    _detailsListView.ItemContainerGenerator.StatusChanged += (s, e) =>
                    {
                        if (_detailsListView.ItemContainerGenerator.Status == System.Windows.Controls.Primitives.GeneratorStatus.ContainersGenerated)
                        {
                            var targetStyle = _detailsListView.Resources["GlassItemStyle"] as Style;
                            if (targetStyle != null)
                            {
                                foreach (var item in _detailsListView.Items)
                                {
                                    var container = _detailsListView.ItemContainerGenerator.ContainerFromItem(item) as System.Windows.Controls.ListViewItem;
                                    if (container != null && container.Style != targetStyle)
                                    {
                                        container.Style = targetStyle;
                                    }
                                }
                            }
                        }
                    };
                }
                catch { }

                GridView gridView = new GridView();
                gridView.Columns.Add(new GridViewColumn { Header = "#", DisplayMemberBinding = new System.Windows.Data.Binding("Index"), Width = 30 });

                var nameFactory = new FrameworkElementFactory(typeof(StackPanel));
                nameFactory.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);
                var imgFactory = new FrameworkElementFactory(typeof(Image));
                imgFactory.SetBinding(Image.SourceProperty, new System.Windows.Data.Binding("Icon"));
                imgFactory.SetValue(Image.WidthProperty, 16.0);
                imgFactory.SetValue(Image.HeightProperty, 16.0);
                imgFactory.SetValue(Image.MarginProperty, new Thickness(0, 0, 5, 0));
                var txtFactory = new FrameworkElementFactory(typeof(TextBlock));
                txtFactory.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("Name"));
                txtFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
                nameFactory.AppendChild(imgFactory);
                nameFactory.AppendChild(txtFactory);

                gridView.Columns.Add(new GridViewColumn { Header = "Name", CellTemplate = new DataTemplate { VisualTree = nameFactory }, Width = 160 });
                gridView.Columns.Add(new GridViewColumn { Header = "Date modified", DisplayMemberBinding = new System.Windows.Data.Binding("DateModified"), Width = 110 });
                gridView.Columns.Add(new GridViewColumn { Header = "Type", DisplayMemberBinding = new System.Windows.Data.Binding("Type"), Width = 90 });
                gridView.Columns.Add(new GridViewColumn { Header = "Size", DisplayMemberBinding = new System.Windows.Data.Binding("Size"), Width = 70 });

                _detailsListView.View = gridView;

                // --- STEP 8: Load Saved Columns & Watch for Width/Order Changes ---
                LoadColumnState(gridView);

                gridView.Columns.CollectionChanged += (s, e) => TriggerColumnSave();
                foreach (var col in gridView.Columns)
                {
                    if (col is System.ComponentModel.INotifyPropertyChanged npc)
                    {
                        npc.PropertyChanged += (s, e) =>
                        {
                            if (e.PropertyName == nameof(GridViewColumn.Width)) TriggerColumnSave();
                        };
                    }
                }
                // ------------------------------------------------------------------

                // --- LIVE BINDING & NATIVE NAVIGATION ENGINE ---
                Action<PortalItemModel, bool> HandleItemInteraction = (selected, isCtrlPressed) =>
                {
                    if (selected == null) return;

                    if (isCtrlPressed)
                    {
                        if (selected.IsFolder)
                        {
                            // 1. Resolve Live Window & Frame to guarantee Nav Bar updates
                            var win = Window.GetWindow(_parentScrollViewer) as NonActivatingWindow ?? Window.GetWindow(_wpcont) as NonActivatingWindow;
                            string fId = win?.Tag?.ToString() ?? _frame.Id?.ToString();
                            var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == fId) ?? _frame;

                            // 2. Trigger native navigation and explicit UI refresh
                            Framemanager.NavigatePortalFrame(liveFrame, selected.FullPath);
                            if (win != null)
                            {
                                Framemanager.RefreshPortalNavBar(win, liveFrame);
                                var dockPanel = (win.Content as Border)?.Child as DockPanel;
                                dockPanel?.UpdateLayout(); // Force DockPanel to recalculate top bar boundaries
                            }
                        }
                        return;
                    }

                    Framemanager.LaunchItem(new StackPanel(), selected.FullPath, selected.IsFolder);
                };

                _detailsListView.PreviewMouseLeftButtonUp += (s, e) =>
                {
                    var row = GetClickedListViewItem(e.OriginalSource as DependencyObject);
                    if (row?.Content is PortalItemModel selected)
                    {
                        bool isCtrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                        if (isCtrl)
                        {
                            HandleItemInteraction(selected, true);
                            e.Handled = true;
                            return;
                        }

                        if (SettingsManager.SingleClickToLaunch)
                        {
                            HandleItemInteraction(selected, false);
                        }
                    }
                };

                _detailsListView.MouseDoubleClick += (s, e) =>
                {
                    if (SettingsManager.SingleClickToLaunch) return;
                    var row = GetClickedListViewItem(e.OriginalSource as DependencyObject);
                    if (row?.Content is PortalItemModel selected)
                    {
                        bool isCtrl = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control;
                        if (!isCtrl) HandleItemInteraction(selected, false);
                    }
                };

                _detailsListView.Loaded += AttachGridViewHeaderMenu;

                // --- STEP 7: Attach Header Click Sorting & Initial Arrow Indicator ---
                _detailsListView.AddHandler(GridViewColumnHeader.ClickEvent, new RoutedEventHandler((s, e) =>
                {
                    if (e.OriginalSource is GridViewColumnHeader headerClicked && headerClicked.Role != GridViewColumnHeaderRole.Padding)
                    {
                        SortDetailsViewColumn(headerClicked);
                    }
                }));

                _detailsListView.Loaded += (s, e) =>
                {
                    if (_detailsListView.View is GridView gv && gv.Columns.Count > 1)
                    {
                        int colIndex = _sortMode == 1 ? 2 : (_sortMode == 2 ? 3 : (_sortMode == 3 ? 4 : 1));
                        if (colIndex < gv.Columns.Count && gv.Columns[colIndex].Header != null)
                        {
                            string baseName = gv.Columns[colIndex].Header.ToString().Replace(" ▲", "").Replace(" ▼", "");
                            gv.Columns[colIndex].Header = baseName + (_sortAscending ? " ▲" : " ▼");
                            _currentSortColumn = gv.Columns[colIndex];
                        }
                    }
                };
                // ----------------------------------------------------------------------
            } // <-- Close the (if _detailsListView == null) check here!

            // --- ALWAYS EXECUTE ON SWITCH: Ensure ScrollViewer swaps to Details View ---
            if (_parentScrollViewer != null)
            {
                if (_parentScrollViewer.Content != _detailsListView)
                {
                    _parentScrollViewer.Content = _detailsListView;
                }
                // The parent scroll viewer must be completely DISABLED on BOTH axes so it doesn't 
                // grant infinite layout space or intercept the wheel, allowing the ListView's internal scroll viewer to work.
                _parentScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
                _parentScrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
                _parentScrollViewer.Margin = new Thickness(0);
            }

            // --- SCROLLBAR FIX: Apply setting to the ListView's internal ScrollViewer ---
            if (_detailsListView != null)
            {
                var visibility = SettingsManager.DisableFrameScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto;
                ScrollViewer.SetVerticalScrollBarVisibility(_detailsListView, visibility);
                ScrollViewer.SetHorizontalScrollBarVisibility(_detailsListView, visibility);
            }

            _wpcont.Visibility = Visibility.Collapsed;
            _detailsListView.Visibility = Visibility.Visible;
            _isDetailsView = true;
        }

        /// <summary>
        /// Public hook to dynamically refresh scrollbar visibility on-the-fly when settings change.
        /// </summary>
        public void UpdateScrollbarVisibility()
        {
            _dispatcher.Invoke(() =>
            {
                var visibility = SettingsManager.DisableFrameScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto;

                if (_isDetailsView && _detailsListView != null)
                {
                    // Ensure outer container remains completely disabled to bind ListView horizontal limits
                    if (_parentScrollViewer != null)
                    {
                        _parentScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
                        _parentScrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
                    }

                    // 1. Update WPF attached properties ONLY (Preserves internal MouseWheel TemplateBindings!)
                    ScrollViewer.SetVerticalScrollBarVisibility(_detailsListView, visibility);
                    ScrollViewer.SetHorizontalScrollBarVisibility(_detailsListView, visibility);

                    // 2. Safely force the live visual update without destroying bindings
                    _detailsListView.InvalidateMeasure();
                    var internalSv = FindVisualChild<ScrollViewer>(_detailsListView);
                    if (internalSv != null)
                    {
                        internalSv.InvalidateMeasure();
                        internalSv.UpdateLayout();
                    }
                    _detailsListView.UpdateLayout();
                }
                else if (!_isDetailsView && _parentScrollViewer != null)
                {
                    _parentScrollViewer.VerticalScrollBarVisibility = visibility;

                    // IMPORTANT: Icons View (WrapPanel) MUST have horizontal scrolling Disabled, 
                    // otherwise the icons will stretch infinitely instead of wrapping to the next line!
                    _parentScrollViewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
                }
            });
        }
        private ListViewItem GetClickedListViewItem(DependencyObject source)
        {
            while (source != null && !(source is ListViewItem))
            {
                // Abort if they clicked a column header or scrollbar
                if (source is GridViewColumnHeader || source is System.Windows.Controls.Primitives.ScrollBar) return null;
                source = VisualTreeHelper.GetParent(source);
            }
            return source as ListViewItem;
        }

        private void AttachGridViewHeaderMenu(object sender, RoutedEventArgs e)
        {
            GridView gridView = _detailsListView?.View as GridView;
            if (gridView == null) return;

            ContextMenu headerMenu = new ContextMenu();

            Style themedStyle = GetThemedContextMenuStyle();
            if (themedStyle != null) headerMenu.Style = themedStyle;

            foreach (var col in gridView.Columns)
            {
                string cleanName = col.Header?.ToString().Replace(" ▲", "").Replace(" ▼", "") ?? "";
                if (string.IsNullOrEmpty(cleanName)) continue;

                var targetCol = col;
                // Double.IsNaN check guarantees auto-sized columns show as checked!
                MenuItem colItem = new MenuItem { Header = cleanName, IsCheckable = true, IsChecked = double.IsNaN(targetCol.Width) || targetCol.Width > 0 };
                colItem.Click += (ms, me) =>
                {
                    targetCol.Width = colItem.IsChecked ? Double.NaN : 0;
                    TriggerColumnSave(); // Save visibility change
                };
                headerMenu.Items.Add(colItem);
            }

            headerMenu.Items.Add(new Separator());

            MenuItem resetSortItem = new MenuItem { Header = "Reset sorting" };
            resetSortItem.Click += (ms, me) => { _sortMode = 0; SortContents(); };
            headerMenu.Items.Add(resetSortItem);

            MenuItem resetColsItem = new MenuItem { Header = "Reset column layout" };
            resetColsItem.Click += (ms, me) =>
            {
                string defaultLayout = "#:30;Name:160;Date modified:110;Type:90;Size:70";
                Framemanager.UpdateFrameProperty(_frame, "DetailsViewColumns", defaultLayout, "Reset column layout");
                LoadColumnState(gridView, defaultLayout);
            };
            headerMenu.Items.Add(resetColsItem);

            // --- LIVE THEME FIX ---
            headerMenu.Opened += (ms, me) =>
            {
                try
                {
                    Style liveStyle = Framemanager.GetThemedContextMenuStyle(_frame);
                    if (liveStyle != null) headerMenu.Style = liveStyle;
                }
                catch { }
            };

            var headerRow = FindVisualChild<GridViewHeaderRowPresenter>(_detailsListView);
            if (headerRow != null) headerRow.ContextMenu = headerMenu;
        }

        private static T FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild) return typedChild;
                var found = FindVisualChild<T>(child);
                if (found != null) return found;
            }
            return null;
        }

        // --- DETAILS VIEW ENGINE END ---

        /// <summary>
        /// Live switches between Icons and Details view without destroying the window or fence.
        /// </summary>
        public void SetViewMode(string viewMode)
        {
            _isDetailsView = string.Equals(viewMode, "Details", StringComparison.OrdinalIgnoreCase);

            _dispatcher.Invoke(() =>
            {
                if (_parentScrollViewer != null)
                {
                    if (_isDetailsView)
                    {
                        BuildOrSwitchToDetailsView();
                    }
                    else
                    {
                        // Restore Icon Grid View
                        _parentScrollViewer.Content = _wpcont;
                        _parentScrollViewer.VerticalScrollBarVisibility = SettingsManager.DisableFrameScrollbars ? ScrollBarVisibility.Hidden : ScrollBarVisibility.Auto;
                        _parentScrollViewer.Margin = new Thickness(0);
                        if (_detailsListView != null) _detailsListView.Visibility = Visibility.Collapsed;
                        _wpcont.Visibility = Visibility.Visible;
                    }
                }

                // Clear containers to force a clean re-render into the newly selected view
                _wpcont.Children.Clear();
                if (_detailsListView != null) _detailsListView.Items.Clear();
            });

            // Immediately resync from disk to populate the newly selected view
            TriggerSync(immediate: true);

            // Refresh top navigation bar to ensure Z-order and layout bounds are correct
            _dispatcher.InvokeAsync(() =>
            {
                var win = Window.GetWindow(_parentScrollViewer) as NonActivatingWindow ?? Window.GetWindow(_wpcont) as NonActivatingWindow;
                if (win != null)
                {
                    string fId = win.Tag?.ToString() ?? _frame.Id?.ToString();
                    var liveFrame = FrameDataManager.FrameData.FirstOrDefault(f => f.Id?.ToString() == fId) ?? _frame;
                    Framemanager.RefreshPortalNavBar(win, liveFrame);
                    var dockPanel = (win.Content as Border)?.Child as DockPanel;
                    dockPanel?.UpdateLayout();
                }
            }, System.Windows.Threading.DispatcherPriority.Loaded);

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"Portal view mode live switched to: {viewMode}");
        }

        /// <summary>
        /// Safely switches the monitored folder without destroying the frame window.
        /// </summary>
        public void NavigateTo(string newPath)
        {
            if (string.IsNullOrEmpty(newPath) || !System.IO.Directory.Exists(newPath))
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Cannot navigate to invalid path: {newPath}");
                return;
            }

            try
            {
                _targetFolderPath = newPath;

                if (_watcher != null)
                {
                    _watcher.EnableRaisingEvents = false;
                    _watcher.Path = newPath;
                    _watcher.EnableRaisingEvents = true;
                }

                _dispatcher.Invoke(() =>
                {
                    _wpcont.Children.Clear();
                    if (_detailsListView != null) _detailsListView.Items.Clear();
                });

                InitializeFrameContents();
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Navigated portal frame to: {newPath}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Error navigating to {newPath}: {ex.Message}");
            }
        }

        public void Dispose()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
            }
            if (_debounceTimer != null)
            {
                _debounceTimer.Stop();
                _debounceTimer.Tick -= ProcessPendingEvents;
            }
        }
    }
}
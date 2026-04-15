using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Desktop_Fences
{
    public class AutoOrganizeForm : Window
    {
        private ListBox _rulesList;
        private ScrollViewer _editorScroll;
        private StackPanel _editorPanel;

        private CheckBox _chkIsEnabled;
        private TextBox _txtName;
        private ComboBox _cmbCondition;
        private Border _customExtBorder;
        private TextBox _txtCustomExt;
        private TextBox _txtNameContains;
        private ComboBox _cmbTarget;
        private CheckBox _chkAutoCreate;
        private ComboBox _cmbConflict;

        private TextBlock _txtLastRun;

        private OrganizeRule _selectedRule;
        private List<OrganizeRule> _localRules;
        private bool _isLoadingRule = false; // Prevents dropdowns from overwriting data during load

        private readonly Dictionary<string, string> _presets = new Dictionary<string, string>
        {
            { "Images", "*.jpg; *.jpeg; *.png; *.gif; *.bmp; *.webp" },
            { "Documents", "*.doc*; *.pdf; *.txt; *.rtf; *.xls*; *.ppt*" },
            { "Executables", "*.exe; *.bat; *.msi; *.cmd; *.ps1" },
            { "Archives", "*.zip; *.rar; *.7z; *.tar; *.gz" },
            { "Custom Extensions...", "" }
        };

        public AutoOrganizeForm()
        {
            AutoOrganizeManager.Pause(); // Pause while editing rules!
                                         // Clone rules for safe editing
            _localRules = AutoOrganizeManager.Rules.Select(r => new OrganizeRule
            {
                Id = r.Id,
                Name = r.Name,
                Extensions = r.Extensions,
                NameContains = r.NameContains,
                TargetFolderPath = r.TargetFolderPath,
                ConflictAction = r.ConflictAction,
                AutoCreateFence = r.AutoCreateFence,
                IsEnabled = r.IsEnabled,
                Priority = r.Priority,
                LastRun = r.LastRun // <--- FIX: Actually copy the timestamp!
            }).OrderBy(r => r.Priority).ToList();

            InitializeModernComponent();
            RefreshList();

            // Select first rule if available
            if (_localRules.Count > 0)
                _rulesList.SelectedIndex = 0;
        }

        private void InitializeModernComponent()
        {
            Title = "Smart Desktop Rules";
            Width = 850; // Increased Width
            Height = 700; // Increased Height
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
            WindowStyle = WindowStyle.None;
            AllowsTransparency = false;

            Grid mainContainer = new Grid
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                Margin = new Thickness(8)
            };

            Border mainBorder = new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(0)
            };

            Grid rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header (50px)
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Footer

            CreateModernHeader(rootGrid);
            CreateModernContent(rootGrid);
            CreateModernFooter(rootGrid);

            mainBorder.Child = rootGrid;
            mainContainer.Children.Add(mainBorder);
            Content = mainContainer;

            this.KeyDown += AutoOrganizeForm_KeyDown;
            this.Focusable = true;
            this.Focus();
        }

        private void AutoOrganizeForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !_txtCustomExt.IsFocused && !_txtName.IsFocused && !_txtNameContains.IsFocused)
            {
                Save_Click(this, new RoutedEventArgs());
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                AutoOrganizeManager.Resume();
                DialogResult = false;
                Close();
                e.Handled = true;
            }
        }

        private SolidColorBrush GetAccentColorBrush()
        {
            try
            {
                return new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor));
            }
            catch
            {
                return new SolidColorBrush(Color.FromRgb(66, 133, 244)); // Fallback
            }
        }

        private void CreateModernHeader(Grid rootGrid)
        {
            Border headerBorder = new Border
            {
                Height = 50,
                Background = GetAccentColorBrush(),
                CornerRadius = new CornerRadius(0)
            };

            headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) DragMove(); };

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            TextBlock titleText = new TextBlock
            {
                Text = "Smart Desktop Rules",
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };

            Button closeButton = new Button
            {
                Content = "✕",
                Width = 32,
                Height = 32,
                FontSize = 12,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                Background = Brushes.Transparent,
                BorderBrush = Brushes.Transparent,
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 9, 0),
                VerticalAlignment = VerticalAlignment.Center
            };

            closeButton.MouseEnter += (s, e) => closeButton.Background = new SolidColorBrush(Color.FromArgb(50, 255, 255, 255));
            closeButton.MouseLeave += (s, e) => closeButton.Background = Brushes.Transparent;
            closeButton.Click += (s, e) => { AutoOrganizeManager.Resume(); DialogResult = false; Close(); };

            headerGrid.Children.Add(titleText);
            headerGrid.Children.Add(closeButton);
            Grid.SetColumn(closeButton, 1);

            headerBorder.Child = headerGrid;
            Grid.SetRow(headerBorder, 0);
            rootGrid.Children.Add(headerBorder);
        }

        private void CreateModernContent(Grid rootGrid)
        {
            Border contentBorder = new Border
            {
                Background = Brushes.White,
                Padding = new Thickness(20, 15, 20, 10)
            };

            Grid contentGrid = new Grid();
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(280) }); // Increased left panel slightly
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(20) }); // Gutter
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            // --- LEFT PANEL: LIST & BUTTONS ---
            Grid leftPanel = new Grid();
            leftPanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            leftPanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            Border listBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(4)
            };

            DataTemplate template = new DataTemplate(typeof(OrganizeRule));
            FrameworkElementFactory textBlock = new FrameworkElementFactory(typeof(TextBlock));
            textBlock.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding("Name"));
            textBlock.SetValue(TextBlock.ForegroundProperty, new SolidColorBrush(Color.FromRgb(150, 150, 150))); // Disabled gray color

            DataTrigger trigger = new DataTrigger { Binding = new System.Windows.Data.Binding("IsEnabled"), Value = true };
            trigger.Setters.Add(new Setter(TextBlock.FontWeightProperty, FontWeights.Bold));
            trigger.Setters.Add(new Setter(TextBlock.ForegroundProperty, new SolidColorBrush(Color.FromRgb(34, 139, 34)))); // Enabled dark green color

            template.VisualTree = textBlock;
            template.Triggers.Add(trigger);

            _rulesList = new ListBox
            {
                ItemTemplate = template,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontSize = 13
            };
            _rulesList.SelectionChanged += (s, e) => LoadRuleEditor((OrganizeRule)_rulesList.SelectedItem);
            listBorder.Child = _rulesList;
            Grid.SetRow(listBorder, 0);
            leftPanel.Children.Add(listBorder);

            StackPanel listButtons = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 12, 0, 0) };

            Grid topButtons = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            topButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            topButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            topButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnUp = CreateModernSecondaryButton("▲ Up"); btnUp.Click += (s, e) => MoveRule(-1);
            Button btnDown = CreateModernSecondaryButton("▼ Down"); btnDown.Click += (s, e) => MoveRule(1);
            Grid.SetColumn(btnUp, 0); Grid.SetColumn(btnDown, 2);
            topButtons.Children.Add(btnUp); topButtons.Children.Add(btnDown);

            Grid botButtons = new Grid();
            botButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            botButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(8) });
            botButtons.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Button btnAdd = new Button { Content = "+ Add Rule", Height = 32, FontSize = 12, FontWeight = FontWeights.Medium, Cursor = Cursors.Hand, Foreground = Brushes.White };
            btnAdd.Background = new SolidColorBrush(Color.FromRgb(34, 139, 34)); // Dark Green
            btnAdd.BorderThickness = new Thickness(0);
            btnAdd.Click += (s, e) => AddNewRule();

            Button btnRemove = new Button { Content = "− Remove", Height = 32, FontSize = 12, FontWeight = FontWeights.Medium, Cursor = Cursors.Hand, Foreground = Brushes.White };
            btnRemove.Background = new SolidColorBrush(Color.FromRgb(255, 140, 0)); // Dark Orange
            btnRemove.BorderThickness = new Thickness(0);
            btnRemove.Click += (s, e) => RemoveRule();
            Grid.SetColumn(btnAdd, 0); Grid.SetColumn(btnRemove, 2);
            botButtons.Children.Add(btnAdd); botButtons.Children.Add(btnRemove);

            listButtons.Children.Add(topButtons);
            listButtons.Children.Add(botButtons);
            Grid.SetRow(listButtons, 1);
            leftPanel.Children.Add(listButtons);

            Grid.SetColumn(leftPanel, 0);
            contentGrid.Children.Add(leftPanel);

            // --- RIGHT PANEL: EDITOR WITH SCROLLBAR ---
            _editorScroll = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(0, 0, 10, 0)
            };

            _editorPanel = new StackPanel { IsEnabled = false };

            // --- GROUP 1: Rule General Info ---
            Border group1 = new Border { Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Margin = new Thickness(0, 0, 0, 12), Padding = new Thickness(15) };
            StackPanel sp1 = new StackPanel();

            _chkIsEnabled = new CheckBox { Content = "Enable this rule", FontSize = 14, FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(34, 139, 34)), Margin = new Thickness(0, 0, 0, 5) };
            _chkIsEnabled.Checked += (s, e) => { if (!_isLoadingRule && _selectedRule != null) { _selectedRule.IsEnabled = true; _rulesList.Items.Refresh(); } };
            _chkIsEnabled.Unchecked += (s, e) => { if (!_isLoadingRule && _selectedRule != null) { _selectedRule.IsEnabled = false; _rulesList.Items.Refresh(); } };
            sp1.Children.Add(_chkIsEnabled);

            _txtLastRun = new TextBlock { FontSize = 12, Foreground = Brushes.Gray, FontStyle = FontStyles.Italic, Margin = new Thickness(0, 0, 0, 15) };
            sp1.Children.Add(_txtLastRun);

            sp1.Children.Add(new TextBlock { Text = "Rule Name:", FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 0, 0, 8) });
            _txtName = new TextBox { FontSize = 13, Padding = new Thickness(8, 6, 8, 6), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Background = Brushes.White, Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)) };
            _txtName.TextChanged += (s, e) => { if (!_isLoadingRule && _selectedRule != null) { _selectedRule.Name = _txtName.Text; _rulesList.Items.Refresh(); } };
            sp1.Children.Add(_txtName);

            group1.Child = sp1;
            _editorPanel.Children.Add(group1);

            // --- GROUP 2: Match Conditions ---
            Border group2 = new Border { Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Margin = new Thickness(0, 0, 0, 12), Padding = new Thickness(15) };
            StackPanel sp2 = new StackPanel();

            sp2.Children.Add(new TextBlock { Text = "If File Extension Is:", FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 0, 0, 8) });
            _cmbCondition = new ComboBox { FontSize = 13, Padding = new Thickness(8, 6, 8, 6), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Background = Brushes.White, Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)) };
            foreach (var p in _presets.Keys) _cmbCondition.Items.Add(p);
            _cmbCondition.SelectionChanged += CmbCondition_SelectionChanged;
            sp2.Children.Add(_cmbCondition);

            _customExtBorder = new Border { Visibility = Visibility.Collapsed };
            StackPanel customExtStack = new StackPanel();
            customExtStack.Children.Add(new TextBlock { Text = "Custom Extensions (e.g., *.zip; *.rar):", FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 12, 0, 8) });
            _txtCustomExt = new TextBox { FontSize = 13, Padding = new Thickness(8, 6, 8, 6), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Background = Brushes.White, Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)) };
            _txtCustomExt.TextChanged += (s, e) => { if (!_isLoadingRule && _selectedRule != null && _cmbCondition.SelectedItem?.ToString() == "Custom Extensions...") _selectedRule.Extensions = _txtCustomExt.Text; };
            customExtStack.Children.Add(_txtCustomExt);
            _customExtBorder.Child = customExtStack;
            sp2.Children.Add(_customExtBorder);

            sp2.Children.Add(new TextBlock { Text = "And Filename Contains (Optional, leave blank for any):", FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 12, 0, 8) });
            _txtNameContains = new TextBox { FontSize = 13, Padding = new Thickness(8, 6, 8, 6), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Background = Brushes.White, Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)) };
            _txtNameContains.TextChanged += (s, e) => { if (!_isLoadingRule && _selectedRule != null) _selectedRule.NameContains = _txtNameContains.Text; };
            sp2.Children.Add(_txtNameContains);

            group2.Child = sp2;
            _editorPanel.Children.Add(group2);

            // --- GROUP 3: Actions ---
            Border group3 = new Border { Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), CornerRadius = new CornerRadius(6), Margin = new Thickness(0, 0, 0, 12), Padding = new Thickness(15) };
            StackPanel sp3 = new StackPanel();

            sp3.Children.Add(new TextBlock { Text = "Move To:", FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 0, 0, 8) });
            _cmbTarget = new ComboBox { FontSize = 13, Padding = new Thickness(8, 6, 8, 6), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Background = Brushes.White, Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)) };
            LoadPortalFencesIntoDropdown();
            _cmbTarget.SelectionChanged += CmbTarget_SelectionChanged;
            sp3.Children.Add(_cmbTarget);

            _chkAutoCreate = new CheckBox { Content = "Generate new Portal Fence for this folder", FontSize = 12, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(4, 6, 0, 6), Visibility = Visibility.Collapsed };
            _chkAutoCreate.Checked += (s, e) => { if (!_isLoadingRule && _selectedRule != null) _selectedRule.AutoCreateFence = true; };
            _chkAutoCreate.Unchecked += (s, e) => { if (!_isLoadingRule && _selectedRule != null) _selectedRule.AutoCreateFence = false; };
            sp3.Children.Add(_chkAutoCreate);

            sp3.Children.Add(new TextBlock { Text = "If File Already Exists:", FontSize = 12, FontWeight = FontWeights.Medium, Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)), Margin = new Thickness(0, 12, 0, 8) });
            _cmbConflict = new ComboBox { FontSize = 13, Padding = new Thickness(8, 6, 8, 6), BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)), BorderThickness = new Thickness(1), Background = Brushes.White, Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)) };
            _cmbConflict.Items.Add("Auto-Rename (e.g., File (1).jpg)");
            _cmbConflict.Items.Add("Overwrite (Send old file to Recycle Bin)");
            _cmbConflict.Items.Add("Skip (Leave on Desktop)");
            _cmbConflict.SelectionChanged += (s, e) => { if (!_isLoadingRule && _selectedRule != null) _selectedRule.ConflictAction = (RuleConflictAction)_cmbConflict.SelectedIndex; };
            sp3.Children.Add(_cmbConflict);

            group3.Child = sp3;
            _editorPanel.Children.Add(group3);

            _editorScroll.Content = _editorPanel;
            Grid.SetColumn(_editorScroll, 2);
            contentGrid.Children.Add(_editorScroll);

            contentBorder.Child = contentGrid;
            Grid.SetRow(contentBorder, 1);
            rootGrid.Children.Add(contentBorder);
        }

        private void CreateModernFooter(Grid rootGrid)
        {
            Border footerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Padding = new Thickness(20, 16, 20, 16),
                CornerRadius = new CornerRadius(0)
            };

            StackPanel buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            Button closeButton = CreateModernSecondaryButton("Close");
            closeButton.MinWidth = 80;
            closeButton.Margin = new Thickness(0, 0, 10, 0);
            closeButton.Click += (s, e) => { AutoOrganizeManager.Resume(); DialogResult = false; Close(); };

            Button saveButton = new Button
            {
                Content = "Save Rules",
                Height = 36,
                MinWidth = 100,
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Cursor = Cursors.Hand,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(16, 0, 16, 0),
                Background = GetAccentColorBrush(),
                Foreground = Brushes.White
            };
            saveButton.Click += Save_Click;

            buttonPanel.Children.Add(closeButton);
            buttonPanel.Children.Add(saveButton);
            footerBorder.Child = buttonPanel;

            Grid.SetRow(footerBorder, 2);
            rootGrid.Children.Add(footerBorder);
        }

        private Button CreateModernPrimaryButton(string text)
        {
            return new Button
            {
                Content = text,
                Height = 32,
                FontSize = 12,
                FontWeight = FontWeights.Medium,
                Cursor = Cursors.Hand,
                BorderThickness = new Thickness(1),
                Background = new SolidColorBrush(Color.FromRgb(66, 133, 244)),
                Foreground = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244))
            };
        }

        private Button CreateModernSecondaryButton(string text)
        {
            return new Button
            {
                Content = text,
                Height = 32,
                FontSize = 12,
                FontWeight = FontWeights.Medium,
                Cursor = Cursors.Hand,
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224))
            };
        }

        private void LoadPortalFencesIntoDropdown()
        {
            _cmbTarget.Items.Clear();

            // FIX: Pull live data and use the correct "Title" and "Path" schema properties!
            foreach (var f in FenceManager.GetFenceData())
            {
                try
                {
                    string fType = f is Newtonsoft.Json.Linq.JObject j ? j["ItemsType"]?.ToString() : f.GetType().GetProperty("ItemsType")?.GetValue(f)?.ToString();
                    if (fType == "Portal")
                    {
                        string fTitle = f is Newtonsoft.Json.Linq.JObject jn ? jn["Title"]?.ToString() : f.GetType().GetProperty("Title")?.GetValue(f)?.ToString();
                        string fPath = f is Newtonsoft.Json.Linq.JObject jp ? jp["Path"]?.ToString() : f.GetType().GetProperty("Path")?.GetValue(f)?.ToString();

                        if (!string.IsNullOrEmpty(fPath))
                        {
                            ComboBoxItem item = new ComboBoxItem { Content = $"Portal fence: {fTitle}", Tag = fPath };
                            _cmbTarget.Items.Add(item);
                        }
                    }
                }
                catch { }
            }
            _cmbTarget.Items.Add(new ComboBoxItem { Content = "Browse for Folder...", Tag = "BROWSE", FontWeight = FontWeights.Bold });
        }

        private void RefreshList()
        {
            _rulesList.ItemsSource = null;
            _rulesList.ItemsSource = _localRules;
            if (_selectedRule != null) _rulesList.SelectedItem = _selectedRule;
        }

        private void AddNewRule()
        {
            // Inherit the last selected options on the screen to be more user-friendly
            string newExt = "*.jpg; *.jpeg; *.png; *.gif; *.bmp; *.webp";
            if (_cmbCondition.SelectedItem != null)
            {
                string selected = _cmbCondition.SelectedItem.ToString();
                newExt = selected == "Custom Extensions..." ? _txtCustomExt.Text : _presets[selected];
            }

            string newTarget = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (_cmbTarget.SelectedItem is ComboBoxItem item && item.Tag != null && item.Tag.ToString() != "BROWSE")
            {
                newTarget = item.Tag.ToString();
            }

            RuleConflictAction newConflict = RuleConflictAction.Rename;
            if (_cmbConflict.SelectedIndex >= 0)
            {
                newConflict = (RuleConflictAction)_cmbConflict.SelectedIndex;
            }

            var rule = new OrganizeRule
            {
                Name = "New Rule",
                Extensions = newExt,
                TargetFolderPath = newTarget,
                ConflictAction = newConflict,
                AutoCreateFence = _chkAutoCreate.IsChecked == true,
                IsEnabled = true
            };

            _localRules.Add(rule);
            RefreshList();
            _rulesList.SelectedItem = rule;
        }

        private void RemoveRule()
        {
            if (_selectedRule == null) return;
            _localRules.Remove(_selectedRule);
            _selectedRule = null;
            _editorPanel.IsEnabled = false;
            RefreshList();
        }

        private void MoveRule(int direction)
        {
            if (_selectedRule == null) return;
            int idx = _localRules.IndexOf(_selectedRule);
            if (idx + direction < 0 || idx + direction >= _localRules.Count) return;

            _localRules.RemoveAt(idx);
            _localRules.Insert(idx + direction, _selectedRule);

            for (int i = 0; i < _localRules.Count; i++) _localRules[i].Priority = i;
            RefreshList();
        }

        private void LoadRuleEditor(OrganizeRule rule)
        {
            _isLoadingRule = true; // Lock events from firing during load
            _selectedRule = rule;
            _editorPanel.IsEnabled = rule != null;
            if (rule == null)
            {
                _txtName.Text = ""; _txtCustomExt.Text = ""; _txtNameContains.Text = ""; _chkAutoCreate.Visibility = Visibility.Collapsed;
                _chkIsEnabled.IsChecked = false;
                _isLoadingRule = false;
                return;
            }

            _txtName.Text = rule.Name;
            _chkIsEnabled.IsChecked = rule.IsEnabled;
            _txtNameContains.Text = rule.NameContains ?? "";
            // Add this to update the timestamp!
            _txtLastRun.Text = rule.LastRun.HasValue ? $"Last successful run: {rule.LastRun.Value.ToString("g")}" : "Last successful run: Never";
            _cmbConflict.SelectedIndex = (int)rule.ConflictAction;
            _chkAutoCreate.IsChecked = rule.AutoCreateFence;

            var matchingPreset = _presets.FirstOrDefault(p => p.Value == rule.Extensions).Key;
            if (matchingPreset != null)
            {
                _cmbCondition.SelectedItem = matchingPreset;
                _customExtBorder.Visibility = Visibility.Collapsed;
            }
            else
            {
                _cmbCondition.SelectedItem = "Custom Extensions...";
                _txtCustomExt.Text = rule.Extensions;
                _customExtBorder.Visibility = Visibility.Visible;
            }

            bool found = false;
            foreach (ComboBoxItem item in _cmbTarget.Items)
            {
                if (item.Tag?.ToString() == rule.TargetFolderPath)
                {
                    _cmbTarget.SelectedItem = item;
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                var customItem = new ComboBoxItem { Content = $"Custom path: {rule.TargetFolderPath}", Tag = rule.TargetFolderPath };
                _cmbTarget.Items.Insert(_cmbTarget.Items.Count - 1, customItem);
                _cmbTarget.SelectedItem = customItem;
            }

            _chkAutoCreate.Visibility = (_cmbTarget.SelectedItem as ComboBoxItem)?.Content.ToString().StartsWith("Custom path:") == true ? Visibility.Visible : Visibility.Collapsed;

            _isLoadingRule = false; // Unlock events
        }

        private void CmbCondition_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isLoadingRule || _selectedRule == null || _cmbCondition.SelectedItem == null) return;

            string selected = _cmbCondition.SelectedItem.ToString();

            if (selected == "Custom Extensions...")
            {
                _customExtBorder.Visibility = Visibility.Visible;
                _selectedRule.Extensions = _txtCustomExt.Text;
            }
            else
            {
                _customExtBorder.Visibility = Visibility.Collapsed;
                _selectedRule.Extensions = _presets[selected];
            }
        }

        private void CmbTarget_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isLoadingRule || _selectedRule == null || _cmbTarget.SelectedItem == null) return;
            var selectedItem = (ComboBoxItem)_cmbTarget.SelectedItem;

            if (selectedItem.Tag?.ToString() == "BROWSE")
            {
                using (var d = new System.Windows.Forms.FolderBrowserDialog())
                {
                    if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        _selectedRule.TargetFolderPath = d.SelectedPath;
                        LoadRuleEditor(_selectedRule);
                    }
                    else
                    {
                        LoadRuleEditor(_selectedRule);
                    }
                }
            }
            else
            {
                _selectedRule.TargetFolderPath = selectedItem.Tag.ToString();
                _chkAutoCreate.Visibility = selectedItem.Content.ToString().StartsWith("Custom path:") ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < _localRules.Count; i++) _localRules[i].Priority = i;
            AutoOrganizeManager.Rules = _localRules;
            AutoOrganizeManager.SaveRules();

            // FIX: We removed Stop() and Start() here. 
            // We want the engine to remain paused until you actually close the window.
            // When you close it, Resume() will run and process the new rules instantly.

            // Stay open, but confirm save!
            MessageBoxesManager.ShowOKOnlyMessageBoxForm("Your rules have been saved and applied successfully.", "Saved");
        }
    }
}
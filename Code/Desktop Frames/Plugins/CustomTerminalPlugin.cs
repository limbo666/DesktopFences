using Desktop_Frames.Localization;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Runtime.InteropServices;

namespace Desktop_Frames.Plugins
{
    public class CustomTerminalPlugin : IFramePlugin
    {
        public string PluginId => "CustomTerminal";
        public string DisplayName => "Terminal Emulator";
        public int DevelopmentState => 2; // Set to 1, 2, or 3 based on your testing phase

        private Process _terminalProcess;
        private Grid _rootGrid;
        private RichTextBox _outputBox;
        private TextBox _inputBox;
        private TextBlock _promptBlock;

        // Terminal state
        private List<string> _commandHistory = new List<string>();
        private int _historyIndex = -1;
        private readonly string _historyFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DesktopFrames_TerminalHistory.txt");
        private string _currentPath = "";
        
        // Plugin Settings
        private string _shellType = "powershell.exe"; // "powershell.exe" or "cmd.exe"
        private string _terminalColor = "Green"; // "Green", "White", "Cyan"
        private int _fontSize = 12;
        private string _startupDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
      

        // --- Windows API for Ctrl+C Signal ---
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AttachConsole(uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool FreeConsole();

        [DllImport("kernel32.dll")]
        static extern bool SetConsoleCtrlHandler(IntPtr HandlerRoutine, bool Add);

        [DllImport("kernel32.dll")]
        static extern bool GenerateConsoleCtrlEvent(uint dwCtrlEvent, uint dwProcessGroupId);

        private void SendCtrlC()
        {
            if (_terminalProcess == null || _terminalProcess.HasExited) return;

            AppendOutput("^C");

            // To send a signal, we must detach from current console, attach to the target, send the event, and restore.
            FreeConsole();
            if (AttachConsole((uint)_terminalProcess.Id))
            {
                SetConsoleCtrlHandler(IntPtr.Zero, true); // Prevent host app from catching its own Ctrl+C and dying
                GenerateConsoleCtrlEvent(0, 0);           // 0 is CTRL_C_EVENT
                System.Threading.Thread.Sleep(50);        // Give the console a split-second to process the interrupt
                FreeConsole();
                SetConsoleCtrlHandler(IntPtr.Zero, false);
            }
        }

        public FrameworkElement CreateVisualElement()
        {
            _rootGrid = new Grid
            {
                Background = new SolidColorBrush(Color.FromArgb(200, 10, 10, 10)),
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                // Adjust this global margin to center the whole terminal block inside the frame safely
                Margin = new Thickness(4, 4, 20, 18)
            };

            _rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Output area
            _rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Input area

            // Output Box (Read-only RichTextBox for Syntax Highlighting)
            _outputBox = new RichTextBox
            {
                IsReadOnly = true,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = GetBrushFromColorString(_terminalColor),
                FontFamily = new FontFamily("Consolas"),
                FontSize = _fontSize,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Padding = new Thickness(5),
                Margin = new Thickness(0, 0, 6, 0),
                Document = new FlowDocument { PagePadding = new Thickness(0), LineHeight = 1.5 }
            };

            // Ensure clicking the empty space focuses the input
            _outputBox.PreviewMouseLeftButtonUp += (s, e) =>
            {
                if (_outputBox.Selection.IsEmpty) FocusInput();
            };

            Grid.SetRow(_outputBox, 0);
            _rootGrid.Children.Add(_outputBox);

            // Clicking anywhere else in the background also focuses input
            _rootGrid.MouseLeftButtonUp += (s, e) => FocusInput();

            // Input Area
            Grid inputGrid = new Grid
            {
                Background = new SolidColorBrush(Color.FromArgb(255, 20, 20, 20)),
                Margin = new Thickness(0, 0, 6, 12) // Clears the bottom edge and the resize grip
            };
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            inputGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            Grid.SetRow(inputGrid, 1);

            _promptBlock = new TextBlock
            {
                Text = ">",
                Foreground = _outputBox.Foreground,
                FontFamily = new FontFamily("Consolas"),
                FontSize = _fontSize,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(5, 0, 5, 0),
                TextTrimming = TextTrimming.CharacterEllipsis,
                MaxWidth = 250 // Prevent exceptionally long paths from crushing the input box
            };
            Grid.SetColumn(_promptBlock, 0);
            inputGrid.Children.Add(_promptBlock);

            _inputBox = new TextBox
            {
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = _outputBox.Foreground,
                FontFamily = new FontFamily("Consolas"),
                FontSize = _fontSize,
                CaretBrush = _outputBox.Foreground,
                Padding = new Thickness(0, 5, 5, 5)
            };
            // --- FIX: Instantly trigger focus mode when clicked ---
            _inputBox.PreviewMouseLeftButtonDown += (s, e) => FocusInput();

            // Switch to PreviewKeyDown to intercept keys before the non-activating window swallows them
            _inputBox.PreviewKeyDown += InputBox_PreviewKeyDown;

            Grid.SetColumn(_inputBox, 1);
            inputGrid.Children.Add(_inputBox);

            _rootGrid.Children.Add(inputGrid);

            return _rootGrid;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            LoadHistory();
            ApplySettingsData(settings);
            ApplyUIStyles();
            StartProcess();
        }

        private void StartProcess()
        {
            Cleanup(); // Ensure no orphan processes

            try
            {
                _terminalProcess = new Process();
                _terminalProcess.StartInfo.FileName = _shellType;

                // Apply Startup Directory (fallback to C:\ if invalid)
                if (System.IO.Directory.Exists(_startupDir))
                    _currentPath = _startupDir;
                else
                    _currentPath = "C:\\";

                _terminalProcess.StartInfo.WorkingDirectory = _currentPath;
                UpdatePromptUI();

                _terminalProcess.StartInfo.RedirectStandardOutput = true;



                _terminalProcess.StartInfo.RedirectStandardError = true;
                _terminalProcess.StartInfo.RedirectStandardInput = true;
                _terminalProcess.StartInfo.UseShellExecute = false;
                _terminalProcess.StartInfo.CreateNoWindow = true;
                _terminalProcess.StartInfo.StandardOutputEncoding = Encoding.UTF8;
                _terminalProcess.StartInfo.StandardErrorEncoding = Encoding.UTF8;

                _terminalProcess.OutputDataReceived += (s, e) => AppendOutput(e.Data);
                _terminalProcess.ErrorDataReceived += (s, e) => AppendOutput(e.Data);

                _terminalProcess.Start();
                _terminalProcess.BeginOutputReadLine();
                _terminalProcess.BeginErrorReadLine();

                AppendOutput($"Desktop Frames + Terminal Emulator");
                AppendOutput($"[{_shellType} initialized in {_terminalProcess.StartInfo.WorkingDirectory}]");
                AppendOutput("--------------------------------------------------");
            }
            catch (Exception ex)
            {
                AppendOutput($"[Terminal Initialization Error: {ex.Message}]");
            }
        }


        private void LoadHistory()
        {
            try
            {
                if (File.Exists(_historyFilePath))
                {
                    _commandHistory = File.ReadAllLines(_historyFilePath).ToList();

                    // Cap at 100 commands to prevent bloat
                    if (_commandHistory.Count > 100)
                        _commandHistory = _commandHistory.Skip(_commandHistory.Count - 100).ToList();

                    _historyIndex = _commandHistory.Count;
                }
            }
            catch { /* Fail silently */ }
        }

        private void SaveHistory()
        {
            try { File.WriteAllLines(_historyFilePath, _commandHistory); }
            catch { /* Fail silently */ }
        }

        private void ClearHistory()
        {
            _commandHistory.Clear();
            _historyIndex = -1;
            try { if (File.Exists(_historyFilePath)) File.Delete(_historyFilePath); }
            catch { /* Fail silently */ }
        }

        private void UpdatePromptUI()
        {
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (_promptBlock != null)
                {
                    // Format nicely depending on whether we are in a drive root (C:\) or a subfolder (C:\Folder)
                    _promptBlock.Text = _currentPath.EndsWith("\\") ? $"{_currentPath}>" : $"{_currentPath}\\>";
                }
            });
        }
        private void FocusInput()
        {
            Application.Current.Dispatcher.BeginInvoke(new Action(() => {

                // --- CRITICAL FIX: Bypass NonActivatingWindow Protection ---
                // We ask the Host App to temporarily allow keyboard input for this text box
                try
                {
                    Window parentWin = Window.GetWindow(_rootGrid);
                    if (parentWin != null)
                    {
                        dynamic dynWin = parentWin;
                        dynWin.BeginKeyboardInteractiveEdit(_inputBox);
                    }
                }
                catch { } // Failsafe: Ignore if the host method is unavailable
                // -----------------------------------------------------------

                _inputBox?.Focus();
                if (_inputBox != null) _inputBox.CaretIndex = _inputBox.Text.Length;
            }), DispatcherPriority.Input);
        }

        private void InputBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Detect Ctrl + C
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.C)
            {
                if (_inputBox.SelectionLength == 0)
                {
                    e.Handled = true;
                    SendCtrlC();
                    return;
                }
            }

            // Detect Ctrl + L (Clear Screen Muscle Memory)
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.L)
            {
                e.Handled = true;
                _outputBox.Document.Blocks.Clear();
                return;
            }

            // Fix WPF's native Ctrl + Backspace bug (delete whole word)
            if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.Back)
            {
                e.Handled = true;
                int caret = _inputBox.CaretIndex;
                if (caret > 0)
                {
                    int spaceIdx = _inputBox.Text.LastIndexOf(' ', caret - 2);
                    int cutIndex = spaceIdx < 0 ? 0 : spaceIdx + 1;
                    _inputBox.Text = _inputBox.Text.Remove(cutIndex, caret - cutIndex);
                    _inputBox.CaretIndex = cutIndex;
                }
                return;
            }

            if (e.Key == Key.Enter)
            {
                e.Handled = true; // Stop the beep sound

                string command = _inputBox.Text.Trim();
                string commandToSend = command; // DECLARED HERE FOR GLOBAL SCOPE
                _inputBox.Clear();


                // Only process history and local commands if there is actual text
                if (!string.IsNullOrWhiteSpace(command))
                {
                    // Add to history
                    if (_commandHistory.Count == 0 || _commandHistory[_commandHistory.Count - 1] != command)
                    {
                        _commandHistory.Add(command);
                        SaveHistory();
                    }
                    _historyIndex = _commandHistory.Count;

                    // --- LOCAL COMMAND INTERCEPTION ---
                    string cmdLower = command.ToLower();
                    if (cmdLower == "clear" || cmdLower == "cls")
                    {
                        _outputBox.Document.Blocks.Clear();
                        return; // Do not send to hidden process
                    }
                    if (cmdLower == "exit")
                    {
                        _outputBox.Document.Blocks.Clear();
                        AppendOutput("[Restarting Shell Session...]");
                        StartProcess();
                        return;
                    }

                    // --- PATH TRACKING INTERCEPTION ---
                    if (cmdLower.StartsWith("cd ") || cmdLower == "cd..")
                    {
                        // Auto-correct CMD syntax so PowerShell doesn't fail and desync the path
                        if (cmdLower == "cd..") commandToSend = "cd ..";

                        string targetDir = cmdLower == "cd.." ? ".." : command.Substring(3).Trim();
                        try
                        {
                            string combined = System.IO.Path.Combine(_currentPath, targetDir);
                            string resolved = System.IO.Path.GetFullPath(combined);
                            if (System.IO.Directory.Exists(resolved))
                            {
                                _currentPath = resolved;
                                UpdatePromptUI();
                            }
                        }
                        catch { }
                    }
                    else if (cmdLower.Length == 2 && cmdLower[1] == ':')
                    {
                        try
                        {
                            if (System.IO.Directory.Exists(command))
                            {
                                _currentPath = command.ToUpper() + "\\";
                                UpdatePromptUI();
                            }
                        }
                        catch { }
                    }
                    // ----------------------------------
                }

                if (_terminalProcess != null && !_terminalProcess.HasExited)
                {
                    try
                    {
                        // Note: We completely removed the manual C# AppendOutput echo here. 
                        // The native hidden shell (cmd/powershell) automatically echoes its 
                        // true prompt and the command, preventing double-printing and desyncs!
                        _terminalProcess.StandardInput.WriteLine(commandToSend);
                    }
                    catch (Exception ex)
                    {
                        AppendOutput($"[Input Error: {ex.Message}]");
                    }
                }
                else
                {
                    AppendOutput("[Process exited. Attempting restart...]");
                    StartProcess();
                }
            }
            else if (e.Key == Key.Up)
            {
                e.Handled = true; // Force interception

                if (_historyIndex > 0)
                {
                    _historyIndex--;
                    _inputBox.Text = _commandHistory[_historyIndex];
                    _inputBox.CaretIndex = _inputBox.Text.Length;
                }
            }
            else if (e.Key == Key.Down)
            {
                e.Handled = true; // Force interception

                if (_historyIndex < _commandHistory.Count - 1)
                {
                    _historyIndex++;
                    _inputBox.Text = _commandHistory[_historyIndex];
                    _inputBox.CaretIndex = _inputBox.Text.Length;
                }
                else
                {
                    _historyIndex = _commandHistory.Count;
                    _inputBox.Clear();
                }
            }
        }


        private Paragraph ColorizeLine(string text)
        {
            Paragraph paragraph = new Paragraph { Margin = new Thickness(0) };

            // 1. Determine Base Line Color based on IT/SysAdmin Keywords
            Brush lineBrush = GetBrushFromColorString(_terminalColor); // Default
            string lowerText = text.ToLower();

            if (Regex.IsMatch(lowerText, @"\b(error|exception|fail|failed|denied|not recognized|timed out|could not find)\b"))
            {
                lineBrush = Brushes.Tomato;
            }
            else if (Regex.IsMatch(lowerText, @"\b(warning|unreachable|bad command|not found)\b"))
            {
                lineBrush = Brushes.Gold;
            }
            else if (Regex.IsMatch(lowerText, @"\b(reply from|success|copied|bytes=)\b"))
            {
                lineBrush = Brushes.LightGreen;
            }

            // 2. Extract and Highlight Specific Data (IPs, MACs, Paths)
            // Regex for IPv4
            string ipPattern = @"\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b";
            // Regex for common paths (C:\folder\file)
            string pathPattern = @"[a-zA-Z]:\\[^:\*\?\""<>\|]*";

            var combinedRegex = new Regex($"({ipPattern})|({pathPattern})", RegexOptions.IgnoreCase);
            var matches = combinedRegex.Matches(text);

            if (matches.Count == 0)
            {
                // No special data, just add the whole line
                paragraph.Inlines.Add(new Run(text) { Foreground = lineBrush });
            }
            else
            {
                // Reconstruct the line with highlighted segments
                int lastIndex = 0;
                foreach (Match match in matches)
                {
                    // Add text before the match
                    if (match.Index > lastIndex)
                    {
                        paragraph.Inlines.Add(new Run(text.Substring(lastIndex, match.Index - lastIndex)) { Foreground = lineBrush });
                    }

                    // Add the highlighted match (Cyan for IPs, Light Blue for Paths)
                    Brush highlightBrush = Regex.IsMatch(match.Value, ipPattern) ? Brushes.Cyan : Brushes.LightSkyBlue;
                    paragraph.Inlines.Add(new Run(match.Value) { Foreground = highlightBrush, FontWeight = FontWeights.Bold });

                    lastIndex = match.Index + match.Length;
                }

                // Add remaining text
                if (lastIndex < text.Length)
                {
                    paragraph.Inlines.Add(new Run(text.Substring(lastIndex)) { Foreground = lineBrush });
                }
            }

            return paragraph;
        }

        private void AppendOutput(string text)
        {
            if (text == null) return;

            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                bool isScrolledUp = _outputBox.VerticalOffset < (_outputBox.ExtentHeight - _outputBox.ViewportHeight - 5);

                // Pass text through the color engine
                Paragraph p = ColorizeLine(text);
                _outputBox.Document.Blocks.Add(p);

                // RAM Protection: Prevent the RichTextBox from holding too much history
                if (_outputBox.Document.Blocks.Count > 2000)
                {
                    _outputBox.Document.Blocks.Remove(_outputBox.Document.Blocks.FirstBlock);
                }

                if (!isScrolledUp)
                {
                    _outputBox.ScrollToEnd();
                }
            }, DispatcherPriority.Render);
        }

        private Brush GetBrushFromColorString(string colorStr)
        {
            switch (colorStr.ToLower())
            {
                case "green": return new SolidColorBrush(Color.FromRgb(12, 192, 40)); // Matrix Green
                case "cyan": return Brushes.Cyan;
                case "white": return Brushes.LightGray;
                default: return Brushes.LightGray;
            }
        }

        private void ApplySettingsData(Dictionary<string, object> settings)
        {
            if (settings != null)
            {
                if (settings.ContainsKey("ShellType")) _shellType = settings["ShellType"].ToString();
                if (settings.ContainsKey("TerminalColor")) _terminalColor = settings["TerminalColor"].ToString();
                if (settings.ContainsKey("FontSize") && int.TryParse(settings["FontSize"].ToString(), out int size)) _fontSize = size;
                if (settings.ContainsKey("StartupDir")) _startupDir = settings["StartupDir"].ToString();
            }
        }

        private void ApplyUIStyles()
        {
            if (_outputBox != null)
            {
                _outputBox.Foreground = GetBrushFromColorString(_terminalColor);
                _outputBox.FontSize = _fontSize;
            }
            if (_inputBox != null)
            {
                _inputBox.Foreground = GetBrushFromColorString(_terminalColor);
                _inputBox.CaretBrush = GetBrushFromColorString(_terminalColor);
                _inputBox.FontSize = _fontSize;
            }
            if (_promptBlock != null)
            {
                _promptBlock.Foreground = GetBrushFromColorString(_terminalColor);
                _promptBlock.FontSize = _fontSize;
            }
        }

        public void Pause()
        {
            // Do not kill the process on pause, just allow it to idle in background.
        }

        public void Resume()
        {
            // Focus the input box when the frame comes back into view
            _inputBox?.Focus();
        }

        public void Cleanup()
        {
            try
            {
                if (_terminalProcess != null && !_terminalProcess.HasExited)
                {
                    _terminalProcess.CancelOutputRead();
                    _terminalProcess.CancelErrorRead();
                    _terminalProcess.Kill();
                    _terminalProcess.Dispose();
                }
            }
            catch { /* Fail silently */ }
        }

        // ==========================================
        // STANDARD UI BOILERPLATE FOR DESKTOP FRAMES +
        // ==========================================
        public void ShowSettingsWindow(Window ownerWindow, dynamic frameData)
        {
            SolidColorBrush accentBrush;
            try
            {
                // Note: Assumes Utility and SettingsManager exist in Host application context
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
                Text = Strings.TermSettings,
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
            groupSp.Children.Add(new TextBlock { Text = Strings.TermConsoleOptions, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 10) });

            groupSp.Children.Add(new TextBlock { Text = Strings.TermTargetShell, Margin = new Thickness(0, 0, 0, 5) });
            ComboBox cmbShell = new ComboBox { Margin = new Thickness(0, 0, 0, 15) };
            cmbShell.Items.Add("powershell.exe");
            cmbShell.Items.Add("cmd.exe");
            cmbShell.SelectedItem = _shellType;
            groupSp.Children.Add(cmbShell);

            groupSp.Children.Add(new TextBlock { Text = Strings.LblTextColor, Margin = new Thickness(0, 0, 0, 5) });
            ComboBox cmbColor = new ComboBox { Margin = new Thickness(0, 0, 0, 15) };
            cmbColor.Items.Add("Green");
            cmbColor.Items.Add("White");
            cmbColor.Items.Add("Cyan");
            cmbColor.SelectedItem = _terminalColor;
            groupSp.Children.Add(cmbColor);

            groupSp.Children.Add(new TextBlock { Text = Strings.LblFontSize, Margin = new Thickness(0, 0, 0, 5) });
            TextBox txtFontSize = new TextBox { Text = _fontSize.ToString(), Margin = new Thickness(0, 0, 0, 15) };
            groupSp.Children.Add(txtFontSize);

            groupSp.Children.Add(new TextBlock { Text = Strings.TermStartupDirectory, Margin = new Thickness(0, 0, 0, 5) });

            Grid dirGrid = new Grid { Margin = new Thickness(0, 0, 0, 15) };
            dirGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            dirGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            TextBox txtStartupDir = new TextBox { Text = _startupDir, Margin = new Thickness(0, 0, 5, 0) };
            Grid.SetColumn(txtStartupDir, 0);
            dirGrid.Children.Add(txtStartupDir);

            Button btnBrowse = new Button { Content = "...", Width = 30 };
            btnBrowse.Click += (s, e) =>
            {
                // Native .NET 8 WPF Folder Dialog
                var dialog = new Microsoft.Win32.OpenFolderDialog
                {
                    Title = Strings.TermSelectStartupDir,
                    InitialDirectory = System.IO.Directory.Exists(txtStartupDir.Text) ? txtStartupDir.Text : "C:\\"
                };

                if (dialog.ShowDialog() == true)
                {
                    txtStartupDir.Text = dialog.FolderName;
                }
            };
            Grid.SetColumn(btnBrowse, 1);
            dirGrid.Children.Add(btnBrowse);

            groupSp.Children.Add(dirGrid);

            // Clear History Button
            Button btnClearHistory = new Button
            {
                Content = Strings.TermClearHistory,
                Background = Brushes.White,
                BorderBrush = Brushes.LightGray,
                Padding = new Thickness(5),
                HorizontalAlignment = HorizontalAlignment.Left
            };
            btnClearHistory.Click += (s, e) =>
            {
                ClearHistory();
                btnClearHistory.Content = Strings.TermHistoryCleared;
                btnClearHistory.IsEnabled = false;
            };
            groupSp.Children.Add(btnClearHistory);

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
                Content = Strings.BtnCancel,
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
                Content = Strings.BtnSave,
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
                    { "ShellType", cmbShell.SelectedItem.ToString() },
                    { "TerminalColor", cmbColor.SelectedItem.ToString() },
                    { "FontSize", txtFontSize.Text },
                    { "StartupDir", txtStartupDir.Text }
                };

                if (frameData is Newtonsoft.Json.Linq.JObject jFrame)
                    jFrame["PluginSettings"] = Newtonsoft.Json.Linq.JObject.FromObject(newSettings);
                else
                    ((IDictionary<string, object>)frameData)["PluginSettings"] = newSettings;

                try { FrameDataManager.SaveFrameData(); } catch { }
                ApplySettingsData(newSettings);
                ApplyUIStyles();

                // If shell type changed, restart the process
                if (_shellType != cmbShell.SelectedItem.ToString())
                {
                    _outputBox.Document.Blocks.Clear();
                    StartProcess();
                }

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
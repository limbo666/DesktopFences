using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Desktop_Frames.Plugins
{
    public class CalculatorPlugin : IFramePlugin
    {
        public string PluginId => "MinimalCalculator";
        public string DisplayName => "Minimal Calculator";
        public int DevelopmentState => 1;

        // UI Elements
        private Grid _rootGrid;
        private SolidColorBrush _rootBgBrush;

        private Border _displayCard;
        private TextBox _mainDisplay;
        private TextBlock _operatorSymbol;
        private TextBlock _operationLabelText;
        private TextBlock _liveExpressionText;
        private SolidColorBrush _displayBorderBrush;
        private SolidColorBrush _displayForegroundBrush;

        private Border _historyCard;
        private ListBox _historyListBox;

        private Border _keypadCard;
        private UniformGrid _virtualKeypad;

        // Math State Engine
        private double _currentValue = 0;
        private double _previousValue = 0;
        private double _memoryValue = 0;
        private string _currentOperator = "";
        private bool _isNewInput = true;
        private string _percentageString = ""; // Tracks raw percentage input for history

        // Settings & Configurations
        private List<string> _historyTape = new List<string>();
        private int _maxHistory = 20;

        // Combo Tracking
        private int _clearKeyCount = 0;
        private DateTime _lastClearKeyPress = DateTime.MinValue;
        private bool _showVirtualKeypad = false;
        private bool _showHistoryTape = true;
        private string _displayColor = "White";
        private string _historyColor = "Gray";
        private int _symbolFadeMs = 1500;
        private bool _clearAfterEquals = true; // Matches MS Calculator behavior
        private bool _showOperationLabel = true; // Enables tiny expiring text

        // Animation Timers
        private DispatcherTimer _fadeTimer;

        // ==========================================
        // WINDOWS API: LOW-LEVEL KEYBOARD HOOK
        // ==========================================
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;
        private const int VK_SHIFT = 0x10;

        private LowLevelKeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        // ==========================================
        // CORE INITIALIZATION & UI BUILDER
        // ==========================================
        public FrameworkElement CreateVisualElement()
        {
            _proc = HookCallback;
            _fadeTimer = new DispatcherTimer();

            // Assign the tick handler EXACTLY ONCE to prevent memory leaks
            _fadeTimer.Tick += (s, e) =>
            {
                _fadeTimer.Stop();
                if (_operatorSymbol != null)
                {
                    _operatorSymbol.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(500)));
                }
            };

            // Secure Brush Instantiation (Fixes the Null Reference Crashes)
            _rootBgBrush = new SolidColorBrush(Color.FromArgb(15, 255, 255, 255));
            _displayBorderBrush = new SolidColorBrush(Color.FromArgb(0, 66, 133, 244));
            _displayForegroundBrush = new SolidColorBrush(Colors.White);

            _rootGrid = new Grid
            {
                Background = _rootBgBrush,
                // Thickness(Left, Top, Right, Bottom)
                // Increase the 3rd (Right) and 4th (Bottom) values to pull the entire plugin inward away from the frame edges:
                Margin = new Thickness(5, 5, 25, 10),
                MinWidth = 220
            };

            // HOVER ENGINE & BLINK + FOCUS RING EFFECT
            _rootGrid.MouseEnter += (s, e) =>
            {
                InstallHook();

                ColorAnimation blinkAnim = new ColorAnimation { From = Color.FromArgb(70, 255, 255, 255), To = Color.FromArgb(15, 255, 255, 255), Duration = TimeSpan.FromMilliseconds(400), EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut } };
                _rootBgBrush.BeginAnimation(SolidColorBrush.ColorProperty, blinkAnim);

                // Dynamically grab the selected display color and apply 200/255 opacity for the glow
                Color activeBorderColor = ParseColorSafe(_displayColor, Colors.White);
                activeBorderColor.A = 200;

                ColorAnimation focusAnim = new ColorAnimation { To = activeBorderColor, Duration = TimeSpan.FromMilliseconds(200) };
                _displayBorderBrush.BeginAnimation(SolidColorBrush.ColorProperty, focusAnim);
            };

            _rootGrid.MouseLeave += (s, e) =>
            {
                UninstallHook();

                // Fade out using the selected display color with 0 opacity
                Color inactiveBorderColor = ParseColorSafe(_displayColor, Colors.White);
                inactiveBorderColor.A = 0;

                ColorAnimation unfocusAnim = new ColorAnimation { To = inactiveBorderColor, Duration = TimeSpan.FromMilliseconds(300) };
                _displayBorderBrush.BeginAnimation(SolidColorBrush.ColorProperty, unfocusAnim);
            };

            _rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            _rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            _rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            // --- 1. Main Display Card ---
            _displayCard = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(15, 0, 0, 0)),
                BorderBrush = _displayBorderBrush,
                BorderThickness = new Thickness(1.5),
                CornerRadius = new CornerRadius(6),
                Margin = new Thickness(10, 5, 10, 15), // The Memorized Margin Blueprint
                Height = 80,
                Padding = new Thickness(10)
            };

            Grid displayLayout = new Grid();
            displayLayout.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            displayLayout.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            _operationLabelText = new TextBlock
            {
                FontSize = 10,
                Foreground = Brushes.Gray,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, -5, 0, 0),
                Opacity = 0, // Hidden initially
                IsHitTestVisible = false
            };
            Grid.SetColumnSpan(_operationLabelText, 2);
            displayLayout.Children.Add(_operationLabelText);

            _liveExpressionText = new TextBlock
            {
                FontSize = 13,
                FontFamily = new FontFamily("Lucida Console"),
                VerticalAlignment = VerticalAlignment.Top,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, -2, 0, 0),
                Opacity = 0.85,
                IsHitTestVisible = false
            };
            Grid.SetColumn(_liveExpressionText, 1);
            displayLayout.Children.Add(_liveExpressionText);

            _operatorSymbol = new TextBlock
            {
                FontSize = 20,
                Foreground = Brushes.DeepSkyBlue,
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                TextAlignment = TextAlignment.Center,
                Width = 25, // Fixed width prevents the main display from jumping left/right
                Margin = new Thickness(5, 0, 15, 0),
                Opacity = 0 // Invisible until triggered
            };
            Grid.SetColumn(_operatorSymbol, 0);
            displayLayout.Children.Add(_operatorSymbol);

            _mainDisplay = new TextBox
            {
                Text = "0",
                FontSize = 36,
                FontWeight = FontWeights.Bold,
                FontFamily = new FontFamily("Lucida Console"),
                Foreground = _displayForegroundBrush,
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                TextAlignment = TextAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center,
                IsReadOnly = true,
                Cursor = Cursors.Arrow
            };

            // Context Menu: Display Card
            ContextMenu displayMenu = new ContextMenu();
            MenuItem copyItem = new MenuItem { Header = "Copy" };
            copyItem.Click += (s, e) => { Clipboard.SetText(_mainDisplay.Text); TriggerTextFlash(Color.FromRgb(0, 255, 200)); };

            MenuItem pasteItem = new MenuItem { Header = "Paste" };
            pasteItem.Click += (s, e) =>
            {
                if (Clipboard.ContainsText())
                {
                    string pasteText = Clipboard.GetText().Replace(",", ".");
                    string cleanText = new string(pasteText.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());
                    if (double.TryParse(cleanText, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                    {
                        _currentValue = result; _isNewInput = true; UpdateDisplay(); TriggerTextFlash(Color.FromRgb(0, 255, 200));
                    }
                }
            };

            MenuItem clearItem = new MenuItem { Header = "Clear" };
            clearItem.Click += (s, e) => { ProcessInput("C"); }; // Routes through standard math engine for full reset + flash

            displayMenu.Items.Add(copyItem);
            displayMenu.Items.Add(pasteItem);
            displayMenu.Items.Add(new Separator());
            displayMenu.Items.Add(clearItem);

            _mainDisplay.ContextMenu = displayMenu;

            Grid.SetColumn(_mainDisplay, 1);
            displayLayout.Children.Add(_mainDisplay);

            _displayCard.Child = displayLayout;
            Grid.SetRow(_displayCard, 0);
            _rootGrid.Children.Add(_displayCard);

            // --- 2. History Scrolling Card ---
            _historyCard = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(10, 255, 255, 255)),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(10, 5, 10, 15), // The Memorized Margin Blueprint
                Padding = new Thickness(5),
                Height = 65,
                Visibility = Visibility.Collapsed
            };

            _historyListBox = new ListBox
            {
                Background = Brushes.Transparent,
                BorderThickness = new Thickness(0),
                Foreground = new SolidColorBrush(Colors.Gray),
                FontFamily = new FontFamily("Lucida Console"),
                FontSize = 12,
                HorizontalContentAlignment = HorizontalAlignment.Right,
                Focusable = false,
                Cursor = Cursors.ScrollAll
            };
            ScrollViewer.SetVerticalScrollBarVisibility(_historyListBox, ScrollBarVisibility.Hidden);

            // Context Menu: History Tape
            ContextMenu historyMenu = new ContextMenu();
            MenuItem clearHistoryItem = new MenuItem { Header = "Clear History" };
            clearHistoryItem.Click += (s, e) =>
            {
                _historyTape.Clear();
                SaveState();
                UpdateDisplay();
            };
            historyMenu.Items.Add(clearHistoryItem);

            _historyListBox.ContextMenu = historyMenu;

            _historyCard.Child = _historyListBox;
            Grid.SetRow(_historyCard, 1);
            _rootGrid.Children.Add(_historyCard);

            // --- 3. Virtual Keypad Card ---
            _keypadCard = new Border
            {
                Margin = new Thickness(10, 5, 10, 15), // The Memorized Margin Blueprint
                Visibility = Visibility.Collapsed,
                MinHeight = 200
            };

            _virtualKeypad = new UniformGrid { Columns = 4, Rows = 6 };
            string[] buttons = { "MC", "MR", "M+", "M-", "CE", "C", "⌫", "÷", "7", "8", "9", "×", "4", "5", "6", "-", "1", "2", "3", "+", "%", "0", ".", "=" };
            foreach (string btnText in buttons) _virtualKeypad.Children.Add(CreateButton(btnText));

            _keypadCard.Child = _virtualKeypad;
            Grid.SetRow(_keypadCard, 2);
            _rootGrid.Children.Add(_keypadCard);

            return _rootGrid;
        }

        public void Initialize(FrameworkElement visual, Dictionary<string, object> settings)
        {
            ApplySettingsData(settings);
            LoadState();
            UpdateDisplay();
            ApplyLayoutSettings();
        }

        // ==========================================
        // VISUAL FEEDBACK ENGINE
        // ==========================================
        private Color ParseColorSafe(string colorName, Color fallback)
        {
            try { return (Color)ColorConverter.ConvertFromString(colorName); }
            catch { return fallback; }
        }

        private void TriggerTextFlash(Color flashColor)
        {
            if (_displayForegroundBrush == null) return; // Ultimate crash protection

            Color targetColor = ParseColorSafe(_displayColor, Colors.White);
            ColorAnimation textBlink = new ColorAnimation
            {
                From = flashColor,
                To = targetColor,
                Duration = TimeSpan.FromMilliseconds(800),
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };
            _displayForegroundBrush.BeginAnimation(SolidColorBrush.ColorProperty, textBlink);
        }

        private void TriggerOperationLabel(string text)
        {
            if (!_showOperationLabel || _operationLabelText == null) return;

            // Sever old animations instantly
            _operationLabelText.BeginAnimation(UIElement.OpacityProperty, null);

            _operationLabelText.Text = text;
            _operationLabelText.Opacity = 1;

            DoubleAnimation fade = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromMilliseconds(800),
                BeginTime = TimeSpan.FromMilliseconds(1200) // Stays solid for 1.2s before fading
            };
            _operationLabelText.BeginAnimation(UIElement.OpacityProperty, fade);
        }

        private void TriggerSymbol(string symbol, Color color)
        {
            if (_operatorSymbol == null) return;

            _fadeTimer.Stop();

            // CRITICAL: Release the WPF animation lock so we can manually set Opacity to 1
            _operatorSymbol.BeginAnimation(UIElement.OpacityProperty, null);

            _operatorSymbol.Text = symbol;
            _operatorSymbol.Foreground = new SolidColorBrush(color);
            _operatorSymbol.Opacity = 1;

            if (_symbolFadeMs > 0) // If 0, it stays on screen until cleared or changed
            {
                _fadeTimer.Interval = TimeSpan.FromMilliseconds(_symbolFadeMs);
                _fadeTimer.Start();
            }
        }

        // ==========================================
        // THE HOOK ENGINE (Bypassing WPF Focus)
        // ==========================================
        private void InstallHook()
        {
            if (_hookID == IntPtr.Zero)
            {
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    _hookID = SetWindowsHookEx(WH_KEYBOARD_LL, _proc, GetModuleHandle(curModule.ModuleName), 0);
                }
            }
        }

        private void UninstallHook()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Key key = KeyInterop.KeyFromVirtualKey(vkCode);
                bool isShiftPressed = (GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
                bool isCtrlPressed = (GetAsyncKeyState(0x11) & 0x8000) != 0;

                bool handled = false;

                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (isCtrlPressed && key == Key.C) { Clipboard.SetText(_mainDisplay.Text); TriggerTextFlash(Color.FromRgb(0, 255, 200)); handled = true; return; }
                    if (isCtrlPressed && key == Key.V)
                    {
                        if (Clipboard.ContainsText())
                        {
                            string pasteText = Clipboard.GetText().Replace(",", ".");
                            string cleanText = new string(pasteText.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());
                            if (double.TryParse(cleanText, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                            {
                                _currentValue = result; _isNewInput = true; UpdateDisplay();
                                TriggerTextFlash(Color.FromRgb(0, 255, 200));
                            }
                        }
                        handled = true; return;
                    }

                    handled = true;
                    switch (key)
                    {
                        case Key.NumPad0: case Key.D0: ProcessInput("0"); break;
                        case Key.NumPad1: case Key.D1: ProcessInput("1"); break;
                        case Key.NumPad2: case Key.D2: ProcessInput("2"); break;
                        case Key.NumPad3: case Key.D3: ProcessInput("3"); break;
                        case Key.NumPad4: case Key.D4: ProcessInput("4"); break;
                        case Key.NumPad5: case Key.D5: if (isShiftPressed) ProcessInput("%"); else ProcessInput("5"); break;
                        case Key.NumPad6: case Key.D6: ProcessInput("6"); break;
                        case Key.NumPad7: case Key.D7: ProcessInput("7"); break;
                        case Key.NumPad8: case Key.D8: if (isShiftPressed) ProcessInput("×"); else ProcessInput("8"); break;
                        case Key.NumPad9: case Key.D9: ProcessInput("9"); break;
                        case Key.Decimal: case Key.OemPeriod: case Key.OemComma: ProcessInput("."); break;
                        case Key.Add: ProcessInput("+"); break;
                        case Key.OemPlus: if (isShiftPressed) ProcessInput("+"); else ProcessInput("="); break;
                        case Key.Subtract: case Key.OemMinus: ProcessInput("-"); break;
                        case Key.Multiply: ProcessInput("×"); break;
                        case Key.Divide: case Key.OemQuestion: ProcessInput("÷"); break;
                        case Key.Enter: ProcessInput("="); break;
                        case Key.Back: ProcessInput("⌫"); break;
                        case Key.Delete: ProcessInput("C"); break; // UX Fix: Delete now acts as a full clear
                        case Key.Escape: ProcessInput("C"); break;
                        default: handled = false; break;
                    }
                });

                if (handled) return (IntPtr)1;
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        // ==========================================
        // MATH & LOGIC ENGINE
        // ==========================================
        private void ProcessInput(string input)
        {
            // Strip out memory tags and thousands separators so the math engine reads pure numbers
            string rawText = _mainDisplay.Text.Replace(" [M]", "").Replace(",", "");

            if (char.IsDigit(input, 0) || input == ".")
            {
                if (_isNewInput || rawText == "Infinity" || rawText == "-Infinity" || rawText == "NaN")
                {
                    rawText = input == "." ? "0." : input;
                    _isNewInput = false;
                }
                else
                {
                    if (input == "." && rawText.Contains(".")) return;
                    rawText += input;
                }

                double.TryParse(rawText, NumberStyles.Any, CultureInfo.InvariantCulture, out _currentValue);

                // UX: Live Thousands Separator formatting
                string formattedText = rawText;
                var parts = rawText.Split('.');
                if (long.TryParse(parts[0], out long wholeNumber))
                {
                    formattedText = wholeNumber.ToString("#,##0", CultureInfo.InvariantCulture);
                    if (parts.Length > 1) formattedText += "." + parts[1]; // Preserve trailing decimals perfectly
                }

                _mainDisplay.Text = formattedText + (_memoryValue != 0 ? " [M]" : "");
            }
            else if (input == "⌫")
            {
                // Removed the strict !_isNewInput check so you can backspace loaded states or recent calculation results
                if (rawText.Length > 0 && rawText != "0")
                {
                    if (rawText == "Infinity" || rawText == "-Infinity" || rawText == "NaN")
                    {
                        rawText = "0";
                        _isNewInput = true;
                    }
                    else
                    {
                        rawText = rawText.Substring(0, rawText.Length - 1);
                        if (rawText == "" || rawText == "-")
                        {
                            rawText = "0";
                            _isNewInput = true;
                        }
                        else
                        {
                            // Transition into active editing mode if we backspaced a loaded or calculated number
                            _isNewInput = false;
                        }
                    }
                    double.TryParse(rawText, NumberStyles.Any, CultureInfo.InvariantCulture, out _currentValue);

                    // UX: Live Thousands Separator formatting for Backspace
                    string formattedText = rawText;
                    var parts = rawText.Split('.');
                    if (long.TryParse(parts[0], out long wholeNumber))
                    {
                        formattedText = wholeNumber.ToString("#,##0", CultureInfo.InvariantCulture);
                        if (parts.Length > 1) formattedText += "." + parts[1];
                    }

                    _mainDisplay.Text = formattedText + (_memoryValue != 0 ? " [M]" : "");
                }
            }
            else if (input == "C" || input == "CE")
            {
                // 1. Math Reset (C vs CE)
                if (input == "C")
                {
                    _currentValue = 0; _previousValue = 0; _currentOperator = ""; _isNewInput = true; _percentageString = "";
                    if (_operatorSymbol != null) _operatorSymbol.Opacity = 0;
                    TriggerOperationLabel("Cleared");
                }
                else // CE
                {
                    // Smart UX: If the entry is already clear and CE is pressed again, upgrade to a full 'C' wipe
                    if (_isNewInput && _currentValue == 0)
                    {
                        _previousValue = 0; _currentOperator = ""; _percentageString = "";
                        if (_operatorSymbol != null) _operatorSymbol.Opacity = 0;
                        TriggerOperationLabel("Cleared All");
                    }
                    else
                    {
                        _currentValue = 0; _isNewInput = true; _percentageString = "";
                        TriggerOperationLabel("Entry Cleared");
                    }
                }
                TriggerTextFlash(Color.FromRgb(255, 80, 80));

                // 2. The 5x Combo History Wipe Mechanic
                DateTime now = DateTime.Now;
                if ((now - _lastClearKeyPress).TotalSeconds <= 2)
                {
                    _clearKeyCount++;
                }
                else
                {
                    _clearKeyCount = 1; // Reset combo if too much time passed
                }
                _lastClearKeyPress = now;

                if (_clearKeyCount >= 5)
                {
                    _historyTape.Clear();
                    _clearKeyCount = 0; // Reset after successful wipe
                    SaveState();
                }

                UpdateDisplay();
            }
            else if (input == "MC") { _memoryValue = 0; SaveState(); TriggerTextFlash(Color.FromRgb(255, 180, 0)); UpdateDisplay(); }

            else if (input == "MR") { _currentValue = _memoryValue; _isNewInput = true; TriggerTextFlash(Color.FromRgb(255, 180, 0)); UpdateDisplay(); }
            else if (input == "M+") { _memoryValue += _currentValue; SaveState(); TriggerTextFlash(Color.FromRgb(255, 180, 0)); UpdateDisplay(); }
            else if (input == "M-") { _memoryValue -= _currentValue; SaveState(); TriggerTextFlash(Color.FromRgb(255, 180, 0)); UpdateDisplay(); }
            else if (input == "%")
            {
                // Capture raw user input before conversion
                _percentageString = $"{_currentValue}%";

                if (_currentOperator == "+" || _currentOperator == "-")
                {
                    _currentValue = _previousValue * (_currentValue / 100.0);
                }
                else
                {
                    _currentValue = _currentValue / 100.0;
                }
                _isNewInput = true;
                TriggerOperationLabel("Percentage");
                TriggerTextFlash(Color.FromRgb(255, 180, 0));
                UpdateDisplay();
            }
            else if (input == "=")
            {
                Calculate();
                _currentOperator = "";
                TriggerTextFlash(Color.FromRgb(50, 255, 100));
                TriggerSymbol("=", Colors.LimeGreen);
                UpdateDisplay();

                // If setting is disabled, trick the engine into appending the next digit
                if (!_clearAfterEquals)
                {
                    _isNewInput = false;
                }
            }
            else // Operators (+, -, ×, ÷)
            {
                if (!_isNewInput && !string.IsNullOrEmpty(_currentOperator)) Calculate();

                _previousValue = _currentValue;
                _currentOperator = input;
                _isNewInput = true;

                // UX: Fire the tiny operation label to confirm the memory updated if the user swaps operators
                string opName = "";
                switch (input)
                {
                    case "+": opName = "Addition"; break;
                    case "-": opName = "Subtraction"; break;
                    case "×": opName = "Multiplication"; break;
                    case "÷": opName = "Division"; break;
                }
                TriggerOperationLabel(opName);

                TriggerTextFlash(Color.FromRgb(255, 180, 0));
                TriggerSymbol(input, Colors.DeepSkyBlue);
                UpdateDisplay();
            }
        }

        private void Calculate()
        {
            if (string.IsNullOrEmpty(_currentOperator)) return;

            double result = 0;
            string opName = "";
            switch (_currentOperator)
            {
                case "+": result = _previousValue + _currentValue; opName = "Addition"; break;
                case "-": result = _previousValue - _currentValue; opName = "Subtraction"; break;
                case "×": result = _previousValue * _currentValue; opName = "Multiplication"; break;
                case "÷": result = _previousValue / _currentValue; opName = "Division"; break;
            }

            TriggerOperationLabel(opName);

            // Sanitize micro-artifacts from base-2 floating-point math (e.g. 0.1 + 0.2)
            result = Math.Round(result, 12);

            // Inject percentage notation into history if used, then clear it
            string rightSide = string.IsNullOrEmpty(_percentageString) ? _currentValue.ToString(CultureInfo.InvariantCulture) : $"{_percentageString} ({_currentValue.ToString(CultureInfo.InvariantCulture)})";
            string equation = $"{_previousValue} {_currentOperator} {rightSide} = {result}";

            _percentageString = ""; // Reset tracker

            _historyTape.Add(equation);
            if (_historyTape.Count > _maxHistory) _historyTape.RemoveAt(0);

            _currentValue = result;
            _isNewInput = true;
            SaveState();
        }

        private void UpdateDisplay()
        {
            if (_isNewInput)
            {
                string memIndicator = _memoryValue != 0 ? " [M]" : "";

                // UX: Graceful Error State instead of raw C# Infinity/NaN strings
                if (double.IsInfinity(_currentValue) || double.IsNaN(_currentValue))
                {
                    _mainDisplay.Text = "Error" + memIndicator;
                }
                else
                {
                    // UX: Intelligent Formatting (e.g., 54,547,009) while preserving exact decimals
                    _mainDisplay.Text = _currentValue.ToString("#,##0.##########", CultureInfo.InvariantCulture) + memIndicator;
                }
            }

            if (_liveExpressionText != null)
            {
                if (!string.IsNullOrEmpty(_currentOperator))
                {
                    // UX: Apply thousands separators to the Live Strip to match the main display
                    string leftSide = _previousValue.ToString(CultureInfo.InvariantCulture);
                    var leftParts = leftSide.Split('.');
                    if (long.TryParse(leftParts[0], out long lWhole))
                    {
                        leftSide = lWhole.ToString("#,##0", CultureInfo.InvariantCulture);
                        if (leftParts.Length > 1) leftSide += "." + leftParts[1];
                    }

                    string rightSide = "";
                    if (!_isNewInput || !string.IsNullOrEmpty(_percentageString))
                    {
                        if (!string.IsNullOrEmpty(_percentageString))
                        {
                            rightSide = _percentageString;
                        }
                        else
                        {
                            string rawRight = _currentValue.ToString(CultureInfo.InvariantCulture);
                            var rightParts = rawRight.Split('.');
                            if (long.TryParse(rightParts[0], out long rWhole))
                            {
                                rightSide = rWhole.ToString("#,##0", CultureInfo.InvariantCulture);
                                if (rightParts.Length > 1) rightSide += "." + rightParts[1];
                            }
                            else rightSide = rawRight;
                        }
                    }

                    _liveExpressionText.Text = $"{leftSide} {_currentOperator} {rightSide}".Trim();
                }
                else
                {
                    _liveExpressionText.Text = "";
                }
            }

            if (_historyListBox != null)
            {
                _historyListBox.ItemsSource = null;
                _historyListBox.ItemsSource = _historyTape;
                if (_historyTape.Count > 0) _historyListBox.ScrollIntoView(_historyTape.Last());
            }

            HighlightActiveOperator();
        }

        private void HighlightActiveOperator()
        {
            if (_virtualKeypad == null) return;

            foreach (var child in _virtualKeypad.Children)
            {
                if (child is Button btn)
                {
                    string text = btn.Content?.ToString();
                    if (text == "+" || text == "-" || text == "×" || text == "÷")
                    {
                        // UX: Glow the active operator so the user knows it is waiting for a number
                        if (text == _currentOperator && _isNewInput)
                        {
                            btn.Background = new SolidColorBrush(Color.FromArgb(80, 66, 133, 244));
                            btn.Foreground = Brushes.White;
                            btn.BorderBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244));
                        }
                        else // Reset inactive operators to default
                        {
                            btn.Background = new SolidColorBrush(Color.FromArgb(10, 255, 255, 255));
                            btn.Foreground = new SolidColorBrush(Color.FromRgb(66, 133, 244));
                            btn.BorderBrush = new SolidColorBrush(Color.FromArgb(15, 255, 255, 255));
                        }
                    }
                }
            }
        }

        // ==========================================
        // PERSISTENCE & BOILERPLATE
        // ==========================================
        private void SaveState()
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CalcState.dat");
                string historyStr = string.Join(";;", _historyTape);
                File.WriteAllText(path, $"{_currentValue.ToString(CultureInfo.InvariantCulture)}|{_memoryValue.ToString(CultureInfo.InvariantCulture)}|{historyStr}");
            }
            catch { }
        }

        private void LoadState()
        {
            try
            {
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CalcState.dat");
                if (File.Exists(path))
                {
                    string[] parts = File.ReadAllText(path).Split('|');
                    if (parts.Length > 0 && double.TryParse(parts[0], NumberStyles.Any, CultureInfo.InvariantCulture, out double cv)) _currentValue = cv;
                    if (parts.Length > 1 && double.TryParse(parts[1], NumberStyles.Any, CultureInfo.InvariantCulture, out double mv)) _memoryValue = mv;
                    if (parts.Length == 3 && !string.IsNullOrEmpty(parts[2]))
                        _historyTape = parts[2].Split(new[] { ";;" }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
            }
            catch { }
        }

        public void Pause() => UninstallHook();
        public void Resume() { }
        public void Cleanup() { UninstallHook(); SaveState(); }

        // ==========================================
        // UI: SETTINGS & BUTTON STYLES
        // ==========================================
        private void ApplySettingsData(Dictionary<string, object> settings)
        {
            if (settings != null)
            {
                if (settings.ContainsKey("ShowVirtualKeypad") && bool.TryParse(settings["ShowVirtualKeypad"].ToString(), out bool vk)) _showVirtualKeypad = vk;
                if (settings.ContainsKey("ShowHistoryTape") && bool.TryParse(settings["ShowHistoryTape"].ToString(), out bool ht)) _showHistoryTape = ht;
                if (settings.ContainsKey("DisplayColor")) _displayColor = settings["DisplayColor"].ToString();
                if (settings.ContainsKey("HistoryColor")) _historyColor = settings["HistoryColor"].ToString();
                if (settings.ContainsKey("SymbolFadeMs") && int.TryParse(settings["SymbolFadeMs"].ToString(), out int sf)) _symbolFadeMs = sf;
                if (settings.ContainsKey("ClearAfterEquals") && bool.TryParse(settings["ClearAfterEquals"].ToString(), out bool cae)) _clearAfterEquals = cae;
                if (settings.ContainsKey("ShowOperationLabel") && bool.TryParse(settings["ShowOperationLabel"].ToString(), out bool sol)) _showOperationLabel = sol;
            }
        }

        private void ApplyLayoutSettings()
        {
            if (_historyCard != null) _historyCard.Visibility = _showHistoryTape ? Visibility.Visible : Visibility.Collapsed;
            if (_keypadCard != null) _keypadCard.Visibility = _showVirtualKeypad ? Visibility.Visible : Visibility.Collapsed;
            // Completely rebuild the brush to instantly sever any WPF animation locks
            if (_mainDisplay != null)
            {
                Color baseColor = ParseColorSafe(_displayColor, Colors.White);
                _displayForegroundBrush = new SolidColorBrush(baseColor);
                _mainDisplay.Foreground = _displayForegroundBrush;

                // Dynamically tint the live expression color to stand out (shifts RGB slightly cooler/dimmer)
                Color liveColor = Color.FromRgb(
                    (byte)Math.Max(0, baseColor.R - 40),
                    (byte)Math.Max(0, baseColor.G - 20),
                    (byte)Math.Min(255, baseColor.B + 30)
                );

                // Fallback contrast safety if the display color is near-black
                if (baseColor.R < 50 && baseColor.G < 50 && baseColor.B < 50)
                {
                    liveColor = Color.FromRgb(100, 120, 150);
                }

                if (_liveExpressionText != null)
                {
                    _liveExpressionText.Foreground = new SolidColorBrush(liveColor);
                }
            }

            if (_historyListBox != null)
            {
                _historyListBox.Foreground = new SolidColorBrush(ParseColorSafe(_historyColor, Colors.Gray));
            }
        }

        private Button CreateButton(string text)
        {
            bool isOperator = text == "+" || text == "-" || text == "×" || text == "÷" || text == "=";
            bool isAction = text == "C" || text == "CE" || text == "⌫" || text == "%";
            bool isMemory = text.StartsWith("M");

            SolidColorBrush bgBrush = new SolidColorBrush(Color.FromArgb(10, 255, 255, 255));
            SolidColorBrush fgBrush = Brushes.White;

            if (isOperator) fgBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244));
            if (isAction) fgBrush = Brushes.LightGray;
            if (isMemory) fgBrush = Brushes.Orange;

            Button btn = new Button
            {
                Content = text,
                FontSize = isMemory || isAction ? 14 : 18,
                FontWeight = isOperator ? FontWeights.Bold : FontWeights.Normal,
                Foreground = fgBrush,
                Background = bgBrush,
                BorderThickness = new Thickness(1),
                BorderBrush = new SolidColorBrush(Color.FromArgb(15, 255, 255, 255)),
                Margin = new Thickness(2),
                Cursor = Cursors.Hand,
                Focusable = false
            };
            btn.Click += (s, e) => { ProcessInput(text); };
            return btn;
        }

        public void ShowSettingsWindow(Window ownerWindow, dynamic frameData)
        {
            UninstallHook();

            SolidColorBrush accentBrush;
            try { accentBrush = new SolidColorBrush(Utility.GetColorFromName(SettingsManager.SelectedColor)); }
            catch { accentBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244)); }

            Window win = new Window
            {
                Owner = ownerWindow,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowStyle = WindowStyle.None,
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                SizeToContent = SizeToContent.Height,
                Width = 370
            };

            Border headerBorder = new Border { Height = 50, Background = accentBrush };
            headerBorder.MouseLeftButtonDown += (s, e) => { if (e.ButtonState == MouseButtonState.Pressed) win.DragMove(); };
            Grid headerGrid = new Grid();
            headerGrid.Children.Add(new TextBlock { Text = "Calculator Settings", Foreground = Brushes.White, FontWeight = FontWeights.Bold, VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(15, 0, 0, 0) });
            Button btnClose = new Button { Content = "X", Foreground = Brushes.White, Background = Brushes.Transparent, BorderThickness = new Thickness(0), Width = 40, HorizontalAlignment = HorizontalAlignment.Right };
            btnClose.Click += (s, e) => win.Close();
            headerGrid.Children.Add(btnClose);
            headerBorder.Child = headerGrid;

            Border contentBorder = new Border { Background = Brushes.White, Padding = new Thickness(20) };
            StackPanel contentPanel = new StackPanel();

            CheckBox chkHistory = new CheckBox { Content = "Show Calculation History Tape", IsChecked = _showHistoryTape, Margin = new Thickness(0, 0, 0, 10), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkHistory);

            CheckBox chkKeypad = new CheckBox { Content = "Show Virtual Keypad", IsChecked = _showVirtualKeypad, Margin = new Thickness(0, 0, 0, 10), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkKeypad);

            CheckBox chkClear = new CheckBox { Content = "Clear Display on New Input (After '=')", IsChecked = _clearAfterEquals, Margin = new Thickness(0, 0, 0, 10), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkClear);

            CheckBox chkOpLabel = new CheckBox { Content = "Show Operation Names (e.g. Addition)", IsChecked = _showOperationLabel, Margin = new Thickness(0, 0, 0, 15), FontWeight = FontWeights.SemiBold };
            contentPanel.Children.Add(chkOpLabel);

            string[] availableColors = { "White", "Black", "LightGray", "DarkGray", "Gray", "Cyan", "LimeGreen", "Gold", "Orange", "DeepSkyBlue" };
            // Display Color
            StackPanel displayColorSp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
            displayColorSp.Children.Add(new TextBlock { Text = "Display Color:", Width = 150, VerticalAlignment = VerticalAlignment.Center });
            ComboBox cmbDisplayColor = new ComboBox { Width = 120 };
            foreach (var c in availableColors) cmbDisplayColor.Items.Add(c);
            cmbDisplayColor.SelectedItem = availableColors.Contains(_displayColor) ? _displayColor : "White";
            displayColorSp.Children.Add(cmbDisplayColor);
            contentPanel.Children.Add(displayColorSp);

            // History Color
            StackPanel historyColorSp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
            historyColorSp.Children.Add(new TextBlock { Text = "History Color:", Width = 150, VerticalAlignment = VerticalAlignment.Center });
            ComboBox cmbHistoryColor = new ComboBox { Width = 120 };
            foreach (var c in availableColors) cmbHistoryColor.Items.Add(c);
            cmbHistoryColor.SelectedItem = availableColors.Contains(_historyColor) ? _historyColor : "Gray";
            historyColorSp.Children.Add(cmbHistoryColor);
            contentPanel.Children.Add(historyColorSp);

            // Operator Badge Fade Time
            StackPanel fadeSp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 15) };
            fadeSp.Children.Add(new TextBlock { Text = "Badge Fade Time (ms):", Width = 150, VerticalAlignment = VerticalAlignment.Center });
            ComboBox cmbFade = new ComboBox { Width = 120 };
            cmbFade.Items.Add("500"); cmbFade.Items.Add("1000"); cmbFade.Items.Add("1500"); cmbFade.Items.Add("2500"); cmbFade.Items.Add("0 (Never Fade)");
            cmbFade.SelectedItem = _symbolFadeMs == 0 ? "0 (Never Fade)" : _symbolFadeMs.ToString();
            fadeSp.Children.Add(cmbFade);
            contentPanel.Children.Add(fadeSp);

            Button btnClearMem = new Button { Content = "Clear Memory & History", Padding = new Thickness(5), Margin = new Thickness(0, 10, 0, 0) };
            btnClearMem.Click += (s, e) => { _memoryValue = 0; _historyTape.Clear(); SaveState(); UpdateDisplay(); MessageBox.Show("Cleared.", "Calculator", MessageBoxButton.OK, MessageBoxImage.Information); };
            contentPanel.Children.Add(btnClearMem);

            contentBorder.Child = contentPanel;

            Border footerBorder = new Border { Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)), BorderThickness = new Thickness(0, 1, 0, 0), BorderBrush = Brushes.LightGray, Padding = new Thickness(15) };
            StackPanel footerSp = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            Button btnCancel = new Button { Content = "Cancel", Background = Brushes.White, BorderBrush = Brushes.Gray, Width = 80, Height = 30, Margin = new Thickness(0, 0, 10, 0) };
            btnCancel.Click += (s, e) => win.Close();
            Button btnSave = new Button { Content = "Save", Background = accentBrush, Foreground = Brushes.White, FontWeight = FontWeights.Bold, BorderThickness = new Thickness(0), Width = 80, Height = 30 };
            btnSave.Click += (s, e) =>
            {
                int newFadeMs = 1500;
                if (cmbFade.SelectedItem != null)
                {
                    string selectedFade = cmbFade.SelectedItem.ToString();
                    if (selectedFade == "0 (Never Fade)") newFadeMs = 0;
                    else int.TryParse(selectedFade, out newFadeMs);
                }

                Dictionary<string, object> newSettings = new Dictionary<string, object> {
                    { "ShowVirtualKeypad", chkKeypad.IsChecked == true },
             { "ShowHistoryTape", chkHistory.IsChecked == true },
                    { "ClearAfterEquals", chkClear.IsChecked == true },
                    { "ShowOperationLabel", chkOpLabel.IsChecked == true },
                    { "DisplayColor", cmbDisplayColor.SelectedItem?.ToString() ?? "White" },
                    { "HistoryColor", cmbHistoryColor.SelectedItem?.ToString() ?? "Gray" },
                    { "SymbolFadeMs", newFadeMs }
                };
                if (frameData is Newtonsoft.Json.Linq.JObject jFrame) jFrame["PluginSettings"] = Newtonsoft.Json.Linq.JObject.FromObject(newSettings);
                else ((IDictionary<string, object>)frameData)["PluginSettings"] = newSettings;
                try { FrameDataManager.SaveFrameData(); } catch { }
                ApplySettingsData(newSettings);
                ApplyLayoutSettings();
                UpdateDisplay();
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
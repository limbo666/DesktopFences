using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Desktop_Fences;

namespace Desktop_Fences
{
    public class FenceFocusFormManager : Window
    {
        // --- Win32 API for Z-Order Management ---
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        // ----------------------------------------

        private TextBox searchBox;
        private ListBox fenceListBox;
        private List<FenceItem> allFences;

        // Helper class to store fence data

        // Helper class to store fence data
        private class FenceItem
        {
            public string Title { get; set; }
            public NonActivatingWindow WindowRef { get; set; }
            public override string ToString() => Title;
        }

        public FenceFocusFormManager()
        {
            InitializeModernComponent();
            LoadActiveFences();
        }

        private void InitializeModernComponent()
        {
            Title = "Focus Fence";
            Width = 450;
            Height = 550;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
            WindowStyle = WindowStyle.None;
            AllowsTransparency = false;
            Topmost = true; // Ensure the selector itself opens on top

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
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Footer

            CreateModernHeader(rootGrid);
            CreateModernContent(rootGrid);
            CreateModernFooter(rootGrid);

            mainBorder.Child = rootGrid;
            mainContainer.Children.Add(mainBorder);
            Content = mainContainer;

            this.KeyDown += Window_KeyDown;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                FocusSelectedFence();
                e.Handled = true;
            }
            else if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
                e.Handled = true;
            }
            else if (e.Key == Key.Down && searchBox.IsFocused)
            {
                fenceListBox.Focus();
                if (fenceListBox.Items.Count > 0 && fenceListBox.SelectedIndex == -1)
                    fenceListBox.SelectedIndex = 0;
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

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            TextBlock titleText = new TextBlock
            {
                Text = "Focus Fence",
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
            closeButton.Click += (s, e) => { DialogResult = false; Close(); };

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
                Padding = new Thickness(20, 10, 20, 10)
            };

            StackPanel contentPanel = new StackPanel { Orientation = Orientation.Vertical };

            // Field Section matching EditShortcutWindow
            Border fieldBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(251, 252, 253)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(12)
            };

            Grid fieldGrid = new Grid();
            fieldGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Search Label
            fieldGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Search Box
            fieldGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // ListBox

            TextBlock searchLabel = new TextBlock
            {
                Text = "Search active fences:",
                FontSize = 12,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                Margin = new Thickness(0, 0, 0, 8)
            };

            searchBox = new TextBox
            {
                FontSize = 13,
                Padding = new Thickness(8, 6, 8, 6),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                Margin = new Thickness(0, 0, 0, 12)
            };
            searchBox.TextChanged += SearchBox_TextChanged;

            fenceListBox = new ListBox
            {
                Height = 280, // Fixed height for scrollable area
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Background = Brushes.White,
                FontSize = 13,
                Padding = new Thickness(4)
            };
            fenceListBox.MouseDoubleClick += (s, e) => FocusSelectedFence();

            fieldGrid.Children.Add(searchLabel);
            Grid.SetRow(searchLabel, 0);

            fieldGrid.Children.Add(searchBox);
            Grid.SetRow(searchBox, 1);

            fieldGrid.Children.Add(fenceListBox);
            Grid.SetRow(fenceListBox, 2);

            fieldBorder.Child = fieldGrid;
            contentPanel.Children.Add(fieldBorder);

            contentBorder.Child = contentPanel;
            Grid.SetRow(contentBorder, 1);
            rootGrid.Children.Add(contentBorder);

            // Focus search box on load
            this.Loaded += (s, e) => searchBox.Focus();
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

            Button cancelButton = new Button
            {
                Content = "Cancel",
                Height = 36,
                MinWidth = 80,
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Cursor = Cursors.Hand,
                Background = Brushes.White,
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(16, 0, 16, 0),
                Margin = new Thickness(0, 0, 10, 0)
            };
            cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

            Button focusButton = new Button
            {
                Content = "Focus Fence",
                Height = 36,
                MinWidth = 80,
                FontSize = 13,
                FontWeight = FontWeights.Bold,
                Cursor = Cursors.Hand,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(16, 0, 16, 0),
                Background = GetAccentColorBrush(),
                Foreground = Brushes.White
            };
            focusButton.Click += (s, e) => FocusSelectedFence();

            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(focusButton);
            footerBorder.Child = buttonPanel;

            Grid.SetRow(footerBorder, 2);
            rootGrid.Children.Add(footerBorder);
        }

        private void LoadActiveFences()
        {
            allFences = new List<FenceItem>();

            // Find all active fence windows
            var fenceWindows = Application.Current.Windows.OfType<NonActivatingWindow>().ToList();

            foreach (var win in fenceWindows)
            {
                string title = win.Title;
                if (string.IsNullOrWhiteSpace(title)) title = "Unnamed Fence";

                allFences.Add(new FenceItem { Title = title, WindowRef = win });
            }

            // Alphabetical sort
            allFences = allFences.OrderBy(f => f.Title).ToList();
            UpdateListBox(allFences);
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (allFences == null) return;

            string query = searchBox.Text.ToLower();
            var filtered = allFences.Where(f => f.Title.ToLower().Contains(query)).ToList();
            UpdateListBox(filtered);
        }

        private void UpdateListBox(List<FenceItem> items)
        {
            fenceListBox.Items.Clear();
            foreach (var item in items)
            {
                fenceListBox.Items.Add(item);
            }
            if (fenceListBox.Items.Count > 0)
                fenceListBox.SelectedIndex = 0;
        }

        private void FocusSelectedFence()
        {
            if (fenceListBox.SelectedItem is FenceItem selected)
            {
                NonActivatingWindow targetWin = selected.WindowRef;

                // 1. Close the form FIRST to yield the foreground status back to other apps
                DialogResult = true;
                Close();

                if (targetWin != null)
                {
                    Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        // 1. Yield foreground focus back to the previous app
                        await System.Threading.Tasks.Task.Delay(100);

                        // 2. Wake up fences quietly in the background
                        if (FenceManager._areFencesAutoHidden)
                        {
                            FenceManager.WakeUpFences();
                            await System.Threading.Tasks.Task.Delay(50);
                        }

                        // 3. Grab the target window
                        targetWin.Topmost = true;

                        // 4. THE TRIGGER: This activates the fence, causing Windows to pull ALL fences up.
                        targetWin.Activate();
                        targetWin.Focus();

                        // 5. THE BRUTAL FIX: Let the OS finish its group-pull, then slam the others down.
                        await System.Threading.Tasks.Task.Delay(20);

                        var otherFences = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().Where(w => w != targetWin).ToList();
                        foreach (var f in otherFences)
                        {
                            var hwnd = new System.Windows.Interop.WindowInteropHelper(f).Handle;
                            SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                        }

                        EventHandler deactivatedHandler = null;
                        deactivatedHandler = (s, args) =>
                        {
                            targetWin.Topmost = false;
                            targetWin.Deactivated -= deactivatedHandler;
                            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Fence '{selected.Title}' released from Topmost.");
                        };

                        targetWin.Deactivated += deactivatedHandler;
                    });
                }
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
                return new SolidColorBrush(Color.FromRgb(66, 133, 244));
            }
        }
    }
}
using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using System.Windows.Input;
using Desktop_Fences;
using System.Linq;

namespace Desktop_Fences
{
    public class IconPickerDialog : Window
    {
        private Button _okButton;

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern uint ExtractIconEx(string szFileName, int nIconIndex,
                                       IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

        [DllImport("user32.dll")]
        static extern bool DestroyIcon(IntPtr hIcon);

        public int SelectedIndex { get; private set; } = -1;
        private readonly string _filePath;
        private WrapPanel _iconPanel;

        public IconPickerDialog(string filePath)
        {
            _filePath = filePath;
            InitializeWindow();
            LoadIcons();
        }

        private void InitializeWindow()
        {
            Title = "Select Icon";
            Width = 400;
            Height = 300;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = new SolidColorBrush(Color.FromRgb(248, 249, 250));
            WindowStyle = WindowStyle.None;
            AllowsTransparency = false;



            Border mainBorder = new Border
            {
                Background = Brushes.White,
                CornerRadius = new CornerRadius(0), // No rounded corners like other forms
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(1, 1, 1, 1),
                Margin = new Thickness(0, 0, 0, 0)
            };

            Grid rootGrid = new Grid();
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            rootGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            CreateHeader(rootGrid);
            CreateContent(rootGrid);
            CreateFooter(rootGrid);

            mainBorder.Child = rootGrid;
            Content = mainBorder;
        }

        private void CreateHeader(Grid rootGrid)
        {
            // Header with accent color background (same as CustomizeFenceForm)
            Border headerBorder = new Border
            {
                Height = 50,
                Background = GetAccentColorBrush(), // Use user's selected theme color
                CornerRadius = new CornerRadius(0) // No rounded corners
            };

            Grid headerGrid = new Grid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // Title with white text (same as CustomizeFenceForm)
            TextBlock titleText = new TextBlock
            {
                Text = "Select Icon",
                FontSize = 14,
                FontWeight = FontWeights.Bold, // Bold like CustomizeFenceForm
                Foreground = Brushes.White, // White text on colored background
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(16, 0, 0, 0)
            };

            // Close button with hover effects (same as CustomizeFenceForm)
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

            // Add hover effect like CustomizeFenceForm
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
        private void CreateContent(Grid rootGrid)
        {
            Border contentBorder = new Border
            {
                Margin = new Thickness(20, 10, 20, 10)
            };

            ScrollViewer scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };

            _iconPanel = new WrapPanel();
            scrollViewer.Content = _iconPanel;
            contentBorder.Child = scrollViewer;

            Grid.SetRow(contentBorder, 1);
            rootGrid.Children.Add(contentBorder);
        }

        private void CreateFooter(Grid rootGrid)
        {
            Border footerBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                BorderThickness = new Thickness(0, 1, 0, 0), // Only top border
                Padding = new Thickness(20, 16, 20, 16),
                CornerRadius = new CornerRadius(0) // No rounded corners
            };

            // Use Grid to have text on left and buttons on right
            Grid footerGrid = new Grid();
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) }); // Text area
            footerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto }); // Buttons area

            // Instruction text on the left
            TextBlock instructionText = new TextBlock
            {
                Text = "Select an Icon from the available and click OK",
                FontSize = 13,
                FontWeight = FontWeights.Medium,
                Foreground = new SolidColorBrush(Color.FromRgb(95, 99, 104)),
                VerticalAlignment = VerticalAlignment.Center,
                TextWrapping = TextWrapping.Wrap,
                Height = 36, // Same height as buttons
                Background = new SolidColorBrush(Color.FromRgb(248, 249, 250)), // Same as footer background
                Margin = new Thickness(0, 0, 20, 0) // Space between text and buttons
            };

            // Buttons panel on the right
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
                BorderThickness = new Thickness(1, 1, 1, 1),
                Padding = new Thickness(16, 0, 16, 0),
                Background = Brushes.White,
                Foreground = new SolidColorBrush(Color.FromRgb(32, 33, 36)),
                BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                Margin = new Thickness(0, 0, 10, 0)
            };
            cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

            // OK button with accent color (same as Save button in other forms)
            _okButton = new Button
            {
                Content = "OK",
                Height = 36,
                MinWidth = 80,
                FontSize = 13,
                FontWeight = FontWeights.Bold, // Bold like other forms
                Cursor = Cursors.Hand,
                BorderThickness = new Thickness(0), // No border like other forms
                Padding = new Thickness(16, 0, 16, 0),
                Background = GetAccentColorBrush(), // Use user's selected theme color
                Foreground = Brushes.White,
                IsEnabled = false
            };
            _okButton.Click += (s, e) =>
            {
                if (SelectedIndex >= 0)
                {
                    DialogResult = true;
                    Close();
                }
            };
            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(_okButton);

            // Add text and buttons to grid
            footerGrid.Children.Add(instructionText);
            footerGrid.Children.Add(buttonPanel);
            Grid.SetColumn(instructionText, 0);
            Grid.SetColumn(buttonPanel, 1);

            footerBorder.Child = footerGrid;
            Grid.SetRow(footerBorder, 2);
            rootGrid.Children.Add(footerBorder);
        }

        private void LoadIcons()
        {
            if (string.IsNullOrEmpty(_filePath)) return;

            var iconCount = (int)ExtractIconEx(_filePath, -1, null, null, 0);
            if (iconCount == 0) return;

            for (int i = 0; i < iconCount; i++)
            {
                IntPtr[] hIcon = new IntPtr[1];
                if (ExtractIconEx(_filePath, i, hIcon, null, 1) <= 0) continue;

                if (hIcon[0] != IntPtr.Zero)
                {
                    try
                    {
                        var bitmap = Imaging.CreateBitmapSourceFromHIcon(
                            hIcon[0],
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions()
                        );

                        Button iconButton = new Button
                        {
                            Content = new Image
                            {
                                Source = bitmap,
                                Width = 32,
                                Height = 32
                            },
                            Tag = i,
                            Margin = new Thickness(5, 5, 5, 5),
                            Width = 48,
                            Height = 48,
                            Background = Brushes.Transparent,
                            BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224)),
                            BorderThickness = new Thickness(1, 1, 1, 1)
                        };

                        iconButton.Click += (s, e) =>
                        {
                            // Clear selection from all buttons
                            foreach (Button btn in _iconPanel.Children.OfType<Button>())
                            {
                                btn.Background = Brushes.Transparent;
                                btn.BorderBrush = new SolidColorBrush(Color.FromRgb(218, 220, 224));
                                btn.BorderThickness = new Thickness(1, 1, 1, 1);
                            }

                            // Highlight selected button
                            iconButton.Background = new SolidColorBrush(Color.FromRgb(232, 240, 254));
                            iconButton.BorderBrush = new SolidColorBrush(Color.FromRgb(66, 133, 244));
                            iconButton.BorderThickness = new Thickness(2, 2, 2, 2);

                            SelectedIndex = (int)iconButton.Tag;
                            _okButton.IsEnabled = true;
                        };

                        _iconPanel.Children.Add(iconButton);
                    }
                    finally
                    {
                        DestroyIcon(hIcon[0]);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the user's selected accent color (same method as CustomizeFenceForm)
        /// </summary>
        private SolidColorBrush GetAccentColorBrush()
        {
            try
            {
                string selectedColorName = SettingsManager.SelectedColor;
                var mediaColor = Utility.GetColorFromName(selectedColorName);
                return new SolidColorBrush(mediaColor);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error getting accent color: {ex.Message}");
                // Fallback to blue
                return new SolidColorBrush(Color.FromRgb(66, 133, 244));
            }
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);
            if (e.ButtonState == MouseButtonState.Pressed)
            {
                DragMove();
            }
        }
    }
}
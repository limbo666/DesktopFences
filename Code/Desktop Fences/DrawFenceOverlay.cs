using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Runtime.InteropServices; // Required for GetCursorPos

namespace Desktop_Fences
{
    public class DrawFenceOverlay : Window
    {
        private Point _startPoint;
        private Rectangle _selectionRect;
        private Canvas _canvas;
        private bool _isDragging = false;
        private Border _hintBox;

        // --- NATIVE HELPER FOR ROBUST MOUSE DETECTION ---
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetCursorPos(ref Win32Point pt);

        [StructLayout(LayoutKind.Sequential)]
        internal struct Win32Point { public Int32 X; public Int32 Y; }
        // ------------------------------------------------

        public DrawFenceOverlay()
        {
            // 1. VIRTUAL SCREEN SETUP (Multi-Monitor Support)
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = new SolidColorBrush(Color.FromArgb(1, 0, 0, 0)); // 1/255 opacity
            ShowInTaskbar = false;
            Topmost = true;
            Cursor = Cursors.Cross;

            // Explicitly cover the entire virtual desktop area
            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Left = SystemParameters.VirtualScreenLeft;
            this.Top = SystemParameters.VirtualScreenTop;
            this.Width = SystemParameters.VirtualScreenWidth;
            this.Height = SystemParameters.VirtualScreenHeight;

            _canvas = new Canvas();
            Content = _canvas;

            // 2. HINT MESSAGE
            //_hintBox = new Border
            //{
            //    Background = new SolidColorBrush(Color.FromArgb(200, 0, 0, 0)),
            //    CornerRadius = new CornerRadius(5),
            //    Padding = new Thickness(10),
            //    Child = new TextBlock
            //    {
            //        Text = "Draw a box to create a Fence (Esc to cancel)",
            //        Foreground = Brushes.White,
            //        FontWeight = FontWeights.Bold
            //    },
            //    IsHitTestVisible = false // Pass clicks through
            //};

            // --- REPLACE WITH THIS ---
            // Use the centralized design factory (assuming you put it in Utility)
            // If you put it in FenceManager, change 'Utility' to 'FenceManager'
            _hintBox = MessageBoxesManager.CreateUnifiedMessage("Draw a box to create a Fence (Esc to cancel)");

            _canvas.Children.Add(_hintBox);

          //  _canvas.Children.Add(_hintBox);

            _selectionRect = new Rectangle
            {
                Stroke = Brushes.DeepSkyBlue,
                StrokeThickness = 2,
                Fill = new SolidColorBrush(Color.FromArgb(50, 0, 191, 255)),
                Visibility = Visibility.Collapsed
            };
            _canvas.Children.Add(_selectionRect);

            // 3. ROBUST POSITIONING (Using Windows API)
            this.Loaded += (s, e) =>
            {
                try
                {
                    // Get absolute screen coordinates (works even if WPF hasn't captured mouse yet)
                    Win32Point w32Mouse = new Win32Point();
                    GetCursorPos(ref w32Mouse);

                    // Convert to local Window coordinates
                    // (Global Mouse X) - (Window Left) = Local X
                    double relativeX = w32Mouse.X - this.Left;
                    double relativeY = w32Mouse.Y - this.Top;

                    // Position hint 20px down-right from cursor
                    Canvas.SetLeft(_hintBox, relativeX + 20);
                    Canvas.SetTop(_hintBox, relativeY + 20);
                }
                catch
                {
                    // Fallback to center if something insane happens
                    Canvas.SetLeft(_hintBox, (this.Width - _hintBox.ActualWidth) / 2);
                    Canvas.SetTop(_hintBox, 100);
                }
            };

            this.MouseDown += OnMouseDown;
            this.MouseMove += OnMouseMove;
            this.MouseUp += OnMouseUp;
            this.KeyDown += (s, e) => { if (e.Key == Key.Escape) Close(); };
        }

        private void OnMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                _startPoint = e.GetPosition(_canvas);
                _isDragging = true;

                Canvas.SetLeft(_selectionRect, _startPoint.X);
                Canvas.SetTop(_selectionRect, _startPoint.Y);
                _selectionRect.Width = 0;
                _selectionRect.Height = 0;
                _selectionRect.Visibility = Visibility.Visible;

                _hintBox.Visibility = Visibility.Collapsed;
            }
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                var pos = e.GetPosition(_canvas);
                var x = Math.Min(pos.X, _startPoint.X);
                var y = Math.Min(pos.Y, _startPoint.Y);
                var w = Math.Abs(pos.X - _startPoint.X);
                var h = Math.Abs(pos.Y - _startPoint.Y);

                Canvas.SetLeft(_selectionRect, x);
                Canvas.SetTop(_selectionRect, y);
                _selectionRect.Width = w;
                _selectionRect.Height = h;
            }
        }

        private void OnMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDragging)
            {
                _isDragging = false;

                if (_selectionRect.Width > 50 && _selectionRect.Height > 50)
                {
                    // 4. COORDINATE MAPPING (Global Fix)
                    double globalX = Canvas.GetLeft(_selectionRect) + this.Left;
                    double globalY = Canvas.GetTop(_selectionRect) + this.Top;

                    Rect finalRect = new Rect(globalX, globalY, _selectionRect.Width, _selectionRect.Height);

                    this.Close();

                    FenceManager.CreateFenceFromDraw(finalRect);
                }
                else
                {
                    _selectionRect.Visibility = Visibility.Collapsed;
                    _hintBox.Visibility = Visibility.Visible;

                    // Re-position hint to current mouse up location
                    Point mousePos = e.GetPosition(_canvas);
                    Canvas.SetLeft(_hintBox, mousePos.X + 20);
                    Canvas.SetTop(_hintBox, mousePos.Y + 20);
                }
            }
        }
    }
}
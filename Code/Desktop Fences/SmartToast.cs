using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;

namespace Desktop_Fences
{
    public class SmartToast : Window
    {
        private DispatcherTimer _timer;

        public static void Show(string title, string message)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                new SmartToast(title, message).Show();
            });
        }

        private SmartToast(string title, string message)
        {
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            Topmost = true;
            ShowInTaskbar = false;
            Width = 320;
            Height = 80;

            // Position at bottom right of primary screen
            var workArea = SystemParameters.WorkArea;
            Left = workArea.Right - Width - 20;
            Top = workArea.Bottom - Height - 20;

            Border mainBorder = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(41, 74, 122)), // Smart Desktop Navy Theme
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(15, 10, 15, 10),
                Effect = new System.Windows.Media.Effects.DropShadowEffect { Color = Colors.Black, BlurRadius = 10, Opacity = 0.3, ShadowDepth = 2 }
            };

            StackPanel sp = new StackPanel { Orientation = Orientation.Vertical, VerticalAlignment = VerticalAlignment.Center };
            sp.Children.Add(new TextBlock { Text = title, FontWeight = FontWeights.Bold, Foreground = Brushes.White, FontSize = 14, Margin = new Thickness(0, 0, 0, 5) });
            sp.Children.Add(new TextBlock { Text = message, Foreground = new SolidColorBrush(Color.FromRgb(220, 230, 240)), FontSize = 12, TextTrimming = TextTrimming.CharacterEllipsis });

            mainBorder.Child = sp;
            Content = mainBorder;

            // Fade out and close after 3.5 seconds
            _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3.5) };
            _timer.Tick += (s, e) =>
            {
                _timer.Stop();
                var anim = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(500));
                anim.Completed += (s2, e2) => Close();
                BeginAnimation(OpacityProperty, anim);
            };
            _timer.Start();
        }
    }
}
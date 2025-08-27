using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices; // Added for DeleteObject
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms.VisualStyles;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging; // Added for ImageSource
using IWshRuntimeLibrary;

namespace Desktop_Fences
{
    public static class Utility
    {

        public static Effect CreateIconEffect(IconVisibilityEffect effectType, string fenceColor = null)
        {
            switch (effectType)
            {
                case IconVisibilityEffect.Glow:
                    return new DropShadowEffect
                    {
                        Color = Colors.White,
                        Direction = 0,
                        ShadowDepth = 0,
                        BlurRadius = 8,
                        Opacity = 0.6
                    };

                case IconVisibilityEffect.Shadow:
                    return new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 315,
                        ShadowDepth = 3,
                        BlurRadius = 5,
                        Opacity = 0.7
                    };

                case IconVisibilityEffect.Outline:
                    return new DropShadowEffect
                    {
                        Color = Colors.White,
                        Direction = 0,
                        ShadowDepth = 0,
                        BlurRadius = 2,
                        Opacity = 0.8
                    };

                case IconVisibilityEffect.StrongShadow:
                    return new DropShadowEffect
                    {
                        Color = Colors.Black,
                        Direction = 315,
                        ShadowDepth = 5,
                        BlurRadius = 10,
                        Opacity = 0.9
                    };

                case IconVisibilityEffect.ColoredGlow:
                    var glowColor = string.IsNullOrEmpty(fenceColor)
                        ? Colors.White
                        : GetColorFromName(fenceColor);
                    return new DropShadowEffect
                    {
                        Color = glowColor,
                        Direction = 0,
                        ShadowDepth = 0,
                        BlurRadius = 6,
                        Opacity = 0.5
                    };

                //case IconVisibilityEffect.AngelGlow:
                //    // This would require a different approach with Border/Ellipse
                //    //  return null; // Handle separately if needed
                //    // Brighten effect - makes icons appear brighter/more visible
                //    return new System.Windows.Media.Effects.DropShadowEffect
                //    {
                //        Color = System.Windows.Media.Colors.Blue,
                //        Direction = 0,
                //        ShadowDepth = 0,
                //        BlurRadius = 36,
                //        Opacity = 1.0
                //    };
                case Desktop_Fences.IconVisibilityEffect.AngelGlow:
                    return new System.Windows.Media.Effects.DropShadowEffect
                    {
                        Color = System.Windows.Media.Colors.Snow,
                        Direction = 0,
                        ShadowDepth = 0,
                        BlurRadius = 5,
                        Opacity = 1.0
                    };



                case IconVisibilityEffect.None:
                default:
                    return null;
            }
        }

        // TABS FEATURE: Generate tab color scheme based on fence color
        public static (Color activeTab, Color inactiveTab, Color hoverTab, Color borderColor) GenerateTabColorScheme(string fenceColorName)
        {
            var baseColor = GetColorFromName(fenceColorName);

            // Convert to HSV for better color manipulation
            var hsv = RgbToHsv(baseColor);

            Color activeTab, inactiveTab, hoverTab, borderColor;

            // Handle special cases for very light or very dark colors
            if (hsv.Value < 0.3) // Very dark colors
            {
                // For dark colors, brighten significantly for active tab
                activeTab = HsvToRgb(hsv.Hue, hsv.Saturation * 0.8, Math.Min(hsv.Value + 0.4, 1.0));
                inactiveTab = HsvToRgb(hsv.Hue, hsv.Saturation * 0.3, Math.Min(hsv.Value + 0.6, 0.9));
                hoverTab = HsvToRgb(hsv.Hue, hsv.Saturation * 0.5, Math.Min(hsv.Value + 0.5, 0.95));
                borderColor = HsvToRgb(hsv.Hue, hsv.Saturation * 0.6, Math.Min(hsv.Value + 0.3, 0.8));
            }
            else if (hsv.Value > 0.8 && hsv.Saturation < 0.3) // Very light colors (like White, Beige)
            {
                // For light colors, use deeper variations
                activeTab = HsvToRgb(hsv.Hue, Math.Min(hsv.Saturation + 0.3, 0.6), hsv.Value * 0.7);
                inactiveTab = HsvToRgb(hsv.Hue, Math.Min(hsv.Saturation + 0.1, 0.2), hsv.Value * 0.95);
                hoverTab = HsvToRgb(hsv.Hue, Math.Min(hsv.Saturation + 0.2, 0.4), hsv.Value * 0.8);
                borderColor = HsvToRgb(hsv.Hue, Math.Min(hsv.Saturation + 0.4, 0.7), hsv.Value * 0.6);
            }
            else // Normal colors
            {
                // Active tab: slightly brighter and more saturated
                activeTab = HsvToRgb(hsv.Hue, Math.Min(hsv.Saturation + 0.1, 1.0), Math.Min(hsv.Value + 0.2, 1.0));

                // Inactive tab: much lighter and less saturated
                inactiveTab = HsvToRgb(hsv.Hue, hsv.Saturation * 0.3, Math.Min(hsv.Value + 0.4, 0.95));

                // Hover tab: between active and inactive
                hoverTab = HsvToRgb(hsv.Hue, hsv.Saturation * 0.7, Math.Min(hsv.Value + 0.3, 0.9));

                // Border: slightly darker than active
                borderColor = HsvToRgb(hsv.Hue, Math.Min(hsv.Saturation + 0.2, 1.0), hsv.Value * 0.8);
            }

            return (activeTab, inactiveTab, hoverTab, borderColor);
        }

        // Helper method to convert RGB to HSV
        private static (double Hue, double Saturation, double Value) RgbToHsv(Color color)
        {
            double r = color.R / 255.0;
            double g = color.G / 255.0;
            double b = color.B / 255.0;

            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            double delta = max - min;

            double hue = 0;
            if (delta != 0)
            {
                if (max == r) hue = ((g - b) / delta) % 6;
                else if (max == g) hue = (b - r) / delta + 2;
                else hue = (r - g) / delta + 4;
                hue *= 60;
                if (hue < 0) hue += 360;
            }

            double saturation = max == 0 ? 0 : delta / max;
            double value = max;

            return (hue, saturation, value);
        }

        // Helper method to convert HSV to RGB
        private static Color HsvToRgb(double hue, double saturation, double value)
        {
            double c = value * saturation;
            double x = c * (1 - Math.Abs((hue / 60) % 2 - 1));
            double m = value - c;

            double r, g, b;

            if (hue >= 0 && hue < 60) { r = c; g = x; b = 0; }
            else if (hue >= 60 && hue < 120) { r = x; g = c; b = 0; }
            else if (hue >= 120 && hue < 180) { r = 0; g = c; b = x; }
            else if (hue >= 180 && hue < 240) { r = 0; g = x; b = c; }
            else if (hue >= 240 && hue < 300) { r = x; g = 0; b = c; }
            else { r = c; g = 0; b = x; }

            return Color.FromRgb(
                (byte)Math.Round((r + m) * 255),
                (byte)Math.Round((g + m) * 255),
                (byte)Math.Round((b + m) * 255)
            );
        }

        public static void ApplyTintAndColorToFence(Window fence, string colorName = null)
        {
            var fenceControl = fence.Content as Border; // Matches your structure
            if (fenceControl == null) return;

            string effectiveColor = colorName ?? SettingsManager.SelectedColor; // Fallback to global

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.Settings,
                $"ApplyTintAndColorToFence: TintValue={SettingsManager.TintValue}, Color={effectiveColor}, Opacity={SettingsManager.TintValue / 100.0}");

            // Use TintValue from SettingsManager to determine if tint is applied
            fenceControl.Background = SettingsManager.TintValue > 0
                ? new SolidColorBrush(GetColorFromName(effectiveColor)) { Opacity = SettingsManager.TintValue / 100.0 }
                : Brushes.Transparent;

            // TABS FEATURE: Refresh tab colors when fence color changes
            if (fence is NonActivatingWindow fenceWindow)
            {
                FenceManager.RefreshTabColors(fenceWindow, effectiveColor);
            }
        }


        public static Color GetColorFromName(string colorName)
        {
            return colorName switch
            {

                "Red" => (Color)ColorConverter.ConvertFromString("#9E052E"),
                "Green" => (Color)ColorConverter.ConvertFromString("#06491A"),
                "Teal" => (Color)ColorConverter.ConvertFromString("#008080"),
                "Blue" => (Color)ColorConverter.ConvertFromString("#012162"),
                "Bismark" => (Color)ColorConverter.ConvertFromString("#49697E"),
                "White" => (Color)ColorConverter.ConvertFromString("#F1F1F6"),
                "Beige" => (Color)ColorConverter.ConvertFromString("#C8AD7E"),
                "Gray" => (Color)ColorConverter.ConvertFromString("#6E6E6E"),
                "Black" => (Color)ColorConverter.ConvertFromString("#0b0b0c"),
                "Purple" => (Color)ColorConverter.ConvertFromString("#3a0b50"),
                "Fuchsia" => (Color)ColorConverter.ConvertFromString("#5F093d"),
                "Yellow" => (Color)ColorConverter.ConvertFromString("#C1C708"),
                "Orange" => (Color)ColorConverter.ConvertFromString("#B75433"),
                _ => Colors.Transparent,
                // "Red" => (Color)ColorConverter.ConvertFromString("#c10338"),
                // "Green" => (Color)ColorConverter.ConvertFromString("#005618"),
                // "Blue" => (Color)ColorConverter.ConvertFromString("#012162"),
                //  "White" => (Color)ColorConverter.ConvertFromString("#fdfdff"),
                //  "Gray" => (Color)ColorConverter.ConvertFromString("#3d3d3f"),
                //  "Black" => (Color)ColorConverter.ConvertFromString("#0b0b0c"),
                // "Purple" => (Color)ColorConverter.ConvertFromString("#3a0b50"),
                //  "Yellow" => (Color)ColorConverter.ConvertFromString("#d8da1f"),
                //  _ => Colors.Transparent, 
            };
        }

        public static bool IsExecutableFile(string filePath)
        {
            string[] executableExtensions = { ".exe", ".bat", ".cmd", ".vbs", ".ps1", ".hta", ".msi" };
            if (Path.GetExtension(filePath).ToLower() == ".lnk")
            {
                try
                {
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                    filePath = shortcut.TargetPath;
                }
                catch
                {
                    return false;
                }
            }
            string extension = Path.GetExtension(filePath).ToLower();
            return executableExtensions.Contains(extension);
        }

        public static string GetShortcutTarget(string filePath)
        {
            if (Path.GetExtension(filePath).ToLower() == ".lnk")
            {
                try
                {
                    WshShell shell = new WshShell();
                    IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(filePath);
                    return shortcut.TargetPath;
                }
                catch
                {
                    return null;
                }
            }
            return filePath;
        }

        public static System.Drawing.Image LoadImageFromResources(string resourcePath)
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourcePath))
            {
                if (stream == null)
                {
                    throw new FileNotFoundException($"Resource '{resourcePath}' not found.");
                }
                return System.Drawing.Image.FromStream(stream);
            }
        }
        /// <summary>
        /// Converts a System.Drawing.Icon to a WPF ImageSource.
        /// </summary>
        /// <param name="icon">The icon to convert.</param>
        /// <returns>An ImageSource usable in WPF.</returns>
        public static System.Windows.Media.ImageSource ToImageSource(this System.Drawing.Icon icon)
        {
            using (var bitmap = icon.ToBitmap())
            {
                var hBitmap = bitmap.GetHbitmap();
                try
                {
                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                        hBitmap,
                        IntPtr.Zero,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());
                }
                finally
                {
                    DeleteObject(hBitmap); // Clean up the HBITMAP handle
                }
            }
        }

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);


    }



}
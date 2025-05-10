using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using IWshRuntimeLibrary;
using System.Windows.Media.Imaging; // Added for ImageSource
using System.Runtime.InteropServices; // Added for DeleteObject
using Desktop_Fences; // Ensure this is included for Utility extensions

namespace Desktop_Fences
{
    public static class Utility
    {



        public static void ApplyTintAndColorToFence(Window fence, string colorName = null)
        {
            var fenceControl = fence.Content as Border; // Matches your structure
            if (fenceControl == null) return;

            string effectiveColor = colorName ?? SettingsManager.SelectedColor; // Fallback to global
                                                                                // Use TintValue from SettingsManager to determine if tint is applied
            fenceControl.Background = SettingsManager.TintValue > 0
                ? new SolidColorBrush(GetColorFromName(effectiveColor)) { Opacity = SettingsManager.TintValue / 100.0 }
                : Brushes.Transparent;
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
            string[] executableExtensions = { ".exe", ".bat", ".cmd", ".vbs", ".ps1", ".hta",".msi" };
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
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace Desktop_Fences
{
    public static class WallpaperColorManager
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SystemParametersInfo(int uAction, int uParam, StringBuilder lpvParam, int fuWinIni);
        private const int SPI_GETDESKWALLPAPER = 0x0073;

        public static System.Windows.Media.Color CurrentWallpaperColor { get; private set; } = System.Windows.Media.Color.FromRgb(100, 100, 100);
        public static event EventHandler WallpaperColorChanged;

        private static FileSystemWatcher _wallpaperWatcher;
        private static DateTime _lastUpdate = DateTime.MinValue;

        public static void Initialize()
        {
            UpdateWallpaperColor();

            // Hook 1: Listen for Global Windows Theme changes
            SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;

            // Hook 2: Aggressive Physical File Watcher (Catches Slideshows & "Set as Desktop Background")
            try
            {
                string themesDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Windows\Themes");
                if (Directory.Exists(themesDir))
                {
                    _wallpaperWatcher = new FileSystemWatcher(themesDir)
                    {
                        Filter = "TranscodedWallpaper",
                        NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime,
                        EnableRaisingEvents = true
                    };
                    _wallpaperWatcher.Changed += PhysicalWallpaperFileChanged;
                    _wallpaperWatcher.Created += PhysicalWallpaperFileChanged;
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General, $"Could not start Wallpaper Watcher: {ex.Message}");
            }
        }

        private static void PhysicalWallpaperFileChanged(object sender, FileSystemEventArgs e)
        {
            TriggerUpdate();
        }

        private static void SystemEvents_UserPreferenceChanged(object sender, UserPreferenceChangedEventArgs e)
        {
            if (e.Category == UserPreferenceCategory.Desktop || e.Category == UserPreferenceCategory.General)
            {
                TriggerUpdate();
            }
        }

        private static void TriggerUpdate()
        {
            // Debounce: If Windows fires 5 events in half a second, we only process it once!
            if ((DateTime.Now - _lastUpdate).TotalMilliseconds < 500) return;
            _lastUpdate = DateTime.Now;

            System.Windows.Application.Current?.Dispatcher.BeginInvoke(new Action(async () =>
            {
                // Wait exactly 250ms for Windows to finish saving the file to the hard drive
                await System.Threading.Tasks.Task.Delay(250);
                UpdateWallpaperColor();
                WallpaperColorChanged?.Invoke(null, EventArgs.Empty);
            }));
        }

        public static void UpdateWallpaperColor()
        {
            try
            {
                string wallpaperPath = GetActiveWallpaperPath();

                if (!string.IsNullOrEmpty(wallpaperPath) && File.Exists(wallpaperPath))
                {
                    CurrentWallpaperColor = GetDominantColor(wallpaperPath);
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General, $"Chameleon updated color successfully from: {wallpaperPath}");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Chameleon Error: {ex.Message}");
            }
        }

        private static string GetActiveWallpaperPath()
        {
            StringBuilder wallPaperPath = new StringBuilder(260);
            SystemParametersInfo(SPI_GETDESKWALLPAPER, 260, wallPaperPath, 0);
            string path = wallPaperPath.ToString();

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
                return path;

            string transcodedPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Windows\Themes\TranscodedWallpaper");
            if (File.Exists(transcodedPath))
                return transcodedPath;

            return null;
        }

        private static System.Windows.Media.Color GetDominantColor(string imagePath)
        {
            try
            {
                // ULTRA FIX: Use FileStream with ReadWrite share so Windows never blocks us!
                byte[] imageBytes;
                using (FileStream fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    imageBytes = new byte[fs.Length];
                    fs.Read(imageBytes, 0, imageBytes.Length);
                }

                using (MemoryStream ms = new MemoryStream(imageBytes))
                using (Bitmap originalBmp = new Bitmap(ms))
                using (Bitmap bmp = new Bitmap(originalBmp, new Size(32, 32))) // 4x Faster: Shrink to 1024 total pixels
                {
                    Dictionary<int, int> colorCounts = new Dictionary<int, int>();
                    int maxCount = 0;
                    int dominantColor = 0;

                    for (int x = 0; x < bmp.Width; x++)
                    {
                        for (int y = 0; y < bmp.Height; y++)
                        {
                            System.Drawing.Color p = bmp.GetPixel(x, y);

                            if (p.GetSaturation() < 0.15f || p.GetBrightness() < 0.15f || p.GetBrightness() > 0.85f)
                                continue;

                            int r = (p.R / 16) * 16;
                            int g = (p.G / 16) * 16;
                            int b = (p.B / 16) * 16;

                            int hash = (r << 16) | (g << 8) | b;

                            if (colorCounts.TryGetValue(hash, out int count))
                                colorCounts[hash] = count + 1;
                            else
                                colorCounts[hash] = 1;

                            if (colorCounts[hash] > maxCount)
                            {
                                maxCount = colorCounts[hash];
                                dominantColor = hash;
                            }
                        }
                    }

                    if (maxCount > 0)
                    {
                        byte r = (byte)((dominantColor >> 16) & 0xFF);
                        byte g = (byte)((dominantColor >> 8) & 0xFF);
                        byte b = (byte)(dominantColor & 0xFF);

                        r = (byte)Math.Min(255, r + 20);
                        g = (byte)Math.Min(255, g + 20);
                        b = (byte)Math.Min(255, b + 20);

                        return System.Windows.Media.Color.FromRgb(r, g, b);
                    }

                    return GetAverageColor(bmp);
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General, $"Chameleon image extraction failed: {ex.Message}");
                return System.Windows.Media.Color.FromRgb(70, 90, 110);
            }
        }

        private static System.Windows.Media.Color GetAverageColor(Bitmap bmp)
        {
            long r = 0, g = 0, b = 0;
            int count = bmp.Width * bmp.Height;

            for (int x = 0; x < bmp.Width; x++)
            {
                for (int y = 0; y < bmp.Height; y++)
                {
                    System.Drawing.Color c = bmp.GetPixel(x, y);
                    r += c.R;
                    g += c.G;
                    b += c.B;
                }
            }

            return System.Windows.Media.Color.FromRgb(
                (byte)Math.Min(255, (r / count) + 30),
                (byte)Math.Min(255, (g / count) + 30),
                (byte)Math.Min(255, (b / count) + 30));
        }
    }
}
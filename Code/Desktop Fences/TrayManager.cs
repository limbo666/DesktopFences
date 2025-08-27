using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using WinFormsMouseEventArgs = System.Windows.Forms.MouseEventArgs;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Desktop_Fences
{
    public class TrayManager : IDisposable
    {
        private NotifyIcon _trayIcon;
      
        private bool _disposed;
        public static bool IsStartWithWindows { get; private set; }
        
        private static readonly List<HiddenFence> HiddenFences = new List<HiddenFence>();
    
        private ToolStripMenuItem _showHiddenFencesItem;
        public static TrayManager Instance { get; private set; } // Singleton instance

        private bool _areFencesTempHidden = false;
    
        private List<NonActivatingWindow> _tempHiddenFences = new List<NonActivatingWindow>();

        private bool Showintray = SettingsManager.ShowInTray;

        private const int WM_NCLBUTTONDOWN = 0xA1;

        private const int HT_CAPTION = 0x2;

        private class HiddenFence
        {
            public string Title { get; set; }
            public NonActivatingWindow Window { get; set; }
        }

        public TrayManager()
        {
            IsStartWithWindows = IsInStartupFolder();
            Instance = this; // Set singleton instance
        }
          
        private void OnTrayIconDoubleClick(object sender, EventArgs e)
        {
            if (!_areFencesTempHidden)
            {
                var visibleFences = System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>()
                    .Where(w => w.Visibility == Visibility.Visible &&
                           !HiddenFences.Any(hf => hf.Window == w))
                    .ToList();

                foreach (var fence in visibleFences)
                {
                    fence.Visibility = Visibility.Hidden;
                    _tempHiddenFences.Add(fence);
                }
                _areFencesTempHidden = true;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Temporarily hid {visibleFences.Count} fences.");
            }
            else
            {
                int count = _tempHiddenFences.Count;
                foreach (var fence in _tempHiddenFences)
                {
                    fence.Visibility = Visibility.Visible;
                }
                _tempHiddenFences.Clear();
                _areFencesTempHidden = false;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Restored {count} temporarily hidden fences.");
            }
            UpdateTrayIcon();
        }

        /// <summary>
        /// Handles single click on tray icon - checks for special key combination CTRL+ALT+SHIFT
        /// </summary>
        private void OnTrayIconClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            try
            {
                // Only handle left click
                if (e.Button != MouseButtons.Left) return;

                // Check if CTRL+ALT+SHIFT are all pressed
                bool isCtrlPressed = (Control.ModifierKeys & Keys.Control) == Keys.Control;
                bool isAltPressed = (Control.ModifierKeys & Keys.Alt) == Keys.Alt;
                bool isShiftPressed = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;

                if (isCtrlPressed && isAltPressed && isShiftPressed)
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                        "TrayIcon: CTRL+ALT+SHIFT+Click detected - Exporting registry values");

                    // Execute the registry export function
                    bool success = RegistryHelper.ExportProgramManagementValues();

                    if (success)
                    {
                        // Show notification that export was successful
                        _trayIcon.BalloonTipTitle = "Desktop Fences Plus";
                        _trayIcon.BalloonTipText = "Registry values exported successfully to program folder.";
                        _trayIcon.BalloonTipIcon = ToolTipIcon.Info;
                        _trayIcon.ShowBalloonTip(3000); // Show for 3 seconds

                        LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.General,
                            "TrayIcon: Registry values export completed successfully");
                    }
                    else
                    {
                        // Show error notification
                        _trayIcon.BalloonTipTitle = "Desktop Fences Plus - Error";
                        _trayIcon.BalloonTipText = "Failed to export registry values. Check log for details.";
                        _trayIcon.BalloonTipIcon = ToolTipIcon.Error;
                        _trayIcon.ShowBalloonTip(3000);

                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                            "TrayIcon: Registry values export failed");
                    }
                }
                else
                {
                    // Log debug info about key states (only if at least one modifier is pressed)
                    if (isCtrlPressed || isAltPressed || isShiftPressed)
                    {
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"TrayIcon: Single click with modifiers - Ctrl:{isCtrlPressed}, Alt:{isAltPressed}, Shift:{isShiftPressed}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.General,
                    $"TrayIcon: Error in click handler: {ex.Message}");
            }
        }
        public void InitializeTray()
        {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            _trayIcon = new NotifyIcon
            {
                Icon = Icon.ExtractAssociatedIcon(exePath),
                Visible = true,
                Text = "Desktop Fences"
            };

            _trayIcon.DoubleClick += OnTrayIconDoubleClick;
            _trayIcon.MouseClick += OnTrayIconClick; // Add single click handler

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("About", null, (s, e) => AboutFormManager.ShowAboutForm());
            trayMenu.Items.Add("Options", null, (s, e) => OptionsFormManager.ShowOptionsForm());
          //trayMenu.Items.Add("Registry Test", null, (s, e) => InterCore.ActivateLighthouseSweep()); // TEST NEW OPTIONS FORM
            trayMenu.Items.Add("Reload All Fences", null, async (s, e) =>
            {
                reloadAllFences();

            });
            trayMenu.Items.Add("-");
            _showHiddenFencesItem = new ToolStripMenuItem("Show Hidden Fences")
            {
                Enabled = false
            };
            trayMenu.Items.Add(_showHiddenFencesItem);
            trayMenu.Items.Add("-");
            trayMenu.Items.Add("Exit", null, (s, e) => System.Windows.Application.Current.Shutdown());
            _trayIcon.ContextMenuStrip = trayMenu;

            UpdateHiddenFencesMenu();
            UpdateTrayIcon();
        }

        public static async Task reloadAllFences()
        {
            // Keep WPF Window (original was correct) - just modernize styling with proper namespaces
            var waitWindow = new System.Windows.Window
            {
                Title = "Desktop Fences +",
                Width = 300,
                Height = 150,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                WindowStyle = System.Windows.WindowStyle.None,
                Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(248, 249, 250)),
                AllowsTransparency = true,
                Topmost = true
            };

            // Modern main container
            var mainBorder = new System.Windows.Controls.Border
            {
                Background = System.Windows.Media.Brushes.White,
                CornerRadius = new System.Windows.CornerRadius(8),
                Margin = new System.Windows.Thickness(8),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    Color = System.Windows.Media.Colors.Black,
                    Direction = 270,
                    ShadowDepth = 4,
                    Opacity = 0.15,
                    BlurRadius = 8
                }
            };

            var waitStack = new System.Windows.Controls.StackPanel
            {
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Center,
                Orientation = System.Windows.Controls.Orientation.Vertical
            };

            // App title
            var titleText = new System.Windows.Controls.TextBlock
            {
                Text = "Desktop Fences +",
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 16,
                FontWeight = System.Windows.FontWeights.Medium,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(32, 33, 36)),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                Margin = new System.Windows.Thickness(0, 0, 0, 10)
            };
            waitStack.Children.Add(titleText);

            // Logo with fallback
            var logoImage = new System.Windows.Controls.Image
            {
                Width = 32,
                Height = 32,
                Margin = new System.Windows.Thickness(0, 0, 0, 10),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center
            };

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var resourceStream = assembly.GetManifestResourceStream("Desktop_Fences.Resources.logo1.png");
                if (resourceStream != null)
                {
                    var bitmap = new System.Windows.Media.Imaging.BitmapImage();
                    bitmap.BeginInit();
                    bitmap.StreamSource = resourceStream;
                    bitmap.EndInit();
                    logoImage.Source = bitmap;
                }
                else
                {
                    string exePath = Assembly.GetEntryAssembly().Location;
                    logoImage.Source = System.Drawing.Icon.ExtractAssociatedIcon(exePath).ToImageSource();
                }
            }
            catch
            {
                string exePath = Assembly.GetEntryAssembly().Location;
                logoImage.Source = System.Drawing.Icon.ExtractAssociatedIcon(exePath).ToImageSource();
            }
            waitStack.Children.Add(logoImage);

            // Wait message
            var waitText = new System.Windows.Controls.TextBlock
            {
                Text = "Reloading all fences, please wait...",
                FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                FontSize = 12,
                Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(95, 99, 104)),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center
            };
            waitStack.Children.Add(waitText);

            mainBorder.Child = waitStack;
            waitWindow.Content = mainBorder;
            waitWindow.Show();
            try
            {
                await Task.Run(() =>
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        foreach (var fence in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList())
                        {
                            fence.Close();
                        }
                        FenceManager.ReloadFences();
                    });
                });
            }
            catch (Exception ex)
            {
                MessageBoxesManager.ShowOKOnlyMessageBoxFormStatic($"An error occurred while reloading fences: {ex.Message}", "Error");
            }
            finally
            {
                waitWindow.Close();
            }

            // TrayManager.ShowOKOnlyMessageBoxFormStatic($"An error occurred while reloading fences: {ex.Message}", "Error");
        }
      
        public static void AddHiddenFence(NonActivatingWindow fence)
        {
            if (fence == null || string.IsNullOrEmpty(fence.Title)) return;
            fence.Dispatcher.Invoke(() =>
            {
                fence.Visibility = Visibility.Hidden;
            });

            if (!HiddenFences.Any(f => f.Title == fence.Title))
            {
                HiddenFences.Add(new HiddenFence { Title = fence.Title, Window = fence });
                fence.Visibility = System.Windows.Visibility.Hidden;
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Added fence '{fence.Title}' to hidden list");
                Instance?.UpdateHiddenFencesMenu();
                Instance?.UpdateTrayIcon();
            }
        }

        public static void ShowHiddenFence(string title)
        {
            var hiddenFence = HiddenFences.FirstOrDefault(f => f.Title == title);
            if (hiddenFence != null)
            {
                hiddenFence.Window.Dispatcher.Invoke(() =>
                {
                    hiddenFence.Window.Visibility = Visibility.Visible;
                    hiddenFence.Window.Activate();
                    hiddenFence.Window.Show();
                });

                var fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == title);
                if (fenceData != null)
                {
                    FenceManager.UpdateFenceProperty(fenceData, "IsHidden", "false", $"Showed fence '{title}'");
                }
                HiddenFences.Remove(hiddenFence);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, $"Showed fence '{title}'");
                Instance?.UpdateHiddenFencesMenu();
                Instance?.UpdateTrayIcon();
            }
        }

        public void UpdateHiddenFencesMenu()
        {
            if (_showHiddenFencesItem == null) return;

            _showHiddenFencesItem.DropDownItems.Clear();
            _showHiddenFencesItem.Enabled = HiddenFences.Count > 0;

            foreach (var fence in HiddenFences)
            {
                var menuItem = new ToolStripMenuItem(fence.Title);
                menuItem.Click += (s, e) => ShowHiddenFence(fence.Title);
                _showHiddenFencesItem.DropDownItems.Add(menuItem);
            }
        }
     
        private bool IsInStartupFolder()
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
            return File.Exists(shortcutPath);
        }

        public void ToggleStartWithWindows(bool enable)
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            string workingDir = Path.GetDirectoryName(exePath); // Ensure working directory is extracted

            try
            {
                if (enable && !IsInStartupFolder())
                {
                    Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                    dynamic shell = Activator.CreateInstance(shellType);
                    var shortcut = shell.CreateShortcut(shortcutPath);
                    shortcut.TargetPath = exePath;
                    shortcut.WorkingDirectory = workingDir; // Explicitly set working directory
                    shortcut.Description = "Desktop Fences Startup Shortcut";
                    shortcut.Save();
                    IsStartWithWindows = true;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Added Desktop Fences to Startup folder with working directory: " + workingDir);
                }
                else if (!enable && IsInStartupFolder())
                {
                    File.Delete(shortcutPath);
                    IsStartWithWindows = false;
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "Removed Desktop Fences from Startup folder");
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Failed to toggle Start with Windows: {ex.Message}");
                IsStartWithWindows = IsInStartupFolder();
                throw;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _trayIcon?.Dispose();
            _disposed = true;
        }

        private Icon GenerateIconWithNumber(int count)
        {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            using (var baseIcon = Icon.ExtractAssociatedIcon(exePath))

            using (var bitmap = baseIcon.ToBitmap())
            using (var graphics = Graphics.FromImage(bitmap))
            {
                int circleDiameter = 24;
                int circleX = -4;
                int circleY = -1;

                var circleBrush = new SolidBrush(Color.FromArgb(230, 255, 153, 53));
                graphics.FillEllipse(circleBrush, circleX, circleY, circleDiameter, circleDiameter);

                var font = new Font("Calibri", 26, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
                var textBrush = new SolidBrush(Color.Navy);

                string text = count.ToString();
                var textSize = graphics.MeasureString(text, font);
                float textX = circleX + (circleDiameter - textSize.Width) / 2;
                float textY = circleY + (circleDiameter - textSize.Height) / 2;

                graphics.DrawString(text, font, textBrush, textX, textY);

                return Icon.FromHandle(bitmap.GetHicon());
            }
        }

        public void UpdateTrayIcon()
        {

            if (Showintray == true)
            {

                if (HiddenFences.Count > 0)
                {



                    _trayIcon.Icon = GenerateIconWithNumber(HiddenFences.Count + _tempHiddenFences.Count);
                }
                else
                {
                    string exePath = Process.GetCurrentProcess().MainModule.FileName;
                    _trayIcon.Icon = Icon.ExtractAssociatedIcon(exePath);
                }
            }
            else
            {
                _trayIcon.Icon = null; // Hide the icon if Showintray is false
            }
        }

           

    }
}
using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Collections.Generic;

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

        public void InitializeTray()
        {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            _trayIcon = new NotifyIcon
            {
                Icon = Icon.ExtractAssociatedIcon(exePath),
                Visible = true,
                Text = "Desktop Fences"
            };

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("About", null, (s, e) => ShowAboutForm());
            trayMenu.Items.Add("-", null);
            trayMenu.Items.Add("Options", null, (s, e) => ShowOptionsForm());
            _showHiddenFencesItem = new ToolStripMenuItem("Show Hidden Fences") // Simple constructor
            {
                Enabled = false
            };
            trayMenu.Items.Add(_showHiddenFencesItem);
            trayMenu.Items.Add("Exit", null, (s, e) => System.Windows.Application.Current.Shutdown());
            _trayIcon.ContextMenuStrip = trayMenu;

            UpdateHiddenFencesMenu();
        }

        // Add fence to hidden list and update tray menu
        public static void AddHiddenFence(NonActivatingWindow fence)
        {
            if (fence == null || string.IsNullOrEmpty(fence.Title)) return;

            if (!HiddenFences.Any(f => f.Title == fence.Title))
            {
                HiddenFences.Add(new HiddenFence { Title = fence.Title, Window = fence });
                fence.Visibility = System.Windows.Visibility.Hidden;
                Log($"Added fence '{fence.Title}' to hidden list");
                Instance?.UpdateHiddenFencesMenu();
            }
        }

        // Show hidden fence and update tray menu
        public static void ShowHiddenFence(string title)
        {
            var hiddenFence = HiddenFences.FirstOrDefault(f => f.Title == title);
            if (hiddenFence != null)
            {
                hiddenFence.Window.Visibility = System.Windows.Visibility.Visible;
                var fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == title);
                if (fenceData != null)
                {
                    FenceManager.UpdateFenceProperty(fenceData, "IsHidden", "false", $"Showed fence '{title}'");
                }
                HiddenFences.Remove(hiddenFence);
                Log($"Showed fence '{title}'");
                Instance?.UpdateHiddenFencesMenu();
            }
        }

        // Update tray menu with hidden fences
        private void UpdateHiddenFencesMenu()
        {
            if (_showHiddenFencesItem == null) return;

            _showHiddenFencesItem.DropDownItems.Clear();
            _showHiddenFencesItem.Enabled = HiddenFences.Count > 0;

            foreach (var fence in HiddenFences)
            {
                var menuItem = new ToolStripMenuItem(fence.Title); // Use simple constructor
                menuItem.Click += (s, e) => ShowHiddenFence(fence.Title); // Assign Click event separately
                _showHiddenFencesItem.DropDownItems.Add(menuItem);
            }
        }

        private void ShowOptionsForm()
        {
            try
            {
                using (var frmOptions = new Form())
                {
                    frmOptions.Text = "Options";
                    frmOptions.Size = new Size(260, 560);
                    frmOptions.StartPosition = FormStartPosition.CenterScreen;
                    frmOptions.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frmOptions.MaximizeBox = false;
                    frmOptions.MinimizeBox = false;
                    frmOptions.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

                    var layoutPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        Padding = new Padding(10)
                    };
                    layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                    var groupBoxGeneral = new GroupBox
                    {
                        Text = "General",
                        Dock = DockStyle.Top,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    var generalLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        Padding = new Padding(5)
                    };
                    generalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                    var chkStartWithWindows = new CheckBox
                    {
                        Text = "Start with Windows",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = IsStartWithWindows
                    };
                    chkStartWithWindows.CheckedChanged += (s, e) =>
                    {
                        try
                        {
                            ToggleStartWithWindows(chkStartWithWindows.Checked);
                            Log($"Set Start with Windows to {chkStartWithWindows.Checked} via Options");
                        }
                        catch (Exception ex)
                        {
                            Log($"Error setting Start with Windows: {ex.Message}");
                            System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Startup Toggle Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            chkStartWithWindows.Checked = IsStartWithWindows;
                        }
                    };
                    generalLayout.Controls.Add(chkStartWithWindows, 0, 0);
                    groupBoxGeneral.Controls.Add(generalLayout);
                    layoutPanel.Controls.Add(groupBoxGeneral);

                    var groupBoxSelections = new GroupBox
                    {
                        Text = "Selections",
                        Dock = DockStyle.Top,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    var selectionsLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 2,
                        RowCount = 5,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    selectionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    selectionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

                    var chkEnableSnap = new CheckBox
                    {
                        Text = "Enable snap function",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.IsSnapEnabled
                    };
                    selectionsLayout.Controls.Add(chkEnableSnap, 0, 0);
                    selectionsLayout.SetColumnSpan(chkEnableSnap, 2);

                    var lblTint = new Label { Text = "Tint", Dock = DockStyle.Fill, AutoSize = true };
                    var numTint = new NumericUpDown
                    {
                        Maximum = 100,
                        Minimum = 1,
                        Value = SettingsManager.TintValue,
                        Dock = DockStyle.Fill
                    };
                    selectionsLayout.Controls.Add(lblTint, 0, 1);
                    selectionsLayout.Controls.Add(numTint, 1, 1);

                    var lblColor = new Label { Text = "Color", Dock = DockStyle.Fill, AutoSize = true };
                    var cmbColor = new ComboBox
                    {
                        Dock = DockStyle.Fill,
                        DropDownStyle = ComboBoxStyle.DropDownList
                    };
                    cmbColor.Items.AddRange(new string[] { "Gray", "Black", "White", "Green", "Purple", "Yellow", "Red", "Blue" });
                    cmbColor.SelectedItem = SettingsManager.SelectedColor;
                    selectionsLayout.Controls.Add(lblColor, 0, 2);
                    selectionsLayout.Controls.Add(cmbColor, 1, 2);

                    var lblLaunchEffect = new Label { Text = "Launch Effect", Dock = DockStyle.Fill, AutoSize = true };
                    var cmbLaunchEffect = new ComboBox
                    {
                        Dock = DockStyle.Fill,
                        DropDownStyle = ComboBoxStyle.DropDownList
                    };
                    cmbLaunchEffect.Items.AddRange(new string[] { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate" });
                    cmbLaunchEffect.SelectedIndex = (int)SettingsManager.LaunchEffect;
                    selectionsLayout.Controls.Add(lblLaunchEffect, 0, 3);
                    selectionsLayout.Controls.Add(cmbLaunchEffect, 1, 3);

                    var chkSingleClickToLaunch = new CheckBox
                    {
                        Text = "Single Click to Launch",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.SingleClickToLaunch
                    };
                    selectionsLayout.Controls.Add(chkSingleClickToLaunch, 0, 4);
                    selectionsLayout.SetColumnSpan(chkSingleClickToLaunch, 2);

                    groupBoxSelections.Controls.Add(selectionsLayout);
                    layoutPanel.Controls.Add(groupBoxSelections);

                    var groupBoxTools = new GroupBox
                    {
                        Text = "Tools",
                        Dock = DockStyle.Top,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    var toolsLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    toolsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                    var chkEnableLog = new CheckBox
                    {
                        Text = "Enable log",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.IsLogEnabled
                    };
                    toolsLayout.Controls.Add(chkEnableLog, 0, 0);

                    var btnBackup = new Button
                    {
                        Text = "Backup",
                        AutoSize = true,
                        Width = 80,
                        Height = 30,
                        Anchor = AnchorStyles.None
                    };
                    btnBackup.Click += (s, ev) => BackupManager.BackupData();
                    toolsLayout.Controls.Add(btnBackup, 0, 1);

                    groupBoxTools.Controls.Add(toolsLayout);
                    layoutPanel.Controls.Add(groupBoxTools);

                    var buttonsLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        ColumnCount = 2,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

                    var btnCancel = new Button
                    {
                        Text = "Cancel",
                        AutoSize = true,
                        Width = 80,
                        Height = 30,
                        Anchor = AnchorStyles.Right
                    };
                    btnCancel.Click += (s, ev) => frmOptions.Close();

                    var btnSave = new Button
                    {
                        Text = "Save",
                        AutoSize = true,
                        Width = 80,
                        Height = 30,
                        Anchor = AnchorStyles.Right
                    };
                    btnSave.Click += (s, ev) =>
                    {
                        SettingsManager.IsSnapEnabled = chkEnableSnap.Checked;
                        SettingsManager.TintValue = (int)numTint.Value;
                        SettingsManager.SelectedColor = cmbColor.SelectedItem.ToString();
                        SettingsManager.IsLogEnabled = chkEnableLog.Checked;
                        SettingsManager.SingleClickToLaunch = chkSingleClickToLaunch.Checked;
                        SettingsManager.LaunchEffect = (FenceManager.LaunchEffect)cmbLaunchEffect.SelectedIndex;
                        IsStartWithWindows = chkStartWithWindows.Checked;

                        SettingsManager.SaveSettings();
                        FenceManager.UpdateOptionsAndClickEvents();

                        foreach (var fence in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>())
                        {
                            dynamic fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == fence.Title);
                            if (fenceData != null)
                            {
                                string customColor = fenceData.CustomColor?.ToString();
                                string appliedColor = string.IsNullOrEmpty(customColor) ? SettingsManager.SelectedColor : customColor;
                                Utility.ApplyTintAndColorToFence(fence, appliedColor);
                                Log($"Applied color '{appliedColor}' with global tint '{SettingsManager.TintValue}' to fence '{fence.Title}'");
                            }
                            else
                            {
                                Utility.ApplyTintAndColorToFence(fence, SettingsManager.SelectedColor);
                                Log($"Applied global color '{SettingsManager.SelectedColor}' with global tint '{SettingsManager.TintValue}' to unknown fence '{fence.Title}'");
                            }
                        }

                        frmOptions.Close();
                    };

                    buttonsLayout.Controls.Add(btnCancel, 0, 0);
                    buttonsLayout.Controls.Add(btnSave, 1, 0);
                    layoutPanel.Controls.Add(buttonsLayout);

                    var donatePictureBox = new PictureBox
                    {
                        Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.donate.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Dock = DockStyle.Fill,
                        Height = 25,
                        Cursor = Cursors.Hand
                    };
                    donatePictureBox.Click += (s, e) =>
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = "https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE",
                                UseShellExecute = true
                            });
                        }
                        catch (Exception ex)
                        {
                            Log($"Error opening donation link: {ex.Message}");
                            System.Windows.Forms.MessageBox.Show($"Error opening donation link: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };
                    layoutPanel.Controls.Add(donatePictureBox);

                    var donateLabel = new Label
                    {
                        Text = "Donate to help development",
                        Font = new Font("Tahoma", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(donateLabel);

                    frmOptions.Controls.Add(layoutPanel);
                    frmOptions.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                Log($"Error showing Options form: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ShowAboutForm()
        {
            try
            {
                using (var frmAbout = new Form())
                {
                    frmAbout.Text = "About Desktop Fences +";
                    frmAbout.Size = new Size(400, 600);
                    frmAbout.StartPosition = FormStartPosition.CenterScreen;
                    frmAbout.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frmAbout.MaximizeBox = false;
                    frmAbout.MinimizeBox = false;
                    frmAbout.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

                    var layoutPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        RowCount = 9,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        Padding = new Padding(20)
                    };
                    layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                    var pictureBox = new PictureBox
                    {
                        Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.logo1.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Dock = DockStyle.Fill,
                        Height = 50
                    };
                    layoutPanel.Controls.Add(pictureBox);

                    var labelTitle = new Label
                    {
                        Text = "Desktop Fences +",
                        Font = new Font("Tahoma", 14, System.Drawing.FontStyle.Bold),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(labelTitle);

                    var version = Assembly.GetExecutingAssembly().GetName().Version;
                    var labelVersion = new Label
                    {
                        Text = $"ver {version}",
                        Font = new Font("Tahoma", 10, System.Drawing.FontStyle.Bold),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(labelVersion);

                    var labelMainText = new Label
                    {
                        Text = "Desktop Fences + is an open-source alternative to StarDock's Fences, originally created by HakanKokcu as Birdy Fences.\n\nDesktop fences +, is maintained by Nikos Georgousis, has been enhanced and optimized to give better user experience and stability.\n\n ",
                        Font = new Font("Tahoma", 10),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(labelMainText);

                    var horizontalLine = new Label
                    {
                        BorderStyle = BorderStyle.Fixed3D,
                        Height = 2,
                        Dock = DockStyle.Fill,
                        Margin = new Padding(10)
                    };
                    layoutPanel.Controls.Add(horizontalLine);

                    var labelGitHubText = new Label
                    {
                        Text = "Please visit GitHub for news, updates, and bug reports.",
                        Font = new Font("Tahoma", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(labelGitHubText);

                    var linkLabelGitHub = new LinkLabel
                    {
                        Text = "https://github.com/limbo666/DesktopFences",
                        Font = new Font("Tahoma", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    linkLabelGitHub.LinkClicked += (sender, e) =>
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = "https://github.com/limbo666/DesktopFences",
                                UseShellExecute = true
                            });
                        }
                        catch (Exception ex)
                        {
                            Log($"Error opening GitHub link: {ex.Message}");
                            System.Windows.Forms.MessageBox.Show($"Error opening GitHub link: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };
                    layoutPanel.Controls.Add(linkLabelGitHub);

                    var horizontalLine2 = new Label
                    {
                        BorderStyle = BorderStyle.Fixed3D,
                        Height = 2,
                        Dock = DockStyle.Fill,
                        Margin = new Padding(10)
                    };
                    layoutPanel.Controls.Add(linkLabelGitHub);
                    layoutPanel.Controls.Add(horizontalLine2);

               
                  

                    // Donation Logo
                    var donatePictureBox = new PictureBox
                    {
                        Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.donate.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Dock = DockStyle.Fill,
                        Height = 50, // Adjust based on donate.png size
                        Cursor = Cursors.Hand // Indicate clickability
                    };
                    donatePictureBox.Click += (s, e) =>
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = "https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE",
                                UseShellExecute = true
                            });
                        }
                        catch (Exception ex)
                        {
                            Log($"Error opening donation link: {ex.Message}");
                            System.Windows.MessageBox.Show($"Error opening donation link: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                        }
                    };
                    layoutPanel.Controls.Add(donatePictureBox);

                    // Donation Text
                    var donateLabel = new Label
                    {
                        Text = "Donate to help development",
                        Font = new Font("Tahoma", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(donateLabel);

                    frmAbout.Controls.Add(layoutPanel);
                    frmAbout.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                Log($"Error showing About form: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool IsInStartupFolder()
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
            return File.Exists(shortcutPath);
        }

        private void ToggleStartWithWindows(bool enable)
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
            string exePath = Process.GetCurrentProcess().MainModule.FileName;

            try
            {
                if (enable && !IsInStartupFolder())
                {
                    Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                    dynamic shell = Activator.CreateInstance(shellType);
                    var shortcut = shell.CreateShortcut(shortcutPath);
                    shortcut.TargetPath = exePath;
                    shortcut.Description = "Desktop Fences Startup Shortcut";
                    shortcut.Save();
                    IsStartWithWindows = true;
                    Log("Added Desktop Fences to Startup folder");
                }
                else if (!enable && IsInStartupFolder())
                {
                    File.Delete(shortcutPath);
                    IsStartWithWindows = false;
                    Log("Removed Desktop Fences from Startup folder");
                }
            }
            catch (Exception ex)
            {
                Log($"Failed to toggle Start with Windows: {ex.Message}");
                IsStartWithWindows = IsInStartupFolder();
                throw;
            }
        }

        // Dispose of NotifyIcon to prevent resource leaks
        public void Dispose()
        {
            if (_disposed) return;
            _trayIcon?.Dispose();
            _disposed = true;
        }

        private static void Log(string message)
        {
            if (SettingsManager.IsLogEnabled)
            {
                string logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
            }
        }
    }
}
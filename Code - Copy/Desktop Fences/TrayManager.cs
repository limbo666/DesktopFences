using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;
using System.Linq; // Added for OfType
using System.IO; // Added for File operations

namespace Desktop_Fences
{
    public class TrayManager
    {
        private NotifyIcon _trayIcon;
        private ToolStripMenuItem _startWithWindowsItem; // New field for the menu item

        public void InitializeTray()
        {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            _trayIcon = new NotifyIcon
            {
                Icon = Icon.ExtractAssociatedIcon(exePath),
                Visible = true,
                Text = "Desktop Fences" // Added tooltip for clarity
            };

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("About", null, (s, e) => ShowAboutForm());
            trayMenu.Items.Add("-", null); // Horizontal line
            trayMenu.Items.Add("Options", null, (s, e) => ShowOptionsForm());

            // Add "Start with Windows" menu item
            _startWithWindowsItem = new ToolStripMenuItem("Start with Windows")
            {
                Checked = IsInStartupFolder(), // Initial state
                CheckOnClick = true // Toggle on click
            };
            _startWithWindowsItem.Click += ToggleStartWithWindows;
            trayMenu.Items.Add(_startWithWindowsItem);

            trayMenu.Items.Add("Exit", null, (s, e) => System.Windows.Application.Current.Shutdown());
            _trayIcon.ContextMenuStrip = trayMenu;
        }

        private bool IsInStartupFolder()
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
            return File.Exists(shortcutPath);
        }

        private void ToggleStartWithWindows(object sender, EventArgs e)
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
            string exePath = Process.GetCurrentProcess().MainModule.FileName;

            try
            {
                if (_startWithWindowsItem.Checked)
                {
                    // Add shortcut to Startup folder
                    Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                    dynamic shell = Activator.CreateInstance(shellType);
                    var shortcut = shell.CreateShortcut(shortcutPath);
                    shortcut.TargetPath = exePath;
                    shortcut.Description = "Desktop Fences Startup Shortcut";
                    shortcut.Save();
                    Log("Added Desktop Fences to Startup folder");
                }
                else
                {
                    // Remove shortcut from Startup folder
                    if (File.Exists(shortcutPath))
                    {
                        File.Delete(shortcutPath);
                        Log("Removed Desktop Fences from Startup folder");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Failed to toggle Start with Windows: {ex.Message}");
              System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Startup Toggle Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Revert the checkbox state if the operation fails
                _startWithWindowsItem.Checked = IsInStartupFolder();
            }
        }

        private void Log(string message)
        {
            bool isLogEnabled = SettingsManager.IsLogEnabled;
            if (isLogEnabled)
            {
                string logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Desktop_Fences.log");
                File.AppendAllText(logPath, $"{DateTime.Now}: {message}\n");
            }
        }


        private void ShowOptionsForm()
        {
            try
            {
                using (var frmOptions = new Form())
                {
                    frmOptions.Text = "Options";
                    frmOptions.Size = new System.Drawing.Size(260, 400);
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
                        Dock = DockStyle.Bottom,
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
                        // Save global settings
                        SettingsManager.IsSnapEnabled = chkEnableSnap.Checked;
                        SettingsManager.TintValue = (int)numTint.Value;
                        SettingsManager.SelectedColor = cmbColor.SelectedItem.ToString();
                        SettingsManager.IsLogEnabled = chkEnableLog.Checked;
                        SettingsManager.SingleClickToLaunch = chkSingleClickToLaunch.Checked;
                        SettingsManager.LaunchEffect = (FenceManager.LaunchEffect)cmbLaunchEffect.SelectedIndex;

                        SettingsManager.SaveSettings();

                        // Update FenceManager options
                        FenceManager.UpdateOptionsAndClickEvents();

                        // Apply settings to all fences, respecting custom colors but applying global tint
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
                                // Fallback for fences not in _fenceData (shouldn’t happen)
                                Utility.ApplyTintAndColorToFence(fence, SettingsManager.SelectedColor);
                                Log($"Applied global color '{SettingsManager.SelectedColor}' with global tint '{SettingsManager.TintValue}' to unknown fence '{fence.Title}' (not found in fence data)");
                            }
                        }

                        frmOptions.Close();
                    };

                    buttonsLayout.Controls.Add(btnCancel, 0, 0);
                    buttonsLayout.Controls.Add(btnSave, 1, 0);
                    layoutPanel.Controls.Add(buttonsLayout);

                    frmOptions.Controls.Add(layoutPanel);
                    frmOptions.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
        // end of replacement
        // end of replacement

        private void ShowAboutForm()
        {
            // Unchanged from your original
            try
            {
                using (var frmAbout = new Form())
                {
                    frmAbout.Text = "About Desktop Fences +";
                    frmAbout.Size = new System.Drawing.Size(400, 450);
                    frmAbout.StartPosition = FormStartPosition.CenterScreen;
                    frmAbout.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frmAbout.MaximizeBox = false;
                    frmAbout.MinimizeBox = false;
                    frmAbout.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

                    var layoutPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        RowCount = 7,
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
                        Height = 100
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
                            System.Windows.MessageBox.Show($"Error opening GitHub link: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                        }
                    };
                    layoutPanel.Controls.Add(linkLabelGitHub);

                    frmAbout.Controls.Add(layoutPanel);
                    frmAbout.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"An error occurred: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        public void Dispose()
        {
            if (_trayIcon != null)
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
                _trayIcon = null;
            }
        }
    }
}
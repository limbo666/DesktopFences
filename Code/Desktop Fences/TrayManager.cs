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
            trayMenu.Items.Add("Options", null, (s, e) => ShowOptionsForm());
           // trayMenu.Items.Add("Diagnostics", null, (s, e) => ShowDiagnosticsForm());

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
            UpdateTrayIcon(); // Ensure the tray icon is updated on initialization
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
                Instance?.UpdateTrayIcon(); // Update the tray icon
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
                Instance?.UpdateTrayIcon(); // Update the tray icon
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


        public void ShowDiagnosticsForm()
        {
            try
            {
                using (var frmDiagnostics = new Form())
                {
                    frmDiagnostics.Text = "Desktop Fences + Diagnostics";
                    frmDiagnostics.Size = new Size(300, 300);
                    frmDiagnostics.StartPosition = FormStartPosition.CenterScreen;
                    frmDiagnostics.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frmDiagnostics.MaximizeBox = false;
                    frmDiagnostics.MinimizeBox = false;
                    frmDiagnostics.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

                    var layoutPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        Padding = new Padding(10),
                        BorderStyle = BorderStyle.Fixed3D
                    };
                    layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));


                    var groupBoxInfo = new GroupBox
                    {
                        Text = "Info",
                        Dock = DockStyle.Top,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    var infoLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 2,
                        RowCount = 5, // Set the number of rows to 5
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };

                    infoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    infoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

                    // Create arrays for labels and values for easier management
                    string[] variableNames = { "Variable1", "Variable2", "Variable3", "Variable4", "Variable5" };
                    string[] variableValues = { "Value1", "Value2", "Value3", "Value4", "Value5" };

                    // Loop through and create labels for variables and values.
                    for (int i = 0; i < 5; i++) // Loop 5 times for the 5 rows
                    {
                        var lblVariable = new Label
                        {
                            Text = variableNames[i],
                            Dock = DockStyle.Fill,
                            AutoSize = true,
                        };
                        infoLayout.Controls.Add(lblVariable, 0, i); // Add variable name label

                        var lblValue = new Label
                        {
                            Text = variableValues[i],
                            Dock = DockStyle.Fill,
                            //  AutoSize = true, // Removed AutoSize
                        };
                        infoLayout.Controls.Add(lblValue, 1, i); // Add variable value label.
                    }

                    groupBoxInfo.Controls.Add(infoLayout); // Add the infoLayout to the groupBoxInfo
                    layoutPanel.Controls.Add(groupBoxInfo);    // Add the groupBoxInfo to the layoutPanel.

                    // Add the blinking label
                    var blinkingLabel = new Label
                    {
                        Text = "Timed check",
                        Dock = DockStyle.Top, // Or any other suitable DockStyle
                        AutoSize = true,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = new Font("Arial", 12, FontStyle.Bold),
                        ForeColor = Color.Red, // Initial color
                        BackColor = Color.Transparent
                    };
                    layoutPanel.Controls.Add(blinkingLabel);

                    // Create and configure the timer
                    Timer blinkTimer = new Timer();
                    blinkTimer.Interval = 500; // Blink every 500 milliseconds (0.5 seconds)
                    bool isLabelVisible = true;

                    blinkTimer.Tick += (sender, e) =>
                    {
                        if (isLabelVisible)
                        {
                            blinkingLabel.ForeColor = Color.Red; // Change to red
                        }
                        else
                        {
                            blinkingLabel.ForeColor = Color.Green; // Change to green
                        }
                        isLabelVisible = !isLabelVisible; // Toggle visibility state
                    };

                    // Start the timer when the form is shown
                    frmDiagnostics.Shown += (sender, e) =>
                    {
                        blinkTimer.Start();
                    };

                    // Stop the timer when the form is closing
                    frmDiagnostics.FormClosing += (sender, e) =>
                    {
                        blinkTimer.Stop();
                        blinkTimer.Dispose(); // Important: Dispose of the timer to release resources
                    };


                    frmDiagnostics.Controls.Add(layoutPanel);
                    frmDiagnostics.Show();
                }
            }
            catch (Exception ex)
            {
                Log($"Error showing Diagnostics form: {ex.Message}");
                System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        public void ShowOptionsForm()
        {
            try
            {
                using (var frmOptions = new Form())
                {
                    frmOptions.Text = "Desktop Fences + Options";
                    frmOptions.Size = new Size(320, 540);
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
                    //var chkSingleClickToLaunch = new CheckBox
                    //{
                    //    Text = "Single Click to Launch",
                    //    Dock = DockStyle.Fill,
                    //    AutoSize = true,
                    //    Checked = SettingsManager.SingleClickToLaunch
                    //};
                    //selectionsLayout.Controls.Add(chkSingleClickToLaunch, 0, 4);
                    //selectionsLayout.SetColumnSpan(chkSingleClickToLaunch, 2);

                    var chkEnableSnap = new CheckBox
                    {
                        Text = "Enable snap function",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.IsSnapEnabled
                    };
                    selectionsLayout.Controls.Add(chkEnableSnap, 0, 1);
                    selectionsLayout.SetColumnSpan(chkEnableSnap, 2);

                    var chkSingleClickToLaunch = new CheckBox
                    {
                        Text = "Single Click to Launch",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.SingleClickToLaunch
                    };
                    selectionsLayout.Controls.Add(chkSingleClickToLaunch, 0, 0);
                    selectionsLayout.SetColumnSpan(chkSingleClickToLaunch, 2);

                    var lblTint = new Label { Text = "Tint", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
                    var numTint = new NumericUpDown
                    {
                        Maximum = 100,
                        Minimum = 1,
                        Value = SettingsManager.TintValue,
                        Dock = DockStyle.Fill,
                        MaximumSize = new Size(80, 0), // Set maximum width to 150 pixels, height is unrestricted
                        Anchor = AnchorStyles.Right
                    };
                    selectionsLayout.Controls.Add(lblTint, 0, 2);
                    selectionsLayout.Controls.Add(numTint, 1, 2);

                    var lblColor = new Label { Text = "Color", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
                    var cmbColor = new ComboBox
                    {
                        Dock = DockStyle.Fill,
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        MaximumSize = new Size(80, 0), // Set maximum width to 150 pixels, height is unrestricted
                        Anchor = AnchorStyles.Right
                    };
                    cmbColor.Items.AddRange(new string[] { "Gray", "Black", "White", "Green", "Purple", "Yellow", "Red", "Blue" });

                    cmbColor.SelectedItem = SettingsManager.SelectedColor;
                    selectionsLayout.Controls.Add(lblColor, 0, 3);
                    selectionsLayout.Controls.Add(cmbColor, 1, 3);

                    var lblLaunchEffect = new Label { Text = "Launch Effect", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
                    var cmbLaunchEffect = new ComboBox
                    {
                        Dock = DockStyle.Fill,
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        MaximumSize = new Size(80, 0),// Set maximum width to 150 pixels, height is unrestricted

                        Anchor = AnchorStyles.Right
                    };
                    cmbLaunchEffect.Items.AddRange(new string[] { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate" });

                    cmbLaunchEffect.SelectedIndex = (int)SettingsManager.LaunchEffect;
                    selectionsLayout.Controls.Add(lblLaunchEffect, 0, 4);
                    selectionsLayout.Controls.Add(cmbLaunchEffect, 1, 4);


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

                    var btnBackup = new Button
                    {
                        Text = "Backup",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right
                    };
                    btnBackup.Click += (s, ev) => BackupManager.BackupData();
                    toolsLayout.Controls.Add(btnBackup, 0, 0);
                    //var btnRestore = new Button
                    //{
                    //    Text = "Restore",
                    //    AutoSize = false,
                    //    Width = 85,
                    //    Height = 25,
                    //    Anchor = AnchorStyles.Right
                    //};


                    var chkEnableLog = new CheckBox
                    {
                        Text = "Enable logging",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.IsLogEnabled
                    };
                    toolsLayout.Controls.Add(chkEnableLog, 0, 1);



                    groupBoxTools.Controls.Add(toolsLayout);
                    layoutPanel.Controls.Add(groupBoxTools);

                    var buttonsLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        ColumnCount = 3, // Three columns: spacer, Cancel button, Save button
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };

                    // Define column styles
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F)); // Spacer column (60% of width)
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F)); // Cancel button (20% of width)
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F)); // Save button (20% of width)

                    // Cancel button
                    var btnCancel = new Button
                    {
                        // Dock = DockStyle.Fill,
                        Text = "Cancel",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right
                    };
                    btnCancel.Click += (s, ev) => frmOptions.Close();

                    // Save button
                    var btnSave = new Button
                    {
                        // Dock = DockStyle.Fill,
                        Text = "Save",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
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

                    // Add controls to the layout
                    var emptySpacer = new Label { Dock = DockStyle.Fill }; // Spacer column
                    buttonsLayout.Controls.Add(emptySpacer, 0, 0); // Add spacer to the first column
                    buttonsLayout.Controls.Add(btnCancel, 1, 0); // Add Cancel button to the second column
                    buttonsLayout.Controls.Add(btnSave, 2, 0); // Add Save button to the third column

                    // Add the buttons layout to the main layout panel
                    layoutPanel.Controls.Add(buttonsLayout);
                    var donatePictureBox = new PictureBox
                    {
                        Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.donate.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Dock = DockStyle.Fill,
                        Height = 35,
                        Cursor = Cursors.Hand
                    };


                    var tempLabel1 = new Label
                    {
                        Text = "☘",
                        ForeColor = Color.FromArgb(250, 0, 200, 0),
                        Font = new Font("Tahoma", 12),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Anchor = AnchorStyles.Right
                    };
                    layoutPanel.Controls.Add(tempLabel1);


                    var tempLabel2 = new Label
                    {
                        //Text = "❀",
                        Text = " ",
                        ForeColor = Color.FromArgb(250, 210, 20, 120),
                        Font = new Font("Tahoma", 12),
                        TextAlign = ContentAlignment.MiddleRight,
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Anchor = AnchorStyles.Right
                    };
                    layoutPanel.Controls.Add(tempLabel2);

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

        public void ShowAboutForm()
        {
            try
            {
                using (var frmAbout = new Form())
                {
                    frmAbout.Text = "About Desktop Fences +";
                    frmAbout.Size = new Size(400, 580);
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
                        Height = 120
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


                    var label1Text = new Label
                    {
                        Text = " ",
                        Font = new Font("Tahoma", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(label1Text);

                    // Donation Logo
                    var donatePictureBox = new PictureBox
                    {
                        Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.donate.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Dock = DockStyle.Fill,
                        Height = 40, // Adjust based on donate.png size
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
                            System.Windows.Forms.MessageBox.Show($"Error opening donation link: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        private Icon GenerateIconWithNumber(int count)
        {
            // Base icon (default tray icon)
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            using (var baseIcon = Icon.ExtractAssociatedIcon(exePath))
            using (var bitmap = baseIcon.ToBitmap())
            using (var graphics = Graphics.FromImage(bitmap))
            {
                // Define the circle's size and position
                int circleDiameter = 24; // Diameter of the circle
                int circleX = -4; // X position (top-left corner)
                int circleY = -1; // Y position (top-left corner)

                // Draw the whiteish circle
                var circleBrush = new SolidBrush(Color.FromArgb(230, 255, 153, 53)); // Semi-transparent white
                graphics.FillEllipse(circleBrush, circleX, circleY, circleDiameter, circleDiameter);

                // Draw the number overlay
                var font = new Font("Calibri", 26, FontStyle.Bold, GraphicsUnit.Pixel);
                var textBrush = new SolidBrush(Color.Navy);

                // Measure the size of the text to center it within the circle
                string text = count.ToString();
                var textSize = graphics.MeasureString(text, font);
                float textX = circleX + (circleDiameter - textSize.Width) / 2;
                float textY = circleY + (circleDiameter - textSize.Height) / 2;

                graphics.DrawString(text, font, textBrush, textX, textY);

                // Convert the bitmap back to an icon
                return Icon.FromHandle(bitmap.GetHicon());
            }
        }
        private void UpdateTrayIcon()
        {
            if (HiddenFences.Count > 0)
            {
                // Generate an icon with the number of hidden fences
                _trayIcon.Icon = GenerateIconWithNumber(HiddenFences.Count);
            }
            else
            {
                // Revert to the default icon
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
                _trayIcon.Icon = Icon.ExtractAssociatedIcon(exePath);
            }
        }
    }
}


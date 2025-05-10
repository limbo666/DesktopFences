using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using System.Windows;
using Desktop_Fences;

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


        private System.Drawing.Color ConvertToDrawingColor(System.Windows.Media.Color mediaColor)
        {
            return System.Drawing.Color.FromArgb(mediaColor.A, mediaColor.R, mediaColor.G, mediaColor.B);
        }


       
        private void OnTrayIconDoubleClick(object sender, EventArgs e)
        {
            if (!_areFencesTempHidden)
            {
                // Hide all visible fences not already in HiddenFences
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
                Log($"Temporarily hid {visibleFences.Count} fences.");
            }
            else
            {
                // Restore all temp-hidden fences
                int count = _tempHiddenFences.Count;
                foreach (var fence in _tempHiddenFences)
                {
                    fence.Visibility = Visibility.Visible;
                }
                _tempHiddenFences.Clear();
                _areFencesTempHidden = false;
                Log($"Restored {count} temporarily hidden fences.");
            }
            UpdateTrayIcon(); // Optional: Update icon if needed
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



            // Inside InitializeTray(), after creating _trayIcon:
            _trayIcon.DoubleClick += OnTrayIconDoubleClick;

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("About", null, (s, e) => ShowAboutForm());
            trayMenu.Items.Add("Options", null, (s, e) => ShowOptionsForm());
           // trayMenu.Items.Add("Restore Backup...", null, (s, e) => RestoreBackup());
            // trayMenu.Items.Add("Diagnostics", null, (s, e) => ShowDiagnosticsForm());
            //trayMenu.Items.Add("testmsgbox", null, (s, e) =>
            

            trayMenu.Items.Add("Reload All Fences", null, async (s, e) =>
            {
                // Create and show the wait window
                var waitWindow = new Window
                {
                    Title = "Please Wait",
                    Content = new System.Windows.Controls.Label
                    {
                        Content = "Reloading all fences, please wait...",
                        HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                        VerticalAlignment = System.Windows.VerticalAlignment.Center,
                        FontFamily = new System.Windows.Media.FontFamily("Segoe UI"),
                        FontSize = 10
                    },
                    Width = 300,
                    Height = 150,
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    WindowStyle = WindowStyle.ToolWindow,
                    ResizeMode = ResizeMode.NoResize,
                    Topmost = true
                };

                waitWindow.Show();

                try
                {
                    await Task.Run(() =>
                    {
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            // Close all fence windows
                            foreach (var fence in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList())
                            {
                                fence.Close();
                            }

                            // Clear heart references
                          //  _heartTextBlocks.Clear();  // Add this line if exists in your code

                            // Reload fences
                            FenceManager.ReloadFences();
                        });
                    });
                }
                catch (Exception ex)
                {
                    // Updated the problematic line to use System.Windows.MessageBoxButton and System.Windows.MessageBoxImage
                    System.Windows.MessageBox.Show(
                        $"An error occurred while reloading fences: {ex.Message}",
                        "Error",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Error
                    );
                }
                finally
                {
                    // Close the wait window
                    waitWindow.Close();
                }
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
            UpdateTrayIcon(); // Ensure the tray icon is updated on initialization
        }

        // Add fence to hidden list and update tray menu
        public static void AddHiddenFence(NonActivatingWindow fence)
        {
            if (fence == null || string.IsNullOrEmpty(fence.Title)) return;
            // Ensure window is actually hidden
            fence.Dispatcher.Invoke(() =>
            {
                fence.Visibility = Visibility.Hidden;
            });

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
        //public static void ShowHiddenFence(string title)
        //{
        //    var hiddenFence = HiddenFences.FirstOrDefault(f => f.Title == title);
        //    if (hiddenFence != null)
        //    {
        //        hiddenFence.Window.Visibility = System.Windows.Visibility.Visible;
        //        var fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == title);
        //        if (fenceData != null)
        //        {
        //            FenceManager.UpdateFenceProperty(fenceData, "IsHidden", "false", $"Showed fence '{title}'");
        //        }
        //        HiddenFences.Remove(hiddenFence);
        //        Log($"Showed fence '{title}'");
        //        Instance?.UpdateHiddenFencesMenu();
        //        Instance?.UpdateTrayIcon(); // Update the tray icon
        //    }
        //}

        public static void ShowHiddenFence(string title)
        {
            var hiddenFence = HiddenFences.FirstOrDefault(f => f.Title == title);
            if (hiddenFence != null)
            {
                hiddenFence.Window.Dispatcher.Invoke(() =>
                {
                    hiddenFence.Window.Visibility = Visibility.Visible;
                    hiddenFence.Window.Activate(); // Ensure window is activated
                  //  hiddenFence.Window.Topmost = true; // Bring to front
                 //   hiddenFence.Window.Topmost = false; // Reset to allow normal behavior
                    hiddenFence.Window.Show();
                });

                var fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == title);
                if (fenceData != null)
                {
                    FenceManager.UpdateFenceProperty(fenceData, "IsHidden", "false", $"Showed fence '{title}'");
                }
                HiddenFences.Remove(hiddenFence);
                Log($"Showed fence '{title}'");
                Instance?.UpdateHiddenFencesMenu();
                Instance?.UpdateTrayIcon();
            }
        }

        // Update tray menu with hidden fences
        public void UpdateHiddenFencesMenu()
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
                    frmDiagnostics.Size = new System.Drawing.Size(300, 300);
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
                        Font = new Font("Segoe UI", 12, System.Drawing.FontStyle.Bold),
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



        private Color InvertColor(Color color)
        {
            return Color.FromArgb(color.A, 255 - color.R, 255 - color.G, 255 - color.B);
        }
        private Color SimilarColor(Color color, int offset)
        {
            // Define an offset for cycling (e.g., 50 for demonstration)
            //int offset = 100;

            // Calculate new color components, cycling within the 0-255 range
            int newR = (color.R + offset) % 256;
            int newG = (color.G + offset) % 256;
            int newB = (color.B + offset) % 256;

            // Return the new color with the same alpha value
            return Color.FromArgb(color.A, newR, newG, newB);
        }

        public bool ShowCustomMessageBoxForm()
        {
            bool result = false;

            try
            {
                using (var frmCustomMessageBox = new Form())
                {
                    frmCustomMessageBox.Text = "Confirm delete";
                    frmCustomMessageBox.Size = new System.Drawing.Size(350, 135); 
                    frmCustomMessageBox.StartPosition = FormStartPosition.CenterScreen;
                    frmCustomMessageBox.FormBorderStyle = FormBorderStyle.None;
                    frmCustomMessageBox.MaximizeBox = false;
                    frmCustomMessageBox.MinimizeBox = false;
                    frmCustomMessageBox.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);
                   // frmCustomMessageBox.BackColor = ColorTranslator.FromHtml(SettingsManager.SelectedColor);

                    // Get the color from Utility.GetColorFromName and convert it
                    string selectedColorName = SettingsManager.SelectedColor; // Example source of the color name
                    var mediaColor = Utility.GetColorFromName(selectedColorName);
                    var drawingColor = ConvertToDrawingColor(mediaColor);
                   // System.Windows.Forms.MessageBox.Show($"Color: {drawingColor}");
                   
                    frmCustomMessageBox.BackColor = drawingColor;

                    // Create header panel
                    var headerPanel = new Panel
                    {
                        Height = 25,
                        Dock = DockStyle.Top,
                        // BackColor = Color.DarkSlateBlue, 
                        BackColor = SimilarColor(frmCustomMessageBox.BackColor, 200),
                        Padding = new Padding(5, 0, 0, 0)
                    };

                    // Add title label to header
                    var lblTitle = new Label
                    {
                        Text = frmCustomMessageBox.Text,
                       // ForeColor = Color.White,
                       ForeColor =frmCustomMessageBox.BackColor,
                        Font = new Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleLeft
                    };
                    headerPanel.Controls.Add(lblTitle);

                    // Add header to form
                    frmCustomMessageBox.Controls.Add(headerPanel);

                    // Position the form on the monitor where the mouse is located
                    var mousePosition = Cursor.Position;
                    var screen = Screen.FromPoint(mousePosition);
                    frmCustomMessageBox.StartPosition = FormStartPosition.Manual;
                    frmCustomMessageBox.Location = new System.Drawing.Point(
                        screen.WorkingArea.Left + (screen.WorkingArea.Width - frmCustomMessageBox.Width) / 2,
                        screen.WorkingArea.Top + (screen.WorkingArea.Height - frmCustomMessageBox.Height) / 2
                    );

                    var layoutPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        RowCount = 2,
                        ColumnCount = 1,
                        Padding = new Padding(10),
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        Margin = new Padding(0, 25, 0, 0) // Add margin to account for header
                    };
                    
                    layoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 70F)); // Message
                    layoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30F)); // Buttons
                

                    var lblMessage = new Label
                    {
                        Text = "Are you sure you want to remove this fence?",
                        ForeColor = InvertColor(frmCustomMessageBox.BackColor),
                        Font = new Font("Segoe", 8),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = false
                    };
                    layoutPanel.Controls.Add(lblMessage, 0, 0);

                    var buttonsPanel = new FlowLayoutPanel
                    {
                        FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        WrapContents = false,
                      //  BackColor = SimilarColor(frmCustomMessageBox.BackColor, 100),
                        Anchor = AnchorStyles.Right
                     
                    };

                    var btnYes = new Button
                    {
                        Font = new Font("Segoe", 8, System.Drawing.FontStyle.Bold),
                        Text = "Yes",
                                            Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right,
                        BackColor = SimilarColor(frmCustomMessageBox.BackColor, 100)
                        //ForeColor = InvertColor(btnYes.BackColor)
                        //  BackColor = Color.FromArgb(250, 200, 200, 200)
                    };
                    btnYes.ForeColor = SimilarColor(btnYes.BackColor,140);
                    btnYes.Click += (s, ev) =>
                    {
                        result = true;
                        frmCustomMessageBox.Close();
                    };

                    var btnNo = new Button
                    {
                        Font = new Font("Segoe", 8),
                        Text = "No",
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right,
                        BackColor = SimilarColor(frmCustomMessageBox.BackColor, 100)
                        //  BackColor = Color.FromArgb(250, 200, 200, 200)
                    };
                    btnNo.ForeColor = SimilarColor(btnYes.BackColor, 140);
                    btnNo.Click += (s, ev) =>
                    {
                        result = false;
                        frmCustomMessageBox.Close();
                    };

                    buttonsPanel.Controls.Add(btnNo);
                    buttonsPanel.Controls.Add(btnYes);
                   

                    layoutPanel.Controls.Add(buttonsPanel, 0, 1);

                    frmCustomMessageBox.Controls.Add(layoutPanel);
                    frmCustomMessageBox.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                Log($"Error showing CustomMessageBox form: {ex.Message}");
              System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        public void ShowOptionsForm()
        {
            try
            {
                using (var frmOptions = new Form())
                {
                    frmOptions.Text = "Desktop Fences + Options";
                    frmOptions.Size = new System.Drawing.Size(320, 540);
                    frmOptions.StartPosition = FormStartPosition.CenterScreen;
                    frmOptions.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frmOptions.MaximizeBox = false;
                    frmOptions.MinimizeBox = false;
                    frmOptions.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

                    // Create a ToolTip component
                    var toolTip = new ToolTip
                    {
                        AutoPopDelay = 5000, // Time tooltip remains visible (ms)
                        InitialDelay = 500,  // Time before tooltip appears (ms)
                        ReshowDelay = 500,   // Time between subsequent tooltips (ms)
                        ShowAlways = true    // Show tooltips even if form is inactive
                    };

                    var layoutPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        Padding = new Padding(10)
                    };
                    layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                    // General GroupBox
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

                    // Set tooltip for chkStartWithWindows
                    toolTip.SetToolTip(chkStartWithWindows, "Enable to launch Desktop Fences + automatically when Windows starts.");

                    // Selections GroupBox
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
                        MaximumSize = new System.Drawing.Size(80, 0),
                        Anchor = AnchorStyles.Right
                    };
                    selectionsLayout.Controls.Add(lblTint, 0, 2);
                    selectionsLayout.Controls.Add(numTint, 1, 2);

                    var lblColor = new Label { Text = "Color", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
                    var cmbColor = new ComboBox
                    {
                        Dock = DockStyle.Fill,
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        MaximumSize = new System.Drawing.Size(80, 0),
                        Anchor = AnchorStyles.Right
                    };
                    cmbColor.Items.AddRange(new string[] { "Gray", "Black", "White", "Beige", "Green", "Purple", "Fuchsia", "Yellow", "Orange", "Red", "Blue", "Bismark" });
                    cmbColor.SelectedItem = SettingsManager.SelectedColor;
                    selectionsLayout.Controls.Add(lblColor, 0, 3);
                    selectionsLayout.Controls.Add(cmbColor, 1, 3);

                    var lblLaunchEffect = new Label { Text = "Launch Effect", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
                    var cmbLaunchEffect = new ComboBox
                    {
                        Dock = DockStyle.Fill,
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        MaximumSize = new System.Drawing.Size(80, 0),
                        Anchor = AnchorStyles.Right
                    };
                    cmbLaunchEffect.Items.AddRange(new string[] { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate", "GrowAndFly", "Pulse", "Elastic", "Flip3D", "Spiral" });
                    cmbLaunchEffect.SelectedIndex = (int)SettingsManager.LaunchEffect;
                    selectionsLayout.Controls.Add(lblLaunchEffect, 0, 4);
                    selectionsLayout.Controls.Add(cmbLaunchEffect, 1, 4);

                    groupBoxSelections.Controls.Add(selectionsLayout);
                    layoutPanel.Controls.Add(groupBoxSelections);

                    // Set tooltips for Selections controls
                    toolTip.SetToolTip(chkSingleClickToLaunch, "Enable to launch fence items with a single click instead of a double click.");
                    toolTip.SetToolTip(chkEnableSnap, "Enable to allow fences to snap to screen edges or other fences.");
                    toolTip.SetToolTip(numTint, "Adjust the tint level for fence backgrounds (1-100).");
                    toolTip.SetToolTip(cmbColor, "Select the background color for all fences.");
                    toolTip.SetToolTip(cmbLaunchEffect, "Choose an animation effect when launching fence items.");

                    // Tools GroupBox
                    var groupBoxTools = new GroupBox
                    {
                        Text = "Tools",
                        Dock = DockStyle.Top,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    var toolsLayout1 = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        ColumnCount = 5,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
                    toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
                    toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
                    toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
                    toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));

                    var btnBackup = new Button
                    {
                        Text = "Backup",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right
                    };
                    btnBackup.Click += (s, ev) => BackupManager.BackupData();
                    toolsLayout1.Controls.Add(btnBackup, 1, 0);

                    var btnRestore = new Button
                    {
                        Text = "Restore...",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Left
                    };
                    btnRestore.Click += (s, ev) => RestoreBackup();
                    toolsLayout1.Controls.Add(btnRestore, 3, 0);

                    var toolsLayout2 = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        ColumnCount = 2,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    toolsLayout2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));
                    toolsLayout2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));

                    var lblSpacer = new Label
                    {
                        Text = "----------------------------------------------------",
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    toolsLayout2.Controls.Add(lblSpacer, 0, 1);
                    toolsLayout2.SetColumnSpan(lblSpacer, 2);

                    var chkEnableLog = new CheckBox
                    {
                        Text = "Enable logging",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.IsLogEnabled
                    };
                    toolsLayout2.Controls.Add(chkEnableLog, 0, 2);

                    groupBoxTools.Controls.Add(toolsLayout2);
                    groupBoxTools.Controls.Add(toolsLayout1);
                    layoutPanel.Controls.Add(groupBoxTools);

                    // Set tooltips for Tools controls
                    toolTip.SetToolTip(btnBackup, "Create a backup of your fence settings and data.");
                    toolTip.SetToolTip(btnRestore, "Restore fence settings and data from a backup file.");
                    toolTip.SetToolTip(chkEnableLog, "Enable to log application events for troubleshooting.");

                    // Buttons Layout
                    var buttonsLayout = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        ColumnCount = 3,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));
                    buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F));

                    var btnCancel = new Button
                    {
                        Text = "Cancel",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right
                    };
                    btnCancel.Click += (s, ev) => frmOptions.Close();

                    var btnSave = new Button
                    {
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

                    var emptySpacer = new Label { Dock = DockStyle.Fill };
                    buttonsLayout.Controls.Add(emptySpacer, 0, 0);
                    buttonsLayout.Controls.Add(btnCancel, 1, 0);
                    buttonsLayout.Controls.Add(btnSave, 2, 0);

                    layoutPanel.Controls.Add(buttonsLayout);

                    // Set tooltips for buttons
                    toolTip.SetToolTip(btnCancel, "Close the options form without saving changes.");
                    toolTip.SetToolTip(btnSave, "Save changes and apply them to all fences.");

                    // Bottom Panel (Donation Section)
                    var bottomPanel = new Panel
                    {
                        Dock = DockStyle.Bottom,
                        Height = 80,
                        Padding = new Padding(10),
                        BackColor = SimilarColor(frmOptions.BackColor, 220)
                    };

                    var donateTablePanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        ColumnCount = 1,
                        RowCount = 2,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    donateTablePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 70F));
                    donateTablePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30F));

                    var donatePictureBox = new PictureBox
                    {
                        Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.donate.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Height = 55,
                        Width = 75,
                        Cursor = Cursors.Hand,
                        Anchor = AnchorStyles.Top
                    };

                    var donateLabel = new Label
                    {
                        Text = "Donate to help development",
                        Font = new Font("Segoe UI", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        AutoSize = true,
                        Anchor = AnchorStyles.Top
                    };

                    donateTablePanel.Controls.Add(donatePictureBox, 0, 0);
                    donateTablePanel.Controls.Add(donateLabel, 0, 1);
                    bottomPanel.Controls.Add(donateTablePanel);
                    layoutPanel.Controls.Add(bottomPanel);

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
                            System.Windows.Forms.MessageBox.Show($"Error opening donation link: {ex.Message}",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    };

                    // Set tooltips for donation controls
                    toolTip.SetToolTip(donatePictureBox, "Click to donate via PayPal and support Desktop Fences + development.");
                    toolTip.SetToolTip(donateLabel, "Click the image above to donate via PayPal and support Desktop Fences + development.");

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
        //public void ShowOptionsForm()
        //{
        //    try
        //    {
        //        using (var frmOptions = new Form())
        //        {
        //            frmOptions.Text = "Desktop Fences + Options";
        //            frmOptions.Size = new System.Drawing.Size(320, 540);
        //            frmOptions.StartPosition = FormStartPosition.CenterScreen;
        //            frmOptions.FormBorderStyle = FormBorderStyle.FixedDialog;
        //            frmOptions.MaximizeBox = false;
        //            frmOptions.MinimizeBox = false;
        //            frmOptions.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

        //            // Create a ToolTip component
        //            var toolTip = new ToolTip
        //            {
        //                AutoPopDelay = 5000, // Time tooltip remains visible (ms)
        //                InitialDelay = 500,  // Time before tooltip appears (ms)
        //                ReshowDelay = 500,   // Time between subsequent tooltips (ms)
        //                ShowAlways = true    // Show tooltips even if form is inactive
        //            };


        //            var layoutPanel = new TableLayoutPanel
        //            {
        //                Dock = DockStyle.Fill,
        //                ColumnCount = 1,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink,
        //                Padding = new Padding(10)
        //            };
        //            layoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        //            var groupBoxGeneral = new GroupBox
        //            {
        //                Text = "General",
        //                Dock = DockStyle.Top,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };
        //            var generalLayout = new TableLayoutPanel
        //            {
        //                Dock = DockStyle.Fill,
        //                ColumnCount = 1,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink,
        //                Padding = new Padding(5)
        //            };
        //            generalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        //            var chkStartWithWindows = new CheckBox
        //            {
        //                Text = "Start with Windows",
        //                Dock = DockStyle.Fill,
        //                AutoSize = true,
        //                Checked = IsStartWithWindows

        //            };
        //            chkStartWithWindows.CheckedChanged += (s, e) =>
        //            {
        //                try
        //                {
        //                    ToggleStartWithWindows(chkStartWithWindows.Checked);
        //                    Log($"Set Start with Windows to {chkStartWithWindows.Checked} via Options");
        //                }
        //                catch (Exception ex)
        //                {
        //                    Log($"Error setting Start with Windows: {ex.Message}");
        //                    System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Startup Toggle Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                    chkStartWithWindows.Checked = IsStartWithWindows;
        //                }
        //            };
        //            generalLayout.Controls.Add(chkStartWithWindows, 0, 0);
        //            groupBoxGeneral.Controls.Add(generalLayout);
        //            layoutPanel.Controls.Add(groupBoxGeneral);

        //            // Set tooltip for chkStartWithWindows
        //            toolTip.SetToolTip(chkStartWithWindows, "Enable to launch Desktop Fences automatically when Windows starts.");

        //            var groupBoxSelections = new GroupBox
        //            {
        //                Text = "Selections",
        //                Dock = DockStyle.Top,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };
        //            var selectionsLayout = new TableLayoutPanel
        //            {
        //                Dock = DockStyle.Fill,
        //                ColumnCount = 2,
        //                RowCount = 5,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };
        //            selectionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        //            selectionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        //            //var chkSingleClickToLaunch = new CheckBox
        //            //{
        //            //    Text = "Single Click to Launch",
        //            //    Dock = DockStyle.Fill,
        //            //    AutoSize = true,
        //            //    Checked = SettingsManager.SingleClickToLaunch
        //            //};
        //            //selectionsLayout.Controls.Add(chkSingleClickToLaunch, 0, 4);
        //            //selectionsLayout.SetColumnSpan(chkSingleClickToLaunch, 2);

        //            var chkEnableSnap = new CheckBox
        //            {
        //                Text = "Enable snap function",
        //                Dock = DockStyle.Fill,
        //                AutoSize = true,
        //                Checked = SettingsManager.IsSnapEnabled
        //            };
        //            selectionsLayout.Controls.Add(chkEnableSnap, 0, 1);
        //            selectionsLayout.SetColumnSpan(chkEnableSnap, 2);

        //            var chkSingleClickToLaunch = new CheckBox
        //            {
        //                Text = "Single Click to Launch",
        //                Dock = DockStyle.Fill,
        //                AutoSize = true,
        //                Checked = SettingsManager.SingleClickToLaunch
        //            };
        //            selectionsLayout.Controls.Add(chkSingleClickToLaunch, 0, 0);
        //            selectionsLayout.SetColumnSpan(chkSingleClickToLaunch, 2);

        //            var lblTint = new Label { Text = "Tint", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
        //            var numTint = new NumericUpDown
        //            {
        //                Maximum = 100,
        //                Minimum = 1,
        //                Value = SettingsManager.TintValue,
        //                Dock = DockStyle.Fill,
        //                MaximumSize = new System.Drawing.Size(80, 0), // Set maximum width to 150 pixels, height is unrestricted
        //                Anchor = AnchorStyles.Right
        //            };
        //            selectionsLayout.Controls.Add(lblTint, 0, 2);
        //            selectionsLayout.Controls.Add(numTint, 1, 2);

        //            var lblColor = new Label { Text = "Color", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
        //            var cmbColor = new ComboBox
        //            {
        //                Dock = DockStyle.Fill,
        //                DropDownStyle = ComboBoxStyle.DropDownList,
        //                MaximumSize = new System.Drawing.Size(80, 0), // Set maximum width to 150 pixels, height is unrestricted
        //                Anchor = AnchorStyles.Right
        //            };
        //            cmbColor.Items.AddRange(new string[] { "Gray", "Black", "White","Beige", "Green", "Purple","Fuchsia", "Yellow","Orange", "Red", "Blue", "Bismark" });

        //            cmbColor.SelectedItem = SettingsManager.SelectedColor;
        //            selectionsLayout.Controls.Add(lblColor, 0, 3);
        //            selectionsLayout.Controls.Add(cmbColor, 1, 3);

        //            var lblLaunchEffect = new Label { Text = "Launch Effect", Dock = DockStyle.Fill, AutoSize = true, Anchor = AnchorStyles.Right };
        //            var cmbLaunchEffect = new ComboBox
        //            {
        //                Dock = DockStyle.Fill,
        //                DropDownStyle = ComboBoxStyle.DropDownList,
        //                MaximumSize = new System.Drawing.Size(80, 0),// Set maximum width to 150 pixels, height is unrestricted

        //                Anchor = AnchorStyles.Right
        //            };
        //            cmbLaunchEffect.Items.AddRange(new string[] { "Zoom", "Bounce", "FadeOut", "SlideUp", "Rotate", "Agitate", "GrowAndFly", "Pulse", "Elastic", "Flip3D", "Spiral" });





        //            cmbLaunchEffect.SelectedIndex = (int)SettingsManager.LaunchEffect;
        //            selectionsLayout.Controls.Add(lblLaunchEffect, 0, 4);
        //            selectionsLayout.Controls.Add(cmbLaunchEffect, 1, 4);


        //            groupBoxSelections.Controls.Add(selectionsLayout);
        //            layoutPanel.Controls.Add(groupBoxSelections);

        //            var groupBoxTools = new GroupBox
        //            {
        //                Text = "Tools",
        //                Dock = DockStyle.Top,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };
        //            var toolsLayout1 = new TableLayoutPanel
        //            {
        //                Dock = DockStyle.Top,
        //                ColumnCount = 5,
        //                //RowCount=3,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };
        //            toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
        //            toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
        //            toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));
        //            toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
        //            toolsLayout1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10F));


        //            var btnBackup = new Button
        //            {
        //                Text = "Backup",
        //                AutoSize = false,
        //                Width = 85,
        //                Height = 25,
        //                Anchor = AnchorStyles.Right
        //            };
        //            btnBackup.Click += (s, ev) => BackupManager.BackupData();
        //            toolsLayout1.Controls.Add(btnBackup, 1, 0);

        //            var btnRestore = new Button
        //            {
        //                Text = "Restore...",
        //                AutoSize = false,
        //                Width = 85,
        //                Height = 25,
        //                Anchor = AnchorStyles.Left
        //            };
        //            btnRestore.Click += (s, ev)  => RestoreBackup();
        //            toolsLayout1.Controls.Add(btnRestore, 3, 0);



        //            var toolsLayout2 = new TableLayoutPanel
        //            {
        //                Dock = DockStyle.Top,
        //                ColumnCount = 2,
        //                //RowCount=3,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };
        //            toolsLayout2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));
        //            toolsLayout2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));

        //            var lblSpacer = new Label
        //            {
        //                Text = "----------------------------------------------------",
        //                Dock = DockStyle.Fill,
        //                AutoSize = true

        //            };
        //            toolsLayout2.Controls.Add(lblSpacer, 0,1 );
        //            toolsLayout2.SetColumnSpan(lblSpacer, 2);

        //            var chkEnableLog = new CheckBox
        //            {
        //                Text = "Enable logging",
        //                Dock = DockStyle.Fill,
        //                AutoSize = true,
        //                Checked = SettingsManager.IsLogEnabled
        //            };
        //            toolsLayout2.Controls.Add(chkEnableLog, 0, 2);


        //            groupBoxTools.Controls.Add(toolsLayout2);
        //            groupBoxTools.Controls.Add(toolsLayout1);


        //            layoutPanel.Controls.Add(groupBoxTools);

        //            var buttonsLayout = new TableLayoutPanel
        //            {
        //                Dock = DockStyle.Top,
        //                ColumnCount = 3, // Three columns: spacer, Cancel button, Save button
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };

        //            // Define column styles
        //            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F)); // Spacer column (60% of width)
        //            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F)); // Cancel button (20% of width)
        //            buttonsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25F)); // Save button (20% of width)

        //            // Cancel button
        //            var btnCancel = new Button
        //            {
        //                // Dock = DockStyle.Fill,
        //                Text = "Cancel",
        //                AutoSize = false,
        //                Width = 85,
        //                Height = 25,
        //                Anchor = AnchorStyles.Right
        //            };
        //            btnCancel.Click += (s, ev) => frmOptions.Close();

        //            // Save button
        //            var btnSave = new Button
        //            {
        //                // Dock = DockStyle.Fill,
        //                Text = "Save",
        //                AutoSize = false,
        //                Width = 85,
        //                Height = 25,
        //                Anchor = AnchorStyles.Right
        //            };
        //            btnSave.Click += (s, ev) =>
        //            {
        //                SettingsManager.IsSnapEnabled = chkEnableSnap.Checked;
        //                SettingsManager.TintValue = (int)numTint.Value;
        //                SettingsManager.SelectedColor = cmbColor.SelectedItem.ToString();
        //                SettingsManager.IsLogEnabled = chkEnableLog.Checked;
        //                SettingsManager.SingleClickToLaunch = chkSingleClickToLaunch.Checked;
        //                SettingsManager.LaunchEffect = (FenceManager.LaunchEffect)cmbLaunchEffect.SelectedIndex;
        //                IsStartWithWindows = chkStartWithWindows.Checked;

        //                SettingsManager.SaveSettings();
        //                FenceManager.UpdateOptionsAndClickEvents();

        //                foreach (var fence in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>())
        //                {
        //                    dynamic fenceData = FenceManager.GetFenceData().FirstOrDefault(f => f.Title == fence.Title);
        //                    if (fenceData != null)
        //                    {
        //                        string customColor = fenceData.CustomColor?.ToString();
        //                        string appliedColor = string.IsNullOrEmpty(customColor) ? SettingsManager.SelectedColor : customColor;
        //                        Utility.ApplyTintAndColorToFence(fence, appliedColor);
        //                        Log($"Applied color '{appliedColor}' with global tint '{SettingsManager.TintValue}' to fence '{fence.Title}'");
        //                    }
        //                    else
        //                    {
        //                        Utility.ApplyTintAndColorToFence(fence, SettingsManager.SelectedColor);
        //                        Log($"Applied global color '{SettingsManager.SelectedColor}' with global tint '{SettingsManager.TintValue}' to unknown fence '{fence.Title}'");
        //                    }
        //                }

        //                frmOptions.Close();
        //            };

        //            // Add controls to the layout
        //            var emptySpacer = new Label { Dock = DockStyle.Fill }; // Spacer column
        //            buttonsLayout.Controls.Add(emptySpacer, 0, 0); // Add spacer to the first column
        //            buttonsLayout.Controls.Add(btnCancel, 1, 0); // Add Cancel button to the second column
        //            buttonsLayout.Controls.Add(btnSave, 2, 0); // Add Save button to the third column

        //            // Add the buttons layout to the main layout panel
        //            layoutPanel.Controls.Add(buttonsLayout);

        //            // Replace your existing bottom panel and donation section with this code
        //            var bottomPanel = new Panel
        //            {
        //                Dock = DockStyle.Bottom,
        //                Height = 80,  // Adjusted height to better fit vertical layout
        //                Padding = new Padding(10),
        //                BackColor = SimilarColor(frmOptions.BackColor,220)
        //            };

        //            // Create a TableLayoutPanel for vertical alignment
        //            var donateTablePanel = new TableLayoutPanel
        //            {
        //                Dock = DockStyle.Fill,
        //                ColumnCount = 1,
        //                RowCount = 2,
        //                AutoSize = true,
        //                AutoSizeMode = AutoSizeMode.GrowAndShrink
        //            };


        //            donateTablePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 70F));
        //            donateTablePanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30F));


        //            // Create the donation picture box
        //            var donatePictureBox = new PictureBox
        //            {
        //                Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.donate.png"),
        //                SizeMode = PictureBoxSizeMode.Zoom,
        //                Height = 55,
        //                Width = 75,
        //                Cursor = Cursors.Hand,
        //                Anchor = AnchorStyles.Top  
        //            };

        //            // Create the donation label
        //            var donateLabel = new Label
        //            {
        //                Text = "Donate to help development",
        //                Font = new Font("Segoe UI", 9),
        //                TextAlign = ContentAlignment.MiddleCenter,
        //                AutoSize = true,
        //                Anchor = AnchorStyles.Top
        //            };

        //            // Add controls to table panel - image first, then label
        //            donateTablePanel.Controls.Add(donatePictureBox, 0, 0);
        //            donateTablePanel.Controls.Add(donateLabel, 0, 1);

        //            // Center the table panel in the bottom panel
        //            bottomPanel.Controls.Add(donateTablePanel);

        //            // Add bottom panel to the main layout
        //            layoutPanel.Controls.Add(bottomPanel);

        //            // Add the click handler for the donation image
        //            donatePictureBox.Click += (s, e) =>
        //            {
        //                try
        //                {
        //                    Process.Start(new ProcessStartInfo
        //                    {
        //                        FileName = "https://www.paypal.com/donate/?hosted_button_id=M8H4M4R763RBE",
        //                        UseShellExecute = true
        //                    });
        //                }
        //                catch (Exception ex)
        //                {
        //                    Log($"Error opening donation link: {ex.Message}");
        //                    System.Windows.Forms.MessageBox.Show($"Error opening donation link: {ex.Message}",
        //                                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                }
        //            };

        //            frmOptions.Controls.Add(layoutPanel);
        //            frmOptions.ShowDialog();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Log($"Error showing Options form: {ex.Message}");
        //        System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}

        public void ShowAboutForm()
        {
            try
            {
                using (var frmAbout = new Form())
                {
                    frmAbout.Text = "About Desktop Fences +";
                    frmAbout.Size = new System.Drawing.Size(400, 580);
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
                        Font = new Font("Segoe UI", 14, System.Drawing.FontStyle.Bold),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(labelTitle);

                    var version = Assembly.GetExecutingAssembly().GetName().Version;
                    var labelVersion = new Label
                    {
                        Text = $"ver {version}",
                        Font = new Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(labelVersion);

                    var labelMainText = new Label
                    {
                        Text = "Desktop Fences + is an open-source alternative to StarDock's Fences, originally created by HakanKokcu as Birdy Fences.\n\nDesktop fences +, is maintained by Nikos Georgousis, has been enhanced and optimized to give better user experience and stability.\n\n ",
                        Font = new Font("Segoe UI", 10),
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
                        Font = new Font("Segoe UI", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = true
                    };
                    layoutPanel.Controls.Add(labelGitHubText);

                    var linkLabelGitHub = new LinkLabel
                    {
                        Text = "https://github.com/limbo666/DesktopFences",
                        Font = new Font("Segoe UI", 9),
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
                        Font = new Font("Segoe UI", 9),
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
                        Dock = DockStyle.Bottom,
                        Height = 40, // Adjust based on donate.png size
                        Cursor = Cursors.Hand // Indicate clickability
                    }; var btnBackup = new Button
                    {
                        Text = "Backup",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right
                    };
                    //btnBackup.Click += (s, ev) => BackupManager.BackupData();
                    //toolsLayout1.Controls.Add(btnBackup, 3, 0);
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
                        Font = new Font("Segoe UI", 9),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Bottom,
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
                var font = new Font("Calibri", 26, System.Drawing.FontStyle.Bold, GraphicsUnit.Pixel);
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
        public void UpdateTrayIcon()
        {
            if (HiddenFences.Count > 0)
            {

                // Use an icon indicating hidden fences (existing + temp)
                _trayIcon.Icon = GenerateIconWithNumber(HiddenFences.Count + _tempHiddenFences.Count);

                //// Generate an icon with the number of hidden fences
                //_trayIcon.Icon = GenerateIconWithNumber(HiddenFences.Count);
            }
            else
            {
                // Revert to the default icon
                string exePath = Process.GetCurrentProcess().MainModule.FileName;
                _trayIcon.Icon = Icon.ExtractAssociatedIcon(exePath);
            }
        }


        private void  RestoreBackup()
        {
            var waitWindow = new Window
            {
                Title = "Please Wait",
                Content = new System.Windows.Controls.Label
                {
                    Content = "Restoring backup, please wait...",
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center
                },
                Width = 300,
                Height = 150,
                WindowStyle = WindowStyle.ToolWindow,
                ResizeMode = ResizeMode.NoResize,
                Topmost = true
            };

            // Get the program's root directory
            string rootDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            // Option 1: Open directly in Backups subfolder (if it exists)
            string backupsDir = Path.Combine(rootDir, "Backups");
            string initialPath = Directory.Exists(backupsDir) ? backupsDir : rootDir;


            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select backup folder to restore from",
                ShowNewFolderButton = false,
                   // Set initial directory to executable path
               //  RootFolder = Environment.SpecialFolder.MyComputer,
                SelectedPath = initialPath,  // This will focus the dialog on the desired folder
                
            };

       

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                waitWindow.Show();

                try
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        FenceManager.RestoreFromBackup(dialog.SelectedPath);
                        // Close all existing windows
                        foreach (var window in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList())
                        {
                            window.Close();
                        }
                        // Reload fences
                        FenceManager.LoadAndCreateFences(new TargetChecker(1000));
                    });
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Restore failed: {ex.Message}", "Error",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    waitWindow.Close();
                }
            }
        }

    }
}


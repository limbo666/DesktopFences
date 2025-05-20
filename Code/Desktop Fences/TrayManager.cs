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
using Microsoft.VisualBasic;
using System.Diagnostics.Eventing.Reader;
using System.Media;

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
                int count = _tempHiddenFences.Count;
                foreach (var fence in _tempHiddenFences)
                {
                    fence.Visibility = Visibility.Visible;
                }
                _tempHiddenFences.Clear();
                _areFencesTempHidden = false;
                Log($"Restored {count} temporarily hidden fences.");
            }
            UpdateTrayIcon();
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

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("About", null, (s, e) => ShowAboutForm());
            trayMenu.Items.Add("Options", null, (s, e) => ShowOptionsForm());
            trayMenu.Items.Add("Reload All Fences", null, async (s, e) =>
            {
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
                    //System.Windows.MessageBox.Show(
                    //    $"An error occurred while reloading fences: {ex.Message}",
                    //    "Error",
                    //    System.Windows.MessageBoxButton.OK,
                    //    System.Windows.MessageBoxImage.Error
                    //);
                    ShowOKOnlyMessageBoxForm($"An error occurred while reloading fences: {ex.Message}", "Error");
                }
                finally
                {
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
            UpdateTrayIcon();
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
                Log($"Added fence '{fence.Title}' to hidden list");
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
                Log($"Showed fence '{title}'");
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
                        RowCount = 5,
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };

                    infoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    infoLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

                    string[] variableNames = { "Variable1", "Variable2", "Variable3", "Variable4", "Variable5" };
                    string[] variableValues = { "Value1", "Value2", "Value3", "Value4", "Value5" };

                    for (int i = 0; i < 5; i++)
                    {
                        var lblVariable = new Label
                        {
                            Text = variableNames[i],
                            Dock = DockStyle.Fill,
                            AutoSize = true,
                        };
                        infoLayout.Controls.Add(lblVariable, 0, i);

                        var lblValue = new Label
                        {
                            Text = variableValues[i],
                            Dock = DockStyle.Fill,
                        };
                        infoLayout.Controls.Add(lblValue, 1, i);
                    }

                    groupBoxInfo.Controls.Add(infoLayout);
                    layoutPanel.Controls.Add(groupBoxInfo);

                    var blinkingLabel = new Label
                    {
                        Text = "Timed check",
                        Dock = DockStyle.Top,
                        AutoSize = true,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = new Font("Segoe UI", 12, System.Drawing.FontStyle.Bold),
                        ForeColor = Color.Red,
                        BackColor = Color.Transparent
                    };
                    layoutPanel.Controls.Add(blinkingLabel);

                    Timer blinkTimer = new Timer();
                    blinkTimer.Interval = 500;
                    bool isLabelVisible = true;

                    blinkTimer.Tick += (sender, e) =>
                    {
                        if (isLabelVisible)
                        {
                            blinkingLabel.ForeColor = Color.Red;
                        }
                        else
                        {
                            blinkingLabel.ForeColor = Color.Green;
                        }
                        isLabelVisible = !isLabelVisible;
                    };

                    frmDiagnostics.Shown += (sender, e) =>
                    {
                        blinkTimer.Start();
                    };

                    frmDiagnostics.FormClosing += (sender, e) =>
                    {
                        blinkTimer.Stop();
                        blinkTimer.Dispose();
                    };

                    frmDiagnostics.Controls.Add(layoutPanel);
                    frmDiagnostics.Show();
                }
            }
            catch (Exception ex)
            {
                Log($"Error showing Diagnostics form: {ex.Message}");
                // System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowOKOnlyMessageBoxForm($"An error occurred: {ex.Message}", "Error");
            }
        }

        private Color InvertColor(Color color)
        {
            return Color.FromArgb(color.A, 255 - color.R, 255 - color.G, 255 - color.B);
        }

        private Color SimilarColor(Color color, int offset)
        {
            int newR = (color.R + offset) % 256;
            int newG = (color.G + offset) % 256;
            int newB = (color.B + offset) % 256;
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

                    string selectedColorName = SettingsManager.SelectedColor;
                    var mediaColor = Utility.GetColorFromName(selectedColorName);
                    var drawingColor = ConvertToDrawingColor(mediaColor);
                    //frmCustomMessageBox.BackColor = drawingColor;
                    frmCustomMessageBox.BackColor = SimilarColor(drawingColor, 18);


                    var headerPanel = new Panel
                    {
                        Height = 25,
                        Dock = DockStyle.Top,
                        BackColor = SimilarColor(frmCustomMessageBox.BackColor, 200),
                        Padding = new Padding(5, 0, 0, 0)
                    };

                    var lblTitle = new Label
                    {
                        Text = frmCustomMessageBox.Text,
                        ForeColor = frmCustomMessageBox.BackColor,
                        Font = new Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleLeft
                    };
                    headerPanel.Controls.Add(lblTitle);

                    frmCustomMessageBox.Controls.Add(headerPanel);

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
                        Margin = new Padding(0, 25, 0, 0)
                    };

                    layoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 70F));
                    layoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30F));
                    //  System.Windows.Forms.MessageBox.Show($"Backcolor: {frmCustomMessageBox.BackColor}", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    var lblMessage = new Label
                    {
                        Text = "Are you sure you want to remove this fence?",
                        ForeColor = InvertColor(frmCustomMessageBox.BackColor),
                        Font = new Font("Segoe", 8),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = false
                    };

                    ///select costom for colors for specific colors
                    Color argbGray = Color.FromArgb(255, 110, 110, 110);
                    Color argbOrange = Color.FromArgb(255, 183, 84, 51);
                    Color argbBismark = Color.FromArgb(255, 73, 105, 126);
                    if (argbGray.ToArgb() == frmCustomMessageBox.BackColor.ToArgb())
                    {
                        // System.Windows.Forms.MessageBox.Show("Exact ARGB match.");
                        lblMessage.ForeColor = Color.FromArgb(255, 0, 0, 0);
                    }
                    else if (argbOrange.ToArgb() == frmCustomMessageBox.BackColor.ToArgb())
                    {
                        // System.Windows.Forms.MessageBox.Show("Exact ARGB match.");
                        lblMessage.ForeColor = Color.FromArgb(255, 0, 0, 120);
                    }
                    else if (argbBismark.ToArgb() == frmCustomMessageBox.BackColor.ToArgb())
                    {
                        //System.Windows.Forms.MessageBox.Show("Exact ARGB match.");
                        lblMessage.ForeColor = Color.FromArgb(255, 10, 0, 250);
                    }
                 
                    



               
                    layoutPanel.Controls.Add(lblMessage, 0, 0);

                    var buttonsPanel = new FlowLayoutPanel
                    {
                        FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        WrapContents = false,
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
                    };
                    btnYes.ForeColor = SimilarColor(btnYes.BackColor, 140);
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
                    // Play the ding.wav sound from embedded resource
                    try
                    {
                        using (Stream soundStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Desktop_Fences.Resources.ding.wav"))
                        {
                            if (soundStream != null)
                            {
                                using (SoundPlayer player = new SoundPlayer(soundStream))
                                {
                                    player.Play();
                                }
                            }
                            else
                            {
                                Log("Sound resource 'ding.wav' not found.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Error playing sound: {ex.Message}");
                    }
                    frmCustomMessageBox.TopMost = true;
                   
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



        public void ShowOKOnlyMessageBoxForm(string msgboxMessage, string msgboxTitle)
        {
           // bool result = false;

            try
            {
                using (var frmCustomOKOnlyBox = new Form())
                {
                    frmCustomOKOnlyBox.Text = msgboxTitle ;
                    frmCustomOKOnlyBox.Size = new System.Drawing.Size(350, 135);
                    frmCustomOKOnlyBox.StartPosition = FormStartPosition.CenterScreen;
                    frmCustomOKOnlyBox.FormBorderStyle = FormBorderStyle.None;
                    frmCustomOKOnlyBox.MaximizeBox = false;
                    frmCustomOKOnlyBox.MinimizeBox = false;
                    frmCustomOKOnlyBox.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

                    string selectedColorName = SettingsManager.SelectedColor;
                    var mediaColor = Utility.GetColorFromName(selectedColorName);
                    var drawingColor = ConvertToDrawingColor(mediaColor);
                    //frmCustomOKOnlyBox.BackColor = drawingColor;
                    frmCustomOKOnlyBox.BackColor = SimilarColor(drawingColor, 18); // a liitle different to stnadf out from the rest controls if appears above fences with tint to 100

                    var headerPanel = new Panel
                    {
                        Height = 25,
                        Dock = DockStyle.Top,
                        BackColor = SimilarColor(frmCustomOKOnlyBox.BackColor, 200),
                        Padding = new Padding(5, 0, 0, 0)
                    };

                    var lblTitle = new Label
                    {
                        Text = frmCustomOKOnlyBox.Text,
                        ForeColor = frmCustomOKOnlyBox.BackColor,
                        Font = new Font("Segoe UI", 9, System.Drawing.FontStyle.Bold),
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleLeft
                    };
                    headerPanel.Controls.Add(lblTitle);

                    frmCustomOKOnlyBox.Controls.Add(headerPanel);

                    var mousePosition = Cursor.Position;
                    var screen = Screen.FromPoint(mousePosition);
                    frmCustomOKOnlyBox.StartPosition = FormStartPosition.Manual;
                    frmCustomOKOnlyBox.Location = new System.Drawing.Point(
                        screen.WorkingArea.Left + (screen.WorkingArea.Width - frmCustomOKOnlyBox.Width) / 2,
                        screen.WorkingArea.Top + (screen.WorkingArea.Height - frmCustomOKOnlyBox.Height) / 2
                    );

                    var layoutPanel = new TableLayoutPanel
                    {
                        Dock = DockStyle.Fill,
                        RowCount = 2,
                        ColumnCount = 1,
                        Padding = new Padding(10),
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                        Margin = new Padding(0, 25, 0, 0)
                    };

                    layoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 70F));
                    layoutPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30F));

                    var lblMessage = new Label
                    {
                        Text = msgboxMessage,
                        ForeColor = InvertColor(frmCustomOKOnlyBox.BackColor),
                        Font = new Font("Segoe", 8),
                        TextAlign = ContentAlignment.MiddleCenter,
                        Dock = DockStyle.Fill,
                        AutoSize = false
                    };

                    ///select costom for colors for specific colors
                    Color argbGray = Color.FromArgb(255, 110, 110, 110);
                    Color argbOrange = Color.FromArgb(255, 183, 84, 51);
                    Color argbBismark = Color.FromArgb(255, 73, 105, 126);
                    if (argbGray.ToArgb() == frmCustomOKOnlyBox.BackColor.ToArgb())
                    {
                        // System.Windows.Forms.MessageBox.Show("Exact ARGB match.");
                        lblMessage.ForeColor = Color.FromArgb(255, 0, 0, 0);
                    }
                    else if (argbOrange.ToArgb() == frmCustomOKOnlyBox.BackColor.ToArgb())
                    {
                        // System.Windows.Forms.MessageBox.Show("Exact ARGB match.");
                        lblMessage.ForeColor = Color.FromArgb(255, 0, 0, 120);
                    }
                    else if (argbBismark.ToArgb() == frmCustomOKOnlyBox.BackColor.ToArgb())
                    {
                        //System.Windows.Forms.MessageBox.Show("Exact ARGB match.");
                        lblMessage.ForeColor = Color.FromArgb(255, 10, 0, 250);
                    }



                    layoutPanel.Controls.Add(lblMessage, 0, 0);

                    var buttonsPanel = new FlowLayoutPanel
                    {
                        FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft,
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        WrapContents = false,
                        Anchor = AnchorStyles.Right
                    };

                    var btnYes = new Button
                    {
                        Font = new Font("Segoe", 8, System.Drawing.FontStyle.Bold),
                        Text = "Yes",
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right,
                        BackColor = SimilarColor(frmCustomOKOnlyBox.BackColor, 100)
                    };
                    btnYes.ForeColor = SimilarColor(btnYes.BackColor, 140);
                    btnYes.Click += (s, ev) =>
                    {
                       // result = true;
                        frmCustomOKOnlyBox.Close();
                    };

                    var btnOK = new Button
                    {
                        Font = new Font("Segoe", 8),
                        Text = "OK",
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right,
                        BackColor = SimilarColor(frmCustomOKOnlyBox.BackColor, 100)
                    };
                    btnOK.ForeColor = SimilarColor(btnYes.BackColor, 140);
                    btnOK.Click += (s, ev) =>
                    {
                        //  result = false;
                        frmCustomOKOnlyBox.Close();
                    };

                    buttonsPanel.Controls.Add(btnOK);
                    // buttonsPanel.Controls.Add(btnYes);

                    layoutPanel.Controls.Add(buttonsPanel, 0, 1);
                                        frmCustomOKOnlyBox.Controls.Add(layoutPanel);

                    // Play the ding.wav sound from embedded resource
                    try
                    {
                        using (Stream soundStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Desktop_Fences.Resources.ding.wav"))
                        {
                            if (soundStream != null)
                            {
                                using (SoundPlayer player = new SoundPlayer(soundStream))
                                {
                                    player.Play();
                                }
                            }
                            else
                            {
                                Log("Sound resource 'ding.wav' not found.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"Error playing sound: {ex.Message}");
                    }
                    frmCustomOKOnlyBox.TopMost= true;
                    frmCustomOKOnlyBox.ShowDialog();
                  

                }
            }
            catch (Exception ex)
            {
                Log($"Error showing CustomOKOnlyMessageBox form: {ex.Message}");
                 System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //ShowOKOnlyMessageBoxForm($"An error occurred: {ex.Message}", "Error");

            }

           // return result;
        }



        public void ShowOptionsForm()
        {
            try
            {
                using (var frmOptions = new Form())
                {
                    frmOptions.Text = "Desktop Fences + Options";
                    frmOptions.Size = new System.Drawing.Size(320, 560);
                    frmOptions.StartPosition = FormStartPosition.CenterScreen;
                    frmOptions.FormBorderStyle = FormBorderStyle.FixedDialog;
                    frmOptions.MaximizeBox = false;
                    frmOptions.MinimizeBox = false;
                    frmOptions.AutoScaleMode = AutoScaleMode.Dpi;
                    frmOptions.AutoScaleDimensions = new SizeF(96F, 96F);
                    frmOptions.Icon = Icon.ExtractAssociatedIcon(Process.GetCurrentProcess().MainModule.FileName);

                    var toolTip = new ToolTip
                    {
                        AutoPopDelay = 5000,
                        InitialDelay = 500,
                        ReshowDelay = 500,
                        ShowAlways = true
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
                            //System.Windows.Forms.MessageBox.Show($"Error: {ex.Message}", "Startup Toggle Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ShowOKOnlyMessageBoxForm($"Error: {ex.Message}", "Startup Toggle Error");
                            chkStartWithWindows.Checked = IsStartWithWindows;
                        }
                    };
                    generalLayout.Controls.Add(chkStartWithWindows, 0, 0);
                    groupBoxGeneral.Controls.Add(generalLayout);
                    layoutPanel.Controls.Add(groupBoxGeneral);

                    toolTip.SetToolTip(chkStartWithWindows, "Enable to launch Desktop Fences + automatically when Windows starts.");

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
                        RowCount = 6, // Increased to accommodate new checkbox
                        AutoSize = true,
                        AutoSizeMode = AutoSizeMode.GrowAndShrink
                    };
                    selectionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
                    selectionsLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

                    var chkSingleClickToLaunch = new CheckBox
                    {
                        Text = "Single Click to Launch",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.SingleClickToLaunch
                    };
                    selectionsLayout.Controls.Add(chkSingleClickToLaunch, 0, 0);
                    selectionsLayout.SetColumnSpan(chkSingleClickToLaunch, 2);

                    var chkEnableSnap = new CheckBox
                    {
                        Text = "Enable Snap Near Fences",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.IsSnapEnabled
                    };
                    selectionsLayout.Controls.Add(chkEnableSnap, 0, 1);
                    selectionsLayout.SetColumnSpan(chkEnableSnap, 2);

                    var chkEnableDimensionSnap = new CheckBox
                    {
                        Text = "Enable Dimension Snap",
                        Dock = DockStyle.Fill,
                        AutoSize = true,
                        Checked = SettingsManager.EnableDimensionSnap
                    };
                    selectionsLayout.Controls.Add(chkEnableDimensionSnap, 0, 2);
                    selectionsLayout.SetColumnSpan(chkEnableDimensionSnap, 2);

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
                    selectionsLayout.Controls.Add(lblTint, 0, 3);
                    selectionsLayout.Controls.Add(numTint, 1, 3);

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
                    selectionsLayout.Controls.Add(lblColor, 0, 4);
                    selectionsLayout.Controls.Add(cmbColor, 1, 4);

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
                    selectionsLayout.Controls.Add(lblLaunchEffect, 0, 5);
                    selectionsLayout.Controls.Add(cmbLaunchEffect, 1, 5);

                    groupBoxSelections.Controls.Add(selectionsLayout);
                    layoutPanel.Controls.Add(groupBoxSelections);

                    toolTip.SetToolTip(chkSingleClickToLaunch, "Enable to launch fence items with a single click instead of a double click.");
                    toolTip.SetToolTip(chkEnableSnap, "Enable to allow fences to snap to screen edges or other fences.");
                    toolTip.SetToolTip(chkEnableDimensionSnap, "Enable to snap fence dimensions to the nearest multiple of 10 when resizing.");
                    toolTip.SetToolTip(numTint, "Adjust the tint level for fence backgrounds (1-100).");
                    toolTip.SetToolTip(cmbColor, "Select the background color for all fences.");
                    toolTip.SetToolTip(cmbLaunchEffect, "Choose an animation effect when launching fence items.");

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

                    toolTip.SetToolTip(btnBackup, "Create a backup of your fence settings and data.");
                    toolTip.SetToolTip(btnRestore, "Restore fence settings and data from a backup file.");
                    toolTip.SetToolTip(chkEnableLog, "Enable to log application events for troubleshooting.");

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
                        SettingsManager.EnableDimensionSnap = chkEnableDimensionSnap.Checked; // Save new setting
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

                    toolTip.SetToolTip(btnCancel, "Close the options form without saving changes.");
                    toolTip.SetToolTip(btnSave, "Save changes and apply them to all fences.");

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
                            // System.Windows.Forms.MessageBox.Show($"Error opening donation link: {ex.Message}",
                            //     "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ShowOKOnlyMessageBoxForm($"Error opening donation link: {ex.Message}", "Error");
                        }
                    };

                    toolTip.SetToolTip(donatePictureBox, "Click to donate via PayPal and support Desktop Fences + development.");
                    toolTip.SetToolTip(donateLabel, "Click the image above to donate via PayPal and support Desktop Fences + development.");

                    frmOptions.Controls.Add(layoutPanel);
                    frmOptions.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                Log($"Error showing Options form: {ex.Message}");
                //  System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowOKOnlyMessageBoxForm($"An error occurred: {ex.Message}", "Error");
            }
        }

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
                    frmAbout.AutoScaleMode = AutoScaleMode.Dpi;
                    frmAbout.AutoScaleDimensions = new SizeF(96F, 96F);
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
                            // System.Windows.Forms.MessageBox.Show($"Error opening GitHub link: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ShowOKOnlyMessageBoxForm($"Error opening GitHub link: {ex.Message}", "Error");
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

                    var donatePictureBox = new PictureBox
                    {
                        Image = Utility.LoadImageFromResources("Desktop_Fences.Resources.donate.png"),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Dock = DockStyle.Bottom,
                        Height = 40,
                        Cursor = Cursors.Hand
                    };
                    var btnBackup = new Button
                    {
                        Text = "Backup",
                        AutoSize = false,
                        Width = 85,
                        Height = 25,
                        Anchor = AnchorStyles.Right
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
                            //   System.Windows.Forms.MessageBox.Show($"Error opening donation link: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ShowOKOnlyMessageBoxForm($"Error opening donation link: {ex.Message}", "Error");
                        }
                    };
                    layoutPanel.Controls.Add(donatePictureBox);

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
                //  System.Windows.Forms.MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ShowOKOnlyMessageBoxForm($"An error occurred: {ex.Message}", "Error");
            }
        }

        private bool IsInStartupFolder()
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
            return File.Exists(shortcutPath);
        }

        //private void ToggleStartWithWindows(bool enable)
        //{
        //    string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        //    string shortcutPath = Path.Combine(startupPath, "Desktop Fences.lnk");
        //    string exePath = Process.GetCurrentProcess().MainModule.FileName;

        //    try
        //    {
        //        if (enable && !IsInStartupFolder())
        //        {
        //            Type shellType = Type.GetTypeFromProgID("WScript.Shell");
        //            dynamic shell = Activator.CreateInstance(shellType);
        //            var shortcut = shell.CreateShortcut(shortcutPath);
        //            shortcut.WorkingDirectory = Path.GetDirectoryName(exePath); 
        //            shortcut.TargetPath = exePath;
        //            shortcut.Description = "Desktop Fences Startup Shortcut";
        //            shortcut.Save();
        //            IsStartWithWindows = true;
        //            Log("Added Desktop Fences to Startup folder");
        //        }
        //        else if (!enable && IsInStartupFolder())
        //        {
        //            File.Delete(shortcutPath);
        //            IsStartWithWindows = false;
        //            Log("Removed Desktop Fences from Startup folder");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Log($"Failed to toggle Start with Windows: {ex.Message}");
        //        IsStartWithWindows = IsInStartupFolder();
        //        throw;
        //    }
        //}


        private void ToggleStartWithWindows(bool enable)
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
                    Log("Added Desktop Fences to Startup folder with working directory: " + workingDir);
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

        private void RestoreBackup()
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

            string rootDir = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            string backupsDir = Path.Combine(rootDir, "Backups");
            string initialPath = Directory.Exists(backupsDir) ? backupsDir : rootDir;

            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select backup folder to restore from",
                ShowNewFolderButton = false,
                SelectedPath = initialPath,
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                waitWindow.Show();

                try
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        FenceManager.RestoreFromBackup(dialog.SelectedPath);
                        foreach (var window in System.Windows.Application.Current.Windows.OfType<NonActivatingWindow>().ToList())
                        {
                            window.Close();
                        }
                        FenceManager.LoadAndCreateFences(new TargetChecker(1000));
                    });
                }
                catch (Exception ex)
                {
                    // System.Windows.MessageBox.Show($"Restore failed: {ex.Message}", "Error",
                    //                  MessageBoxButton.OK, MessageBoxImage.Error);
                    ShowOKOnlyMessageBoxForm($"Restore failed: {ex.Message}", "Error");
                }
                finally
                {
                    waitWindow.Close();
                }
            }
        }
    }
}
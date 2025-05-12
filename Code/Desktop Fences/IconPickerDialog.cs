using System;
using System.Windows;
using System.Windows.Controls;
using System.Drawing; // System.Drawing.Icon
using System.Runtime.InteropServices;
using System.Windows.Interop; // Imaging
using System.Windows.Media.Imaging; // BitmapSource
using System.Windows.Media; // ImageSource

public class IconPickerDialog : Window
{
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    static extern uint ExtractIconEx(string szFileName, int nIconIndex,
                                   IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

    [DllImport("user32.dll")]
    static extern bool DestroyIcon(IntPtr hIcon);

    public int SelectedIndex { get; private set; } = -1;
    private readonly string _filePath;

    public IconPickerDialog(string filePath)
    {
        _filePath = filePath;
        InitializeWindow();
        LoadIcons();
    }

    private void InitializeWindow()
    {
        Title = "Select Icon";
        Width = 400;
        Height = 300;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;

        var scrollViewer = new ScrollViewer();
        var wrapPanel = new WrapPanel();
        scrollViewer.Content = wrapPanel;
        Content = scrollViewer;
    }

    private void LoadIcons()
    {
        if (string.IsNullOrEmpty(_filePath)) return;
        if (string.IsNullOrEmpty(_filePath)) return;

        var iconCount = (int)ExtractIconEx(_filePath, -1, null, null, 0);
        if (iconCount == 0) return;

        var panel = ((ScrollViewer)Content).Content as WrapPanel;


        for (int i = 0; i < iconCount; i++)
        {
            IntPtr[] hIcon = new IntPtr[1];
            if (ExtractIconEx(_filePath, i, hIcon, null, 1) <= 0) continue;

            if (hIcon[0] != IntPtr.Zero)
            {
                try
                {
                    // Create bitmap directly from handle
                    var bitmap = Imaging.CreateBitmapSourceFromHIcon(
                        hIcon[0],
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions()
                    );

                    var button = new Button
                    {
                        Content = new System.Windows.Controls.Image
                        {
                            Source = bitmap,
                            Width = 32,
                            Height = 32
                        },
                        Tag = i, // Correct index storage
                        Margin = new Thickness(5)
                    };

                    button.Click += (s, e) =>
                    {
                        var clickedButton = (Button)s;
                        Console.WriteLine($"Clicked button Tag: {clickedButton.Tag}");
                        SelectedIndex = (int)clickedButton.Tag;
                        DialogResult = true;
                    };

                    panel.Children.Add(button);
                }
                finally
                {
                    // Destroy handle after creating bitmap
                    DestroyIcon(hIcon[0]);
                }
            }
        }
    }
}
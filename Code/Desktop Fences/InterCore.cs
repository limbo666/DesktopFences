using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Desktop_Fences
{

    // InterCore system for Desktop Fences - handles special interactive features and animations.
    // Provides enhanced user experience through dynamic visual effects and hidden functionalities.

    public static class InterCore
    {
        #region Private Fields

        // Simple key combination tracking
        private static readonly HashSet<Key> _currentlyPressed = new HashSet<Key>();
        private static DispatcherTimer _keyReleaseTimer;
        private static bool _isDanceActive = false;
        private static bool _isGravityActive = false;
        private static readonly Dictionary<StackPanel, Point> _originalIconPositions = new Dictionary<StackPanel, Point>();
        private static Window _sparkleOverlay;

        // Global key hook variables
        private static GlobalKeyboardHook _globalKeyHook;

        #endregion

        #region Global Keyboard Hook

        private class GlobalKeyboardHook
        {
            private const int WH_KEYBOARD_LL = 13;
            private const int HC_ACTION = 0;
            private const int WM_KEYDOWN = 0x0100;
            private const int WM_SYSKEYDOWN = 0x0104;

            private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
            private LowLevelKeyboardProc _proc = HookCallback;
            private IntPtr _hookID = IntPtr.Zero;

            public delegate void KeyPressedEventHandler(Key key);
            public static event KeyPressedEventHandler KeyPressed;

            public GlobalKeyboardHook()
            {
                _hookID = SetHook(_proc);
            }

            public void Dispose()
            {
                UnhookWindowsHookEx(_hookID);
            }

            private IntPtr SetHook(LowLevelKeyboardProc proc)
            {
                using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
                using (var curModule = curProcess.MainModule)
                {
                    return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                        GetModuleHandle(curModule.ModuleName), 0);
                }
            }

            private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
            {
                if (nCode >= HC_ACTION && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    Key key = KeyInterop.KeyFromVirtualKey(vkCode);
                    KeyPressed?.Invoke(key);
                }
                return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
            }

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr SetWindowsHookEx(int idHook,
                LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
                IntPtr wParam, IntPtr lParam);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr GetModuleHandle(string lpModuleName);
        }

        #endregion

        #region Public Methods


        // Initializes the InterCore system - call this during application startup

        public static void Initialize()
        {
            try
            {
                // Set up key release timer for combination detection
                _keyReleaseTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(500) // Reset after 500ms of no key presses
                };
                _keyReleaseTimer.Tick += (s, e) =>
                {
                    _currentlyPressed.Clear();
                    _keyReleaseTimer.Stop();
                };

                // Set up global keyboard hook
                _globalKeyHook = new GlobalKeyboardHook();
                GlobalKeyboardHook.KeyPressed += OnGlobalKeyPressed;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore system initialized with global key hook");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error during initialization: {ex.Message}");
            }
        }


        // Processes fence title changes for special triggers

        // <param name="fence">The fence object being renamed</param>
        // <param name="newTitle">The new title</param>
        // <param name="originalTitle">The original title before change</param>
        // <returns>The final title to use (may be reverted for special triggers)</returns>
        public static string ProcessTitleChange(dynamic fence, string newTitle, string originalTitle)
        {
            try
            {
                // Check for limbo666 trigger
                if (string.Equals(newTitle, "limbo666", StringComparison.OrdinalIgnoreCase))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: limbo666 trigger activated");
                    ActivateSparkleEffect();
                    return originalTitle; // Revert to original title
                }

                return newTitle; // No special trigger, use new title
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error processing title change: {ex.Message}");
                return newTitle;
            }
        }


        // Cleans up any active effects - call this during application shutdown

        public static void Cleanup()
        {
            try
            {
                _keyReleaseTimer?.Stop();
                _sparkleOverlay?.Close();
                _originalIconPositions.Clear();

                // Clean up global key hook
                _globalKeyHook?.Dispose();
                _globalKeyHook = null;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore system cleaned up");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error during cleanup: {ex.Message}");
            }
        }

        #endregion

        #region Private Methods - Global Key Handling


        // Global key press handler for all keyboard input

        private static void OnGlobalKeyPressed(Key key)
        {
            try
            {
                // Only process if Desktop Fences is the active application or has fences visible
                var currentApp = Application.Current;
                if (currentApp == null) return;

                var fenceWindows = currentApp.Windows.OfType<NonActivatingWindow>().ToList();
                if (!fenceWindows.Any()) return;

                // Process the key input
                ProcessKeyInput(key);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error in global key handler: {ex.Message}");
            }
        }


        // Processes keyboard input for special sequences and commands

        // <param name="key">The pressed key</param>
        private static void ProcessKeyInput(Key key)
        {
            try
            {
                // Add to currently pressed keys
                _currentlyPressed.Add(key);

                // Reset the timer
                _keyReleaseTimer.Stop();
                _keyReleaseTimer.Start();

                // Check for Dance Party trigger (Ctrl+Alt+D)
                if (key == Key.D &&
                    _currentlyPressed.Contains(Key.LeftCtrl) &&
                    _currentlyPressed.Contains(Key.LeftAlt))
                {
                    ActivateDanceParty();
                }

                // Handle Gravity Drop shortcut (Ctrl+Shift+G)
                if (key == Key.G &&
                    (_currentlyPressed.Contains(Key.LeftCtrl) || _currentlyPressed.Contains(Key.RightCtrl)) &&
                    (_currentlyPressed.Contains(Key.LeftShift) || _currentlyPressed.Contains(Key.RightShift)))
                {
                    ActivateGravityDrop();
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error processing key input: {ex.Message}");
            }
        }

        #endregion

        //#region Private Methods - Dance Party

        //private static void ActivateDanceParty()
        //{
        //    if (_isDanceActive) return;

        //    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: Dance Party activated!");
        //    _isDanceActive = true;

        //    try
        //    {
        //        // Get all fence icons and make them dance
        //        var fenceWindows = Application.Current.Windows.OfType<NonActivatingWindow>();
        //        var allIcons = new List<StackPanel>();

        //        foreach (var window in fenceWindows)
        //        {
        //            var wrapPanel = FindWrapPanel(window);
        //            if (wrapPanel != null)
        //            {
        //                allIcons.AddRange(wrapPanel.Children.OfType<StackPanel>());
        //            }
        //        }

        //        // Create bounce animation for each icon
        //        foreach (var icon in allIcons)
        //        {
        //            var bounceAnimation = new DoubleAnimationUsingKeyFrames();

        //            // Add keyframes with individual easing
        //            var easing = new BounceEase { EasingMode = EasingMode.EaseOut };
        //            bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
        //            bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(-20, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))) { EasingFunction = easing });
        //            bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))) { EasingFunction = easing });
        //            bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(-15, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.9))) { EasingFunction = easing });
        //            bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(1.2))) { EasingFunction = easing });

        //            bounceAnimation.RepeatBehavior = new RepeatBehavior(TimeSpan.FromSeconds(10));

        //            // Create transform if it doesn't exist
        //            if (icon.RenderTransform == null || icon.RenderTransform == Transform.Identity)
        //            {
        //                icon.RenderTransform = new TranslateTransform();
        //            }

        //            var transform = icon.RenderTransform as TranslateTransform ?? new TranslateTransform();
        //            icon.RenderTransform = transform;

        //            transform.BeginAnimation(TranslateTransform.YProperty, bounceAnimation);
        //        }

        //        // Reset flag after animation completes
        //        var resetTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
        //        resetTimer.Tick += (s, e) =>
        //        {
        //            _isDanceActive = false;
        //            resetTimer.Stop();
        //            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "InterCore: Dance Party effect ended");
        //        };
        //        resetTimer.Start();
        //    }
        //    catch (Exception ex)
        //    {
        //        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error in Dance Party activation: {ex.Message}");
        //        _isDanceActive = false;
        //    }
        //}

        //#endregion

        #region Private Methods - Dance Party

        private static void ActivateDanceParty()
        {
            if (_isDanceActive) return;

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: Dance Party activated!");
            _isDanceActive = true;

            try
            {
                // Play MIDI music (simple generated tune)
                // PlayDanceMusic();
                PlayHappyTune2();

                // Get all fence icons and make them dance
                var fenceWindows = Application.Current.Windows.OfType<NonActivatingWindow>();
                var allIcons = new List<StackPanel>();

                foreach (var window in fenceWindows)
                {
                    var wrapPanel = FindWrapPanel(window);
                    if (wrapPanel != null)
                    {
                        allIcons.AddRange(wrapPanel.Children.OfType<StackPanel>());
                    }
                }

                // Create bounce animation for each icon
                foreach (var icon in allIcons)
                {
                    var bounceAnimation = new DoubleAnimationUsingKeyFrames();

                    // Add keyframes with individual easing
                    var easing = new BounceEase { EasingMode = EasingMode.EaseOut };
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(-20, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))) { EasingFunction = easing });
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))) { EasingFunction = easing });
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(-15, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.9))) { EasingFunction = easing });
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(1.2))) { EasingFunction = easing });

                    bounceAnimation.RepeatBehavior = new RepeatBehavior(TimeSpan.FromSeconds(10));

                    // Create transform if it doesn't exist
                    if (icon.RenderTransform == null || icon.RenderTransform == Transform.Identity)
                    {
                        icon.RenderTransform = new TranslateTransform();
                    }

                    var transform = icon.RenderTransform as TranslateTransform ?? new TranslateTransform();
                    icon.RenderTransform = transform;

                    transform.BeginAnimation(TranslateTransform.YProperty, bounceAnimation);
                }

                // Reset flag after animation completes
                var resetTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
                resetTimer.Tick += (s, e) =>
                {
                    _isDanceActive = false;
                    resetTimer.Stop();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "InterCore: Dance Party effect ended");
                };
                resetTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error in Dance Party activation: {ex.Message}");
                _isDanceActive = false;
            }
        }

        #endregion


        private static void PlayHappyTune2()
        {
            try
            {
                var midiOut = new NAudio.Midi.MidiOut(0);
                midiOut.Send(NAudio.Midi.MidiMessage.ChangePatch(12, 1).RawData); //12 Marimba // 24 Pianio

                var thread = new Thread(() =>
                {
                    // Chord progression: C - G - Am - F
                    int[][] chords = {
                new[] { 60, 64, 67 }, // C
                new[] { 67, 71, 74 }, // G
                new[] { 69, 72, 76 }, // Am
                new[] { 65, 69, 72 }  // F
            };

                    for (int i = 0; i < 14; i++) // iterations
                    {
                        // Play chord
                        foreach (var note in chords[i % chords.Length])
                        {
                            midiOut.Send(NAudio.Midi.MidiMessage.StartNote(note, 90, 1).RawData);
                        }
                        Thread.Sleep(300);

                        // Staccato rhythm
                        for (int j = 0; j < 4; j++)
                        {
                            int rootNote = chords[i % chords.Length][0];
                            midiOut.Send(NAudio.Midi.MidiMessage.StopNote(rootNote, 0, 1).RawData);
                            midiOut.Send(NAudio.Midi.MidiMessage.StartNote(rootNote + (j % 2 == 0 ? 0 : 7), 110, 1).RawData);
                            Thread.Sleep(100);
                        }

                        // Release chord
                        foreach (var note in chords[i % chords.Length])
                        {
                            midiOut.Send(NAudio.Midi.MidiMessage.StopNote(note, 0, 1).RawData);
                        }
                    }

                    midiOut.Dispose();
                });

                thread.IsBackground = true;
                thread.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                    $"InterCore: Couldn't play happy tune 2: {ex.Message}");
            }
        }


        #region Private Methods - Gravity Drop

        private static void ActivateGravityDrop()
        {
            if (_isGravityActive) return;

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: Gravity Drop activated!");
            _isGravityActive = true;

            try
            {
                // Get all fence icons
                var fenceWindows = Application.Current.Windows.OfType<NonActivatingWindow>();
                var allIcons = new List<StackPanel>();
                _originalIconPositions.Clear();

                foreach (var window in fenceWindows)
                {
                    var wrapPanel = FindWrapPanel(window);
                    if (wrapPanel != null)
                    {
                        var icons = wrapPanel.Children.OfType<StackPanel>().ToList();
                        allIcons.AddRange(icons);

                        // Store original positions
                        foreach (var icon in icons)
                        {
                            _originalIconPositions[icon] = new Point(icon.Margin.Left, icon.Margin.Top);
                        }
                    }
                }

                // Create gravity effect for each icon
                foreach (var icon in allIcons)
                {
                    var random = new Random();
                    var fallDistance = 500 + random.Next(100); // Random fall distance
                    var fallDuration = 1.5 + random.NextDouble() * 0.5; // Random fall speed

                    // Create fall animation
                    var fallAnimation = new DoubleAnimation
                    {
                        From = 0,
                        To = fallDistance,
                        Duration = TimeSpan.FromSeconds(fallDuration),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
                    };

                    // Create bounce back animation
                    var bounceAnimation = new DoubleAnimation
                    {
                        From = fallDistance,
                        To = 0,
                        Duration = TimeSpan.FromSeconds(0.8),
                        EasingFunction = new BounceEase { EasingMode = EasingMode.EaseOut, Bounces = 3 },
                        BeginTime = TimeSpan.FromSeconds(fallDuration)
                    };

                    // Create transform if it doesn't exist
                    if (icon.RenderTransform == null || icon.RenderTransform == Transform.Identity)
                    {
                        icon.RenderTransform = new TranslateTransform();
                    }

                    var transform = icon.RenderTransform as TranslateTransform ?? new TranslateTransform();
                    icon.RenderTransform = transform;

                    // Start fall animation
                    transform.BeginAnimation(TranslateTransform.YProperty, fallAnimation);

                    // Queue bounce animation
                    var bounceTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(fallDuration) };
                    bounceTimer.Tick += (s, e) =>
                    {
                        transform.BeginAnimation(TranslateTransform.YProperty, bounceAnimation);
                        bounceTimer.Stop();
                    };
                    bounceTimer.Start();
                }

                // Reset flag after all animations complete
                var resetTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(4) };
                resetTimer.Tick += (s, e) =>
                {
                    _isGravityActive = false;
                    _originalIconPositions.Clear();
                    resetTimer.Stop();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "InterCore: Gravity Drop effect ended");
                };
                resetTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error in Gravity Drop activation: {ex.Message}");
                _isGravityActive = false;
            }
        }

        #endregion

        #region Private Methods - Epic Fireworks (limbo666)

        private static void ActivateSparkleEffect()
        {
            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: Epic Fireworks effect activated for limbo666!");

            try
            {
                // Create fullscreen overlay window
                _sparkleOverlay = new Window
                {
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = Brushes.Transparent,
                    Topmost = true,
                    ShowInTaskbar = false,
                    WindowState = WindowState.Maximized,
                    IsHitTestVisible = false // Allow clicks to pass through
                };

                var canvas = new Canvas();
                _sparkleOverlay.Content = canvas;

                // Create multiple firework launching points
                var random = new Random();
                var fireworkCount = 32; // Number of fireworks to launch

                // Launch fireworks from bottom of screen at different times
                for (int i = 0; i < fireworkCount; i++)
                {
                    var delay = TimeSpan.FromMilliseconds(random.Next(0, 10000)); // Spread over 10 seconds
                    var launchTimer = new DispatcherTimer { Interval = delay };

                    launchTimer.Tick += (s, e) =>
                    {
                        LaunchFirework(canvas, random);
                        ((DispatcherTimer)s).Stop();
                    };
                    launchTimer.Start();
                }

                _sparkleOverlay.Show();

                // Close overlay after 8 seconds (longer for fireworks show)
                var closeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(11) };
                closeTimer.Tick += (s, e) =>
                {
                    _sparkleOverlay?.Close();
                    _sparkleOverlay = null;
                    closeTimer.Stop();
                    LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI, "InterCore: Epic Fireworks show ended");
                };
                closeTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error in Epic Fireworks activation: {ex.Message}");
                _sparkleOverlay?.Close();
                _sparkleOverlay = null;
            }
        }

        private static void LaunchFirework(Canvas canvas, Random random)
        {
            // Random launch position at bottom of screen
            var launchX = random.Next(100, (int)SystemParameters.PrimaryScreenWidth - 100);
            var launchY = (int)SystemParameters.PrimaryScreenHeight - 50;

            // Random explosion position in upper area
            var explodeX = launchX + random.Next(-200, 200);
            var explodeY = random.Next(100, (int)SystemParameters.PrimaryScreenHeight / 2);

            // Create rocket trail
            CreateRocketTrail(canvas, launchX, launchY, explodeX, explodeY, random);

            // Schedule explosion after rocket travel time
            var travelTime = 1.0 + random.NextDouble() * 0.5; // 1-1.5 seconds
            var explodeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(travelTime) };
            explodeTimer.Tick += (s, e) =>
            {
                CreateFireworkExplosion(canvas, explodeX, explodeY, random);
                ((DispatcherTimer)s).Stop();
            };
            explodeTimer.Start();
        }

        private static void CreateRocketTrail(Canvas canvas, double startX, double startY, double endX, double endY, Random random)
        {
            // Create rocket particle
            var rocket = new Ellipse
            {
                Width = 4,
                Height = 8,
                Fill = new SolidColorBrush(Colors.Orange),
                Effect = new DropShadowEffect
                {
                    Color = Colors.Yellow,
                    BlurRadius = 8,
                    ShadowDepth = 0
                }
            };

            Canvas.SetLeft(rocket, startX);
            Canvas.SetTop(rocket, startY);
            canvas.Children.Add(rocket);

            // Create rocket trail animation
            var duration = TimeSpan.FromSeconds(1.0 + random.NextDouble() * 0.5);
            var moveXAnimation = new DoubleAnimation
            {
                From = startX,
                To = endX,
                Duration = duration,
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            var moveYAnimation = new DoubleAnimation
            {
                From = startY,
                To = endY,
                Duration = duration,
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
            };

            // Add trailing sparks
            CreateTrail(canvas, startX, startY, endX, endY, duration.TotalSeconds);

            rocket.BeginAnimation(Canvas.LeftProperty, moveXAnimation);
            rocket.BeginAnimation(Canvas.TopProperty, moveYAnimation);

            // Remove rocket after animation
            var removeTimer = new DispatcherTimer { Interval = duration };
            removeTimer.Tick += (s, e) =>
            {
                canvas.Children.Remove(rocket);
                ((DispatcherTimer)s).Stop();
            };
            removeTimer.Start();
        }

        private static void CreateTrail(Canvas canvas, double startX, double startY, double endX, double endY, double duration)
        {
            var random = new Random();
            var trailParticles = 15;

            for (int i = 0; i < trailParticles; i++)
            {
                var delay = (duration / trailParticles) * i;
                var progress = (double)i / trailParticles;

                var particleX = startX + (endX - startX) * progress;
                var particleY = startY + (endY - startY) * progress;

                var trailTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(delay) };
                trailTimer.Tick += (s, e) =>
                {
                    var trail = new Ellipse
                    {
                        Width = 2,
                        Height = 2,
                        Fill = new SolidColorBrush(Color.FromArgb(150, 255, 165, 0)), // Semi-transparent orange
                        Effect = new BlurEffect { Radius = 1 }
                    };

                    Canvas.SetLeft(trail, particleX + random.Next(-3, 3));
                    Canvas.SetTop(trail, particleY + random.Next(-3, 3));
                    canvas.Children.Add(trail);

                    // Fade out trail particle
                    var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.8));
                    trail.BeginAnimation(UIElement.OpacityProperty, fadeOut);

                    // Remove trail particle
                    var removeTrail = new DispatcherTimer { Interval = TimeSpan.FromSeconds(0.8) };
                    removeTrail.Tick += (s2, e2) =>
                    {
                        canvas.Children.Remove(trail);
                        ((DispatcherTimer)s2).Stop();
                    };
                    removeTrail.Start();

                    ((DispatcherTimer)s).Stop();
                };
                trailTimer.Start();
            }
        }

        private static void CreateFireworkExplosion(Canvas canvas, double centerX, double centerY, Random random)
        {
            // Choose explosion type
            var explosionTypes = new[] { "Burst", "Ring", "Willow" };
            var explosionType = explosionTypes[random.Next(explosionTypes.Length)];

            // Choose color scheme
            var colorSchemes = new[]
            {
                new[] { Colors.Red, Colors.Orange, Colors.Yellow },
                new[] { Colors.Blue, Colors.Cyan, Colors.White },
                new[] { Colors.Green, Colors.Lime, Colors.Yellow },
                new[] { Colors.Purple, Colors.Magenta, Colors.Pink },
                new[] { Colors.Gold, Colors.Orange, Colors.White }
            };
            var colors = colorSchemes[random.Next(colorSchemes.Length)];

            switch (explosionType)
            {
                case "Burst":
                    CreateBurstExplosion(canvas, centerX, centerY, colors, random);
                    break;
                case "Ring":
                    CreateRingExplosion(canvas, centerX, centerY, colors, random);
                    break;
                case "Willow":
                    CreateWillowExplosion(canvas, centerX, centerY, colors, random);
                    break;
            }
        }

        private static void CreateBurstExplosion(Canvas canvas, double centerX, double centerY, Color[] colors, Random random)
        {
            var particleCount = 60 + random.Next(40); // 60-100 particles

            for (int i = 0; i < particleCount; i++)
            {
                var angle = (2 * Math.PI * i) / particleCount + random.NextDouble() * 0.5; // Add randomness
                var velocity = 80 + random.Next(120); // Random velocity
                var size = 3 + random.Next(5);

                var particle = new Ellipse
                {
                    Width = size,
                    Height = size,
                    Fill = new SolidColorBrush(colors[random.Next(colors.Length)]),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.White,
                        BlurRadius = size * 2,
                        ShadowDepth = 0
                    }
                };

                Canvas.SetLeft(particle, centerX);
                Canvas.SetTop(particle, centerY);
                canvas.Children.Add(particle);

                // Calculate end position
                var endX = centerX + Math.Cos(angle) * velocity;
                var endY = centerY + Math.Sin(angle) * velocity;

                // Movement animation with gravity
                var moveXAnimation = new DoubleAnimation
                {
                    From = centerX,
                    To = endX,
                    Duration = TimeSpan.FromSeconds(2.5),
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                };

                var moveYAnimation = new DoubleAnimation
                {
                    From = centerY,
                    To = endY + 100, // Add gravity effect
                    Duration = TimeSpan.FromSeconds(2.5),
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
                };

                // Fade out animation
                var fadeAnimation = new DoubleAnimation
                {
                    From = 1.0,
                    To = 0,
                    Duration = TimeSpan.FromSeconds(2.5),
                    BeginTime = TimeSpan.FromSeconds(0.3)
                };

                particle.BeginAnimation(Canvas.LeftProperty, moveXAnimation);
                particle.BeginAnimation(Canvas.TopProperty, moveYAnimation);
                particle.BeginAnimation(UIElement.OpacityProperty, fadeAnimation);

                // Remove particle after animation
                var removeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
                removeTimer.Tick += (s, e) =>
                {
                    canvas.Children.Remove(particle);
                    ((DispatcherTimer)s).Stop();
                };
                removeTimer.Start();
            }
        }

        private static void CreateRingExplosion(Canvas canvas, double centerX, double centerY, Color[] colors, Random random)
        {
            var particleCount = 36; // Perfect circle
            var radius = 120 + random.Next(80);

            for (int i = 0; i < particleCount; i++)
            {
                var angle = (2 * Math.PI * i) / particleCount;
                var size = 4 + random.Next(3);

                var particle = new Ellipse
                {
                    Width = size,
                    Height = size,
                    Fill = new SolidColorBrush(colors[random.Next(colors.Length)]),
                    Effect = new DropShadowEffect
                    {
                        Color = Colors.White,
                        BlurRadius = 10,
                        ShadowDepth = 0
                    }
                };

                Canvas.SetLeft(particle, centerX);
                Canvas.SetTop(particle, centerY);
                canvas.Children.Add(particle);

                var endX = centerX + Math.Cos(angle) * radius;
                var endY = centerY + Math.Sin(angle) * radius;

                var moveXAnimation = new DoubleAnimation(centerX, endX, TimeSpan.FromSeconds(1.5));
                var moveYAnimation = new DoubleAnimation(centerY, endY + 50, TimeSpan.FromSeconds(2.0));
                var fadeAnimation = new DoubleAnimation(1.0, 0, TimeSpan.FromSeconds(2.0)) { BeginTime = TimeSpan.FromSeconds(0.5) };

                particle.BeginAnimation(Canvas.LeftProperty, moveXAnimation);
                particle.BeginAnimation(Canvas.TopProperty, moveYAnimation);
                particle.BeginAnimation(UIElement.OpacityProperty, fadeAnimation);

                var removeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2.5) };
                removeTimer.Tick += (s, e) =>
                {
                    canvas.Children.Remove(particle);
                    ((DispatcherTimer)s).Stop();
                };
                removeTimer.Start();
            }
        }

        private static void CreateWillowExplosion(Canvas canvas, double centerX, double centerY, Color[] colors, Random random)
        {
            var streamCount = 12 + random.Next(8);

            for (int stream = 0; stream < streamCount; stream++)
            {
                var angle = (2 * Math.PI * stream) / streamCount;
                var particlesPerStream = 15 + random.Next(10);

                for (int i = 0; i < particlesPerStream; i++)
                {
                    var delay = i * 0.05; // Stagger particles in stream
                    var distance = (i + 1) * (20 + random.Next(15));

                    var delayTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(delay) };
                    delayTimer.Tick += (s, e) =>
                    {
                        var particle = new Ellipse
                        {
                            Width = 3,
                            Height = 3,
                            Fill = new SolidColorBrush(colors[random.Next(colors.Length)]),
                            Effect = new DropShadowEffect { Color = Colors.Gold, BlurRadius = 6, ShadowDepth = 0 }
                        };

                        Canvas.SetLeft(particle, centerX);
                        Canvas.SetTop(particle, centerY);
                        canvas.Children.Add(particle);

                        var endX = centerX + Math.Cos(angle) * distance;
                        var endY = centerY + Math.Sin(angle) * distance + distance * 0.8; // Drooping effect

                        var moveXAnimation = new DoubleAnimation(centerX, endX, TimeSpan.FromSeconds(3.0));
                        var moveYAnimation = new DoubleAnimation(centerY, endY, TimeSpan.FromSeconds(3.0))
                        {
                            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
                        };
                        var fadeAnimation = new DoubleAnimation(1.0, 0, TimeSpan.FromSeconds(2.5)) { BeginTime = TimeSpan.FromSeconds(0.5) };

                        particle.BeginAnimation(Canvas.LeftProperty, moveXAnimation);
                        particle.BeginAnimation(Canvas.TopProperty, moveYAnimation);
                        particle.BeginAnimation(UIElement.OpacityProperty, fadeAnimation);

                        var removeTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3.5) };
                        removeTimer.Tick += (s2, e2) =>
                        {
                            canvas.Children.Remove(particle);
                            ((DispatcherTimer)s2).Stop();
                        };
                        removeTimer.Start();

                        ((DispatcherTimer)s).Stop();
                    };
                    delayTimer.Start();
                }
            }
        }

        #endregion

        #region Helper Methods


        // Finds the WrapPanel containing icons in a fence window

        private static WrapPanel FindWrapPanel(DependencyObject parent, int depth = 0, int maxDepth = 10)
        {
            if (parent == null || depth > maxDepth)
                return null;

            if (parent is WrapPanel wrapPanel)
                return wrapPanel;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                var result = FindWrapPanel(child, depth + 1, maxDepth);
                if (result != null)
                    return result;
            }

            return null;
        }

        private static Brush GetRandomSparkleColor(Random random)
        {
            var colors = new[]
            {
                Colors.Gold, Colors.Yellow, Colors.Orange, Colors.Red,
                Colors.Pink, Colors.Magenta, Colors.Cyan, Colors.LightBlue,
                Colors.White, Colors.Silver, Colors.Lime, Colors.Violet
            };

            return new SolidColorBrush(colors[random.Next(colors.Length)]);
        }

        #endregion
    }
}
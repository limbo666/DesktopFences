using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Media;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Desktop_Fences
{
    // InterCore system for Desktop Fences - handles special interactive features and animations.
    public static class InterCore
    {
        #region Private Fields

        private static bool _isDanceActive = false;
        private static bool _isGravityActive = false;
        private static readonly Dictionary<StackPanel, Point> _originalIconPositions = new Dictionary<StackPanel, Point>();
        private static Window _sparkleOverlay;

        // --- FIX: Add these missing dictionaries for Legendary Mode ---
        private static readonly Dictionary<string, Storyboard> _legendaryEffects = new Dictionary<string, Storyboard>();
        private static readonly Dictionary<string, Brush> _originalBorders = new Dictionary<string, Brush>();
        private static readonly Dictionary<string, Thickness> _originalBorderThicknesses = new Dictionary<string, Thickness>();
        private static readonly Dictionary<string, Effect> _originalEffects = new Dictionary<string, Effect>();
        // -------------------------------------------------------------

        #endregion

        #region Public Methods

        // Initializes the InterCore system - call this during application startup
        public static void Initialize()
        {
            try
            {
                // REGISTER HOTKEYS from GlobalHotkeyManager
                // Logic moved to GlobalHotkeyManager for better code management
                GlobalHotkeyManager.DancePartyTriggered += (s, e) => ActivateDanceParty();
                GlobalHotkeyManager.GravityDropTriggered += (s, e) => ActivateGravityDrop();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore system initialized and subscribed to global hotkeys");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error during initialization: {ex.Message}");
            }
        }

        // Cleans up any active effects - call this during application shutdown
        public static void Cleanup()
        {
            try
            {
                _sparkleOverlay?.Close();
                _originalIconPositions.Clear();

                // Unsubscribe to prevent leaks (though static longevity matches app life)
                GlobalHotkeyManager.DancePartyTriggered -= (s, e) => ActivateDanceParty();
                GlobalHotkeyManager.GravityDropTriggered -= (s, e) => ActivateGravityDrop();

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore system cleaned up");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error during cleanup: {ex.Message}");
            }
        }

        // Processes fence title changes for special triggers (e.g. "limbo666")
        public static string ProcessTitleChange(dynamic fence, string newTitle, string originalTitle)
        {
            try
            {
                // 1. Limbo666 Easter Egg (One-time trigger, Reverts title)
                if (string.Equals(newTitle, "limbo666", StringComparison.OrdinalIgnoreCase))
                {
                    LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: limbo666 trigger activated");
                    ActivateSparkleEffect();
                    return originalTitle; // Revert
                }

                // 2. The "Nikos" Legendary Mode (Persistent, Keeps title)
                bool isLegendary = string.Equals(newTitle, "Nikos", StringComparison.OrdinalIgnoreCase) ||
                                   string.Equals(newTitle, "Nikos Georgousis", StringComparison.OrdinalIgnoreCase) ||
                                   (!string.IsNullOrEmpty(newTitle) && newTitle.Contains(">:"));

                // We need the window to apply effects. Find it by ID.
                string fenceId = fence.Id?.ToString();
                var win = Application.Current.Windows.OfType<NonActivatingWindow>().FirstOrDefault(w => w.Tag?.ToString() == fenceId);

                if (win != null)
                {
                    if (isLegendary)
                    {
                        ActivateLegendaryMode(win, fenceId);
                    }
                    else
                    {
                        // If it WAS legendary but user changed name to "Work", deactivate it
                        DeactivateLegendaryMode(win, fenceId);
                    }
                }

                return newTitle; // Keep the new name (e.g. "Nikos")
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error processing title change: {ex.Message}");
                return newTitle;
            }
        }

        #region Legendary Mode (Nikos)

        private static void ActivateLegendaryMode(Window win, string fenceId)
        {
            if (_legendaryEffects.ContainsKey(fenceId)) return; // Already active

            try
            {
                var border = win.Content as Border;
                if (border == null) return;

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"InterCore: Activating Legendary Mode for {fenceId}");

                // 1. Save Original State
                _originalBorders[fenceId] = border.BorderBrush;
                _originalBorderThicknesses[fenceId] = border.BorderThickness;
                _originalEffects[fenceId] = win.Effect;

                // 2. Create "Mary-Go-Round" Gradient Brush
                // A vibrant rainbow gradient
                var rainbowBrush = new LinearGradientBrush
                {
                    StartPoint = new Point(0, 0),
                    EndPoint = new Point(1, 1),
                    GradientStops = new GradientStopCollection
                    {
                        new GradientStop(Colors.Red, 0.0),
                        new GradientStop(Colors.Gold, 0.2),
                        new GradientStop(Colors.Lime, 0.4),
                        new GradientStop(Colors.Cyan, 0.6),
                        new GradientStop(Colors.Magenta, 0.8),
                        new GradientStop(Colors.Red, 1.0)
                    },
                    RelativeTransform = new RotateTransform(0, 0.5, 0.5) // Rotate around center
                };

                // 3. Apply New Visuals
                border.BorderBrush = rainbowBrush;
                border.BorderThickness = new Thickness(4); // Make it thick enough to see

                // Add a glowing outer effect
                win.Effect = new DropShadowEffect
                {
                    Color = Colors.Cyan,
                    BlurRadius = 20,
                    ShadowDepth = 0,
                    Opacity = 0.8
                };

                // 4. Create Animation (Infinite Rotation)
                DoubleAnimation rotateAnim = new DoubleAnimation
                {
                    From = 0,
                    To = 360,
                    Duration = TimeSpan.FromSeconds(3), // Speed of rotation
                    RepeatBehavior = RepeatBehavior.Forever
                };

                // Apply animation to the brush's transform
                // Note: We must register the name to target it, or apply directly
                rainbowBrush.RelativeTransform.BeginAnimation(RotateTransform.AngleProperty, rotateAnim);

                // Store a dummy storyboard ref or just track ID to know it's active
                // Since we applied animation directly to the brush property, we don't strictly need a Storyboard object,
                // but we put a placeholder here to mark it as "Active" in our dictionary.
                _legendaryEffects[fenceId] = new Storyboard();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error activating Legendary Mode: {ex.Message}");
            }
        }

        private static void DeactivateLegendaryMode(Window win, string fenceId)
        {
            if (!_legendaryEffects.ContainsKey(fenceId)) return; // Not active

            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, $"InterCore: Deactivating Legendary Mode for {fenceId}");

                var border = win.Content as Border;
                if (border != null)
                {
                    // Restore Border
                    if (_originalBorders.ContainsKey(fenceId)) border.BorderBrush = _originalBorders[fenceId];
                    if (_originalBorderThicknesses.ContainsKey(fenceId)) border.BorderThickness = _originalBorderThicknesses[fenceId];
                }

                // Restore Effect
                if (_originalEffects.ContainsKey(fenceId)) win.Effect = _originalEffects[fenceId];
                else win.Effect = null;

                // Clean up memory
                _legendaryEffects.Remove(fenceId);
                _originalBorders.Remove(fenceId);
                _originalBorderThicknesses.Remove(fenceId);
                _originalEffects.Remove(fenceId);
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error deactivating Legendary Mode: {ex.Message}");
            }
        }

        #endregion

        #endregion

        #region Private Methods - Dance Party (Ctrl+Alt+D)

        private static void ActivateDanceParty()
        {
            if (_isDanceActive) return;

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: Dance Party activated!");
            _isDanceActive = true;

            try
            {
                // Play MIDI music
                PlayHappyTune2();

                // Get all fence icons
                var fenceWindows = Application.Current.Windows.OfType<NonActivatingWindow>();
                var allIcons = new List<StackPanel>();

                foreach (var window in fenceWindows)
                {
                    // Use FenceUtilities to find panels (or local helper if preferred)
                    var wrapPanel = FenceUtilities.FindWrapPanel(window);
                    if (wrapPanel != null)
                    {
                        allIcons.AddRange(wrapPanel.Children.OfType<StackPanel>());
                    }
                }

                // Create bounce animation
                foreach (var icon in allIcons)
                {
                    var bounceAnimation = new DoubleAnimationUsingKeyFrames();
                    var easing = new BounceEase { EasingMode = EasingMode.EaseOut };
                    
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(-20, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))) { EasingFunction = easing });
                    bounceAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))) { EasingFunction = easing });
                    
                    bounceAnimation.RepeatBehavior = new RepeatBehavior(TimeSpan.FromSeconds(10));

                    if (icon.RenderTransform == null || icon.RenderTransform == Transform.Identity)
                        icon.RenderTransform = new TranslateTransform();

                    var transform = icon.RenderTransform as TranslateTransform ?? new TranslateTransform();
                    icon.RenderTransform = transform;

                    transform.BeginAnimation(TranslateTransform.YProperty, bounceAnimation);
                }

                // Reset flag
                var resetTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };
                resetTimer.Tick += (s, e) =>
                {
                    _isDanceActive = false;
                    resetTimer.Stop();
                };
                resetTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error in Dance Party: {ex.Message}");
                _isDanceActive = false;
            }
        }

        private static void PlayHappyTune2()
        {
            if (SettingsManager.EnableSounds == false) return;

            try
            {
                var midiOut = new NAudio.Midi.MidiOut(0);
                midiOut.Send(NAudio.Midi.MidiMessage.ChangePatch(12, 1).RawData); 

                var thread = new Thread(() =>
                {
                    int[][] chords = { new[] { 60, 64, 67 }, new[] { 67, 71, 74 }, new[] { 69, 72, 76 }, new[] { 65, 69, 72 } };

                    for (int i = 0; i < 14; i++) 
                    {
                        foreach (var note in chords[i % chords.Length])
                            midiOut.Send(NAudio.Midi.MidiMessage.StartNote(note, 90, 1).RawData);
                        
                        Thread.Sleep(300);

                        // Staccato rhythm
                        for (int j = 0; j < 4; j++)
                        {
                            int rootNote = chords[i % chords.Length][0];
                            midiOut.Send(NAudio.Midi.MidiMessage.StopNote(rootNote, 0, 1).RawData);
                            midiOut.Send(NAudio.Midi.MidiMessage.StartNote(rootNote + (j % 2 == 0 ? 0 : 7), 110, 1).RawData);
                            Thread.Sleep(100);
                        }
                        foreach (var note in chords[i % chords.Length])
                            midiOut.Send(NAudio.Midi.MidiMessage.StopNote(note, 0, 1).RawData);
                    }
                    midiOut.Dispose();
                });

                thread.IsBackground = true;
                thread.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, $"InterCore: Audio error: {ex.Message}");
            }
        }

        #endregion

        #region Private Methods - Gravity Drop (Ctrl+Shift+G)

        private static void ActivateGravityDrop()
        {
            if (_isGravityActive) return;

            LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI, "InterCore: Gravity Drop activated!");
            _isGravityActive = true;

            try
            {
                var fenceWindows = Application.Current.Windows.OfType<NonActivatingWindow>();
                var allIcons = new List<StackPanel>();
                _originalIconPositions.Clear();

                foreach (var window in fenceWindows)
                {
                    var wrapPanel = FenceUtilities.FindWrapPanel(window);
                    if (wrapPanel != null)
                        allIcons.AddRange(wrapPanel.Children.OfType<StackPanel>());
                }

                foreach (var icon in allIcons)
                {
                    var random = new Random();
                    var fallDistance = 500 + random.Next(100);
                    var fallDuration = 1.5 + random.NextDouble() * 0.5;

                    var fallAnimation = new DoubleAnimation
                    {
                        From = 0, To = fallDistance, Duration = TimeSpan.FromSeconds(fallDuration),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseIn }
                    };

                    var bounceAnimation = new DoubleAnimation
                    {
                        From = fallDistance, To = 0, Duration = TimeSpan.FromSeconds(0.8),
                        EasingFunction = new BounceEase { EasingMode = EasingMode.EaseOut, Bounces = 3 },
                        BeginTime = TimeSpan.FromSeconds(fallDuration)
                    };

                    if (icon.RenderTransform == null || icon.RenderTransform == Transform.Identity)
                        icon.RenderTransform = new TranslateTransform();

                    var transform = icon.RenderTransform as TranslateTransform ?? new TranslateTransform();
                    icon.RenderTransform = transform;

                    transform.BeginAnimation(TranslateTransform.YProperty, fallAnimation);

                    var bounceTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(fallDuration) };
                    bounceTimer.Tick += (s, e) =>
                    {
                        transform.BeginAnimation(TranslateTransform.YProperty, bounceAnimation);
                        bounceTimer.Stop();
                    };
                    bounceTimer.Start();
                }

                var resetTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(4) };
                resetTimer.Tick += (s, e) =>
                {
                    _isGravityActive = false;
                    _originalIconPositions.Clear();
                    resetTimer.Stop();
                };
                resetTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"InterCore: Error in Gravity Drop: {ex.Message}");
                _isGravityActive = false;
            }
        }

        #endregion

        #region Private Methods - Lighthouse & Sparkle

        // (These methods remain unchanged, they are called by other mechanisms, not hotkeys)
        #region Private Methods - Lighthouse Sweep (Single Instance Effect)

        private static bool _isLighthouseSweepActive = false;

        /// <summary>
        /// Activates lighthouse sweep effect across all visible fences
        /// Called when registry monitor detects another instance attempt
        /// Creates a wave of golden glow that sweeps across all fences
        /// </summary>
        /// 

        public static void ActivateLighthouseSweep()
        {
            if (_isLighthouseSweepActive)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    "InterCore: Lighthouse sweep already active, skipping");
                return;
            }

            _isLighthouseSweepActive = true;

            try
            {
                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    "InterCore: Lighthouse Sweep effect activated - another instance detected!");

                // Get all visible fence windows
                var allFences = Application.Current.Windows.OfType<NonActivatingWindow>()
                    .Where(w => w.Visibility == Visibility.Visible)
                    .OrderBy(w => w.Left) // Order by horizontal position for wave effect
                    .ToList();

                if (allFences.Count == 0)
                {
                    LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI,
                        "InterCore: No visible fences found for Lighthouse Sweep");
                    _isLighthouseSweepActive = false;
                    return;
                }

                LogManager.Log(LogManager.LogLevel.Info, LogManager.LogCategory.UI,
                    $"InterCore: Starting Lighthouse Sweep on {allFences.Count} fences");

                // Store original opacities
                var originalOpacities = new Dictionary<NonActivatingWindow, double>();
                foreach (var fence in allFences)
                {
                    originalOpacities[fence] = fence.Opacity;

                    // If fence has high tint (low opacity), fade it to 0.4
                    if (fence.Opacity > 0.4)
                    {
                        var fadeOut = new DoubleAnimation
                        {
                            To = 0.4,
                            Duration = TimeSpan.FromMilliseconds(400), // smooth transition
                            FillBehavior = FillBehavior.HoldEnd
                        };
                        fence.BeginAnimation(UIElement.OpacityProperty, fadeOut);

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"InterCore: Fading opacity for fence '{fence.Title}' from {originalOpacities[fence]:F2} to 0.40");
                    }
                }
                PlaySweepSound();
                // Create wave effect across fences
                for (int i = 0; i < allFences.Count; i++)
                {
                    var fence = allFences[i];
                    int delay = i * 150; // 150ms delay between each fence for wave effect

                    var delayTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(delay) };
                    delayTimer.Tick += (s, e) =>
                    {
                        ApplyLighthouseGlowToFence(fence);
                        ((DispatcherTimer)s).Stop();
                    };
                    delayTimer.Start();
                }

                // Restore original opacities with fade-in
                var restoreTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(3) };
                restoreTimer.Tick += (s, e) =>
                {
                    try
                    {
                        foreach (var fence in allFences)
                        {
                            if (originalOpacities.ContainsKey(fence))
                            {
                                var fadeIn = new DoubleAnimation
                                {
                                    To = originalOpacities[fence],
                                    Duration = TimeSpan.FromMilliseconds(400),
                                    FillBehavior = FillBehavior.HoldEnd
                                };
                                fence.BeginAnimation(UIElement.OpacityProperty, fadeIn);

                                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                                    $"InterCore: Restoring opacity for fence '{fence.Title}' to {originalOpacities[fence]:F2}");
                            }
                        }

                        _isLighthouseSweepActive = false;
                        restoreTimer.Stop();

                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            "InterCore: Lighthouse Sweep effect completed and tint restored");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                            $"InterCore: Error restoring opacities: {ex.Message}");
                        _isLighthouseSweepActive = false;
                    }
                };
                restoreTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"InterCore: Error in Lighthouse Sweep activation: {ex.Message}");
                _isLighthouseSweepActive = false;
            }
        }

        /// <summary>
        /// Applies lighthouse glow effect to a single fence window
        /// Creates bright golden glow that pulses 3 times around the fence border
        /// </summary>
        private static void ApplyLighthouseGlowToFence(NonActivatingWindow fence)
        {
            try
            {
                // Create bright golden glow effect
                var lighthouseGlow = new DropShadowEffect
                {
                    Color = Color.FromRgb(255, 215, 0), // Gold color
                    BlurRadius = 25,
                    ShadowDepth = 0,
                    Opacity = 0
                };

                // Store original effect to restore later
                var originalEffect = fence.Effect;
                fence.Effect = lighthouseGlow;

                // Create pulsing animation - 3 strong pulses
                var pulseAnimation = new DoubleAnimation
                {
                    From = 0,
                    To = 0.9, // Bright but not overwhelming
                    Duration = TimeSpan.FromMilliseconds(200),
                    AutoReverse = true,
                    RepeatBehavior = new RepeatBehavior(3) // Pulse 3 times
                };

                // Restore original effect when animation completes
                pulseAnimation.Completed += (s, e) =>
                {
                    try
                    {
                        fence.Effect = originalEffect;
                        LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                            $"InterCore: Lighthouse glow completed on fence '{fence.Title}'");
                    }
                    catch (Exception ex)
                    {
                        LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                            $"InterCore: Error restoring fence effect: {ex.Message}");
                    }
                };

                // Start the glow animation
                lighthouseGlow.BeginAnimation(DropShadowEffect.OpacityProperty, pulseAnimation);

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    $"InterCore: Applied lighthouse glow to fence '{fence.Title}'");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"InterCore: Error applying lighthouse glow to fence '{fence?.Title}': {ex.Message}");
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






        /// <summary>
        /// Activates lighthouse sweep effect across all visible fences
        /// Called when registry monitor detects another instance attempt
        /// Creates a wave of golden glow that sweeps across all fences
        /// </summary>
        /// 


        #endregion
        #region Private Methods - Lighthouse Sweep (Single Instance Effect)

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

        private static void PlaySweepSound()
        {
            if (SettingsManager.EnableSounds == false)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.UI,
                    "InterCore: Sound is muted, skipping sweep sound playback");
                return;
            }
            try
            {
                using (Stream soundStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Desktop_Fences.Resources.sweep-sound-effect-240243.wav"))
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
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.UI, "Sound resource 'ding.wav' not found.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI, $"Error playing sound: {ex.Message}");
            }
        }


        #endregion


        #endregion
    }
}
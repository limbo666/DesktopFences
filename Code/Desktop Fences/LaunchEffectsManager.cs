using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Desktop_Fences
{

    // Manages visual launch effects for icons when they are activated

    public static class LaunchEffectsManager
    {

        // Available launch effect types

        public enum LaunchEffect
        {
            Zoom,        // First effect
            Bounce,      // Scale up and down a few times
            FadeOut,     // Fade out
            SlideUp,     // Slide up and return
            Rotate,      // Spin 360 degrees
            Agitate,     // Shake back and forth
            GrowAndFly,  // Grow and fly away
            Pulse,       // Pulsing
            Elastic,     // Elastic
            Flip3D,      // Flip 3D
            Spiral,      // Spiral

            Shockwave,   // Ripple effect expanding outward
            Matrix,      // Digital rain effect with glitch
            Supernova,   // Explosive burst with particles
            Teleport     // Sci-fi teleportation effect

        }


        //  specified launch effect on the given StackPanel (icon)
        // <param name="iconStackPanel">The icon StackPanel to animate</param>
        // <param name="effect">The effect to apply</param>
        public static void ExecuteLaunchEffect(StackPanel iconStackPanel, LaunchEffect effect)
        {
            if (iconStackPanel == null)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                    "Cannot execute launch effect: iconStackPanel is null");
                return;
            }

            try
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"Executing launch effect: {effect} on icon");

                // Ensure transform is set up
                EnsureTransformSetup(iconStackPanel);

                // Get transform group components
                var transformGroup = (TransformGroup)iconStackPanel.RenderTransform;
                var scaleTransform = (ScaleTransform)transformGroup.Children[0];
                var translateTransform = (TranslateTransform)transformGroup.Children[1];
                var rotateTransform = (RotateTransform)transformGroup.Children[2];

                // Execute the specific effect
                switch (effect)
                {
                    case LaunchEffect.Zoom:
                        ExecuteZoomEffect(scaleTransform);
                        break;

                    case LaunchEffect.Bounce:
                        ExecuteBounceEffect(scaleTransform);
                        break;

                    case LaunchEffect.FadeOut:
                        ExecuteFadeOutEffect(iconStackPanel);
                        break;

                    case LaunchEffect.SlideUp:
                        ExecuteSlideUpEffect(translateTransform);
                        break;

                    case LaunchEffect.Rotate:
                        ExecuteRotateEffect(rotateTransform);
                        break;

                    case LaunchEffect.Agitate:
                        ExecuteAgitateEffect(translateTransform);
                        break;

                    case LaunchEffect.GrowAndFly:
                        ExecuteGrowAndFlyEffect(iconStackPanel, scaleTransform, translateTransform);
                        break;

                    case LaunchEffect.Pulse:
                        ExecutePulseEffect(iconStackPanel, scaleTransform);
                        break;

                    case LaunchEffect.Elastic:
                        ExecuteElasticEffect(scaleTransform);
                        break;

                    case LaunchEffect.Flip3D:
                        ExecuteFlip3DEffect(iconStackPanel, scaleTransform, rotateTransform);
                        break;

                    case LaunchEffect.Spiral:
                        ExecuteSpiralEffect(scaleTransform, rotateTransform);
                        break;

                    case LaunchEffect.Shockwave:
                        ExecuteShockwaveEffect(iconStackPanel, scaleTransform);
                        break;

                    case LaunchEffect.Matrix:
                        ExecuteMatrixEffect(iconStackPanel, scaleTransform, rotateTransform);
                        break;

                    case LaunchEffect.Supernova:
                        ExecuteSupernovaEffect(iconStackPanel, scaleTransform);
                        break;

                    case LaunchEffect.Teleport:
                        ExecuteTeleportEffect(iconStackPanel, scaleTransform, translateTransform);
                        break;



                    default:
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                            $"Unknown launch effect: {effect}");
                        break;
                }

                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"Successfully executed launch effect: {effect}");
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.General,
                    $"Error executing launch effect {effect}: {ex.Message}");
                // Don't fail the launch if effect fails
            }
        }

        #region Private Effect Implementation Methods


        // Ensures the StackPanel has the proper transform setup for animations

        private static void EnsureTransformSetup(StackPanel stackPanel)
        {
            if (stackPanel.RenderTransform == null || !(stackPanel.RenderTransform is TransformGroup))
            {
                stackPanel.RenderTransform = new TransformGroup
                {
                    Children = new TransformCollection
                    {
                        new ScaleTransform(1, 1),
                        new TranslateTransform(0, 0),
                        new RotateTransform(0)
                    }
                };
                stackPanel.RenderTransformOrigin = new Point(0.5, 0.5);
            }
        }


        //  Zoom effect

        private static void ExecuteZoomEffect(ScaleTransform scaleTransform)
        {
            var zoomScale = new DoubleAnimation(1, 1.2, TimeSpan.FromSeconds(0.1)) { AutoReverse = true };
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, zoomScale);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, zoomScale);
        }


        //  Bounce effect

        private static void ExecuteBounceEffect(ScaleTransform scaleTransform)
        {
            var bounceScale = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.6)
            };
            bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));
            bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
            bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.2, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
            bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
            bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));
            bounceScale.KeyFrames.Add(new LinearDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, bounceScale);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, bounceScale);
        }


        //  FadeOut effect

        private static void ExecuteFadeOutEffect(StackPanel stackPanel)
        {
            var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.2)) { AutoReverse = true };
            stackPanel.BeginAnimation(UIElement.OpacityProperty, fade);
        }


        //  SlideUp effect

        private static void ExecuteSlideUpEffect(TranslateTransform translateTransform)
        {
            var slideUp = new DoubleAnimation(0, -20, TimeSpan.FromSeconds(0.2)) { AutoReverse = true };
            translateTransform.BeginAnimation(TranslateTransform.YProperty, slideUp);
        }


        //  Rotate effect

        private static void ExecuteRotateEffect(RotateTransform rotateTransform)
        {
            var rotate = new DoubleAnimation(0, 360, TimeSpan.FromSeconds(0.4))
            {
                EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
            };
            rotateTransform.BeginAnimation(RotateTransform.AngleProperty, rotate);
        }


        //  Agitate effect

        private static void ExecuteAgitateEffect(TranslateTransform translateTransform)
        {
            var agitateTranslate = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.7)
            };
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(-10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(-10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(-10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(10, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
            agitateTranslate.KeyFrames.Add(new LinearDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.7))));
            translateTransform.BeginAnimation(TranslateTransform.XProperty, agitateTranslate);
        }


        //  GrowAndFly effect

        private static void ExecuteGrowAndFlyEffect(StackPanel stackPanel, ScaleTransform scaleTransform, TranslateTransform translateTransform)
        {
            var growAnimation = new DoubleAnimation(1, 1.2, TimeSpan.FromSeconds(0.2));
            growAnimation.Completed += (s, _) =>
            {
                // After growing, start the fly away animation
                var shrinkAnimation = new DoubleAnimation(1.2, 0.05, TimeSpan.FromSeconds(0.3));
                var moveUpAnimation = new DoubleAnimation(0, -50, TimeSpan.FromSeconds(0.3));

                shrinkAnimation.Completed += (s2, _2) =>
                {
                    // Make the icon invisible
                    stackPanel.Opacity = 0;

                    // Remove all animations to allow direct property setting
                    scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, null);
                    scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, null);
                    translateTransform.BeginAnimation(TranslateTransform.YProperty, null);

                    // Reset transform values
                    scaleTransform.ScaleX = 1.0;
                    scaleTransform.ScaleY = 1.0;
                    translateTransform.Y = 0;

                    // Small delay before showing the icon again
                    var restoreTimer = new DispatcherTimer
                    {
                        Interval = TimeSpan.FromSeconds(0.1)
                    };

                    restoreTimer.Tick += (timerSender, timerArgs) =>
                    {
                        // Restore opacity
                        var restoreAnimation = new DoubleAnimation(0, 1, TimeSpan.FromSeconds(0.1));
                        stackPanel.BeginAnimation(UIElement.OpacityProperty, restoreAnimation);

                        // Stop and cleanup timer
                        restoreTimer.Stop();
                    };

                    restoreTimer.Start();
                };

                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, shrinkAnimation);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, shrinkAnimation);
                translateTransform.BeginAnimation(TranslateTransform.YProperty, moveUpAnimation);
            };

            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, growAnimation);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, growAnimation);
        }


        //  Pulse effect

        private static void ExecutePulseEffect(StackPanel stackPanel, ScaleTransform scaleTransform)
        {
            // Creates a pulsing effect with color change
            var pulseAnimation = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.8),
                AutoReverse = false
            };
            pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
            pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
            pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
            pulseAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));

            // Optional: Add color animation if icon supports it (assuming it's a Path or Shape)
            if (stackPanel.Children.Count > 0 && stackPanel.Children[0] is Shape shape)
            {
                var originalBrush = shape.Fill as SolidColorBrush;
                if (originalBrush != null)
                {
                    var colorAnimation = new ColorAnimation(
                        Colors.Red,
                        TimeSpan.FromSeconds(0.4))
                    {
                        AutoReverse = true
                    };
                    shape.Fill.BeginAnimation(SolidColorBrush.ColorProperty, colorAnimation);
                }
            }

            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, pulseAnimation);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, pulseAnimation);
        }


        //  Elastic effect

        private static void ExecuteElasticEffect(ScaleTransform scaleTransform)
        {
            // Creates a stretchy, elastic effect
            var elasticX = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.8)
            };
            elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.5, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
            elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(0.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
            elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
            elasticX.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));

            var elasticY = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.8)
            };
            elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(0.7, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
            elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(1.2, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));
            elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(0.9, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
            elasticY.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));

            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, elasticX);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, elasticY);
        }


        //  Flip3D effect

        private static void ExecuteFlip3DEffect(StackPanel stackPanel, ScaleTransform scaleTransform, RotateTransform rotateTransform)
        {
            // Creates a 3D flip effect
            var flipAnimation = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.6)
            };

            flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(90, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.15))));
            flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(270, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.45))));
            flipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(360, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));

            // For X-axis flip
            scaleTransform.CenterX = stackPanel.ActualWidth / 2;
            scaleTransform.CenterY = stackPanel.ActualHeight / 2;

            // We use a scale animation to create the flip illusion
            var scaleFlipAnimation = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.6)
            };
            scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.15))));
            scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.45))));
            scaleFlipAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));

            rotateTransform.BeginAnimation(RotateTransform.AngleProperty, flipAnimation);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, scaleFlipAnimation);
        }


        //  Spiral effect

        private static void ExecuteSpiralEffect(ScaleTransform scaleTransform, RotateTransform rotateTransform)
        {
            // Combines rotation with a zoom effect
            var spiralRotate = new DoubleAnimation(0, 720, TimeSpan.FromSeconds(0.7))
            {
                EasingFunction = new BackEase { Amplitude = 0.3, EasingMode = EasingMode.EaseInOut }
            };

            var spiralScale = new DoubleAnimationUsingKeyFrames
            {
                Duration = TimeSpan.FromSeconds(0.7)
            };
            spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
            spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(0.7, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
            spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));
            spiralScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.7))));

            rotateTransform.BeginAnimation(RotateTransform.AngleProperty, spiralRotate);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, spiralScale);
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, spiralScale);
        }
        // CORRECTED Launch Effects with proper DPI scaling and positioning
        // Replace the four effect methods in LaunchEffectsManager.cs with these corrected versions:


        // Gets DPI scale factor for proper positioning

        private static double GetDpiScaleFactor()
        {
            try
            {
                using (var graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    return graphics.DpiX / 96.0; // Standard DPI is 96
                }
            }
            catch
            {
                return 1.0; // Fallback to no scaling
            }
        }


        // Gets the absolute screen position of a UI element

        private static Point GetElementScreenPosition(FrameworkElement element)
        {
            try
            {
                // Get the element's position relative to the window
                var elementPoint = element.TranslatePoint(new Point(0, 0), Application.Current.MainWindow);

                // Get the window's screen position
                var window = Window.GetWindow(element);
                if (window != null)
                {
                    // Calculate absolute screen position
                    return new Point(
                        window.Left + elementPoint.X + (element.ActualWidth / 2),
                        window.Top + elementPoint.Y + (element.ActualHeight / 2)
                    );
                }

                return new Point(0, 0);
            }
            catch
            {
                return new Point(0, 0);
            }
        }


        // Creates a positioned overlay window for effects

        private static Window CreateEffectOverlay(StackPanel iconStackPanel, double width, double height)
        {
            try
            {
                // Get the parent window (fence window)
                var parentWindow = Window.GetWindow(iconStackPanel) as NonActivatingWindow;
                if (parentWindow == null) return null;

                // Get icon's position within the parent window
                var iconPosition = iconStackPanel.TranslatePoint(new Point(0, 0), parentWindow);

                // Calculate center position of the icon
                double iconCenterX = iconPosition.X + (iconStackPanel.ActualWidth / 2);
                double iconCenterY = iconPosition.Y + (iconStackPanel.ActualHeight / 2);

                // Create overlay window
                var overlay = new Window
                {
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = Brushes.Transparent,
                    ShowInTaskbar = false,
                    Topmost = true,
                    Width = width,
                    Height = height,
                    IsHitTestVisible = false,
                    Owner = parentWindow // Important: set owner to fence window
                };

                // Position overlay centered on the icon within the parent window
                overlay.Left = parentWindow.Left + iconCenterX - (width / 2);
                overlay.Top = parentWindow.Top + iconCenterY - (height / 2);

                return overlay;
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error creating effect overlay: {ex.Message}");
                return null;
            }
        }


        //  Shockwave effect - Creates expanding rings like a ripple in water (CORRECTED)

        private static void ExecuteShockwaveEffect(StackPanel stackPanel, ScaleTransform scaleTransform)
        {
            try
            {
                // Create overlay window for the effect
                var overlay = CreateEffectOverlay(stackPanel, 200, 200);
                if (overlay == null) return;

                // Create container for shockwave rings
                var shockwaveContainer = new Canvas
                {
                    Width = 200,
                    Height = 200,
                    Background = Brushes.Transparent
                };

                overlay.Content = shockwaveContainer;
                overlay.Show();

                // Create 3 expanding rings
                for (int i = 0; i < 3; i++)
                {
                    var ring = new Ellipse
                    {
                        Width = 10,
                        Height = 10,
                        Stroke = new SolidColorBrush(Color.FromArgb(150, 0, 200, 255)),
                        StrokeThickness = 3,
                        Fill = Brushes.Transparent
                    };

                    // Center the ring in the canvas
                    Canvas.SetLeft(ring, 95);
                    Canvas.SetTop(ring, 95);
                    shockwaveContainer.Children.Add(ring);

                    // Animate each ring with delay
                    var expandAnimation = new DoubleAnimation(10, 180, TimeSpan.FromSeconds(0.8))
                    {
                        BeginTime = TimeSpan.FromSeconds(i * 0.15),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                    };

                    var fadeAnimation = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(0.8))
                    {
                        BeginTime = TimeSpan.FromSeconds(i * 0.15),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                    };

                    // Apply animations
                    ring.BeginAnimation(FrameworkElement.WidthProperty, expandAnimation);
                    ring.BeginAnimation(FrameworkElement.HeightProperty, expandAnimation);
                    ring.BeginAnimation(UIElement.OpacityProperty, fadeAnimation);

                    // Update position as ring expands to keep it centered
                    var repositionAnimation = new DoubleAnimation(95, 5, TimeSpan.FromSeconds(0.8))
                    {
                        BeginTime = TimeSpan.FromSeconds(i * 0.15),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                    };
                    ring.BeginAnimation(Canvas.LeftProperty, repositionAnimation);
                    ring.BeginAnimation(Canvas.TopProperty, repositionAnimation);
                }

                // Icon scale pulse
                var iconPulse = new DoubleAnimationUsingKeyFrames
                {
                    Duration = TimeSpan.FromSeconds(1.0)
                };
                iconPulse.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                iconPulse.KeyFrames.Add(new EasingDoubleKeyFrame(1.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
                iconPulse.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));

                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, iconPulse);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, iconPulse);

                // Clean up overlay after animation
                var cleanupTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(1.5)
                };
                cleanupTimer.Tick += (s, e) =>
                {
                    overlay?.Close();
                    cleanupTimer.Stop();
                };
                cleanupTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in Shockwave effect: {ex.Message}");
            }
        }


        //  Matrix effect - Digital glitch with cascading particles (CORRECTED)

        private static void ExecuteMatrixEffect(StackPanel stackPanel, ScaleTransform scaleTransform, RotateTransform rotateTransform)
        {
            try
            {
                var random = new Random();

                // Create overlay window for the effect
                var overlay = CreateEffectOverlay(stackPanel, 150, 150);
                if (overlay == null) return;

                // Create container for matrix particles
                var matrixContainer = new Canvas
                {
                    Width = 150,
                    Height = 150,
                    Background = Brushes.Transparent
                };

                overlay.Content = matrixContainer;
                overlay.Show();

                // Create digital "rain" particles
                for (int i = 0; i < 15; i++)
                {
                    var particle = new TextBlock
                    {
                        Text = random.Next(0, 2).ToString(), // Random 0 or 1
                        FontFamily = new FontFamily("Consolas"),
                        FontSize = 12,
                        Foreground = new SolidColorBrush(Color.FromArgb(200, 0, 255, 0)),
                        Effect = new DropShadowEffect
                        {
                            Color = Color.FromRgb(0, 255, 60),
                            BlurRadius = 3,
                            ShadowDepth = 0
                        }
                    };

                    double startX = random.NextDouble() * 130;
                    double startY = random.NextDouble() * 50 - 25;

                    Canvas.SetLeft(particle, startX);
                    Canvas.SetTop(particle, startY);
                    matrixContainer.Children.Add(particle);

                    // Animate particles falling
                    var fallAnimation = new DoubleAnimation(startY, 150, TimeSpan.FromSeconds(0.8 + random.NextDouble() * 0.4))
                    {
                        BeginTime = TimeSpan.FromSeconds(random.NextDouble() * 0.3)
                    };

                    var fadeAnimation = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(0.5))
                    {
                        BeginTime = TimeSpan.FromSeconds(0.6 + random.NextDouble() * 0.3)
                    };

                    particle.BeginAnimation(Canvas.TopProperty, fallAnimation);
                    particle.BeginAnimation(UIElement.OpacityProperty, fadeAnimation);

                    // Randomly change the character during animation
                    var changeTimer = new DispatcherTimer
                    {
                        Interval = TimeSpan.FromMilliseconds(100)
                    };
                    int changes = 0;
                    changeTimer.Tick += (s, e) =>
                    {
                        particle.Text = random.Next(0, 2).ToString();
                        changes++;
                        if (changes > 8)
                        {
                            changeTimer.Stop();
                        }
                    };
                    changeTimer.Start();
                }

                // Icon glitch effect
                var glitchAnimation = new DoubleAnimationUsingKeyFrames
                {
                    Duration = TimeSpan.FromSeconds(1.0)
                };
                glitchAnimation.KeyFrames.Add(new DiscreteDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                glitchAnimation.KeyFrames.Add(new DiscreteDoubleKeyFrame(0.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));
                glitchAnimation.KeyFrames.Add(new DiscreteDoubleKeyFrame(1.2, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.15))));
                glitchAnimation.KeyFrames.Add(new DiscreteDoubleKeyFrame(0.9, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.25))));
                glitchAnimation.KeyFrames.Add(new DiscreteDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.35))));
                glitchAnimation.KeyFrames.Add(new DiscreteDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.5))));

                // Quick rotation glitches
                var rotationGlitch = new DoubleAnimationUsingKeyFrames
                {
                    Duration = TimeSpan.FromSeconds(1.0)
                };
                rotationGlitch.KeyFrames.Add(new DiscreteDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                rotationGlitch.KeyFrames.Add(new DiscreteDoubleKeyFrame(5, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));
                rotationGlitch.KeyFrames.Add(new DiscreteDoubleKeyFrame(-3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.2))));
                rotationGlitch.KeyFrames.Add(new DiscreteDoubleKeyFrame(2, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
                rotationGlitch.KeyFrames.Add(new DiscreteDoubleKeyFrame(0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.4))));

                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, glitchAnimation);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, glitchAnimation);
                rotateTransform.BeginAnimation(RotateTransform.AngleProperty, rotationGlitch);

                // Cleanup
                var cleanupTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.8) };
                cleanupTimer.Tick += (s, e) =>
                {
                    overlay?.Close();
                    cleanupTimer.Stop();
                };
                cleanupTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in Matrix effect: {ex.Message}");
            }
        }


        //  Supernova effect - Explosive burst with radiating particles (CORRECTED)

        private static void ExecuteSupernovaEffect(StackPanel stackPanel, ScaleTransform scaleTransform)
        {
            try
            {
                var random = new Random();

                // Create overlay window for the effect
                var overlay = CreateEffectOverlay(stackPanel, 250, 250);
                if (overlay == null) return;

                // Create explosion container
                var explosionContainer = new Canvas
                {
                    Width = 250,
                    Height = 250,
                    Background = Brushes.Transparent
                };

                overlay.Content = explosionContainer;
                overlay.Show();

                // Create central flash
                var flash = new Ellipse
                {
                    Width = 20,
                    Height = 20,
                    Fill = new RadialGradientBrush
                    {
                        GradientStops = new GradientStopCollection
                {
                    new GradientStop(Color.FromArgb(255, 255, 255, 100), 0.0),
                    new GradientStop(Color.FromArgb(255, 255, 200, 0), 0.3),
                    new GradientStop(Color.FromArgb(100, 255, 100, 0), 0.7),
                    new GradientStop(Colors.Transparent, 1.0)
                }
                    }
                };

                // Center the flash in the canvas
                Canvas.SetLeft(flash, 115);
                Canvas.SetTop(flash, 115);
                explosionContainer.Children.Add(flash);

                // Flash expansion
                var flashExpand = new DoubleAnimation(20, 200, TimeSpan.FromSeconds(0.6))
                {
                    EasingFunction = new ExponentialEase { EasingMode = EasingMode.EaseOut }
                };
                var flashFade = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(0.8));

                flash.BeginAnimation(FrameworkElement.WidthProperty, flashExpand);
                flash.BeginAnimation(FrameworkElement.HeightProperty, flashExpand);
                flash.BeginAnimation(UIElement.OpacityProperty, flashFade);

                // Update flash position as it expands
                var flashReposition = new DoubleAnimation(115, 25, TimeSpan.FromSeconds(0.6))
                {
                    EasingFunction = new ExponentialEase { EasingMode = EasingMode.EaseOut }
                };
                flash.BeginAnimation(Canvas.LeftProperty, flashReposition);
                flash.BeginAnimation(Canvas.TopProperty, flashReposition);

                // Create radiating particles
                for (int i = 0; i < 20; i++)
                {
                    var particle = new Ellipse
                    {
                        Width = 4,
                        Height = 4,
                        Fill = new SolidColorBrush(Color.FromArgb(255,
                            (byte)(200 + random.Next(56)),
                            (byte)(100 + random.Next(156)),
                            (byte)random.Next(100)))
                    };

                    double angle = (360.0 / 20.0) * i + random.NextDouble() * 20 - 10; // Add some randomness
                    double distance = 80 + random.NextDouble() * 40;

                    double endX = 125 + Math.Cos(angle * Math.PI / 180) * distance;
                    double endY = 125 + Math.Sin(angle * Math.PI / 180) * distance;

                    // Start at center
                    Canvas.SetLeft(particle, 125);
                    Canvas.SetTop(particle, 125);
                    explosionContainer.Children.Add(particle);

                    // Animate particles outward
                    var moveXAnimation = new DoubleAnimation(125, endX, TimeSpan.FromSeconds(0.8))
                    {
                        BeginTime = TimeSpan.FromSeconds(0.1),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                    };
                    var moveYAnimation = new DoubleAnimation(125, endY, TimeSpan.FromSeconds(0.8))
                    {
                        BeginTime = TimeSpan.FromSeconds(0.1),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                    };
                    var particleFade = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(0.6))
                    {
                        BeginTime = TimeSpan.FromSeconds(0.3)
                    };

                    particle.BeginAnimation(Canvas.LeftProperty, moveXAnimation);
                    particle.BeginAnimation(Canvas.TopProperty, moveYAnimation);
                    particle.BeginAnimation(UIElement.OpacityProperty, particleFade);
                }

                // Icon supernova effect
                var supernovaScale = new DoubleAnimationUsingKeyFrames
                {
                    Duration = TimeSpan.FromSeconds(1.2)
                };
                supernovaScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                supernovaScale.KeyFrames.Add(new EasingDoubleKeyFrame(0.3, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.1))));
                supernovaScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
                supernovaScale.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));

                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, supernovaScale);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, supernovaScale);

                // Cleanup
                var cleanupTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
                cleanupTimer.Tick += (s, e) =>
                {
                    overlay?.Close();
                    cleanupTimer.Stop();
                };
                cleanupTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in Supernova effect: {ex.Message}");
            }
        }


        //  Teleport effect - Sci-fi beam up/down effect with energy rings (CORRECTED)

        private static void ExecuteTeleportEffect(StackPanel stackPanel, ScaleTransform scaleTransform, TranslateTransform translateTransform)
        {
            try
            {
                // Create overlay window for the effect
                var overlay = CreateEffectOverlay(stackPanel, 120, 200);
                if (overlay == null) return;

                // Create teleport container
                var teleportContainer = new Canvas
                {
                    Width = 120,
                    Height = 200,
                    Background = Brushes.Transparent
                };

                overlay.Content = teleportContainer;
                overlay.Show();

                // Create energy beam
                var beam = new Rectangle
                {
                    Width = 50,
                    Height = 150,
                    Fill = new LinearGradientBrush
                    {
                        StartPoint = new Point(0, 0),
                        EndPoint = new Point(0, 1),
                        GradientStops = new GradientStopCollection
                {
                    new GradientStop(Colors.Transparent, 0.0),
                    new GradientStop(Color.FromArgb(100, 0, 255, 255), 0.2),
                    new GradientStop(Color.FromArgb(150, 100, 200, 255), 0.5),
                    new GradientStop(Color.FromArgb(100, 0, 255, 255), 0.8),
                    new GradientStop(Colors.Transparent, 1.0)
                }
                    },
                    Effect = new BlurEffect { Radius = 2 }
                };

                // Center the beam in the canvas
                Canvas.SetLeft(beam, 35);
                Canvas.SetTop(beam, 25);
                teleportContainer.Children.Add(beam);

                // Create scanning rings
                for (int i = 0; i < 5; i++)
                {
                    var ring = new Ellipse
                    {
                        Width = 60,
                        Height = 8,
                        Stroke = new SolidColorBrush(Color.FromArgb(200, 0, 255, 255)),
                        StrokeThickness = 2,
                        Fill = new SolidColorBrush(Color.FromArgb(50, 0, 255, 255)),
                        Effect = new DropShadowEffect
                        {
                            Color = Color.FromRgb(0, 255, 255),
                            BlurRadius = 5,
                            ShadowDepth = 0
                        }
                    };

                    Canvas.SetLeft(ring, 30);
                    Canvas.SetTop(ring, 200); // Start below
                    teleportContainer.Children.Add(ring);

                    // Animate rings moving up
                    var moveUpAnimation = new DoubleAnimation(200, -10, TimeSpan.FromSeconds(1.0))
                    {
                        BeginTime = TimeSpan.FromSeconds(i * 0.15),
                        EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseInOut }
                    };

                    var ringFade = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(0.3))
                    {
                        BeginTime = TimeSpan.FromSeconds(0.7 + i * 0.15)
                    };

                    ring.BeginAnimation(Canvas.TopProperty, moveUpAnimation);
                    ring.BeginAnimation(UIElement.OpacityProperty, ringFade);
                }

                // Beam fade in and out
                var beamFadeIn = new DoubleAnimation(0.0, 1.0, TimeSpan.FromSeconds(0.2));
                var beamFadeOut = new DoubleAnimation(1.0, 0.0, TimeSpan.FromSeconds(0.3))
                {
                    BeginTime = TimeSpan.FromSeconds(0.9)
                };

                beam.BeginAnimation(UIElement.OpacityProperty, beamFadeIn);

                var beamFadeOutTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(0.9) };
                beamFadeOutTimer.Tick += (s, e) =>
                {
                    beam.BeginAnimation(UIElement.OpacityProperty, beamFadeOut);
                    beamFadeOutTimer.Stop();
                };
                beamFadeOutTimer.Start();

                // Icon teleport effect
                var teleportIconAnimation = new DoubleAnimationUsingKeyFrames
                {
                    Duration = TimeSpan.FromSeconds(1.2)
                };
                teleportIconAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                teleportIconAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
                teleportIconAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
                teleportIconAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.8))));
                teleportIconAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(1.2))));

                // Icon shimmer during teleport
                var shimmerAnimation = new DoubleAnimationUsingKeyFrames
                {
                    Duration = TimeSpan.FromSeconds(1.2)
                };
                shimmerAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0))));
                shimmerAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.2, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.3))));
                shimmerAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(0.8, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.6))));
                shimmerAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.1, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(0.9))));
                shimmerAnimation.KeyFrames.Add(new EasingDoubleKeyFrame(1.0, KeyTime.FromTimeSpan(TimeSpan.FromSeconds(1.2))));

                stackPanel.BeginAnimation(UIElement.OpacityProperty, teleportIconAnimation);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleXProperty, shimmerAnimation);
                scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, shimmerAnimation);

                // Cleanup
                var cleanupTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
                cleanupTimer.Tick += (s, e) =>
                {
                    overlay?.Close();
                    cleanupTimer.Stop();
                };
                cleanupTimer.Start();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.UI,
                    $"Error in Teleport effect: {ex.Message}");
            }
        }

        #endregion
    }
}
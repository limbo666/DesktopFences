using System.Collections.Generic;
using System.Windows;

namespace Desktop_Frames.Plugins
{
    public interface IFramePlugin
    {
        string PluginId { get; }
        string DisplayName { get; }

        // 1 = Finalized, 2 = Experimental, 3 = In Development
        int DevelopmentState { get; }

        FrameworkElement CreateVisualElement();
        void Initialize(FrameworkElement visual, Dictionary<string, object> settings);

        void Pause();
        void Resume();
        void Cleanup();

        // --- NEW: Decoupled Settings Interface ---
        // The host passes its window (for centering) and the frame data.
        // The plugin draws its own window, updates the data, and saves.
        void ShowSettingsWindow(Window ownerWindow, dynamic frameData);
    }
}
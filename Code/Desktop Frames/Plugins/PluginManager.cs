using System;
using System.Collections.Generic;

namespace Desktop_Frames.Plugins
{
    /// <summary>
    /// Central registry for all internal Desktop Frames plugins.
    /// Acts as a factory to ensure every frame gets its own isolated WPF visual instance.
    /// </summary>
    public static class PluginManager
    {
        // Stores the factory functions to create fresh instances of a plugin
        private static readonly Dictionary<string, Func<IFramePlugin>> _pluginFactories = new Dictionary<string, Func<IFramePlugin>>();

        // Stores display names for the Customize Frame Form dropdowns
        private static readonly Dictionary<string, string> _pluginNames = new Dictionary<string, string>();

        // Stores the development tier state for each plugin
        private static readonly Dictionary<string, int> _pluginStates = new Dictionary<string, int>();

        /// <summary>
        /// Called once during application startup (e.g., inside FrameManager.LoadAndCreateFrames)
        /// </summary>
        public static void Initialize()
        {
            _pluginFactories.Clear();
            _pluginNames.Clear();

            // ---------------------------------------------------------
            // REGISTER ALL INTERNAL PLUGINS HERE
            // ---------------------------------------------------------
            RegisterPlugin("PictureSlideshow", "Photo Frame", () => new PictureSlideshowPlugin());
            RegisterPlugin("Calculator", "Calculator", () => new CalculatorPlugin());
            RegisterPlugin("VuMeter", "VU Meter", () => new VUMeterPlugin());
            RegisterPlugin("IPInfo", "IP Info", () => new SystemIpPlugin());
            RegisterPlugin("SystemPerformancePlugin", "System Performance Gauges", () => new SystemPerformancePlugin());
            RegisterPlugin("CustomTerminal", "Terminal Emulator", () => new CustomTerminalPlugin());
 
     
        }

        /// <summary>
        /// Registers a new plugin into the internal ecosystem.
        /// </summary>
        public static void RegisterPlugin(string pluginId, string displayName, Func<IFramePlugin> factory)
        {
            if (!_pluginFactories.ContainsKey(pluginId))
            {
                _pluginFactories[pluginId] = factory;
                _pluginNames[pluginId] = displayName;

                try
                {
                    // Briefly instantiate to read its declared Development State
                    var tempInstance = factory();
                    _pluginStates[pluginId] = tempInstance.DevelopmentState;
                }
                catch
                {
                    _pluginStates[pluginId] = 3; // Failsafe: Assume most unstable if error occurs
                }
            }
        }

        /// <summary>
        /// Generates a brand new, isolated instance of the requested plugin.
        /// </summary>
        public static IFramePlugin CreatePluginInstance(string pluginId)
        {
            if (string.IsNullOrEmpty(pluginId)) return null;

            if (_pluginFactories.TryGetValue(pluginId, out var factory))
            {
                return factory(); // Executes the "() => new Plugin()" rule
            }

            return null;
        }

        /// <summary>
        /// Returns a list of plugins filtered by the user's allowed Development State level.
        /// Level 0 = Returns empty (Menu disabled)
        /// </summary>
        public static Dictionary<string, string> GetAvailablePlugins()
        {
            int userAllowedLevel = SettingsManager.PluginAvailabilityLevel;
            var allowedPlugins = new Dictionary<string, string>();

            if (userAllowedLevel == 0) return allowedPlugins;

            foreach (var kvp in _pluginNames)
            {
                string pId = kvp.Key;
                int pluginLevel = _pluginStates.ContainsKey(pId) ? _pluginStates[pId] : 3;

                // Only allow plugins whose state is LESS THAN OR EQUAL TO the user's allowed level
                if (pluginLevel <= userAllowedLevel)
                {
                    allowedPlugins.Add(pId, kvp.Value);
                }
            }

            return allowedPlugins;
        }
    }
}
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace Desktop_Fences
{
    /// <summary>
    /// Centralized utility methods for Desktop Fences
    /// Extracted from FenceManager for better code organization and reusability
    /// Contains standalone helper methods with minimal dependencies
    /// </summary>
    public static class FenceUtilities
    {
      

        #region Tab Naming - Used by: FenceManager.AddNewTab
        // Herb names for random tab naming
        private static readonly string[] herbNames = {
            "Yarrow", "Mullein", "Plantain", "Mugwort", "Nettle", "Chicory", "Celandine",
            "Sorrel", "Agrimony", "Betony", "Coltsfoot", "Eyebright", "Meadowsweet",
            "Wormwood", "Lovage", "Valerian", "Comfrey", "Bistort", "Fleabane", "Tansy",
            "Feverfew", "Pennyroyal", "Goldenrod", "Boneset", "Selfheal", "Cleavers",
            "Vervain", "Rue", "Horehound", "Arnica"
        };

        // Random instance for herb naming
        private static readonly Random herbNameRandom = new Random();

        /// <summary>
        /// Generates random herb names for new tabs
        /// Used by: FenceManager.AddNewTab
        /// Category: Tab Management
        /// </summary>
        public static string GenerateRandomHerbName()
        {
            return herbNames[herbNameRandom.Next(herbNames.Length)];
        }
        #endregion

        #region Visual Tree Navigation - Used by: Multiple files (centralized from duplicates)
        /// <summary>
        /// Finds WrapPanel in visual tree with depth protection
        /// Used by: FenceManager.RefreshIconClickHandlers, IconDragDropManager, InterCore
        /// Category: Visual Tree Navigation
        /// Centralized from multiple file duplicates
        /// </summary>
        public static WrapPanel FindWrapPanel(DependencyObject parent, int depth = 0, int maxDepth = 10)
        {
            // Prevent infinite recursion
            if (parent == null || depth > maxDepth)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.IconHandling,
                    $"FindWrapPanel: Reached max depth {maxDepth} or null parent at depth {depth}");
                return null;
            }

            // Check if current element is a WrapPanel
            if (parent is WrapPanel wrapPanel)
            {
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"FindWrapPanel: Found WrapPanel at depth {depth}");
                return wrapPanel;
            }

            // Recurse through visual tree
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                    $"FindWrapPanel: Checking child {i} at depth {depth}, type: {child?.GetType()?.Name ?? "null"}");
                var result = FenceUtilities.FindWrapPanel(child, depth + 1, maxDepth);
                if (result != null)
                {
                    return result;
                }
            }

            LogManager.Log(LogManager.LogLevel.Debug, LogManager.LogCategory.General,
                $"FindWrapPanel: No WrapPanel found under parent {parent?.GetType()?.Name ?? "null"} at depth {depth}");
            return null;
        }
        #endregion


    }
}


// CHECKPOINT: Step 1 Complete!
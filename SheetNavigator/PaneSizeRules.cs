using System;

namespace SheetNavigator
{
    /// <summary>
    /// The size rules applied to every pane width and floating height read from or written to the
    /// settings. Pure functions with no Excel dependency, so they can be tested outside Excel.
    /// </summary>
    internal static class PaneSizeRules
    {
        /// <summary>
        /// Ceiling for widths, on both save and load, to reject a garbage value in the settings file.
        /// There is no floor: Excel enforces its own minimum whenever a width is set.
        /// </summary>
        public const int MaxWidth = 400;

        /// <summary>
        /// Ceiling for floating heights, on both save and load. A floating pane has no maximum in
        /// Excel, so a garbage value could push its bottom edge off the screen.
        /// </summary>
        public const int MaxHeight = 1200;

        /// <summary>
        /// A width of zero or less (a garbage value) is the fallback; anything above the ceiling is the ceiling.
        /// </summary>
        public static int ClampWidth(int width, int fallback)
        {
            if (width <= 0) return fallback;
            return Math.Min(MaxWidth, width);
        }

        /// <summary>
        /// A height of zero or less (a garbage value) is the fallback; anything above the ceiling is the ceiling.
        /// </summary>
        public static int ClampHeight(int height, int fallback)
        {
            if (height <= 0) return fallback;
            return Math.Min(MaxHeight, height);
        }
    }
}

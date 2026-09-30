using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace SheetNavigator
{
    /// <summary>
    /// Adds a "Worksheets" toggle button to Excel's View tab. The button shows or hides the
    /// active window's pane and stays pressed while the pane is visible.
    /// Callbacks are named in Ribbon.xml and must stay public.
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private const string WorksheetsToggleId = "SheetNavigatorWorksheetsToggle";
        private Office.IRibbonUI ribbonUI;

        /// <summary>
        /// Returns the Ribbon XML embedded in the assembly.
        /// </summary>
        public string GetCustomUI(string ribbonId)
        {
            return GetResourceText("SheetNavigator.Ribbon.xml");
        }

        /// <summary>
        /// Keeps the Ribbon reference needed to refresh the button later.
        /// </summary>
        public void OnLoad(Office.IRibbonUI ribbon)
        {
            ribbonUI = ribbon;
        }

        /// <summary>
        /// Button click: toggles the pane.
        /// </summary>
        public void OnWorksheetsToggle(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.ToggleSidebar();
            RefreshToggleState();
        }

        /// <summary>
        /// Button pressed state: true while the active window's pane is shown.
        /// </summary>
        public bool GetWorksheetsPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.IsSidebarVisible;
        }

        /// <summary>
        /// Asks Excel to re-query the button's pressed state.
        /// </summary>
        public void RefreshToggleState()
        {
            ribbonUI?.InvalidateControl(WorksheetsToggleId);
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (string.Equals(resourceName, name, StringComparison.OrdinalIgnoreCase))
                {
                    using (StreamReader reader = new StreamReader(asm.GetManifestResourceStream(name)))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }
    }
}

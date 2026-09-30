using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
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
        /// Button click: shows or hides the pane of the window the button belongs to,
        /// matching the state the button now displays.
        /// </summary>
        public void OnWorksheetsToggle(Office.IRibbonControl control, bool pressed)
        {
            Globals.ThisAddIn.SetSidebarVisible(WindowOf(control), pressed);
            RefreshToggleState();
        }

        /// <summary>
        /// Button pressed state: true while the pane of the button's own window is shown.
        /// </summary>
        public bool GetWorksheetsPressed(Office.IRibbonControl control)
        {
            return Globals.ThisAddIn.IsSidebarVisibleIn(WindowOf(control));
        }

        /// <summary>
        /// The Excel window whose Ribbon raised the callback; each window has its own Ribbon.
        /// </summary>
        private static Excel.Window WindowOf(Office.IRibbonControl control)
        {
            try
            {
                return control?.Context as Excel.Window;
            }
            catch { return null; }
        }

        /// <summary>
        /// Asks Excel to re-query the button's pressed state. Safe to call at any time.
        /// </summary>
        public void RefreshToggleState()
        {
            try
            {
                ribbonUI?.InvalidateControl(WorksheetsToggleId);
            }
            catch { /* The Ribbon may be gone while a window closes */ }
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

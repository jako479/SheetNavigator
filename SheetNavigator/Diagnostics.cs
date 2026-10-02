using System;
using System.IO;
using System.Windows.Forms;

namespace SheetNavigator
{
    /// <summary>
    /// Appends errors and any unhandled exception to %TEMP%\SheetNavigator.log.
    /// Never throws; logging must not be able to break the add-in.
    /// </summary>
    internal static class Diagnostics
    {
        /// <summary>Once the log passes this size, only its newest <see cref="KeepBytes"/> are kept.</summary>
        private const long MaxLogBytes = 1024 * 1024;
        private const int KeepBytes = 256 * 1024;

        private static readonly string logPath = Path.Combine(Path.GetTempPath(), "SheetNavigator.log");

        /// <summary>
        /// Appends one timestamped line. Swallows every error, since logging must never break the add-in.
        /// </summary>
        public static void Write(string message)
        {
            try
            {
                TrimIfLarge();
                File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {message}{Environment.NewLine}");
            }
            catch { }
        }

        /// <summary>
        /// Records exceptions that nothing else caught, so a crash leaves its name and stack behind.
        /// </summary>
        public static void HookUnhandledExceptions()
        {
            try
            {
                AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
                    Write("UNHANDLED (AppDomain): " + e.ExceptionObject);

                Application.ThreadException += (sender, e) =>
                    Write("UNHANDLED (WinForms): " + e.Exception);
            }
            catch { }
        }

        /// <summary>
        /// Keeps a repeating error from growing the file without limit.
        /// </summary>
        private static void TrimIfLarge()
        {
            FileInfo info = new FileInfo(logPath);
            if (!info.Exists || info.Length <= MaxLogBytes) return;

            byte[] all = File.ReadAllBytes(logPath);
            byte[] tail = new byte[KeepBytes];
            Array.Copy(all, all.Length - KeepBytes, tail, 0, KeepBytes);
            File.WriteAllBytes(logPath, tail);
        }
    }
}

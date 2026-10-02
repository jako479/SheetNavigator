using System;
using System.IO;
using System.Windows.Forms;

namespace SheetNavigator
{
    /// <summary>
    /// Appends every caught error and any unhandled exception to %TEMP%\SheetNavigator.log.
    /// Never throws; logging must not be able to break the add-in.
    /// </summary>
    internal static class Diagnostics
    {
        /// <summary>Once the log passes this size, only its newest <see cref="KeepBytes"/> are kept.</summary>
        private const long MaxLogBytes = 1024 * 1024;
        private const int KeepBytes = 256 * 1024;

        private static readonly string logPath = Path.Combine(Path.GetTempPath(), "SheetNavigator.log");
        private static readonly object gate = new object();

        /// <summary>The last message written, and how many identical messages have been dropped since.</summary>
        private static string lastMessage;
        private static int repeatCount;

        /// <summary>
        /// Appends one timestamped line. A message identical to the previous one is counted instead of
        /// written, and the count goes out once a different message arrives, so a failure that repeats
        /// on every timer tick fills two lines, not the file.
        /// </summary>
        public static void Write(string message)
        {
            try
            {
                lock (gate)
                {
                    if (string.Equals(message, lastMessage, StringComparison.Ordinal))
                    {
                        repeatCount++;
                        return;
                    }

                    WriteRepeatCount();
                    lastMessage = message;
                    AppendLine(message);
                }
            }
            catch { /* Nowhere left to report a logging failure */ }
        }

        /// <summary>
        /// Writes the count of any repeats still pending, for a shutdown that would otherwise lose it.
        /// </summary>
        public static void Flush()
        {
            try
            {
                lock (gate)
                {
                    WriteRepeatCount();
                }
            }
            catch { /* Nowhere left to report a logging failure */ }
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
            catch (Exception ex) { Write("Unhandled exception hookup failed: " + ex); }
        }

        /// <summary>
        /// Writes how many times the last message repeated, if it did, and resets the count. The caller holds the lock.
        /// </summary>
        private static void WriteRepeatCount()
        {
            if (repeatCount == 0) return;

            int count = repeatCount;
            repeatCount = 0;
            AppendLine($"Last message repeated {count} {(count == 1 ? "time" : "times")}");
        }

        /// <summary>
        /// Appends one timestamped line, trimming the file first if it has grown large. The caller holds the lock.
        /// </summary>
        private static void AppendLine(string line)
        {
            TrimIfLarge();
            File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} {line}{Environment.NewLine}");
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

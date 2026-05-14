using System;
using System.Diagnostics;

namespace PPTSafeCommands.Utilities
{
    /// <summary>
    /// Provides simple logging functionality.
    /// </summary>
    public class Logger
    {
        private readonly string _className;

        public Logger(Type type)
        {
            _className = type.Name;
        }

        /// <summary>
        /// Logs informational messages.
        /// </summary>
        public void Info(string message)
        {
            WriteLog("INFO", message);
        }

        /// <summary>
        /// Logs warning messages.
        /// </summary>
        public void Warn(string message)
        {
            WriteLog("WARN", message);
        }

        /// <summary>
        /// Logs error messages.
        /// </summary>
        public void Error(
            string message,
            Exception exception = null)
        {
            if (exception != null)
            {
                message =
                    message +
                    Environment.NewLine +
                    exception.Message +
                    Environment.NewLine +
                    exception.StackTrace;
            }

            WriteLog("ERROR", message);
        }

        /// <summary>
        /// Writes log output.
        /// </summary>
        private void WriteLog(
            string logLevel,
            string message)
        {
            string logMessage =
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} " +
                $"[{logLevel}] " +
                $"[{_className}] " +
                $"{message}";

            Debug.WriteLine(logMessage);

            Console.WriteLine(logMessage);
        }
    }
}
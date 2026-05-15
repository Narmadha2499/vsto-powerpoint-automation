using System;
using System.IO;
using System.Text;

namespace Sample_exercise.Utilities
{
    public sealed class Logger
    {
        private static readonly object LockObject =
            new object();

        private readonly string _loggerName;

        private static string _logFilePath;

        public Logger(Type type)
        {
            _loggerName = type.Name;
        }

        public static void Initialize(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        public static void LogError(Exception exception)
        {
            var logger = new Logger(typeof(Logger));
            logger.Error("An unhandled exception occurred", exception);
        }

        public void Info(string message)
        {
            WriteLog(
                "INFO",
                message);
        }

        public void Error(
            string message,
            Exception exception = null)
        {
            if (exception != null)
            {
                message +=
                    Environment.NewLine +
                    exception.Message +
                    Environment.NewLine +
                    exception.StackTrace;
            }

            WriteLog(
                "ERROR",
                message);
        }

        private void WriteLog(
            string level,
            string message)
        {
            if (string.IsNullOrWhiteSpace(_logFilePath))
            {
                return;
            }

            string logMessage =
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") +
                " | " +
                level +
                " | " +
                _loggerName +
                " | " +
                message;

            lock (LockObject)
            {
                File.AppendAllText(
                    _logFilePath,
                    logMessage + Environment.NewLine,
                    Encoding.UTF8);
            }
        }
    }
}
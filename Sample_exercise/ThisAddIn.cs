using Sample_exercise.Utilities;
using System;
using System.Diagnostics;
using System.IO;

namespace Sample_exercise
{
    public partial class ThisAddIn
    {
        private Logger _logger;

        private void ThisAddIn_Startup( object sender, EventArgs e)
        {
            try
            {
                InitializeLogger();

                _logger.Info(
                    "Sample_exercise Add-In started successfully.");
            }
            catch (Exception exception)
            {
                Debug.WriteLine(
                    "Startup failed: " +
                    exception.Message);
            }
        }

        private void ThisAddIn_Shutdown( object sender, EventArgs e)
        {
            try
            {
                _logger?.Info(
                    "Add-In shutdown initiated.");
            }
            catch (Exception exception)
            {
                Debug.WriteLine(
                    "Shutdown failed: " +
                    exception.Message);
            }
        }

        private void InitializeLogger()
        {
            string logFilePath =
                Path.Combine(
                    Path.GetTempPath(),
                    "SampleExerciseLog.txt");

            Logger.Initialize(logFilePath);

            _logger =
                new Logger(typeof(ThisAddIn));

            _logger.Info(
                "Logger initialized successfully.");
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;

            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}
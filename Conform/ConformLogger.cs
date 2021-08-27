using System;
using ASCOM.Standard.Interfaces;
using ASCOM.Standard.Utilities;
using static ConformU.Globals;

namespace ConformU
{
    public class ConformLogger : TraceLogger, ILogger
    {
        public ConformLogger(string logFileName, string logFilePath, string loggerName, bool enabled) : base(logFileName, logFilePath, loggerName, enabled)
        {
            Console.WriteLine($"ConformLogger - Creating log file - Name: {logFileName}, Path: {logFilePath}, Type: {loggerName}");
            base.IdentifierWidth = TEST_NAME_WIDTH;
        }

        public void LogMessage(string id, MessageLevel logLevel, string message)
        {
            string formattedMessage = logLevel.ToString().PadRight(9) + message;
            LogMessage(id, formattedMessage);
        }

        public void LogDebug(string id, string message)
        {
            if (base.LoggingLevel==LogLevel.Debug)
            {
                base.LogMessage(id, message);
            }
        }

        public new void LogMessage(string id, string message)
        {
            Console.WriteLine($"{id,-TEST_NAME_WIDTH} {message}");
            base.LogMessage(id, message);
        }

        public new void Log(LogLevel level, string message)
        {
            base.Log(level, message);
        }

    }
}
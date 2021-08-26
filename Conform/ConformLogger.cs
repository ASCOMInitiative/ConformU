using System;
using ASCOM.Standard.Interfaces;
using ASCOM.Standard.Utilities;

namespace ConformU
{
    public class ConformLogger : TraceLogger, ASCOM.Standard.Interfaces.ILogger
    {
        private const int ID_WIDTH = 30;

        public ConformLogger(string logFileName, string logFilePath, string loggerName, bool enabled) : base(logFileName, logFilePath, loggerName, enabled)
        {
            Debug = false;
            Console.WriteLine($"ConformLogger - Creating log file - Name: {logFileName}, Path: {logFilePath}, Type: {loggerName}");
            base.IdentifierWidth = ID_WIDTH;
        }

        public void LogMessage(string id, MessageLevel logLevel, string message)
        {
            string formattedMessage = logLevel.ToString().PadRight(9) + message;
            LogMessage(id, formattedMessage);
        }

        public void LogDebug(string id, string message)
        {
            if (Debug)
            {
                base.LogMessage(id, message);
            }
        }

        public bool Debug { get; set; }

        public new void LogMessage(string id, string message)
        {
            Console.WriteLine($"{id.PadRight(ID_WIDTH)} {message}");
            base.LogMessage(id, message);
        }

        public new void Log(LogLevel level, string message)
        {
            base.Log(level, message);
        }

    }
}
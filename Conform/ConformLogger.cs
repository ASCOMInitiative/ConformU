using System;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ASCOM.Standard.Utilities;
using static ConformU.ConformConstants;

namespace ConformU
{
    public class ConformLogger : TraceLogger
    {
        public ConformLogger(string logFileName, string logFilePath, string loggerName, bool enabled) : base(logFileName, logFilePath, loggerName, enabled)
        {
            Debug = false;
            Console.WriteLine($"CONFORMLOGGER INIT - Log file name: {logFileName}, Log file path: {logFilePath}, Logger name: {loggerName}");
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
            Console.WriteLine($"{id.PadRight(25)} {message}");
            base.LogMessage(id, message);
        }

    }
}
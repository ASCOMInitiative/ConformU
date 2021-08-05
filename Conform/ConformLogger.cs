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

namespace ConformU
{
    public class ConformLogger:TraceLogger
    {
        public ConformLogger(string loggerName, bool enabled):base(loggerName,enabled)
        {
            Debug = false;
        }

        public void LogDebug(string id, string message)
        {
            if (Debug)
            {
                base.LogMessage(id, message);
            }
        }

        public bool Debug { get; set; }
    }
}
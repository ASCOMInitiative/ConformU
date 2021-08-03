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

public class ConformLogger
{
    private readonly TraceLogger TL;

    public ConformLogger(string loggerName, bool enabled)
    {
        Console.WriteLine("***** Creating Trace Logger");
        TL = new TraceLogger(loggerName, enabled);
        TraceLogger = TL;

        Console.WriteLine("***** Trace Logger created OK");
        Debug = false;
    }

    public void LogMessage(string id, string message)
    {
        TL.LogMessage(id, message);
        Console.WriteLine($"##### {id} - {message}");
    }
    public TraceLogger TraceLogger { get; set; }
    public void LogDebug(string id, string message)
    {
        if (Debug)
        {
            TL.LogMessage(id, message);
            Console.WriteLine($"##### {id} - {message}");
        }
    }

    public bool Debug { get; set; }
}
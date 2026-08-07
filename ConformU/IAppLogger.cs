
using ASCOM.Common.Interfaces;
using System;
using ILogger = ASCOM.Common.Interfaces.ILogger;
using LogLevel = ASCOM.Common.Interfaces.LogLevel;

namespace ConformU
{
    public interface IAppLogger : ITraceLogger, ILogger, IDisposable
    {
        new void LogMessage(string method, string message);
        void LogDebug(string method, string message);

    }
}

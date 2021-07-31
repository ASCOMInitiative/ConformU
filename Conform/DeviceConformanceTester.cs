using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using System.Threading.Tasks;

public enum MessageLevel
{
    None = 0,
    Debug = 1,
    Comment = 2,
    Info = 3,
    OK = 4,
    Warning = 5,
    Issue = 6,
    Error = 7,
    Always = 8
}

public class DeviceConformanceTester : IDisposable
{
    public DeviceConformanceTester(ConformConfiguration configuration)
    {
    }

    public event MessageEventHandler OutputChanged;
    private void LogMessage(string id, string message)
    {
        if (OutputChanged is not null)
        {
            MessageEventArgs args = new()
            {
                Id = id,
                Message = message
            };
            OutputChanged(this, args);
        }
    }


    private bool disposedValue;

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            disposedValue = true;
        }
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ConformU
{
    public class DriverHost : IDisposable
    {
        readonly ConformLogger logger;
        readonly Settings conformSettings;
        dynamic driverObject;
        private bool disposedValue;

        public DriverHost(Settings conformSettings, ConformLogger logger)
        {
            this.conformSettings = conformSettings;
            this.logger = logger;

            try
            {
                logger?.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating ProgID: {conformSettings.ComDevice.ProgId}");
                Type driverType = Type.GetTypeFromProgID(conformSettings.ComDevice.ProgId);
                logger?.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating Type: {driverType}");
                driverObject = Activator.CreateInstance(driverType);
                // comment
            }
            catch (Exception ex)
            {
                logger.LogMessage("CreateDevice", MessageLevel.Error, $"Exception creating driver: {ex}");
                throw;
            }
        }
        public dynamic DriverObject
        {
            get
            {
                return driverObject;
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    int remainingObjectCount, loopCount;

                    try
                    {
                        if (conformSettings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.Debug, $"About to set Connected False.");
                        driverObject.Connected = false;
                    }
                    catch { }

                    try
                    {
                        if (conformSettings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                        driverObject.Dispose();
                    }
                    catch { }

                    try
                    {
                        loopCount = 0;
                        do
                        {
                            loopCount += 1;
                            remainingObjectCount = Marshal.ReleaseComObject(driverObject);
                            if (conformSettings.Debug) logger?.LogMessage("Dispose", MessageLevel.Debug, $"Released COM driver. Remaining object count: {remainingObjectCount}.");

                        }
                        while (remainingObjectCount > 0 & loopCount <= 20);
                    }
                    catch { }

                    try
                    {
                        driverObject = null;
                        GC.Collect();
                    }
                    catch { }

                }
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
}

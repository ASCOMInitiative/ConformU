using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ASCOM.Standard.AlpacaClients;

namespace ConformU
{
    public class SafetyMonitorFacade : ISafetyMonitor, IDisposable
    {
        private bool disposedValue;

        private dynamic driver; // COM driver object
        private readonly Settings settings; // Conform configuration settings
        private readonly ILogger logger;

        public SafetyMonitorFacade(Settings conformSettings, ILogger logger)
        {
            settings = conformSettings;
            this.logger = logger;
        }

        public void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case ConformConstants.DeviceTechnology.Alpaca:
                        driver = new SafetyMonitor("http", settings.AlpacaDevice.IpAddress, settings.AlpacaDevice.IpPort, settings.AlpacaDevice.AlpacaDeviceNumber, logger);
                        break;

                    case ConformConstants.DeviceTechnology.COM:
                        Type driverType = Type.GetTypeFromProgID(settings.ComDevice.ProgId);
                        driver = Activator.CreateInstance(driverType);
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"TelescopeFacade:CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"TelescopeFacade:CreateDevice - Exception: {ex}");
            }

        }

        public bool IsSafe => throw new NotImplementedException();

        public bool Connected { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public string Description => throw new NotImplementedException();

        public string DriverInfo => throw new NotImplementedException();

        public string DriverVersion => throw new NotImplementedException();

        public short InterfaceVersion => throw new NotImplementedException();

        public string Name => throw new NotImplementedException();

        public IList<string> SupportedActions => throw new NotImplementedException();

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                disposedValue = true;
            }
        }

        // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
        // ~SafetyMonitorFacade()
        // {
        //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        public string Action(string ActionName, string ActionParameters)
        {
            throw new NotImplementedException();
        }

        public void CommandBlind(string Command, bool Raw = false)
        {
            throw new NotImplementedException();
        }

        public bool CommandBool(string Command, bool Raw = false)
        {
            throw new NotImplementedException();
        }

        public string CommandString(string Command, bool Raw = false)
        {
            throw new NotImplementedException();
        }
    }
}

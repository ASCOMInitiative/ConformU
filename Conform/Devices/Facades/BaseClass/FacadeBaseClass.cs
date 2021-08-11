using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using static ConformU.ConformConstants;

namespace ConformU
{
    public class FacadeBaseClass : IDisposable
    {
        internal dynamic driver; // COM driver object
        internal readonly Settings settings; // Conform configuration settings
        internal readonly ConformLogger logger;
        private bool disposedValue;

        #region New and Dispose

        public FacadeBaseClass(Settings conformSettings, ConformLogger logger)
        {
            settings = conformSettings;
            this.logger = logger;
            try
            {
                logger?.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating ProgID: {settings.ComDevice.ProgId}");
                Type driverType = Type.GetTypeFromProgID(settings.ComDevice.ProgId);
                logger?.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating Type: {driverType}");
                driver = Activator.CreateInstance(driverType);
            }
            catch (Exception ex)
            {
                logger.LogMessage("CreateDevice", MessageLevel.Error, $"Exception creating driver: {ex}");
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    switch (settings.DeviceTechnology)
                    {
                        case DeviceTechnology.Alpaca:
                            try
                            {
                                if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.Debug, $"About to set Connected False.");
                                driver.Connected = false;
                            }
                            catch { }

                            try
                            {
                                if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                                driver.Dispose();
                            }
                            catch { }
                            break;

                        case DeviceTechnology.COM:
                            int remainingObjectCount, loopCount;

                            try
                            {
                                if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.Debug, $"About to set Connected False.");
                                driver.Connected = false;
                            }
                            catch { }

                            try
                            {
                                if (settings.DisplayMethodCalls) logger?.LogMessage("Dispose", MessageLevel.Debug, $"About to call Dispose method.");
                                driver.Dispose();
                            }
                            catch { }

                            try
                            {
                                loopCount = 0;
                                do
                                {
                                    loopCount += 1;
                                    remainingObjectCount = Marshal.ReleaseComObject(driver);
                                    if (settings.Debug) logger?.LogMessage("Dispose", MessageLevel.Debug, $"Released COM driver. Remaining object count: {remainingObjectCount}.");

                                }
                                while (!(remainingObjectCount <= 0 | loopCount == 20));
                            }
                            catch { }

                            try
                            {
                                driver = null;
                                GC.Collect();
                            }
                            catch { }

                            break;
                    }
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put clean-up code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        #endregion

        #region Common Members

        public bool Connected { get => driver.Connected; set => driver.Connected = value; }

        public string Description => driver.Description;

        public string DriverInfo => driver.DriverInfo;

        public string DriverVersion => driver.DriverVersion;

        public short InterfaceVersion => driver.InterfaceVersion;

        public string Name => driver.Name;

        public IList<string> SupportedActions
        {
            get
            {
                List<string> supportedActions = new();
                foreach (string action in driver.SupportedActions)
                {
                    supportedActions.Add(action);
                }
                return supportedActions;
            }
        }

        public string Action(string ActionName, string ActionParameters)
        {
            return driver.Action(ActionName, ActionParameters);
        }

        public void CommandBlind(string Command, bool Raw = false)
        {
            driver.CommandBlind(Command, Raw);
        }

        public bool CommandBool(string Command, bool Raw = false)
        {
            return driver.CommandBool(Command, Raw);
        }

        public string CommandString(string Command, bool Raw = false)
        {
            return driver.CommandString(Command, Raw);
        }

        #endregion

    }
}
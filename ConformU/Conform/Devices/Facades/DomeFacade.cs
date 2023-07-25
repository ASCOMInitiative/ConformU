using ASCOM.Common.DeviceInterfaces;
#if WINDOWS
using System.Windows.Forms;
#endif


namespace ConformU
{
    public class DomeFacade : FacadeBaseClass, IDomeV2
    {
        // Create the test device in the facade base class
        public DomeFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger)
        {
#if WINDOWS
            Form f = new();
#endif
        }

        #region Interface implementation

        public double Altitude
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Altitude);
            }
        }

        public bool AtHome
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.AtHome);
            }
        }

        public bool AtPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.AtPark);
            }
        }

        public double Azimuth
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Azimuth);
            }
        }

        public bool CanFindHome
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanFindHome);
            }
        }

        public bool CanPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanPark);
            }
        }

        public bool CanSetAltitude
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetAltitude);
            }
        }

        public bool CanSetAzimuth
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetAzimuth);
            }
        }

        public bool CanSetPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetPark);
            }
        }

        public bool CanSetShutter
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetShutter);
            }
        }

        public bool CanSlave
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSlave);
            }
        }

        public bool CanSyncAzimuth
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSyncAzimuth);
            }
        }

        public ShutterState ShutterStatus
        {
            get
            {
                return (ShutterState)FunctionNoParameters<object>(() => driver.ShutterStatus);
            }
        }

        public bool Slaved { get => FunctionNoParameters<bool>(() => driver.Slaved); set => Method1Parameter((i) => driver.Slaved = i, value); }

        public bool Slewing
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.Slewing);
            }
        }

        public void AbortSlew()
        {
            MethodNoParameters(() => driver.AbortSlew());
        }

        public void CloseShutter()
        {
            MethodNoParameters(() => driver.CloseShutter());
        }

        public void FindHome()
        {
            MethodNoParameters(() => driver.FindHome());
        }

        public void OpenShutter()
        {
            MethodNoParameters(() => driver.OpenShutter());
        }

        public void Park()
        {
            MethodNoParameters(() => driver.Park());
        }

        public void SetPark()
        {
            MethodNoParameters(() => driver.SetPark());
        }

        public void SlewToAltitude(double Altitude)
        {
            Method1Parameter((i) => driver.SlewToAltitude(i), Altitude);
        }

        public void SlewToAzimuth(double Azimuth)
        {
            Method1Parameter((i) => driver.SlewToAzimuth(i), Azimuth);
        }

        public void SyncToAzimuth(double Azimuth)
        {
            Method1Parameter((i) => driver.SyncToAzimuth(i), Azimuth);
        }

        #endregion

    }
}

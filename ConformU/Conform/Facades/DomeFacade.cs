using ASCOM.Common.DeviceInterfaces;
#if WINDOWS
using System.Windows.Forms;
#endif


namespace ConformU
{
    public class DomeFacade : FacadeBaseClass, IDomeV3
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
                return FunctionNoParameters<double>(() => Driver.Altitude);
            }
        }

        public bool AtHome
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.AtHome);
            }
        }

        public bool AtPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.AtPark);
            }
        }

        public double Azimuth
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.Azimuth);
            }
        }

        public bool CanFindHome
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanFindHome);
            }
        }

        public bool CanPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanPark);
            }
        }

        public bool CanSetAltitude
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetAltitude);
            }
        }

        public bool CanSetAzimuth
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetAzimuth);
            }
        }

        public bool CanSetPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetPark);
            }
        }

        public bool CanSetShutter
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetShutter);
            }
        }

        public bool CanSlave
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSlave);
            }
        }

        public bool CanSyncAzimuth
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSyncAzimuth);
            }
        }

        public ShutterState ShutterStatus
        {
            get
            {
                return (ShutterState)FunctionNoParameters<object>(() => Driver.ShutterStatus);
            }
        }

        public bool Slaved { get => FunctionNoParameters<bool>(() => Driver.Slaved); set => Method1Parameter((i) => Driver.Slaved = i, value); }

        public bool Slewing
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.Slewing);
            }
        }

        public void AbortSlew()
        {
            MethodNoParameters(() => Driver.AbortSlew());
        }

        public void CloseShutter()
        {
            MethodNoParameters(() => Driver.CloseShutter());
        }

        public void FindHome()
        {
            MethodNoParameters(() => Driver.FindHome());
        }

        public void OpenShutter()
        {
            MethodNoParameters(() => Driver.OpenShutter());
        }

        public void Park()
        {
            MethodNoParameters(() => Driver.Park());
        }

        public void SetPark()
        {
            MethodNoParameters(() => Driver.SetPark());
        }

        public void SlewToAltitude(double altitude)
        {
            Method1Parameter((i) => Driver.SlewToAltitude(i), altitude);
        }

        public void SlewToAzimuth(double azimuth)
        {
            Method1Parameter((i) => Driver.SlewToAzimuth(i), azimuth);
        }

        public void SyncToAzimuth(double azimuth)
        {
            Method1Parameter((i) => Driver.SyncToAzimuth(i), azimuth);
        }

        #endregion

    }
}

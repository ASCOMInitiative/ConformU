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

        public double Altitude => (double)FunctionNoParameters(() => driver.Altitude);

        public bool AtHome => (bool)FunctionNoParameters(() => driver.AtHome);

        public bool AtPark => (bool)FunctionNoParameters(() => driver.AtPark);

        public double Azimuth => (double)FunctionNoParameters(() => driver.Azimuth);

        public bool CanFindHome => (bool)FunctionNoParameters(() => driver.CanFindHome);

        public bool CanPark => (bool)FunctionNoParameters(() => driver.CanPark);

        public bool CanSetAltitude => (bool)FunctionNoParameters(() => driver.CanSetAltitude);

        public bool CanSetAzimuth => (bool)FunctionNoParameters(() => driver.CanSetAzimuth);

        public bool CanSetPark => (bool)FunctionNoParameters(() => driver.CanSetPark);

        public bool CanSetShutter => (bool)FunctionNoParameters(() => driver.CanSetShutter);

        public bool CanSlave => (bool)FunctionNoParameters(() => driver.CanSlave);

        public bool CanSyncAzimuth => (bool)FunctionNoParameters(() => driver.CanSyncAzimuth);

        public ShutterState ShutterStatus => (ShutterState)FunctionNoParameters(() => driver.ShutterStatus);

        public bool Slaved { get => (bool)FunctionNoParameters(() => driver.Slaved); set => Method1Parameter((i) => driver.Slaved = i, value); }

        public bool Slewing => (bool)FunctionNoParameters(() => driver.Slewing);

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

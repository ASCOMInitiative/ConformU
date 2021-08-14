using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class DomeFacade : FacadeBaseClass, IDomeV2
    {
        // Create the test device in the facade base class
        public DomeFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public double Altitude => driver.Altitude;

        public bool AtHome => driver.AtHome;

        public bool AtPark => driver.AtPark;

        public double Azimuth => driver.Azimuth;

        public bool CanFindHome => driver.CanFindHome;

        public bool CanPark => driver.CanPark;

        public bool CanSetAltitude => driver.CanSetAltitude;

        public bool CanSetAzimuth => driver.CanSetAzimuth;

        public bool CanSetPark => driver.CanSetPark;

        public bool CanSetShutter => driver.CanSetShutter;

        public bool CanSlave => driver.CanSlave;

        public bool CanSyncAzimuth => driver.CanSyncAzimuth;

        public ShutterState ShutterStatus => (ShutterState)driver.ShutterStatus;

        public bool Slaved { get => driver.Slaved; set => driver.Slaved = value; }

        public bool Slewing => driver.Slewing;

        public void AbortSlew()
        {
            driver.AbortSlew();
        }

        public void CloseShutter()
        {
            driver.CloseShutter();
        }

        public void FindHome()
        {
            driver.FindHome();
        }

        public void OpenShutter()
        {
            driver.OpenShutter();
        }

        public void Park()
        {
            driver.Park();
        }

        public void SetPark()
        {
            driver.SetPark();
        }

        public void SlewToAltitude(double Altitude)
        {
            driver.SlewToAltitude(Altitude);
        }

        public void SlewToAzimuth(double Azimuth)
        {
            driver.SlewToAzimuth(Azimuth);
        }

        public void SyncToAzimuth(double Azimuth)
        {
            driver.SyncToAzimuth(Azimuth);
        }

        #endregion

    }
}

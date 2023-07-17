using ASCOM.Common.DeviceInterfaces;
using System;
/* Unmerged change from project 'ConformU (net5.0)'
Before:
using System.Runtime.InteropServices;
using ASCOM.Alpaca.Clients;
After:
using System.Runtime.InteropServices;
*/


namespace ConformU
{
    public class TelescopeFacade : FacadeBaseClass, ITelescopeV4, IDisposable
    {

        // Create the test device in the facade base class
        public TelescopeFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region ITelescopeV3 interface implementation

        public AlignmentMode AlignmentMode => (AlignmentMode)FunctionNoParameters(() => driver.AlignmentMode);

        public double Altitude => (double)FunctionNoParameters(() => driver.Altitude);

        public double ApertureArea => (double)FunctionNoParameters(() => driver.ApertureArea);

        public double ApertureDiameter => (double)FunctionNoParameters(() => driver.ApertureDiameter);

        public bool AtHome => (bool)FunctionNoParameters(() => driver.AtHome);

        public bool AtPark => (bool)FunctionNoParameters(() => driver.AtPark);

        public double Azimuth => (double)FunctionNoParameters(() => driver.Azimuth);

        public bool CanFindHome => (bool)FunctionNoParameters(() => driver.CanFindHome);

        public bool CanPark => (bool)FunctionNoParameters(() => driver.CanPark);

        public bool CanPulseGuide => (bool)FunctionNoParameters(() => driver.CanPulseGuide);

        public bool CanSetDeclinationRate => (bool)FunctionNoParameters(() => driver.CanSetDeclinationRate);

        public bool CanSetGuideRates => (bool)FunctionNoParameters(() => driver.CanSetGuideRates);

        public bool CanSetPark => (bool)FunctionNoParameters(() => driver.CanSetPark);

        public bool CanSetPierSide => (bool)FunctionNoParameters(() => driver.CanSetPierSide);

        public bool CanSetRightAscensionRate => (bool)FunctionNoParameters(() => driver.CanSetRightAscensionRate);

        public bool CanSetTracking => (bool)FunctionNoParameters(() => driver.CanSetTracking);

        public bool CanSlew => (bool)FunctionNoParameters(() => driver.CanSlew);

        public bool CanSlewAltAz => (bool)FunctionNoParameters(() => driver.CanSlewAltAz);

        public bool CanSlewAltAzAsync => (bool)FunctionNoParameters(() => driver.CanSlewAltAzAsync);

        public bool CanSlewAsync => (bool)FunctionNoParameters(() => driver.CanSlewAsync);

        public bool CanSync => (bool)FunctionNoParameters(() => driver.CanSync);

        public bool CanSyncAltAz => (bool)FunctionNoParameters(() => driver.CanSyncAltAz);

        public bool CanUnpark => (bool)FunctionNoParameters(() => driver.CanUnpark);

        public double Declination => (double)FunctionNoParameters(() => driver.Declination);

        public double DeclinationRate { get => (double)FunctionNoParameters(() => driver.DeclinationRate); set => Method1Parameter((i) => driver.DeclinationRate = i, value); }
        public bool DoesRefraction { get => (bool)FunctionNoParameters(() => driver.DoesRefraction); set => Method1Parameter((i) => driver.DoesRefraction = i, value); }

        public EquatorialCoordinateType EquatorialSystem => (EquatorialCoordinateType)FunctionNoParameters(() => driver.EquatorialSystem);

        public double FocalLength => (double)FunctionNoParameters(() => driver.FocalLength);

        public double GuideRateDeclination { get => (double)FunctionNoParameters(() => driver.GuideRateDeclination); set => Method1Parameter((i) => driver.GuideRateDeclination = i, value); }
        public double GuideRateRightAscension { get => (double)FunctionNoParameters(() => driver.GuideRateRightAscension); set => Method1Parameter((i) => driver.GuideRateRightAscension = i, value); }

        public bool IsPulseGuiding => (bool)FunctionNoParameters(() => driver.IsPulseGuiding);

        public double RightAscension => (double)FunctionNoParameters(() => driver.RightAscension);

        public double RightAscensionRate { get => (double)FunctionNoParameters(() => driver.RightAscensionRate); set => Method1Parameter((i) => driver.RightAscensionRate = i, value); }
        public PointingState SideOfPier { get => (PointingState)FunctionNoParameters(() => driver.SideOfPier); set => Method1Parameter((i) => driver.SideOfPier = i, value); }

        public double SiderealTime => (double)FunctionNoParameters(() => driver.SiderealTime);

        public double SiteElevation { get => (double)FunctionNoParameters(() => driver.SiteElevation); set => Method1Parameter((i) => driver.SiteElevation = i, value); }
        public double SiteLatitude { get => (double)FunctionNoParameters(() => driver.SiteLatitude); set => Method1Parameter((i) => driver.SiteLatitude = i, value); }
        public double SiteLongitude { get => (double)FunctionNoParameters(() => driver.SiteLongitude); set => Method1Parameter((i) => driver.SiteLongitude = i, value); }

        public bool Slewing => (bool)FunctionNoParameters(() => driver.Slewing);

        public short SlewSettleTime { get => (short)FunctionNoParameters(() => driver.SlewSettleTime); set => Method1Parameter((i) => driver.SlewSettleTime = i, value); }
        public double TargetDeclination { get => (double)FunctionNoParameters(() => driver.TargetDeclination); set => Method1Parameter((i) => driver.TargetDeclination = i, value); }
        public double TargetRightAscension { get => (double)FunctionNoParameters(() => driver.TargetRightAscension); set => Method1Parameter((i) => driver.TargetRightAscension = i, value); }
        public bool Tracking { get => (bool)FunctionNoParameters(() => driver.Tracking); set => Method1Parameter((i) => driver.Tracking = i, value); }
        public DriveRate TrackingRate { get => (DriveRate)FunctionNoParameters(() => driver.TrackingRate); set => Method1Parameter((i) => driver.TrackingRate = i, value); }

        public ITrackingRates TrackingRates
        {
            get
            {
                return new TrackingRatesFacade(driver, this);
            }
        }

        public DateTime UTCDate { get => (DateTime)FunctionNoParameters(() => driver.UTCDate); set => Method1Parameter((i) => driver.UTCDate = i, value); }

        public void AbortSlew()
        {
            MethodNoParameters(() => driver.AbortSlew());
        }

        public IAxisRates AxisRates(TelescopeAxis Axis)
        {
            return new AxisRatesFacade(Axis, driver, this, logger);
        }

        public bool CanMoveAxis(TelescopeAxis Axis)
        {
            return (bool)Function1Parameter((i) => driver.CanMoveAxis(i), Axis);
        }

        public PointingState DestinationSideOfPier(double RightAscension, double Declination)
        {
            return (PointingState)Function2Parameters((i, j) => driver.DestinationSideOfPier(i, j), RightAscension, Declination);
        }

        public void FindHome()
        {
            MethodNoParameters(() => driver.FindHome());
        }

        public void MoveAxis(TelescopeAxis Axis, double Rate)
        {
            Method2Parameters((i, j) => driver.MoveAxis(i, j), Axis, Rate);
        }

        public void Park()
        {
            MethodNoParameters(() => driver.Park());
        }

        public void PulseGuide(GuideDirection Direction, int Duration)
        {
            Method2Parameters((i, j) => driver.PulseGuide(i, j), Direction, Duration);
        }

        public void SetPark()
        {
            MethodNoParameters(() => driver.SetPark());
        }

        public void SlewToAltAz(double Azimuth, double Altitude)
        {
            Method2Parameters((i, j) => driver.SlewToAltAz(i, j), Azimuth, Altitude);
        }

        public void SlewToAltAzAsync(double Azimuth, double Altitude)
        {
            Method2Parameters((i, j) => driver.SlewToAltAzAsync(i, j), Azimuth, Altitude);
        }

        public void SlewToCoordinates(double RightAscension, double Declination)
        {
            Method2Parameters((i, j) => driver.SlewToCoordinates(i, j), RightAscension, Declination);
        }

        public void SlewToCoordinatesAsync(double RightAscension, double Declination)
        {
            Method2Parameters((i, j) => driver.SlewToCoordinatesAsync(i, j), RightAscension, Declination);
        }

        public void SlewToTarget()
        {
            MethodNoParameters(() => driver.SlewToTarget());
        }

        public void SlewToTargetAsync()
        {
            MethodNoParameters(() => driver.SlewToTargetAsync());
        }

        public void SyncToAltAz(double Azimuth, double Altitude)
        {
            Method2Parameters((i, j) => driver.SyncToAltAz(i, j), Azimuth, Altitude);
        }

        public void SyncToCoordinates(double RightAscension, double Declination)
        {
            Method2Parameters((i, j) => driver.SyncToCoordinates(i, j), RightAscension, Declination);
        }

        public void SyncToTarget()
        {
            MethodNoParameters(() => driver.SyncToTarget());
        }

        public void Unpark()
        {
            MethodNoParameters(() => driver.Unpark());
        }

        #endregion

        #region ITelescopeV4 interface implementation

        public bool OperationComplete => (bool)FunctionNoParameters(() => driver.OperationComplete);

        public bool InterruptionComplete => (bool)FunctionNoParameters(() => driver.InterruptionComplete); //InterruptionComplete
                                                                                                           //InterruptionComplete

        #endregion

    }
}

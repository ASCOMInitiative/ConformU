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

        public AlignmentMode AlignmentMode
        {
            get
            {
                return (AlignmentMode)FunctionNoParameters<object>(() => driver.AlignmentMode);
            }
        }

        public double Altitude
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Altitude);
            }
        }

        public double ApertureArea
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.ApertureArea);
            }
        }

        public double ApertureDiameter
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.ApertureDiameter);
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

        public bool CanPulseGuide
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanPulseGuide);
            }
        }

        public bool CanSetDeclinationRate
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetDeclinationRate);
            }
        }

        public bool CanSetGuideRates
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetGuideRates);
            }
        }

        public bool CanSetPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetPark);
            }
        }

        public bool CanSetPierSide
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetPierSide);
            }
        }

        public bool CanSetRightAscensionRate
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetRightAscensionRate);
            }
        }

        public bool CanSetTracking
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetTracking);
            }
        }

        public bool CanSlew
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSlew);
            }
        }

        public bool CanSlewAltAz
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSlewAltAz);
            }
        }

        public bool CanSlewAltAzAsync
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSlewAltAzAsync);
            }
        }

        public bool CanSlewAsync
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSlewAsync);
            }
        }

        public bool CanSync
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSync);
            }
        }

        public bool CanSyncAltAz
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSyncAltAz);
            }
        }

        public bool CanUnpark
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanUnpark);
            }
        }

        public double Declination
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Declination);
            }
        }

        public double DeclinationRate { get => FunctionNoParameters<double>(() => driver.DeclinationRate); set => Method1Parameter((i) => driver.DeclinationRate = i, value); }
        public bool DoesRefraction { get => FunctionNoParameters<bool>(() => driver.DoesRefraction); set => Method1Parameter((i) => driver.DoesRefraction = i, value); }

        public EquatorialCoordinateType EquatorialSystem
        {
            get
            {
                return (EquatorialCoordinateType)FunctionNoParameters<object>(() => driver.EquatorialSystem);
            }
        }

        public double FocalLength
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.FocalLength);
            }
        }

        public double GuideRateDeclination { get => FunctionNoParameters<double>(() => driver.GuideRateDeclination); set => Method1Parameter((i) => driver.GuideRateDeclination = i, value); }
        public double GuideRateRightAscension { get => FunctionNoParameters<double>(() => driver.GuideRateRightAscension); set => Method1Parameter((i) => driver.GuideRateRightAscension = i, value); }

        public bool IsPulseGuiding
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.IsPulseGuiding);
            }
        }

        public double RightAscension
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.RightAscension);
            }
        }

        public double RightAscensionRate { get => FunctionNoParameters<double>(() => driver.RightAscensionRate); set => Method1Parameter((i) => driver.RightAscensionRate = i, value); }
        public PointingState SideOfPier { get => (PointingState)FunctionNoParameters<object>(() => driver.SideOfPier); set => Method1Parameter((i) => driver.SideOfPier = i, value); }

        public double SiderealTime
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.SiderealTime);
            }
        }

        public double SiteElevation { get => FunctionNoParameters<double>(() => driver.SiteElevation); set => Method1Parameter((i) => driver.SiteElevation = i, value); }
        public double SiteLatitude { get => FunctionNoParameters<double>(() => driver.SiteLatitude); set => Method1Parameter((i) => driver.SiteLatitude = i, value); }
        public double SiteLongitude { get => FunctionNoParameters<double>(() => driver.SiteLongitude); set => Method1Parameter((i) => driver.SiteLongitude = i, value); }

        public bool Slewing
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.Slewing);
            }
        }

        public short SlewSettleTime { get => FunctionNoParameters<short>(() => driver.SlewSettleTime); set => Method1Parameter((i) => driver.SlewSettleTime = i, value); }
        public double TargetDeclination { get => FunctionNoParameters<double>(() => driver.TargetDeclination); set => Method1Parameter((i) => driver.TargetDeclination = i, value); }
        public double TargetRightAscension { get => FunctionNoParameters<double>(() => driver.TargetRightAscension); set => Method1Parameter((i) => driver.TargetRightAscension = i, value); }
        public bool Tracking { get => FunctionNoParameters<bool>(() => driver.Tracking); set => Method1Parameter((i) => driver.Tracking = i, value); }
        public DriveRate TrackingRate { get => (DriveRate)FunctionNoParameters<object>(() => driver.TrackingRate); set => Method1Parameter((i) => driver.TrackingRate = i, value); }

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
            return Function1Parameter<bool>((i) => driver.CanMoveAxis(i), Axis);
        }

        public PointingState DestinationSideOfPier(double RightAscension, double Declination)
        {
            return Function2Parameters<PointingState>((i, j) => driver.DestinationSideOfPier(i, j), RightAscension, Declination);
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

        #endregion

    }
}

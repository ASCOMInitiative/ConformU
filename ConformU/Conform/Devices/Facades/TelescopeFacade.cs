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
                return (AlignmentMode)FunctionNoParameters<object>(() => Driver.AlignmentMode);
            }
        }

        public double Altitude
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.Altitude);
            }
        }

        public double ApertureArea
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.ApertureArea);
            }
        }

        public double ApertureDiameter
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.ApertureDiameter);
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

        public bool CanPulseGuide
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanPulseGuide);
            }
        }

        public bool CanSetDeclinationRate
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetDeclinationRate);
            }
        }

        public bool CanSetGuideRates
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetGuideRates);
            }
        }

        public bool CanSetPark
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetPark);
            }
        }

        public bool CanSetPierSide
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetPierSide);
            }
        }

        public bool CanSetRightAscensionRate
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetRightAscensionRate);
            }
        }

        public bool CanSetTracking
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetTracking);
            }
        }

        public bool CanSlew
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSlew);
            }
        }

        public bool CanSlewAltAz
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSlewAltAz);
            }
        }

        public bool CanSlewAltAzAsync
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSlewAltAzAsync);
            }
        }

        public bool CanSlewAsync
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSlewAsync);
            }
        }

        public bool CanSync
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSync);
            }
        }

        public bool CanSyncAltAz
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSyncAltAz);
            }
        }

        public bool CanUnpark
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanUnpark);
            }
        }

        public double Declination
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.Declination);
            }
        }

        public double DeclinationRate { get => FunctionNoParameters<double>(() => Driver.DeclinationRate); set => Method1Parameter((i) => Driver.DeclinationRate = i, value); }
        public bool DoesRefraction { get => FunctionNoParameters<bool>(() => Driver.DoesRefraction); set => Method1Parameter((i) => Driver.DoesRefraction = i, value); }

        public EquatorialCoordinateType EquatorialSystem
        {
            get
            {
                return (EquatorialCoordinateType)FunctionNoParameters<object>(() => Driver.EquatorialSystem);
            }
        }

        public double FocalLength
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.FocalLength);
            }
        }

        public double GuideRateDeclination { get => FunctionNoParameters<double>(() => Driver.GuideRateDeclination); set => Method1Parameter((i) => Driver.GuideRateDeclination = i, value); }
        public double GuideRateRightAscension { get => FunctionNoParameters<double>(() => Driver.GuideRateRightAscension); set => Method1Parameter((i) => Driver.GuideRateRightAscension = i, value); }

        public bool IsPulseGuiding
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.IsPulseGuiding);
            }
        }

        public double RightAscension
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.RightAscension);
            }
        }

        public double RightAscensionRate { get => FunctionNoParameters<double>(() => Driver.RightAscensionRate); set => Method1Parameter((i) => Driver.RightAscensionRate = i, value); }
        public PointingState SideOfPier { get => (PointingState)FunctionNoParameters<object>(() => Driver.SideOfPier); set => Method1Parameter((i) => Driver.SideOfPier = i, value); }

        public double SiderealTime
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.SiderealTime);
            }
        }

        public double SiteElevation { get => FunctionNoParameters<double>(() => Driver.SiteElevation); set => Method1Parameter((i) => Driver.SiteElevation = i, value); }
        public double SiteLatitude { get => FunctionNoParameters<double>(() => Driver.SiteLatitude); set => Method1Parameter((i) => Driver.SiteLatitude = i, value); }
        public double SiteLongitude { get => FunctionNoParameters<double>(() => Driver.SiteLongitude); set => Method1Parameter((i) => Driver.SiteLongitude = i, value); }

        public bool Slewing
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.Slewing);
            }
        }

        public short SlewSettleTime { get => FunctionNoParameters<short>(() => Driver.SlewSettleTime); set => Method1Parameter((i) => Driver.SlewSettleTime = i, value); }
        public double TargetDeclination { get => FunctionNoParameters<double>(() => Driver.TargetDeclination); set => Method1Parameter((i) => Driver.TargetDeclination = i, value); }
        public double TargetRightAscension { get => FunctionNoParameters<double>(() => Driver.TargetRightAscension); set => Method1Parameter((i) => Driver.TargetRightAscension = i, value); }
        public bool Tracking { get => FunctionNoParameters<bool>(() => Driver.Tracking); set => Method1Parameter((i) => Driver.Tracking = i, value); }
        public DriveRate TrackingRate { get => (DriveRate)FunctionNoParameters<object>(() => Driver.TrackingRate); set => Method1Parameter((i) => Driver.TrackingRate = i, value); }

        public ITrackingRates TrackingRates
        {
            get
            {
                return new TrackingRatesFacade(Driver, this);
            }
        }

        public DateTime UTCDate { get => (DateTime)FunctionNoParameters(() => Driver.UTCDate); set => Method1Parameter((i) => Driver.UTCDate = i, value); }

        public void AbortSlew()
        {
            MethodNoParameters(() => Driver.AbortSlew());
        }

        public IAxisRates AxisRates(TelescopeAxis axis)
        {
            return new AxisRatesFacade(axis, Driver, this, Logger);
        }

        public bool CanMoveAxis(TelescopeAxis axis)
        {
            return Function1Parameter<bool>((i) => Driver.CanMoveAxis(i), axis);
        }

        public PointingState DestinationSideOfPier(double rightAscension, double declination)
        {
            return Function2Parameters<PointingState>((i, j) => Driver.DestinationSideOfPier(i, j), rightAscension, declination);
        }

        public void FindHome()
        {
            MethodNoParameters(() => Driver.FindHome());
        }

        public void MoveAxis(TelescopeAxis axis, double rate)
        {
            Method2Parameters((i, j) => Driver.MoveAxis(i, j), axis, rate);
        }

        public void Park()
        {
            MethodNoParameters(() => Driver.Park());
        }

        public void PulseGuide(GuideDirection direction, int duration)
        {
            Method2Parameters((i, j) => Driver.PulseGuide(i, j), direction, duration);
        }

        public void SetPark()
        {
            MethodNoParameters(() => Driver.SetPark());
        }

        public void SlewToAltAz(double azimuth, double altitude)
        {
            Method2Parameters((i, j) => Driver.SlewToAltAz(i, j), azimuth, altitude);
        }

        public void SlewToAltAzAsync(double azimuth, double altitude)
        {
            Method2Parameters((i, j) => Driver.SlewToAltAzAsync(i, j), azimuth, altitude);
        }

        public void SlewToCoordinates(double rightAscension, double declination)
        {
            Method2Parameters((i, j) => Driver.SlewToCoordinates(i, j), rightAscension, declination);
        }

        public void SlewToCoordinatesAsync(double rightAscension, double declination)
        {
            Method2Parameters((i, j) => Driver.SlewToCoordinatesAsync(i, j), rightAscension, declination);
        }

        public void SlewToTarget()
        {
            MethodNoParameters(() => Driver.SlewToTarget());
        }

        public void SlewToTargetAsync()
        {
            MethodNoParameters(() => Driver.SlewToTargetAsync());
        }

        public void SyncToAltAz(double azimuth, double altitude)
        {
            Method2Parameters((i, j) => Driver.SyncToAltAz(i, j), azimuth, altitude);
        }

        public void SyncToCoordinates(double rightAscension, double declination)
        {
            Method2Parameters((i, j) => Driver.SyncToCoordinates(i, j), rightAscension, declination);
        }

        public void SyncToTarget()
        {
            MethodNoParameters(() => Driver.SyncToTarget());
        }

        public void Unpark()
        {
            MethodNoParameters(() => Driver.Unpark());
        }

        #endregion

        #region ITelescopeV4 interface implementation

        #endregion

    }
}

using ASCOM;
using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.Utilities;

namespace ConformU
{
    public class TelescopeFacade : ITelescopeV3, IDisposable
    {
        private dynamic driver; // COM driver object
        private Settings settings; // Conform configuration settings

        public TelescopeFacade(Settings conformSettings)
        {
            settings = conformSettings;
        }

        public void CreateDevice()
        {
            try
            {
                switch (settings.CurrentDeviceTechnology)
                {
                    case ConformConstants.TECHNOLOGY_ALPACA:
                        driver = new Telescope("http", "127.0.0.1", 11111, 0, new TraceLogger("TelescopeFacade", true));
                        break;

                    case ConformConstants.TECHNOLOGY_COM:
                        Type driverType = Type.GetTypeFromProgID(settings.CurrentComDevice.ProgId);
                        driver = Activator.CreateInstance(driverType);
                        break;

                    default:
                        throw new InvalidValueException($"TelescopeFacade:CreateDevice - Unknown technology type: {settings.CurrentDeviceTechnology}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"TelescopeFacade:CreateDevice - Exception: {ex}");
            }

        }

        public AlignmentMode AlignmentMode => (AlignmentMode)driver.AlignmentMode;

        public double Altitude => driver.Altitude;

        public double ApertureArea => driver.ApertureArea;

        public double ApertureDiameter => driver.ApertureDiameter;

        public bool AtHome => driver.AtHome;

        public bool AtPark => driver.AtPark;

        public double Azimuth => driver.Azimuth;

        public bool CanFindHome => driver.CanFindHome;

        public bool CanPark => driver.CanPark;

        public bool CanPulseGuide => driver.CanPulseGuide;

        public bool CanSetDeclinationRate => driver.CanSetDeclinationRate;

        public bool CanSetGuideRates => driver.CanSetGuideRates;

        public bool CanSetPark => driver.CanSetPark;

        public bool CanSetPierSide => driver.CanSetPierSide;

        public bool CanSetRightAscensionRate => driver.CanSetRightAscensionRate;

        public bool CanSetTracking => driver.CanSetTracking;

        public bool CanSlew => driver.CanSlew;

        public bool CanSlewAltAz => driver.CanSlewAltAz;

        public bool CanSlewAltAzAsync => driver.CanSlewAltAzAsync;

        public bool CanSlewAsync => driver.CanSlewAsync;

        public bool CanSync => driver.CanSync;

        public bool CanSyncAltAz => driver.CanSyncAltAz;

        public bool CanUnpark => driver.CanUnpark;

        public double Declination => driver.Declination;

        public double DeclinationRate { get => driver.DeclinationRate; set => driver.DeclinationRate = value; }
        public bool DoesRefraction { get => driver.DoesRefraction; set => driver.DoesRefraction = value; }

        public EquatorialCoordinateType EquatorialSystem => (EquatorialCoordinateType)driver.EquatorialSystem;

        public double FocalLength => driver.FocalLength;

        public double GuideRateDeclination { get => driver.GuideRateDeclination; set => driver.GuideRateDeclination = value; }
        public double GuideRateRightAscension { get => driver.GuideRateRightAscension; set => driver.GuideRateRightAscension = value; }

        public bool IsPulseGuiding => driver.IsPulseGuiding;

        public double RightAscension => driver.RightAscension;

        public double RightAscensionRate { get => driver.RightAscensionRate; set => driver.RightAscensionRate = value; }
        public PointingState SideOfPier { get => (PointingState)driver.SideOfPier; set => driver.SideOfPier = value; }

        public double SiderealTime => driver.SiderealTime;

        public double SiteElevation { get => driver.SiteElevation; set => driver.SiteElevation = value; }
        public double SiteLatitude { get => driver.SiteLatitude; set => driver.SiteLatitude = value; }
        public double SiteLongitude { get => driver.SiteLongitude; set => driver.SiteLongitude = value; }

        public bool Slewing => driver.Slewing;

        public short SlewSettleTime { get => driver.SlewSettleTime; set => driver.SlewSettleTime = value; }
        public double TargetDeclination { get => driver.TargetDeclination; set => driver.TargetDeclination = value; }
        public double TargetRightAscension { get => driver.TargetRightAscension; set => driver.TargetRightAscension = value; }
        public bool Tracking { get => driver.Tracking; set => driver.Tracking = value; }
        public DriveRate TrackingRate { get => (DriveRate)driver.TrackingRate; set => driver.TrackingRate = value; }

        public ITrackingRates TrackingRates
        {
            get
            {
                return new TrackingRatesFacade(driver);
            }
        }

        public DateTime UTCDate { get => driver.UTCDate; set => driver.UTCDate = value; }
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

        public void AbortSlew()
        {
            driver.AbortSlew();
        }

        public string Action(string ActionName, string ActionParameters)
        {
            return driver.Action(ActionName, ActionParameters);
        }

        public IAxisRates AxisRates(TelescopeAxis Axis)
        {
            return new AxisRatesFacade(Axis, driver);
        }

        public bool CanMoveAxis(TelescopeAxis Axis)
        {
            return driver.CanMoveAxis(Axis);
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

        public PointingState DestinationSideOfPier(double RightAscension, double Declination)
        {
            return (PointingState)driver.DestinationSideOfPier(RightAscension, Declination);
        }

        public void Dispose()
        {
            switch (settings.CurrentDeviceTechnology)
            {
                case ConformConstants.TECHNOLOGY_ALPACA:

                    break;
                case ConformConstants.TECHNOLOGY_COM:
                    DisposeAndReleaseObject("Dispose", driver);
                    break;
            }
        }

        public void FindHome()
        {
            driver.FindHome();
        }

        public void MoveAxis(TelescopeAxis Axis, double Rate)
        {
            driver.MoveAxis(Axis, Rate);
        }

        public void Park()
        {
            driver.Park();
        }

        public void PulseGuide(GuideDirection Direction, int Duration)
        {
            driver.PulseGuide(Direction, Duration);
        }

        public void SetPark()
        {
            driver.SetPark();
        }

        public void SlewToAltAz(double Azimuth, double Altitude)
        {
            driver.SlewToAltAz(Azimuth, Altitude);
        }

        public void SlewToAltAzAsync(double Azimuth, double Altitude)
        {
            driver.SlewToAltAzAsync(Azimuth, Altitude);
        }

        public void SlewToCoordinates(double RightAscension, double Declination)
        {
            driver.SlewToCoordinates(RightAscension, Declination);
        }

        public void SlewToCoordinatesAsync(double RightAscension, double Declination)
        {
            driver.SlewToCoordinatesAsync(RightAscension, Declination);
        }

        public void SlewToTarget()
        {
            driver.SlewToTarget();
        }

        public void SlewToTargetAsync()
        {
            driver.SlewToTargetAsync();
        }

        public void SyncToAltAz(double Azimuth, double Altitude)
        {
            driver.SyncToAltAz(Azimuth, Altitude);
        }

        public void SyncToCoordinates(double RightAscension, double Declination)
        {
            driver.SyncToCoordinates(RightAscension, Declination);
        }

        public void SyncToTarget()
        {
            driver.SyncToTarget();
        }

        public void UnPark()
        {
            driver.UnPark();
        }

        public void DisposeAndReleaseObject(string driverName, dynamic ObjectToRelease)
        {
            Type ObjectType;
            int RemainingObjectCount, LoopCount;
            //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  About to release {driverName} driver instance");
            if (settings.DisplayMethodCalls)
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Comment, $"About to release {driverName} driver instance");
                try
                {
                    ObjectType = ObjectToRelease.GetType();
                    //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  Unmarshalling {ObjectType.Name} -  {ObjectType.FullName}");
                }
                catch (Exception ex1)
                {
                    //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  GetType Exception: " + ex1.Message);
                }

            try
            {
                if (settings.DisplayMethodCalls)
                    //LogMsg("DisposeAndReleaseObject", MessageLevel.Comment, "About to set Connected property");
                    ObjectToRelease.Connected = false;
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  Connected successfully set to False");
            }
            catch (Exception ex1)
            {
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Exception setting Connected = False: " + ex1.Message);
            }

            try
            {
                ObjectToRelease.Dispose();
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  Successfully called Dispose()");
            }
            catch (Exception ex1)
            {
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Dispose Exception: " + ex1.Message);
            }

            try
            {
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Releasing COM object");
                LoopCount = 0;
                do
                {
                    LoopCount += 1;
                    RemainingObjectCount = Marshal.ReleaseComObject(ObjectToRelease);
                    //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Remaining object count: " + RemainingObjectCount + ", LoopCount: " + LoopCount);
                }
                while (!(RemainingObjectCount <= 0 | LoopCount == 20));
            }
            catch (Exception ex2)
            {
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  ReleaseComObject Exception: " + ex2.Message);
            }

            try
            {
                ObjectToRelease = null;
                GC.Collect();
            }
            catch (Exception ex3)
            {
                //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Set to nothing Exception: " + ex3.Message);
            }

            //LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  End of ReleaseCOMObject");
        }


    }
}

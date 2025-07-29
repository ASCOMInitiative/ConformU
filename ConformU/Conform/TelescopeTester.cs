using ASCOM;
using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.Tools;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;

namespace ConformU
{
    internal class TelescopeTester : DeviceTesterBaseClass, IDisposable
    {

        #region Variables and Constants

        // TelescopeTest constants
        internal const string ABORT_SLEW = "AbortSlew";
        internal const string AXIS_RATE = "AxisRate";
        internal const string CAN_MOVE_AXIS = "CanMoveAxis";
        internal const string COMMANDXXX = "CommandXXX";
        internal const string DESTINATION_SIDE_OF_PIER = "DestinationSideOfPier";
        internal const string FIND_HOME = "FindHome";
        internal const string MOVE_AXIS = "MoveAxis";
        internal const string PARK_UNPARK = "Park/Unpark";
        internal const string PULSE_GUIDE = "PulseGuide";
        internal const string SLEW_TO_ALTAZ = "SlewToAltAz";
        internal const string SLEW_TO_ALTAZ_ASYNC = "SlewToAltAzAsync";
        internal const string SLEW_TO_TARGET = "SlewToTarget";
        internal const string SLEW_TO_TARGET_ASYNC = "SlewToTargetAsync";
        internal const string SYNC_TO_ALTAZ = "SyncToAltAz";
        internal const string SLEW_TO_COORDINATES = "SlewToCoordinates";
        internal const string SLEW_TO_COORDINATES_ASYNC = "SlewToCoordinatesAsync";
        internal const string SYNC_TO_COORDINATES = "SyncToCoordinates";
        internal const string SYNC_TO_TARGET = "SyncToTarget";

        private const int TRACKING_COMMAND_DELAY = 1000; // Time to wait between changing Tracking state
        private const int MOVE_AXIS_TIME = 2000; // Number of milliseconds for each move axis command
        private const int NUM_AXIS_RATES = 1000;
        private const int AXIS_RATE_MINIMUM = 0; // Mnemonics for the axis rate array second dimension
        private const int AXIS_RATE_MAXIMUM = 1;
        private const int WAIT_FOR_SLEW_MINIMUM_DURATION = 5; // Minimum number of seconds to wait before declaring an asynchronous slew is finished (allows for mounts that don't set IsSlewing straight away!)
        private const int SIDEOFPIER_MERIDIAN_TRACKING_PERIOD = 7 * 60 * 1000; // 7 minutes in milliseconds
        private const int DISPLAY_DECIMAL_DIGITS = 2;
        private const int PULSEGUIDE_MOVEMENT_TIME = 2; // Initialise a pulse guide movement for this number of seconds
        private const int PULSEGUIDE_TIMEOUT_TIME = 6; // Wait up to this number of seconds before timing out a pulse guide command
        private const double BAD_RA_LOW = -1.0d; // Good range is 0.0 to 23.99999
        private const double BAD_RA_HIGH = 25.0d;
        private const double BAD_DEC_LOW = -100.0d; // Good range is -90.0 to +90.0
        private const double BAD_DEC_HIGH = 100.0d;
        private const double BAD_AZIMUTH_LOW = -10.0d; // Good range is 0.0 to 360.0
        private const double BAD_AZIMUTH_HIGH = 370.0d;
        private const double BAD_ALTITUDE_LOW = -100.0d; // Good range is -90.0 to +90.0 (-90.0 to allow the scope tube to be parked pointing vertically downwards)
        private const double BAD_ALTITUDE_HIGH = 100.0d;
        private const double SIDE_OF_PIER_INVALID_LATITUDE = 10.0d; // +- this value is the range of latitudes where side of pier tests will not be conducted
        private const double SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR = 0.5d; // Must be in the range 0..+1.0. Target declinations will be calculated as this fraction of the altitude of the highest pole from the horizon
        private const double SLEW_SYNC_OK_TOLERANCE = 10.0d; // (Arc seconds) Upper limit of Slew or Sync error that will generate an OK output, otherwise generates an Info message detailing distance
        private const double SYNC_SIMULATED_ERROR = 60.0d; // (Arc minutes) Distance that the telescope will be told that it is in error  when the sync command is issued. The error is applied to both RA and DEC axes
        private const double SIDEREAL_SECONDS_TO_SI_SECONDS = 0.99726956631945; // Based on earth sidereal rotation period of 23 hours 56 minutes 4.09053 seconds
        private const double SIDEREAL_RATE = 15.0 / SIDEREAL_SECONDS_TO_SI_SECONDS; //Arc-seconds per SI second
        private const double SITE_ELEVATION_TEST_VALUE = 2385.0; //Arbitrary site elevation write test value
        private const double ARC_SECONDS_TO_RA_SECONDS = 1.0 / 15.0; // Convert angular movement to time based movement

        private bool canFindHome, canPark, canPulseGuide, canSetDeclinationRate, canSetGuideRates, canSetPark, canSetPierside, canSetRightAscensionRate;
        private bool canSetTracking, canSlew, canSlewAltAz, canSlewAltAzAsync, canSlewAsync, canSync, canSyncAltAz, canUnpark, canReadTracking;
        private AlignmentMode alignmentMode;
        private double altitude;
        private double apertureArea;
        private double apertureDiameter;
        private bool atHome;
        private bool atPark;
        private double azimuth;
        private double declination;
        private double declinationRate;
        private bool doesRefraction;
        private EquatorialCoordinateType equatorialSystem;
        private double focalLength;
        private double guideRateDeclination;
        private double guideRateRightAscension;
        private bool isPulseGuiding;
        private double rightAscension;
        private double rightAscensionRate;
        private PointingState sideOfPier;
        private double siderealTimeScope;
        private double siteElevation;
        private double siteLatitude;
        private double siteLongitude;
        private bool slewing;
        private short slewSettleTime;
        private double targetDeclination;
        private double targetRightAscension;
        private bool tracking;
        private DateTime utcDate;
        private bool canMoveAxisPrimary, canMoveAxisSecondary, canMoveAxisTertiary;
        private PointingState destinationSideOfPier, destinationSideOfPierEast, destinationSideOfPierWest;
        private double siderealTimeAscom;
        private DateTime startTime, endTime;
        private bool? canReadSideOfPier = null; // Start out in the "not read" state
        private double targetAltitude, targetAzimuth;
        private bool canReadAltitide, canReadAzimuth, canReadSiderealTime;

        private readonly Dictionary<string, bool> telescopeTests;

        // Axis rate checks
        private readonly double[,] axisRatesPrimaryArray = new double[1001, 2];
        private readonly double[,] axisRatesArray = new double[1001, 2];

        // Helper variables
        private ITelescopeV4 telescopeDevice;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        private readonly Transform transform;

        #endregion

        #region Enums
        private enum CanType
        {
            CanFindHome = 1,
            CanPark = 2,
            CanPulseGuide = 3,
            CanSetDeclinationRate = 4,
            CanSetGuideRates = 5,
            CanSetPark = 6,
            CanSetPierSide = 7,
            CanSetRightAscensionRate = 8,
            CanSetTracking = 9,
            CanSlew = 10,
            CanSlewAsync = 11,
            CanSlewAltAz = 12,
            CanSlewAltAzAsync = 13,
            CanSync = 14,
            CanSyncAltAz = 15,
            CanUnPark = 16,
            CanAbortExposure = 17,
            CanAsymmetricBin = 18,
            CanGetCoolerPower = 19,
            CanSetCcdTemperature = 20,
            CanStopExposure = 21
        }

        private enum OptionalMethodType
        {
            AbortSlew = 1,
            DestinationSideOfPier = 2,
            FindHome = 3,
            MoveAxisPrimary = 4,
            MoveAxisSecondary = 5,
            MoveAxisTertiary = 6,
            PulseGuide = 7,
            SideOfPierWrite = 8
        }

        private enum RequiredMethodType
        {
            TstAxisrates = 1,
            TstCanMoveAxisPrimary = 2,
            TstCanMoveAxisSecondary = 3,
            TstCanMoveAxisTertiary = 4
        }

        private enum ParkedExceptionType
        {
            TstPExcepAbortSlew = 1,
            TstPExcepFindHome = 2,
            TstPExcepMoveAxisPrimary = 3,
            TstPExcepMoveAxisSecondary = 4,
            TstPExcepMoveAxisTertiary = 5,
            TstPExcepSlewToCoordinates = 6,
            TstPExcepSlewToCoordinatesAsync = 7,
            TstPExcepSlewToTarget = 8,
            TstPExcepSlewToTargetAsync = 9,
            TstPExcepSyncToCoordinates = 10,
            TstPExcepSyncToTarget = 11,
            TstPExcepPulseGuide = 12,
            TstExcepTracking = 13
        }
        // Private Enum SyncType
        // End Enum
        private enum SlewSyncType
        {
            SlewToCoordinates = 1,
            SlewToCoordinatesAsync = 2,
            SlewToTarget = 3,
            SlewToTargetAsync = 4,
            SlewToAltAz = 5,
            SlewToAltAzAsync = 6,
            SyncToCoordinates = 7,
            SyncToTarget = 8,
            SyncToAltAz = 9
        }

        private enum PerformanceType
        {
            TstPerfAltitude = 0,
            TstPerfAtHome = 1,
            TstPerfAtPark = 2,
            TstPerfAzimuth = 3,
            TstPerfDeclination = 4,
            TstPerfIsPulseGuiding = 5,
            TstPerfRightAscension = 6,
            TstPerfSideOfPier = 7,
            TstPerfSiderealTime = 8,
            TstPerfSlewing = 9,
            TstPerfUtcDate = 10
        }

        public enum FlipTestType
        {
            DestinationSideOfPier,
            SideOfPier
        }

        private enum InterfaceType
        {
            Telescope,
            TelescopeV2,
            TelescopeV3
        }

        enum RateOffset
        {
            RightAscensionRate,
            DeclinationRate
        }

        #endregion

        #region New and Dispose
        public TelescopeTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, true, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            telescopeTests = settings.TelescopeTests;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
            transform = new Transform();
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        protected override void Dispose(bool disposing)
        {
            LogDebug("Dispose", $"Disposing of device: {disposing} {disposedValue}");
            if (!disposedValue)
            {
                if (disposing)
                {
                    telescopeDevice?.Dispose();
                    telescopeDevice = null;
                    transform?.Dispose();
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        #region Conform Process

        public override void InitialiseTest()
        {
            unchecked
            {
                // Set the error type numbers according to the standards adopted by individual authors.
                // Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
                // messages to driver authors!
                switch (settings.ComDevice.ProgId ?? "")
                {
                    case "Hub.Telescope":
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040401;
                            ExInvalidValue2 = (int)0x80040402;
                            ExInvalidValue3 = (int)0x80040405;
                            ExInvalidValue4 = (int)0x80040402;
                            ExInvalidValue5 = (int)0x80040402;
                            ExInvalidValue6 = (int)0x80040402;
                            ExNotSet1 = (int)0x80040403;
                            break;
                        }

                    case "ASCOM.MI250SA.Telescope":
                    case "Celestron.Telescope":
                    case "ASCOM.MI250.Telescope":
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040401;
                            ExInvalidValue2 = (int)0x80040402;
                            ExInvalidValue3 = (int)0x80040402;
                            ExInvalidValue4 = (int)0x80040402;
                            ExInvalidValue5 = (int)0x80040402;
                            ExInvalidValue6 = (int)0x80040402;
                            ExNotSet1 = (int)0x80040403;
                            break;
                        }

                    case "TemmaLite.Telescope":
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040410;
                            ExInvalidValue2 = (int)0x80040418;
                            ExInvalidValue3 = (int)0x80040418;
                            ExInvalidValue4 = (int)0x80040418;
                            ExInvalidValue5 = (int)0x80040418;
                            ExInvalidValue6 = (int)0x80040418;
                            ExNotSet1 = (int)0x80040417;
                            break;
                        }

                    case "Gemini.Telescope":
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040410;
                            ExInvalidValue2 = (int)0x80040418;
                            ExInvalidValue3 = (int)0x80040419;
                            ExInvalidValue4 = (int)0x80040420;
                            ExInvalidValue5 = (int)0x80040420;
                            ExInvalidValue6 = (int)0x80040420;
                            ExNotSet1 = (int)0x80040417;
                            break;
                        }

                    case "POTH.Telescope":
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040405;
                            ExInvalidValue2 = (int)0x80040406;
                            ExInvalidValue3 = (int)0x80040406;
                            ExInvalidValue4 = (int)0x80040406;
                            ExInvalidValue5 = (int)0x80040406;
                            ExInvalidValue6 = (int)0x80040406;
                            ExNotSet1 = (int)0x80040403;
                            break;
                        }

                    case "ServoCAT.Telescope":
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = ErrorCodes.InvalidValue;
                            ExInvalidValue2 = (int)0x80040405;
                            ExInvalidValue3 = (int)0x80040405;
                            ExInvalidValue4 = (int)0x80040405;
                            ExInvalidValue5 = (int)0x80040405;
                            ExInvalidValue6 = (int)0x80040405;
                            ExNotSet1 = (int)0x80040403;
                            ExNotSet2 = (int)0x80040404; // I'm using the simulator values as the defaults since it is the reference platform
                            break;
                        }

                    default:
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = ErrorCodes.InvalidValue;
                            ExInvalidValue2 = (int)0x80040405;
                            ExInvalidValue3 = (int)0x80040405;
                            ExInvalidValue4 = (int)0x80040405;
                            ExInvalidValue5 = (int)0x80040405;
                            ExInvalidValue6 = (int)0x80040405;
                            ExNotSet1 = (int)0x80040403;
                            break;
                        }
                }
            }

            base.InitialiseTest();
        }

        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        LogInfo("CreateDevice",
                            $"Creating Alpaca device: Access service: {settings.AlpacaConfiguration.AccessServiceType}, IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber} ({(settings.AlpacaConfiguration.TrustUserGeneratedSslCertificates ? "A" : "Not a")}ccepting user generated SSL certificates)");
                        telescopeDevice = new AlpacaTelescope(
                                                    settings.AlpacaConfiguration.AccessServiceType,
                                                    settings.AlpacaDevice.IpAddress,
                                                    settings.AlpacaDevice.IpPort,
                                                    settings.AlpacaDevice.AlpacaDeviceNumber,
                                                    settings.AlpacaConfiguration.EstablishConnectionTimeout,
                                                    settings.AlpacaConfiguration.StandardResponseTimeout,
                                                    settings.AlpacaConfiguration.LongResponseTimeout,
                                                    Globals.CLIENT_NUMBER_DEFAULT,
                                                    settings.AlpacaConfiguration.AccessUserName,
                                                    settings.AlpacaConfiguration.AccessPassword,
                                                    settings.AlpacaConfiguration.StrictCasing,
                                                    settings.TraceAlpacaCalls ? logger : null,
                                                    Globals.USER_AGENT_PRODUCT_NAME,
                                                    Assembly.GetExecutingAssembly().GetName().Version.ToString(4),
                                                    settings.AlpacaConfiguration.TrustUserGeneratedSslCertificates);

                        LogInfo("CreateDevice", $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComAccessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                telescopeDevice = new TelescopeFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                telescopeDevice = new Telescope(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(telescopeDevice, DeviceTypes.Telescope); // Assign the driver to the base class
                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);

                // Validate the interface version
                ValidateInterfaceVersion();
            }
            catch (COMException exCom) when (exCom.ErrorCode == REGDB_E_CLASSNOTREG)
            {
                LogDebug("CreateDevice", $"Exception thrown: {exCom.Message}\r\n{exCom}");

                throw new Exception($"The driver is not registered as a {(Environment.Is64BitProcess ? "64bit" : "32bit")} driver");
            }
            catch (Exception ex)
            {
                LogDebug("CreateDevice", $"Exception thrown: {ex.Message}\r\n{ex}");
                throw; // Re throw exception 
            }
        }

        public override void PreRunCheck()
        {
            // Get into a consistent state
            if (GetInterfaceVersion() > 1)
            {
                try
                {
                    LogCallToDriver("Mount Safety", "About to get AtPark property");
                    if (telescopeDevice.AtPark)
                    {
                        if (canUnpark)
                        {
                            try
                            {
                                LogCallToDriver("Mount Safety", "About to call Unpark method");
                                telescopeDevice.Unpark();
                                LogCallToDriver("Mount Safety", "About to get AtPark property repeatedly");
                                WaitWhile("Waiting for scope to unpark", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                LogInfo("Mount Safety", "Scope was parked, so it has been unparked for testing");
                            }
                            catch (Exception ex)
                            {
                                HandleException("Mount Safety - Unpark", MemberType.Method, Required.MustBeImplemented, ex, "CanUnpark is true");
                            }
                        }
                        else
                        {
                            LogIssue("Mount Safety", "Scope reports that it is parked but CanUnPark is false - please manually unpark the scope");
                        }
                    }
                    else
                    {
                        LogInfo("Mount Safety", "Scope is not parked, continuing testing");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Mount Safety - AtPark", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("Mount Safety",
                    $"Skipping AtPark test as this method is not supported in interface V{GetInterfaceVersion()}");
                try
                {
                    if (canUnpark)
                    {
                        LogCallToDriver("Mount Safety", "About to call Unpark method");
                        telescopeDevice.Unpark();
                        LogCallToDriver("Mount Safety", "About to get AtPark property repeatedly");
                        WaitWhile("Waiting for scope to unpark", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                        LogOk("Mount Safety", "Scope has been unparked for testing");
                    }
                    else
                    {
                        LogOk("Mount Safety", "Scope reports that it cannot unpark, unparking skipped");
                    }
                }
                catch (Exception ex)
                {
                    LogIssue("Mount Safety", $"Driver threw an exception while unparking: {ex.Message}");
                }
            }

            if (!cancellationToken.IsCancellationRequested & canSetTracking)
            {
                LogCallToDriver("Mount Safety", "About to set Tracking property true");
                try
                {
                    telescopeDevice.Tracking = true;
                    LogInfo("Mount Safety", "Scope tracking has been enabled");
                }
                catch (Exception ex)
                {
                    LogInfo("Mount Safety", $"Exception while trying to set Tracking to true: {ex.Message}");
                }
            }

            if (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    LogInfo("TimeCheck", $"PC Time Zone:  {TimeZoneInfo.Local.DisplayName} offset: {TimeZoneInfo.Local.BaseUtcOffset.Hours} hours.");
                    LogInfo("TimeCheck", $"PC UTCDate:    {DateTime.UtcNow:dd-MMM-yyyy HH:mm:ss.fff}");
                }
                catch (Exception ex)
                {
                    LogIssue("TimeCheck", $"Exception reading PC Time: {ex}");
                }

                // v1.0.12.0 Added catch logic for any UTCDate issues
                try
                {
                    LogCallToDriver("TimeCheck", "About to get UTCDate property");
                    DateTime mountTime = telescopeDevice.UTCDate;
                    LogDebug("TimeCheck", $"Mount UTCDate Unformatted: {telescopeDevice.UTCDate}");
                    LogInfo("TimeCheck", $"Mount UTCDate: {telescopeDevice.UTCDate:dd-MMM-yyyy HH:mm:ss.fff}");
                }
                catch (Exception ex)
                {
                    HandleException("UTCDate", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
        }

        public override void ReadCanProperties()
        {
            TelescopeCanTest(CanType.CanFindHome, "CanFindHome");
            TelescopeCanTest(CanType.CanPark, "CanPark");
            TelescopeCanTest(CanType.CanPulseGuide, "CanPulseGuide");
            TelescopeCanTest(CanType.CanSetDeclinationRate, "CanSetDeclinationRate");
            TelescopeCanTest(CanType.CanSetGuideRates, "CanSetGuideRates");
            TelescopeCanTest(CanType.CanSetPark, "CanSetPark");
            TelescopeCanTest(CanType.CanSetPierSide, "CanSetPierSide");
            TelescopeCanTest(CanType.CanSetRightAscensionRate, "CanSetRightAscensionRate");
            TelescopeCanTest(CanType.CanSetTracking, "CanSetTracking");
            TelescopeCanTest(CanType.CanSlew, "CanSlew");
            TelescopeCanTest(CanType.CanSlewAltAz, "CanSlewltAz");
            TelescopeCanTest(CanType.CanSlewAltAzAsync, "CanSlewAltAzAsync");
            TelescopeCanTest(CanType.CanSlewAsync, "CanSlewAsync");
            TelescopeCanTest(CanType.CanSync, "CanSync");
            TelescopeCanTest(CanType.CanSyncAltAz, "CanSyncAltAz");
            TelescopeCanTest(CanType.CanUnPark, "CanUnPark");

            // Make sure that Park and Unpark are either both implemented or both not implemented
            if (canUnpark & !canPark)
                LogIssue("CanUnPark", "CanUnPark is true but CanPark is false - this does not comply with ASCOM specification");

            LogDebug("Can Slewing Checks", $"Interface version: {GetInterfaceVersion()}, Device technology: {settings.DeviceTechnology}, " +
                $"CanSlew: {canSlew}, CanSlewAsync: {canSlewAsync}, CamSlewAltAz: {canSlewAltAz}, CanSlewAltAzAsync: {canSlewAltAzAsync}");

            // For Platform 7 ITelescopeV4 and later interfaces, make sure that sync and async slew methods are tied together for COM devices
            if (IsPlatform7OrLater)
            {
                // Only apply this requirement to COM drivers, Alpaca devices are expected only to implement async slewing, synch is optional
                if (settings.DeviceTechnology == DeviceTechnology.COM) // We are testing a COM driver
                {
                    // Test whether the two equatorial slewing values are both either true or false
                    if ((canSlew & !canSlewAsync) || (!canSlew & canSlewAsync)) // The values are not tied together
                    {
                        LogNewLine();
                        LogIssue("Equatorial Slewing", $"The CanSlew and CanSlewAsync properties are not tied together: CanSlew: {canSlew}, CanSlewAsync: {canSlewAsync}");
                        LogInfo("Equatorial Slewing", $"In ITelescopeV4 and later COM drivers, both synchronous and asynchronous equatorial slewing methods must be either implemented or not implemented.");
                        LogInfo("Equatorial Slewing", $"It is not permissible for one to be implemented while the other is not.");
                    }

                    // Test whether the two Alt/Az slewing values are both either true or false
                    if ((canSlewAltAz & !canSlewAltAzAsync) || (!canSlewAltAz & canSlewAltAzAsync)) // The values are not tied together
                    {
                        LogNewLine();
                        LogIssue("Alt/Az Slewing", $"The CanSlewAltAz and CanSlewAltAzAsync properties are not tied together: CanSlewAltAz: {canSlewAltAz}, CanSlewAltAzAsync: {canSlewAltAzAsync}");
                        LogInfo("Alt/Az Slewing", $"In ITelescopeV4 and later COM drivers, both synchronous and asynchronous alt/az slewing methods must be either implemented or not implemented.");
                        LogInfo("Alt/Az Slewing", $"It is not permissible for one to be implemented while the other is not.");
                    }
                }
            }
        }

        public override void CheckProperties()
        {
            bool lOriginalTrackingState;
            DriveRate driveRate;
            double timeDifference;
            ITrackingRates trackingRates = null;
            dynamic trackingRate;

            // Test TargetDeclination and TargetRightAscension first because these tests will fail if the telescope has been slewed previously.
            // Slews can happen in the extended guide rate tests for example.
            // The test is made here but reported later so that the properties are tested in mostly alphabetical order.

            // TargetDeclination Read - Optional
            Exception targetDeclinationReadException = null;
            try // First read should fail!
            {
                LogCallToDriver("TargetDeclination Read", "About to get TargetDeclination property");
                targetDeclination = TimeFunc<double>("TargetDeclination Read", () => telescopeDevice.TargetDeclination, TargetTime.Fast);
            }
            catch (Exception ex)
            {
                targetDeclinationReadException = ex;
            }

            // TargetRightAscension Read - Optional
            Exception targetRightAscensionReadException = null;
            try // First read should fail!
            {
                LogCallToDriver("TargetRightAscension Read", "About to get TargetRightAscension property");
                targetRightAscension = TimeFunc("TargetRightAscension Read", () => telescopeDevice.TargetRightAscension, TargetTime.Fast);
            }
            catch (Exception ex)
            {
                targetRightAscensionReadException = ex;
            }

            // AlignmentMode - Optional
            try
            {
                LogCallToDriver("AlignmentMode", "About to get AlignmentMode property");
                alignmentMode = TimeFunc("AlignmentMode", () => telescopeDevice.AlignmentMode, TargetTime.Fast);
                LogOk("AlignmentMode", alignmentMode.ToString());
            }
            catch (Exception ex)
            {
                HandleException("AlignmentMode", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Altitude - Optional
            try
            {
                canReadAltitide = false;
                LogCallToDriver("Altitude", "About to get Altitude property");
                altitude = TimeFunc("Altitude", () => telescopeDevice.Altitude, TargetTime.Fast);
                canReadAltitide = true; // Read successfully
                switch (altitude)
                {
                    case var @case when @case < 0.0d:
                        {
                            LogIssue("Altitude", $"Altitude is <0.0 degrees: {altitude.ToDMS().Trim()}");
                            break;
                        }

                    case var case1 when case1 > 90.0000001d:
                        {
                            LogIssue("Altitude", $"Altitude is >90.0 degrees: {altitude.ToDMS().Trim()}");
                            break;
                        }

                    default:
                        {
                            LogOk("Altitude", altitude.ToDMS().Trim());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Altitude", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // ApertureArea - Optional
            try
            {
                LogCallToDriver("ApertureArea", "About to get ApertureArea property");
                apertureArea = TimeFunc("ApertureArea", () => telescopeDevice.ApertureArea, TargetTime.Fast);
                switch (apertureArea)
                {
                    case var case2 when case2 < 0d:
                        {
                            LogIssue("ApertureArea", $"ApertureArea is < 0.0 : {apertureArea}");
                            break;
                        }

                    case 0.0d:
                        {
                            LogInfo("ApertureArea", "ApertureArea is 0.0");
                            break;
                        }

                    default:
                        {
                            LogOk("ApertureArea", apertureArea.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("ApertureArea", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // ApertureDiameter - Optional
            try
            {
                LogCallToDriver("ApertureDiameter", "About to get ApertureDiameter property");
                apertureDiameter = TimeFunc("ApertureDiameter", () => telescopeDevice.ApertureDiameter, TargetTime.Fast);
                switch (apertureDiameter)
                {
                    case var case3 when case3 < 0.0d:
                        {
                            LogIssue("ApertureDiameter", $"ApertureDiameter is < 0.0 : {apertureDiameter}");
                            break;
                        }

                    case 0.0d:
                        {
                            LogInfo("ApertureDiameter", "ApertureDiameter is 0.0");
                            break;
                        }

                    default:
                        {
                            LogOk("ApertureDiameter", apertureDiameter.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("ApertureDiameter", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // AtHome - Required
            if (GetInterfaceVersion() > 1)
            {
                try
                {
                    LogCallToDriver("AtHome", "About to get AtHome property");
                    atHome = TimeFunc("AtHome", () => telescopeDevice.AtHome, TargetTime.Fast);
                    LogOk("AtHome", atHome.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("AtHome", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("AtHome",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // AtPark - Required
            if (GetInterfaceVersion() > 1)
            {
                try
                {
                    LogCallToDriver("AtPark", "About to get AtPark property");
                    atPark = TimeFunc("AtPark", () => telescopeDevice.AtPark, TargetTime.Fast);
                    LogOk("AtPark", atPark.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("AtPark", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("AtPark",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Azimuth - Optional
            try
            {
                canReadAzimuth = false;
                LogCallToDriver("Azimuth", "About to get Azimuth property");
                azimuth = TimeFunc("Azimuth", () => telescopeDevice.Azimuth, TargetTime.Fast);
                canReadAzimuth = true; // Read successfully
                switch (azimuth)
                {
                    case var case4 when case4 < 0.0d:
                        {
                            LogIssue("Azimuth", $"Azimuth is <0.0 degrees: {FormatAzimuth(azimuth).Trim()}");
                            break;
                        }

                    case var case5 when case5 > 360.0000000001d:
                        {
                            LogIssue("Azimuth", $"Azimuth is >360.0 degrees: {FormatAzimuth(azimuth).Trim()}");
                            break;
                        }

                    default:
                        {
                            LogOk("Azimuth", FormatAzimuth(azimuth).Trim());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Azimuth", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Declination - Required
            try
            {
                LogCallToDriver("Declination", "About to get Declination property");
                declination = TimeFunc("Declination", () => telescopeDevice.Declination, TargetTime.Fast);
                switch (declination)
                {
                    case var case6 when case6 < -90.0d:
                    case var case7 when case7 > 90.0d:
                        {
                            LogIssue("Declination", $"Declination is <-90 or >90 degrees: {declination.ToDMS().Trim()}");
                            break;
                        }

                    default:
                        {
                            LogOk("Declination", declination.ToDMS().Trim());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Declination", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // DeclinationRate Read - Mandatory - must return a number even when CanSetDeclinationRate is False
            try
            {
                LogCallToDriver("DeclinationRate Read", "About to get DeclinationRate property");
                declinationRate = TimeFunc("DeclinationRate Read", () => telescopeDevice.DeclinationRate, TargetTime.Fast);
                // Read has been successful
                if (canSetDeclinationRate) // Any value is acceptable
                {
                    switch (declinationRate)
                    {
                        case var case8 when case8 >= 0.0d:
                            LogOk("DeclinationRate Read", declinationRate.ToString("0.00"));
                            break;

                        default:
                            LogIssue("DeclinationRate Read", $"Negative DeclinatioRate: {declinationRate:0.00}");
                            break;
                    }
                }
                else // Only zero is acceptable
                {
                    switch (declinationRate)
                    {
                        case 0.0d:
                            LogOk("DeclinationRate Read", declinationRate.ToString("0.00"));
                            break;

                        default:
                            LogIssue("DeclinationRate Read",
                                $"DeclinationRate is non zero when CanSetDeclinationRate is False {declinationRate:0.00}");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!canSetDeclinationRate)
                    LogIssue("DeclinationRate Read", "DeclinationRate must return 0 even when CanSetDeclinationRate is false.");
                HandleException("DeclinationRate Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // DeclinationRate Write - Optional
            if (GetInterfaceVersion() > 1) // ITelescopeV2 and later
            {
                if (canSetDeclinationRate) // Any value is acceptable
                {
                    SetTest("DeclinationRate Write");

                    // Set up the rate offset Write test by ensuring that the telescope is in sidereal mode
                    try
                    {
                        telescopeDevice.TrackingRate = DriveRate.Sidereal;
                    }
                    catch { } // This is best endeavours, it doesn't matter if it fails at this point

                    // Log the test offset rate if extended tests are enabled
                    if (settings.TelescopeExtendedRateOffsetTests)
                    {
                        LogInfo("DeclinationRate Write", $"Configured offset test duration: {settings.TelescopeRateOffsetTestDuration} seconds.");
                        LogInfo("DeclinationRate Write", $"Configured offset rate (low):  {settings.TelescopeRateOffsetTestLowValue,7:+0.000;-0.000;+0.000} arc seconds per second.");
                        LogInfo("DeclinationRate Write", $"Configured offset rate (high): {settings.TelescopeRateOffsetTestHighValue,7:+0.000;-0.000;+0.000} arc seconds per second.");
                    }

                    if (TestRaDecRate("DeclinationRate Write", "Set rate to   0.000", Axis.Dec, 0.0d, true))
                    {
                        TestRaDecRate("DeclinationRate Write", $"Set rate to {settings.TelescopeRateOffsetTestLowValue,7:+0.000;-0.000;+0.000}", Axis.Dec, settings.TelescopeRateOffsetTestLowValue, true);
                        if (cancellationToken.IsCancellationRequested) //
                            return;

                        TestRaDecRate("DeclinationRate Write", $"Set rate to {-settings.TelescopeRateOffsetTestLowValue,7:+0.000;-0.000;+0.000}", Axis.Dec, -settings.TelescopeRateOffsetTestLowValue, true);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        TestRaDecRate("DeclinationRate Write", $"Set rate to {settings.TelescopeRateOffsetTestHighValue,7:+0.000;-0.000;+0.000}", Axis.Dec, settings.TelescopeRateOffsetTestHighValue, true);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        TestRaDecRate("DeclinationRate Write", $"Set rate to {-settings.TelescopeRateOffsetTestHighValue,7:+0.000;-0.000;+0.000}", Axis.Dec, -settings.TelescopeRateOffsetTestHighValue, true);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        if (TestRaDecRate("DeclinationRate Write", "Set rate to   0.000", Axis.Dec, 0.0d, true))
                            if (cancellationToken.IsCancellationRequested)
                                return;

                        // Run the extended rate offset tests if configured to do so.
                        if (settings.TelescopeExtendedRateOffsetTests)
                        {
                            // Only test if the declination rate can be set
                            if (canSetDeclinationRate) // Declination rate can be set
                            {
                                TestDeclinationOffsetRates("DeclinationRate Write", -9.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                TestDeclinationOffsetRates("DeclinationRate Write", 3.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                TestDeclinationOffsetRates("DeclinationRate Write", 9.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                TestDeclinationOffsetRates("DeclinationRate Write", -3.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                            }
                            else // Declination rate can not be set
                            {
                                LogInfo("DeclinationRate Write", "Skipping extended DeclinationRate tests because CanSetDeclinationRate is false.");
                            }
                        }

                        TestRaDecRate("DeclinationRate Write", "Reset rate to 0.0", Axis.Dec, 0.0d, false); // Reset the rate to zero, skipping the slewing test
                    }
                }
                else // Should generate an error
                {
                    try
                    {
                        LogCallToDriver("DeclinationRate Write", "About to set DeclinationRate property to 0.0");
                        telescopeDevice.DeclinationRate = 0.0d; // Set to a harmless value
                        LogIssue("DeclinationRate", "CanSetDeclinationRate is False but setting DeclinationRate did not generate an error");
                    }
                    catch (Exception ex)
                    {
                        HandleException("DeclinationRate Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetDeclinationRate is False");
                    }
                }
            }
            else // ITelescopeV1
            {
                try
                {
                    LogCallToDriver("DeclinationRate Write", "About to set DeclinationRate property to 0.0");
                    telescopeDevice.DeclinationRate = 0.0d; // Set to a harmless value
                    LogOk("DeclinationRate Write", declinationRate.ToString("0.00"));
                }
                catch (Exception ex)
                {
                    HandleException("DeclinationRate Write", MemberType.Property, Required.Optional, ex, "");
                }
            }
            ClearStatus();

            if (cancellationToken.IsCancellationRequested)
                return;

            // DoesRefraction Read - Optional
            if (GetInterfaceVersion() > 1)
            {
                try
                {
                    LogCallToDriver("DoesRefraction Read", "About to DoesRefraction get property");
                    doesRefraction = TimeFunc("DoesRefraction Read", () => telescopeDevice.DoesRefraction, TargetTime.Fast);
                    LogOk("DoesRefraction Read", doesRefraction.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("DoesRefraction Read", MemberType.Property, Required.Optional, ex, "");
                }

                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("DoesRefraction Read",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // DoesRefraction Write - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (doesRefraction) // Telescope reports DoesRefraction = True so try setting to False
                {
                    try
                    {
                        LogCallToDriver("DoesRefraction Write", "About to set DoesRefraction property false");
                        TimeMethod("DoesRefraction Write", () => telescopeDevice.DoesRefraction = false, TargetTime.Standard);
                        LogOk("DoesRefraction Write", "Can set DoesRefraction to False");
                    }
                    catch (Exception ex)
                    {
                        HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "");
                    }
                }
                else // // Telescope reports DoesRefraction = False so try setting to True
                {
                    try
                    {
                        LogCallToDriver("DoesRefraction Write", "About to set DoesRefraction property true");
                        TimeMethod("DoesRefraction Write", () => telescopeDevice.DoesRefraction = true, TargetTime.Standard);
                        LogOk("DoesRefraction Write", "Can set DoesRefraction to True");
                    }
                    catch (Exception ex)
                    {
                        HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "");
                    }
                }

                // Restore the Telescope's original DoesRefraction state.
                try
                {
                    LogCallToDriver("DoesRefraction Write", $"About to set DoesRefraction property {doesRefraction}");
                    telescopeDevice.DoesRefraction = doesRefraction;
                    LogOk("DoesRefraction Write", $"Restored original DoesRefraction state: {doesRefraction}");
                }
                catch (Exception ex)
                {
                    LogInfo("DoesRefraction Write", $"Exception while restoring the telescope's original DoesRefraction state: {ex.Message}");
                    HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "");
                }
            }
            else
            {
                LogInfo("DoesRefraction Write",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // EquatorialSystem - Required
            if (GetInterfaceVersion() > 1)
            {
                try
                {
                    LogCallToDriver("EquatorialSystem", "About to get EquatorialSystem property");
                    equatorialSystem = TimeFunc("EquatorialSystem", () => telescopeDevice.EquatorialSystem, TargetTime.Fast);
                    LogOk("EquatorialSystem", equatorialSystem.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("EquatorialSystem", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("EquatorialSystem",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // FocalLength - Optional
            try
            {
                LogCallToDriver("FocalLength", "About to get FocalLength property");
                focalLength = TimeFunc("FocalLength", () => telescopeDevice.FocalLength, TargetTime.Fast);
                switch (focalLength)
                {
                    case var case9 when case9 < 0.0d:
                        {
                            LogIssue("FocalLength", $"FocalLength is <0.0 : {focalLength}");
                            break;
                        }

                    case 0.0d:
                        {
                            LogInfo("FocalLength", "FocalLength is 0.0");
                            break;
                        }

                    default:
                        {
                            LogOk("FocalLength", focalLength.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("FocalLength", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // GuideRateDeclination - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (canSetGuideRates) // Can set guide rates so read and write are mandatory
                {
                    try
                    {
                        LogCallToDriver("GuideRateDeclination Read", "About to get GuideRateDeclination property");
                        guideRateDeclination = TimeFunc("GuideRateDeclination Read", () => telescopeDevice.GuideRateDeclination, TargetTime.Fast); // Read guiderateDEC

                        if (guideRateDeclination >= 0.0)
                        {
                            LogOk("GuideRateDeclination Read", $"{guideRateDeclination:0.000000} ({guideRateDeclination.ToDMS()})");

                        }
                        else
                        {
                            LogIssue("GuideRateDeclination Read", $"GuideRateDeclination is < 0.0: {guideRateDeclination:0.000000} ({guideRateDeclination.ToDMS()})");
                        }
                    }
                    catch (Exception ex) // Read failed
                    {
                        HandleException("GuideRateDeclination Read", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True");
                    }

                    try // Read OK so now try to write
                    {
                        LogCallToDriver("GuideRateDeclination Read", $"About to set GuideRateDeclination property to {guideRateDeclination}");
                        TimeMethod("GuideRateDeclination Write", () => telescopeDevice.GuideRateDeclination = guideRateDeclination, TargetTime.Standard);
                        LogOk("GuideRateDeclination Write", "Can write Declination Guide Rate OK");
                    }
                    catch (Exception ex) // Write failed
                    {
                        HandleException("GuideRateDeclination Write", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True");
                    }
                }
                else
                {
                    try // Cannot set guide rates so Read is Optional and may generate an error
                    {
                        LogCallToDriver("GuideRateDeclination Read", "About to get GuideRateDeclination property");
                        guideRateDeclination = telescopeDevice.GuideRateDeclination;
                        switch (guideRateDeclination)
                        {
                            case var case11 when case11 < 0.0d:
                                {
                                    LogIssue("GuideRateDeclination Read",
                                        $"GuideRateDeclination is < 0.0 {guideRateDeclination:0.00}");
                                    break;
                                }

                            default:
                                {
                                    LogOk("GuideRateDeclination Read", guideRateDeclination.ToString("0.00"));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex) // Some other error
                    {
                        HandleException("GuideRateDeclination Read", MemberType.Property, Required.Optional, ex, "");
                    }

                    try // Write should definitely raise an error
                    {
                        LogCallToDriver("GuideRateDeclination Write",
                            $"About to set GuideRateDeclination property to {guideRateDeclination}");
                        telescopeDevice.GuideRateDeclination = guideRateDeclination;
                        LogIssue("GuideRateDeclination Write",
                            $"CanSetGuideRates is false but no exception generated; value returned: {guideRateDeclination:0.00}");
                    }
                    catch (Exception ex) // Some other error so OK
                    {
                        HandleException("GuideRateDeclination Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False");
                    }
                }
            }
            else
            {
                LogInfo("GuideRateDeclination",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // GuideRateRightAscension - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (canSetGuideRates)
                {
                    try
                    {
                        LogCallToDriver("GuideRateRightAscension Read", "About to get GuideRateRightAscension property");
                        guideRateRightAscension = TimeFunc("GuideRateRightAscension Read", () => telescopeDevice.GuideRateRightAscension, TargetTime.Fast); // Read guide rate RA

                        if (guideRateRightAscension >= 0.0)
                        {
                            LogOk("GuideRateRightAscension Read", $"{guideRateRightAscension:0.000000} ({guideRateRightAscension.ToDMS()})");
                        }
                        else
                        {
                            LogIssue("GuideRateRightAscension Read", $"GuideRateRightAscension is < 0.0: {guideRateRightAscension:0.000000} ({guideRateRightAscension.ToDMS()})");
                        }
                    }
                    catch (Exception ex) // Read failed
                    {
                        HandleException("GuideRateRightAscension Read", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True");
                    }

                    try // Read OK so now try to write
                    {
                        LogCallToDriver("GuideRateRightAscension Read",
                            $"About to set GuideRateRightAscension property to {guideRateRightAscension}");
                        TimeMethod("GuideRateRightAscension Write", () => telescopeDevice.GuideRateRightAscension = guideRateRightAscension, TargetTime.Standard);
                        LogOk("GuideRateRightAscension Write", "Can set RightAscension Guide OK");
                    }
                    catch (Exception ex) // Write failed
                    {
                        HandleException("GuideRateRightAscension Write", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True");
                    }
                }
                else
                {
                    try // Cannot set guide rates so read may generate an error
                    {
                        LogCallToDriver("GuideRateRightAscension Read", "About to get GuideRateRightAscension property");
                        guideRateRightAscension = telescopeDevice.GuideRateRightAscension; // Read guiderateRA
                        switch (guideRateDeclination)
                        {
                            case var case13 when case13 < 0.0d:
                                {
                                    LogIssue("GuideRateRightAscension Read",
                                        $"GuideRateRightAscension is < 0.0 {guideRateRightAscension:0.00}");
                                    break;
                                }

                            default:
                                {
                                    LogOk("GuideRateRightAscension Read", guideRateRightAscension.ToString("0.00"));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex) // Some other error
                    {
                        HandleException("GuideRateRightAscension Read", MemberType.Property, Required.Optional, ex, "");
                    }

                    try // Write should definitely raise an error
                    {
                        LogCallToDriver("GuideRateRightAscension Write",
                            $"About to set GuideRateRightAscension property to {guideRateRightAscension}");
                        telescopeDevice.GuideRateRightAscension = guideRateRightAscension;
                        LogIssue("GuideRateRightAscension Write",
                            $"CanSetGuideRates is false but no exception generated; value returned: {guideRateRightAscension:0.00}");
                    }
                    catch (Exception ex) // Some other error so OK
                    {
                        HandleException("GuideRateRightAscension Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False");
                    }
                }
            }
            else
            {
                LogInfo("GuideRateRightAscension",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // IsPulseGuiding - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (canPulseGuide) // Can pulse guide so test if we can successfully read IsPulseGuiding
                {
                    try
                    {
                        LogCallToDriver("IsPulseGuiding", "About to get IsPulseGuiding property");
                        isPulseGuiding = TimeFunc("IsPulseGuiding", () => telescopeDevice.IsPulseGuiding, TargetTime.Fast);
                        LogOk("IsPulseGuiding", isPulseGuiding.ToString());
                    }
                    catch (Exception ex) // Read failed
                    {
                        HandleException("IsPulseGuiding", MemberType.Property, Required.MustBeImplemented, ex, "CanPulseGuide is True");
                    }
                }
                else // Can't pulse guide so confirm that an error is raised
                {
                    try
                    {
                        LogCallToDriver("IsPulseGuiding", "About to get IsPulseGuiding property");
                        isPulseGuiding = telescopeDevice.IsPulseGuiding;
                        LogIssue("IsPulseGuiding", "CanPulseGuide is False but no error was raised on calling IsPulseGuiding");
                    }
                    catch (Exception ex)
                    {
                        HandleException("IsPulseGuiding", MemberType.Property, Required.MustNotBeImplemented, ex, "CanPulseGuide is False");
                    }
                }
            }
            else
            {
                LogInfo("IsPulseGuiding",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // RightAscension - Required
            try
            {
                LogCallToDriver("RightAscension", "About to get RightAscension property");
                rightAscension = TimeFunc("RightAscension", () => telescopeDevice.RightAscension, TargetTime.Fast);
                switch (rightAscension)
                {
                    case var case14 when case14 < 0.0d:
                    case var case15 when case15 >= 24.0d:
                        {
                            LogIssue("RightAscension",
                                $"RightAscension is <0 or >=24 hours: {rightAscension} {rightAscension.ToHMS().Trim()}");
                            break;
                        }

                    default:
                        {
                            LogOk("RightAscension", rightAscension.ToHMS().Trim());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("RightAscension", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // RightAscensionRate Read - Mandatory because read must always return 0.0
            try
            {
                LogCallToDriver("RightAscensionRate Read", "About to get RightAscensionRate property");
                rightAscensionRate = TimeFunc("RightAscensionRate Read", () => telescopeDevice.RightAscensionRate, TargetTime.Fast);
                // Read has been successful
                if (canSetRightAscensionRate) // Any value is acceptable
                {
                    switch (declinationRate)
                    {
                        case var case16 when case16 >= 0.0d:
                            {
                                LogOk("RightAscensionRate Read", rightAscensionRate.ToString("0.00"));
                                break;
                            }

                        default:
                            {
                                LogIssue("RightAscensionRate Read",
                                    $"Negative RightAscensionRate: {rightAscensionRate:0.00}");
                                break;
                            }
                    }
                }
                else // Only zero is acceptable
                {
                    switch (rightAscensionRate)
                    {
                        case 0.0d:
                            {
                                LogOk("RightAscensionRate Read", rightAscensionRate.ToString("0.00"));
                                break;
                            }

                        default:
                            {
                                LogIssue("RightAscensionRate Read",
                                    $"RightAscensionRate is non zero when CanSetRightAscensionRate is False {declinationRate:0.00}");
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!canSetRightAscensionRate)
                    LogInfo("RightAscensionRate Read", "RightAscensionRate must return 0 if CanSetRightAscensionRate is false.");
                HandleException("RightAscensionRate Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // RightAscensionRate Write - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (canSetRightAscensionRate) // Perform several tests starting with proving we can set a rate of 0.0
                {

                    SetTest("RightAscensionRate Write");

                    // Set up the rate offset Write test by ensuring that the telescope is in sidereal mode
                    try
                    {
                        telescopeDevice.TrackingRate = DriveRate.Sidereal;
                    }
                    catch { } // This is best endeavours, it doesn't matter if it fails at this point

                    // Log the test offset rate if extended tests are enabled
                    if (settings.TelescopeExtendedRateOffsetTests)
                    {
                        LogInfo("RightAscensionRate Write", $"Configured offset test duration: {settings.TelescopeRateOffsetTestDuration} seconds.");
                        LogInfo("RightAscensionRate Write", $"Configured offset rate (low):  {settings.TelescopeRateOffsetTestLowValue / 15.0,7:+0.0000;-0.0000;+0.0000} RA seconds per second ({settings.TelescopeRateOffsetTestLowValue,7:+0.000;-0.000;+0.000} arc seconds per second).");
                        LogInfo("RightAscensionRate Write", $"Configured offset rate (high): {settings.TelescopeRateOffsetTestHighValue / 15.0,7:+0.0000;-0.0000;+0.0000} RA seconds per second ({settings.TelescopeRateOffsetTestHighValue,7:+0.000;-0.000;+0.000} arc seconds per second).");
                    }

                    if (TestRaDecRate("RightAscensionRate Write", "Set rate to   0.000", Axis.Ra, 0.0d, true))
                    {
                        TestRaDecRate("RightAscensionRate Write", $"Set rate to {settings.TelescopeRateOffsetTestLowValue * ARC_SECONDS_TO_RA_SECONDS,7:+0.000;-0.000;+0.000}", Axis.Ra, settings.TelescopeRateOffsetTestLowValue * ARC_SECONDS_TO_RA_SECONDS, true);
                        if (cancellationToken.IsCancellationRequested) //
                            return;

                        TestRaDecRate("RightAscensionRate Write", $"Set rate to {-settings.TelescopeRateOffsetTestLowValue * ARC_SECONDS_TO_RA_SECONDS,7:+0.000;-0.000;+0.000}", Axis.Ra, -settings.TelescopeRateOffsetTestLowValue * ARC_SECONDS_TO_RA_SECONDS, true);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        TestRaDecRate("RightAscensionRate Write", $"Set rate to {settings.TelescopeRateOffsetTestHighValue * ARC_SECONDS_TO_RA_SECONDS,7:+0.000;-0.000;+0.000}", Axis.Ra, settings.TelescopeRateOffsetTestHighValue * ARC_SECONDS_TO_RA_SECONDS, true);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        TestRaDecRate("RightAscensionRate Write", $"Set rate to {-settings.TelescopeRateOffsetTestHighValue * ARC_SECONDS_TO_RA_SECONDS,7:+0.000;-0.000;+0.000}", Axis.Ra, -settings.TelescopeRateOffsetTestHighValue * ARC_SECONDS_TO_RA_SECONDS, true);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        if (TestRaDecRate("RightAscensionRate Write", "Set rate to   0.000", Axis.Ra, 0.0d, true))
                            if (cancellationToken.IsCancellationRequested)
                                return;

                        // Run the extended rate offset tests if configured to do so.
                        if (settings.TelescopeExtendedRateOffsetTests)
                        {
                            // Only test if the right ascension rate can be set
                            if (canSetRightAscensionRate) // Right ascension rate can be set
                            {
                                TestRAOffsetRates("RightAscensionRate Write", -9.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                TestRAOffsetRates("RightAscensionRate Write", 3.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                TestRAOffsetRates("RightAscensionRate Write", 9.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                TestRAOffsetRates("RightAscensionRate Write", -3.0);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                            }
                            else // Right ascension rate can not be set
                            {
                                LogInfo("RightAscensionRate Write", "Skipping extended RightAscensionRate tests because CanSetRightAscensionRate is false.");
                            }
                        }

                        TestRaDecRate("RightAscensionRate Write", "Reset rate to 0.0", Axis.Ra, 0.0d, false); // Reset the rate to zero, skipping the slewing test
                    }
                }
                else // Should generate an error
                {
                    try
                    {
                        LogCallToDriver("RightAscensionRate Write", "About to set RightAscensionRate property to 0.00");
                        telescopeDevice.RightAscensionRate = 0.0d; // Set to a harmless value
                        LogIssue("RightAscensionRate Write", "CanSetRightAscensionRate is False but setting RightAscensionRate did not generate an error");
                    }
                    catch (Exception ex)
                    {
                        HandleException("RightAscensionRate Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetRightAscensionRate is False");
                    }
                }
            }
            else // V1 has no Can property so just test anyway, it is optional
            {
                try
                {
                    LogCallToDriver("RightAscensionRate Write", "About to set RightAscensionRate property to 0.00");
                    telescopeDevice.RightAscensionRate = 0.0d; // Set to a harmless value
                    LogOk("RightAscensionRate Write", rightAscensionRate.ToString("0.00"));
                }
                catch (Exception ex)
                {
                    HandleException("RightAscensionRate Write", MemberType.Property, Required.Optional, ex, "");
                }
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            #region Site Elevation Tests

            // SiteElevation Read - Optional
            bool canReadSiteElevation = true;
            try
            {
                LogCallToDriver("SiteElevation Read", "About to get SiteElevation property");
                siteElevation = TimeFunc("SiteElevation Read", () => telescopeDevice.SiteElevation, TargetTime.Fast);
                switch (siteElevation)
                {
                    case var case17 when case17 < -300.0d:
                        {
                            LogIssue("SiteElevation Read", "SiteElevation is <-300m");
                            canReadSiteElevation = false;
                            break;
                        }

                    case var case18 when case18 > 10000.0d:
                        {
                            LogIssue("SiteElevation Read", "SiteElevation is >10,000m");
                            canReadSiteElevation = false;
                            break;
                        }

                    default:
                        {
                            LogOk("SiteElevation Read", siteElevation.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                canReadSiteElevation = false;
                HandleException("SiteElevation Read", MemberType.Property, Required.Optional, ex, "");
            }

            // SiteElevation Write - Invalid low value
            bool canWriteSiteElevation = true;
            try
            {
                LogCallToDriver("SiteElevation Write", "About to set SiteElevation property to -301.0");
                telescopeDevice.SiteElevation = -301.0d;
                LogIssue("SiteElevation Write", "No error generated on set site elevation < -300m");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOk("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation < -300m");
            }

            // SiteElevation Write - Invalid high value
            try
            {
                LogCallToDriver("SiteElevation Write", "About to set SiteElevation property to 100001.0");
                telescopeDevice.SiteElevation = 10001.0d;
                LogIssue("SiteElevation Write", "No error generated on set site elevation > 10,000m");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOk("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation > 10,000m");
            }

            //SiteElevation Write - Current device value 
            try
            {
                if (siteElevation < -300.0d | siteElevation > 10000.0d)
                    siteElevation = 1000d;
                LogCallToDriver("SiteElevation Write", $"About to set SiteElevation property to {siteElevation}");
                TimeMethod("SiteElevation Write", () => telescopeDevice.SiteElevation = siteElevation, TargetTime.Standard); // Restore original value
                LogOk("SiteElevation Write", $"Current value {siteElevation}m written successfully");
            }
            catch (Exception ex)
            {
                canWriteSiteElevation = false;
                HandleException("SiteElevation Write", MemberType.Property, Required.Optional, ex, "");
            }

            // Change the site elevation value
            if (canReadSiteElevation & canWriteSiteElevation & settings.TelescopeExtendedSiteTests)
            {
                try
                {
                    // Set the test value
                    LogCallToDriver("SiteElevation Write", $"About to set SiteElevation property to arbitrary value:{SITE_ELEVATION_TEST_VALUE}");
                    telescopeDevice.SiteElevation = SITE_ELEVATION_TEST_VALUE;

                    // Read the value back
                    LogCallToDriver("SiteElevation Write", "About to get SiteElevation property");
                    double newElevation = telescopeDevice.SiteElevation;

                    // Compare with the expected value
                    if (newElevation == SITE_ELEVATION_TEST_VALUE)
                    {
                        LogOk("SiteElevation Write", $"Test value {SITE_ELEVATION_TEST_VALUE} set and read correctly");
                    }
                    else
                    {
                        LogIssue("SiteElevation Write", $"Test value {SITE_ELEVATION_TEST_VALUE} did not round trip correctly. GET SiteElevations returned: {newElevation} instead of {SITE_ELEVATION_TEST_VALUE}");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("SiteElevation Write", MemberType.Property, Required.MustBeImplemented, ex, "A valid value could not be set");
                }

                // Attempt to restore the original value
                try
                {
                    // Set the original value
                    LogCallToDriver("SiteElevation Write", $"About to restore original SiteElevation property :{siteElevation}");
                    telescopeDevice.SiteElevation = siteElevation;
                    LogOk("SiteElevation Write", $"Successfully restored original site elevation: {siteElevation}.");
                }
                catch (Exception ex)
                {
                    HandleException("SiteElevation Write", MemberType.Property, Required.MustBeImplemented, ex, "The original value could not be restored");
                }
            }
            if (cancellationToken.IsCancellationRequested) return;

            #endregion

            #region Site Latitude Tests

            // SiteLatitude Read - Optional
            bool canReadSiteLatitude = true;
            try
            {
                LogCallToDriver("SiteLatitude Read", "About to get SiteLatitude property");
                siteLatitude = TimeFunc("SiteLatitude Read", () => telescopeDevice.SiteLatitude, TargetTime.Fast);
                switch (siteLatitude)
                {
                    case < -90.0d: // Invalid - less than -90.0
                        LogIssue("SiteLatitude Read", $"SiteLatitude is < -90.0 degrees: {siteLatitude.ToDMS()}");
                        canReadSiteLatitude = false;
                        break;

                    case > 90.0d: // Invalid - more than 90.0
                        LogIssue("SiteLatitude Read", $"SiteLatitude is > 90.0 degrees: {siteLatitude.ToDMS()}");
                        canReadSiteLatitude = false;
                        break;

                    default: // Valid - normal range
                        LogOk("SiteLatitude Read", siteLatitude.ToDMS());
                        break;
                }
            }
            catch (Exception ex)
            {
                canReadSiteLatitude = false;
                HandleException("SiteLatitude Read", MemberType.Property, Required.Optional, ex, "");
            }

            // Only try write tests if a valid latitude value was returned
            if (canReadSiteLatitude) // A valid value was received
            {
                // SiteLatitude Write - Invalid low value
                try
                {
                    LogCallToDriver("SiteLatitude Write", "About to set SiteLatitude property to -91.0");
                    telescopeDevice.SiteLatitude = -91.0d;
                    LogIssue("SiteLatitude Write", "No error generated on set site latitude < -90 degrees");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude < -90 degrees");
                }

                // SiteLatitude Write - Invalid high value
                try
                {
                    LogCallToDriver("SiteLatitude Write", "About to set SiteLatitude property to 91.0");
                    telescopeDevice.SiteLatitude = 91.0d;
                    LogIssue("SiteLatitude Write", "No error generated on set site latitude > 90 degrees");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude > 90 degrees");
                }

                // SiteLatitude Write - Valid value
                bool canWriteSiteLatitude = true;
                try
                {
                    LogCallToDriver("SiteLatitude Write", $"About to set SiteLatitude property to {siteLatitude}");
                    TimeMethod("SiteLatitude Write", () => telescopeDevice.SiteLatitude = siteLatitude, TargetTime.Standard); // Restore original value
                    LogOk("SiteLatitude Write", $"Current value: {siteLatitude.ToDMS()} degrees written successfully");
                }
                catch (Exception ex)
                {
                    canWriteSiteLatitude = false;
                    HandleException("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "");
                }

                // Change the site latitude value
                if (canWriteSiteLatitude & settings.TelescopeExtendedSiteTests)
                {
                    try
                    {
                        // Calculate the new test latitude
                        double testLatitude;

                        switch (siteLatitude)
                        {
                            // Latitude -90 to 70
                            case double d when d <= -70.0:
                                testLatitude = siteLatitude + 10.0;
                                break;

                            // Latitude -70 to 0
                            case double d when d > -70.0 & d <= 0.0:
                                testLatitude = siteLatitude - 10.0;
                                break;

                            // Latitude 0 to 70
                            case double d when d > 0.0 & d <= 70.0:
                                testLatitude = siteLatitude + 10.0;
                                break;

                            // Latitude 70 upwards
                            default:
                                testLatitude = siteLatitude - 10.0;
                                break;
                        }

                        // Set the test value
                        LogCallToDriver("SiteLatitude Write", $"About to set SiteLatitude property to arbitrary value:{testLatitude.ToDMS()}");
                        telescopeDevice.SiteLatitude = testLatitude;

                        // Read the value back
                        LogCallToDriver("SiteLatitude Write", "About to get SiteLatitude property");
                        double newLatitude = telescopeDevice.SiteLatitude;

                        // Compare with the expected value
                        if (newLatitude == testLatitude)
                        {
                            LogOk("SiteLatitude Write", $"Test value {testLatitude.ToDMS()} set and read correctly");
                        }
                        else
                        {
                            LogIssue("SiteLatitude Write", $"Test value {testLatitude.ToDMS()} did not round trip correctly. GET SiteLatitude returned: {newLatitude.ToDMS()} instead of {testLatitude.ToDMS()}");
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("SiteLatitude Write", MemberType.Property, Required.MustBeImplemented, ex, "A valid value could not be set");
                    }

                    // Attempt to restore the original value
                    try
                    {
                        // Set the original value
                        LogCallToDriver("SiteLatitude Write", $"About to restore original SiteLatitude property :{siteLatitude.ToDMS()}");
                        telescopeDevice.SiteLatitude = siteLatitude;
                        LogOk("SiteLatitude Write", $"Successfully restored original site latitude: {siteLatitude.ToDMS()}");
                    }
                    catch (Exception ex)
                    {
                        HandleException("SiteLatitude Write", MemberType.Property, Required.MustBeImplemented, ex, "The original value could not be restored");
                    }
                }
            }
            else // Cannot read
            {
                LogInfo("SiteLatitude Write", $"Tests skipped because SiteLatitude Read is not implemented, returned an invalid value or reported an error.");
            }
            if (cancellationToken.IsCancellationRequested) return;

            #endregion

            #region Site Longitude Tests

            // SiteLongitude Read - Optional
            bool canReadSiteLongitude = true;
            try
            {
                LogCallToDriver("SiteLongitude Read", "About to get SiteLongitude property");
                siteLongitude = TimeFunc("SiteLongitude Read", () => telescopeDevice.SiteLongitude, TargetTime.Fast);
                switch (siteLongitude)
                {
                    case var case21 when case21 < -180.0d:
                        {
                            canReadSiteLongitude = false;
                            LogIssue("SiteLongitude Read", "SiteLongitude is < -180 degrees");
                            break;
                        }

                    case var case22 when case22 > 180.0d:
                        {
                            canReadSiteLongitude = false;
                            LogIssue("SiteLongitude Read", "SiteLongitude is > 180 degrees");
                            break;
                        }

                    default:
                        {
                            LogOk("SiteLongitude Read", siteLongitude.ToDMS());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                canReadSiteLongitude = false;
                HandleException("SiteLongitude Read", MemberType.Property, Required.Optional, ex, "");
            }

            // SiteLongitude Write - Invalid low value
            try
            {
                LogCallToDriver("SiteLongitude Write", "About to set SiteLongitude property to -181.0");
                telescopeDevice.SiteLongitude = -181.0d;
                LogIssue("SiteLongitude Write", "No error generated on set site longitude < -180 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOk("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude < -180 degrees");
            }

            // SiteLongitude Write - Invalid high value
            try
            {
                LogCallToDriver("SiteLongitude Write", "About to set SiteLongitude property to 181.0");
                telescopeDevice.SiteLongitude = 181.0d;
                LogIssue("SiteLongitude Write", "No error generated on set site longitude > 180 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOk("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude > 180 degrees");
            }

            // SiteLongitude Write - Valid value
            bool canWriteSiteLongitude = true;
            try // Valid value
            {
                if (siteLongitude < -180.0d | siteLongitude > 180.0d)
                    siteLongitude = 60.0d;
                LogCallToDriver("SiteLongitude Write", $"About to set SiteLongitude property to {siteLongitude}");
                TimeMethod("SiteLatitude Write", () => telescopeDevice.SiteLongitude = siteLongitude, TargetTime.Standard); // Restore original value
                LogOk("SiteLongitude Write", $"Current value {siteLongitude.ToDMS()} degrees written successfully");
            }
            catch (Exception ex)
            {
                canWriteSiteLongitude = false;
                HandleException("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "");
            }

            // Change the site longitude value
            if (canReadSiteLongitude & canWriteSiteLongitude & settings.TelescopeExtendedSiteTests)
            {
                try
                {
                    // Calculate the new test longitude
                    double testLongitude;

                    switch (siteLongitude)
                    {
                        // Longitude -180 to -150
                        case double d when d <= -150.0:
                            testLongitude = siteLongitude + 10.0;
                            break;

                        // Longitude -150 to 0
                        case double d when d > -150.0 & d <= 0.0:
                            testLongitude = siteLongitude - 10.0;
                            break;

                        // Longitude 0 to 150
                        case double d when d > 0.0 & d <= 150.0:
                            testLongitude = siteLongitude + 10.0;
                            break;

                        // Longitude 150 -180
                        default:
                            testLongitude = siteLongitude - 10.0;
                            break;
                    }

                    // Set the test value
                    LogCallToDriver("SiteLongitude Write", $"About to set SiteLongitude property to arbitrary value:{testLongitude.ToDMS()}");
                    telescopeDevice.SiteLongitude = testLongitude;

                    // Read the value back
                    LogCallToDriver("SiteLongitude Write", "About to get SiteLongitude property");
                    double newLongitude = telescopeDevice.SiteLongitude;

                    // Compare with the expected value
                    if (newLongitude == testLongitude)
                    {
                        LogOk("SiteLongitude Write", $"Test value {testLongitude.ToDMS()} set and read correctly");
                    }
                    else
                    {
                        LogIssue("SiteLongitude Write", $"Test value {testLongitude.ToDMS()} did not round trip correctly. GET SiteLongitude returned: {newLongitude.ToDMS()} instead of {testLongitude.ToDMS()}");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("SiteLongitude Write", MemberType.Property, Required.MustBeImplemented, ex, "A valid value could not be set");
                }

                // Attempt to restore the original value
                try
                {
                    // Set the original value
                    LogCallToDriver("SiteLongitude Write", $"About to restore original SiteLongitude property :{siteLongitude.ToDMS()}");
                    telescopeDevice.SiteLongitude = siteLongitude;
                    LogOk("SiteLongitude Write", $"Successfully restored original site longitude: {siteLongitude.ToDMS()}");
                }
                catch (Exception ex)
                {
                    HandleException("SiteLongitude Write", MemberType.Property, Required.MustBeImplemented, ex, "The original value could not be restored");
                }
            }
            if (cancellationToken.IsCancellationRequested) return;

            #endregion

            // Slewing - Optional
            try
            {
                LogCallToDriver("Slewing", "About to get Slewing property");
                slewing = TimeFunc("Slewing", () => telescopeDevice.Slewing, TargetTime.Fast);
                switch (slewing)
                {
                    case false:
                        {
                            LogOk("Slewing", slewing.ToString());
                            break;
                        }

                    case true:
                        {
                            LogIssue("Slewing", $"Slewing should be false and it reads as {slewing}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Slewing", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // SlewSettleTime Read - Optional
            try
            {
                LogCallToDriver("SlewSettleTime Read", "About to get SlewSettleTime property");
                slewSettleTime = TimeFunc("SlewSettleTime Read", () => telescopeDevice.SlewSettleTime, TargetTime.Fast);
                switch (slewSettleTime)
                {
                    case var case23 when case23 < 0:
                        {
                            LogIssue("SlewSettleTime Read", "SlewSettleTime is < 0 seconds");
                            break;
                        }

                    case var case24 when case24 > (short)Math.Round(30.0d):
                        {
                            LogInfo("SlewSettleTime Read", "SlewSettleTime is > 30 seconds");
                            break;
                        }

                    default:
                        {
                            LogOk("SlewSettleTime Read", slewSettleTime.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("SlewSettleTime Read", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // SlewSettleTime Write - Optional
            try
            {
                LogCallToDriver("SlewSettleTime Write", "About to set SlewSettleTime property to -1");
                telescopeDevice.SlewSettleTime = -1;
                LogIssue("SlewSettleTime Write", "No error generated on set SlewSettleTime < 0 seconds");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOk("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set slew settle time < 0");
            }

            try
            {
                if (slewSettleTime < 0)
                    slewSettleTime = 0;
                LogCallToDriver("SlewSettleTime Write", $"About to set SlewSettleTime property to {slewSettleTime}");
                TimeMethod("SlewSettleTime Write", () => telescopeDevice.SlewSettleTime = slewSettleTime, TargetTime.Standard); // Restore original value
                LogOk("SlewSettleTime Write", $"Legal value {slewSettleTime} seconds written successfully");
            }
            catch (Exception ex)
            {
                HandleException("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // SideOfPier Read - Optional
            if (GetInterfaceVersion() > 1)
            {
                try
                {
                    LogCallToDriver("SideOfPier Read", "About to get SideOfPier property");
                    sideOfPier = TimeFunc("SideOfPier Read", () => telescopeDevice.SideOfPier, TargetTime.Fast);
                    LogOk("SideOfPier Read", sideOfPier.ToString());
                    canReadSideOfPier = true; // Flag that it is OK to read SideOfPier
                }
                catch (Exception ex)
                {
                    HandleException("SideOfPier Read", MemberType.Property, Required.Optional, ex, "");
                }
            }
            else
            {
                LogInfo("SideOfPier Read",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // SideOfPier Write - Optional
            // Moved to methods section as this really is a method rather than a property

            // SiderealTime - Required
            try
            {
                canReadSiderealTime = false;
                LogCallToDriver("SiderealTime", "About to get SiderealTime property");
                siderealTimeScope = TimeFunc("SiderealTime", () => telescopeDevice.SiderealTime, TargetTime.Fast);
                canReadSiderealTime = true;
                siderealTimeAscom = (18.697374558d + 24.065709824419081d * (DateTime.UtcNow.ToOADate() + 2415018.5 - 2451545.0d) + siteLongitude / 15.0d) % 24.0d;
                switch (siderealTimeScope)
                {
                    case var case25 when case25 < 0.0d:
                    case var case26 when case26 >= 24.0d:
                        {
                            LogIssue("SiderealTime", $"SiderealTime is <0 or >=24 hours: {siderealTimeScope.ToHMS()}"); // Valid time returned
                            break;
                        }

                    default:
                        {
                            // Now do a sense check on the received value
                            LogOk("SiderealTime", siderealTimeScope.ToHMS());
                            timeDifference = Math.Abs(siderealTimeScope - siderealTimeAscom); // Get time difference between scope and PC
                                                                                              // Process edge cases where the two clocks are on either side of 0:0:0/24:0:0
                            if (siderealTimeAscom > 23.0d & siderealTimeAscom < 23.999d & siderealTimeScope > 0.0d & siderealTimeScope < 1.0d)
                            {
                                timeDifference = Math.Abs(siderealTimeScope - siderealTimeAscom + 24.0d);
                            }

                            if (siderealTimeScope > 23.0d & siderealTimeScope < 23.999d & siderealTimeAscom > 0.0d & siderealTimeAscom < 1.0d)
                            {
                                timeDifference = Math.Abs(siderealTimeScope - siderealTimeAscom - 24.0d);
                            }

                            switch (timeDifference)
                            {
                                case var case27 when case27 <= 1.0d / 3600.0d: // 1 seconds
                                    {
                                        LogOk("SiderealTime",
                                            $"Scope and ASCOM sidereal times agree to better than 1 second, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        break;
                                    }

                                case var case28 when case28 <= 2.0d / 3600.0d: // 2 seconds
                                    {
                                        LogOk("SiderealTime",
                                            $"Scope and ASCOM sidereal times agree to better than 2 seconds, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        break;
                                    }

                                case var case29 when case29 <= 5.0d / 3600.0d: // 5 seconds
                                    {
                                        LogOk("SiderealTime",
                                            $"Scope and ASCOM sidereal times agree to better than 5 seconds, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        break;
                                    }

                                case var case30 when case30 <= 1.0d / 60.0d: // 1 minute
                                    {
                                        LogOk("SiderealTime",
                                            $"Scope and ASCOM sidereal times agree to better than 1 minute, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        break;
                                    }

                                case var case31 when case31 <= 5.0d / 60.0d: // 5 minutes
                                    {
                                        LogOk("SiderealTime",
                                            $"Scope and ASCOM sidereal times agree to better than 5 minutes, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        break;
                                    }

                                case var case32 when case32 <= 0.5d: // 0.5 an hour
                                    {
                                        LogInfo("SiderealTime",
                                            $"Scope and ASCOM sidereal times are up to 0.5 hour different, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        break;
                                    }

                                case var case33 when case33 <= 1.0d: // 1.0 an hour
                                    {
                                        LogInfo("SiderealTime",
                                            $"Scope and ASCOM sidereal times are up to 1.0 hour different, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        break;
                                    }

                                default:
                                    {
                                        LogIssue("SiderealTime",
                                            $"Scope and ASCOM sidereal times are more than 1 hour apart, Scope: {siderealTimeScope.ToHMS()}, ASCOM: {siderealTimeAscom.ToHMS()}");
                                        return;
                                    }
                            }

                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("SiderealTime", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetDeclination Read - Optional
            try // First read should fail!
            {
                // Test whether the command, which was executed earlier threw an error. if so throw it here.
                if (targetDeclinationReadException is not null)
                    throw targetDeclinationReadException;
                if (settings.TelescopeFirstUseTests)
                {
                    LogIssue("TargetDeclination Read", "Read before write should generate an error and didn't");
                    LogInfo("TargetDeclination Read", "This issue can be suppressed by unchecking Conform's \"Enable first time use tests\" setting in the telescope test section.");
                }
                else
                {
                    switch (targetDeclination)
                    {
                        case var case6 when case6 < -90.0d:
                        case var case7 when case7 > 90.0d:
                            LogIssue("TargetDeclination Read",
                                $"TargetDeclination is <-90 or >90 degrees: {targetDeclination.ToDMS()}");
                            break;

                        default:
                            {
                                LogOk("TargetDeclination Read", targetDeclination.ToDMS());
                                break;
                            }
                    }
                }
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == ExNotSet1 | ex.ErrorCode == ExNotSet2)
            {
                LogOk("TargetDeclination Read", "Not Set exception generated on read before write");
            }
            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == ExNotSet1 | ex.Number == ExNotSet2)
            {
                LogOk("TargetDeclination Read", "Not Set exception generated on read before write");
            }
            catch (Exception ex)
            {
                HandleInvalidOperationExceptionAsOk("TargetDeclination Read", MemberType.Property, Required.Optional, ex, "Incorrect exception received", "InvalidOperationException generated as expected on target read before read");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetDeclination Write - Optional
            LogInfo("TargetDeclination Write", "Tests moved after the SlewToCoordinates tests so that Conform can confirm that target coordinates are set as expected.");

            // TargetRightAscension Read - Optional
            try // First read should fail!
            {
                // Test whether the command, which was executed earlier threw an error. if so throw it here.
                if (targetRightAscensionReadException is not null)
                    throw targetRightAscensionReadException;
                if (settings.TelescopeFirstUseTests)
                {
                    LogIssue("TargetRightAscension Read", "Read before write should generate an error and didn't");
                    LogInfo("TargetRightAscension Read", "This issue can be suppressed by unchecking Conform's \"Enable first time use tests\" setting in the telescope test section.");
                }
                else
                {
                    switch (targetRightAscension)
                    {
                        case var case14 when case14 < 0.0d:
                        case var case15 when case15 >= 24.0d:
                            LogIssue("TargetRightAscension Read",
                                $"TargetRightAscension is <0 or >=24 hours: {targetRightAscension} {targetRightAscension.ToHMS()}");
                            break;

                        default:
                            LogOk("TargetRightAscension Read", targetRightAscension.ToHMS());
                            break;
                    }
                }
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == ExNotSet1 | ex.ErrorCode == ExNotSet2)
            {
                LogOk("TargetRightAscension Read", "Not Set exception generated on read before write");
            }
            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == ExNotSet1 | ex.Number == ExNotSet2)
            {
                LogOk("TargetRightAscension Read", "Not Set exception generated on read before write");
            }
            catch (Exception ex)
            {
                HandleInvalidOperationExceptionAsOk("TargetRightAscension Read", MemberType.Property, Required.Optional, ex, "Incorrect exception received", "InvalidOperationException generated as expected on target read before read");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetRightAscension Write - Optional
            LogInfo("TargetRightAscension Write", "Tests moved after the SlewToCoordinates tests so that Conform can confirm that target coordinates are set as expected.");

            // Tracking Read - Required
            try
            {
                canReadTracking = false;
                LogCallToDriver("Tracking Read", "About to get Tracking property");
                tracking = TimeFunc("Tracking Read", () => telescopeDevice.Tracking, TargetTime.Fast); // Read of tracking state is mandatory
                canReadTracking = true;

                LogOk("Tracking Read", tracking.ToString());
            }
            catch (Exception ex)
            {
                HandleException("Tracking Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Tracking Write - Optional
            lOriginalTrackingState = tracking;
            if (canSetTracking) // Set should work OK
            {
                SetTest("Tracking Write");
                try
                {
                    if (tracking) // OK try turning tracking off
                    {
                        LogCallToDriver("Tracking Write", "About to set Tracking property false");
                        TimeMethod("Tracking Write", () => telescopeDevice.Tracking = false, TargetTime.Standard);
                    }
                    else // OK try turning tracking on
                    {
                        LogCallToDriver("Tracking Write", "About to set Tracking property true");
                        TimeMethod("Tracking Write", () => telescopeDevice.Tracking = true, TargetTime.Standard);
                    }

                    SetAction("Waiting for mount to stabilise");
                    WaitFor(TRACKING_COMMAND_DELAY); // Wait for a short time to allow mounts to implement the tracking state change
                    LogCallToDriver("Tracking Write", "About to get Tracking property");
                    tracking = telescopeDevice.Tracking;
                    if (tracking != lOriginalTrackingState)
                    {
                        LogOk("Tracking Write", tracking.ToString());
                    }
                    else
                    {
                        LogIssue("Tracking Write", $"Tracking didn't change state on write: {tracking}");
                    }
                    LogCallToDriver("Tracking Write", $"About to set Tracking property {lOriginalTrackingState}");
                    telescopeDevice.Tracking = lOriginalTrackingState; // Restore original state
                    SetAction("Waiting for mount to stabilise");
                    WaitFor(TRACKING_COMMAND_DELAY); // Wait for a short time to allow mounts to implement the tracking state change
                }
                catch (Exception ex)
                {
                    HandleException("Tracking Write", MemberType.Property, Required.MustBeImplemented, ex, "CanSetTracking is True");
                }
                ClearStatus();
            }
            else // Can read OK but Set tracking should fail
            {
                try
                {
                    if (tracking) // OK try turning tracking off
                    {
                        LogCallToDriver("Tracking Write", "About to set Tracking property false");
                        telescopeDevice.Tracking = false;
                    }
                    else // OK try turning tracking on
                    {
                        LogCallToDriver("Tracking Write", "About to set Tracking property true");
                        telescopeDevice.Tracking = true;
                    }
                    LogCallToDriver("Tracking Write", "About to get Tracking property");
                    tracking = telescopeDevice.Tracking;
                    LogIssue("Tracking Write", "CanSetTracking is false but no error generated when value is set");
                }
                catch (Exception ex)
                {
                    HandleException("Tracking Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetTracking is False");
                }
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // TrackingRates - Required
            if (GetInterfaceVersion() > 1)
            {
                int lCount = 0;
                try
                {
                    LogCallToDriver("TrackingRates", "About to get TrackingRates property");
                    trackingRates = TimeFunc("Tracking Write", () => telescopeDevice.TrackingRates, TargetTime.Fast);
                    if (trackingRates is null)
                    {
                        LogDebug("TrackingRates", "ERROR: The driver did NOT return an TrackingRates object!");
                    }
                    else
                    {
                        LogDebug("TrackingRates", "OK - the driver returned an TrackingRates object");
                    }

                    lCount = trackingRates.Count; // Save count for use later if no members are returned in the for each loop test
                    LogDebug("TrackingRates Count", lCount.ToString());

                    var loopTo = trackingRates.Count;
                    for (int ii = 1; ii <= loopTo; ii++)
                        LogDebug("TrackingRates Count", $"Found drive rate: {Enum.GetName(typeof(DriveRate), (trackingRates[ii]))}");
                }
                catch (Exception ex)
                {
                    HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (trackingRates is not null)
                {
                    try
                    {
                        IEnumerator lEnum;
                        object lObj;
                        DriveRate lDrv;
                        lEnum = (IEnumerator)trackingRates.GetEnumerator();
                        if (lEnum is null)
                        {
                            LogDebug("TrackingRates Enum", "ERROR: The driver did NOT return an Enumerator object!");
                        }
                        else
                        {
                            LogDebug("TrackingRates Enum", "OK - the driver returned an Enumerator object");
                        }

                        lEnum.Reset();
                        LogDebug("TrackingRates Enum", "Reset Enumerator");
                        while (lEnum.MoveNext())
                        {
                            LogDebug("TrackingRates Enum", "Reading Current");
                            lObj = lEnum.Current;
                            LogDebug("TrackingRates Enum", $"Read Current OK, Type: {lObj.GetType().Name}");
                            lDrv = (DriveRate)lObj;
                            LogDebug("TrackingRates Enum", $"Found drive rate: {Enum.GetName(typeof(DriveRate), lDrv)}");
                        }

                        lEnum.Reset();
                        lEnum = null;

                        // Clean up TrackingRates object
                        if (trackingRates is object)
                        {
                            try
                            {
                                trackingRates.Dispose();
                            }
                            catch
                            {
                            }

                            if (OperatingSystem.IsWindows())
                            {
                                try
                                {
                                    Marshal.ReleaseComObject(trackingRates);
                                }
                                catch
                                {
                                }
                            }

                            trackingRates = null;
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "");
                    }
                }
                else
                {
                    LogInfo("TrackingRates Enum", "Skipped enumerator test because of an issue creating the TrackingRates object");
                }

                try
                {
                    LogCallToDriver("TrackingRates", "About to get TrackingRates property");
                    trackingRates = telescopeDevice.TrackingRates;
                    LogDebug("TrackingRates", $"Read TrackingRates OK, Count: {trackingRates.Count}");
                    int lRateCount = 0;
                    foreach (DriveRate currentLDriveRate in (IEnumerable)trackingRates)
                    {
                        driveRate = currentLDriveRate;
                        LogTestAndMessage("TrackingRates", $"Found drive rate: {driveRate}");
                        lRateCount += 1;
                    }

                    if (lRateCount > 0)
                    {
                        LogOk("TrackingRates", "Drive rates read OK");
                    }
                    else if (lCount > 0) // We did get some members on the first call, but now they have disappeared!
                    {
                        // This can be due to the driver returning the same TrackingRates object on every TrackingRates call but not resetting the iterator pointer
                        LogIssue("TrackingRates", "Multiple calls to TrackingRates returned different answers!");
                        LogInfo("TrackingRates", "");
                        LogInfo("TrackingRates", $"The first call to TrackingRates returned {lCount} drive rates; the next call appeared to return no rates.");
                        LogInfo("TrackingRates", "This can arise when the SAME TrackingRates object is returned on every TrackingRates call.");
                        LogInfo("TrackingRates", "The root cause is usually that the enumeration pointer in the object is set to the end of the");
                        LogInfo("TrackingRates", "collection through the application's use of the first object; subsequent uses see the pointer at the end");
                        LogInfo("TrackingRates", "of the collection, which indicates no more members and is interpreted as meaning the collection is empty.");
                        LogInfo("TrackingRates", "");
                        LogInfo("TrackingRates", "It is recommended to return a new TrackingRates object on each call. Alternatively, you could reset the");
                        LogInfo("TrackingRates", "object's enumeration pointer every time the GetEnumerator method is called.");
                        LogInfo("TrackingRates", "");
                    }
                    else
                    {
                        LogIssue("TrackingRates", "No drive rates returned");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "");
                }

                // Clean up TrackingRates object
                if (trackingRates is object)
                {
                    try
                    {
                        trackingRates.Dispose();
                    }
                    catch
                    {
                    }

                    if (OperatingSystem.IsWindows())
                    {
                        try
                        {
                            Marshal.ReleaseComObject(trackingRates);
                        }
                        catch { }
                    }
                }

                // Test the TrackingRates.Dispose() method
                LogDebug("TrackingRates", "Getting tracking rates");
                trackingRates = telescopeDevice.TrackingRates;
                try
                {
                    LogDebug("TrackingRates", "Disposing tracking rates");
                    trackingRates.Dispose();
                    LogOk("TrackingRates", "Disposed tracking rates OK");
                }
                catch (MissingMemberException)
                {
                    LogOk("TrackingRates", "Dispose member not present");
                }
                catch (Exception ex)
                {
                    LogIssue("TrackingRates",
                        $"TrackingRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: {ex.Message}");
                    LogDebug("TrackingRates.Dispose", $"Exception: {ex}");
                }

                if (OperatingSystem.IsWindows())
                {
                    try
                    {
                        Marshal.ReleaseComObject(trackingRates);
                    }
                    catch { }
                }

                if (cancellationToken.IsCancellationRequested)
                    return;

            }
            else
            {
                LogInfo("TrackingRates", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // TrackingRate Read - Required and Write - Optional. Test after TrackingRates so we know what the valid values are
            if (GetInterfaceVersion() > 1)
            {
                try
                {
                    LogCallToDriver("TrackingRates", "About to get TrackingRates property");
                    trackingRates = telescopeDevice.TrackingRates;

                    // Make sure that we have received a TrackingRates object after the Dispose() method was called
                    if (trackingRates is object) // We did get a TrackingRates object
                    {
                        LogOk("TrackingRates", "Successfully obtained a TrackingRates object after the previous TrackingRates object was disposed");
                        LogCallToDriver("TrackingRate Read", "About to get TrackingRate property");
                        trackingRate = telescopeDevice.TrackingRate;
                        LogOk("TrackingRate Read", trackingRate.ToString());

                        // TrackingRate Write - Optional

                        // Check whether TrackingRate SET is implemented at all
                        bool canSetTrackingRate = false; // Default value to false
                        try
                        {
                            telescopeDevice.TrackingRate = DriveRate.Sidereal;

                            // If we get here tracking rate can be set
                            canSetTrackingRate = true;
                        }
                        catch (InvalidValueException)
                        {
                            // Tracking rate can be set but not to Sidereal rate
                            canSetTrackingRate = true;
                        }
                        catch (Exception ex)
                        {
                            canSetTrackingRate = false;
                            HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "");
                        }

                        if (canSetTrackingRate)
                        {
                            try
                            {
                                // Create a list to hold the supplied drive rates
                                List<DriveRate> driveRates = new List<DriveRate>();

                                // We can read TrackingRate so now test trying to set each tracking rate in turn
                                LogDebug("TrackingRate Write", "About to enumerate tracking rates object");
                                foreach (DriveRate currentDriveRate in (IEnumerable)trackingRates)
                                {
                                    // Exit if required
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    // Save the supplied drive rate
                                    driveRates.Add(currentDriveRate);

                                    try
                                    {
                                        // Set the tracking rate
                                        LogCallToDriver("TrackingRate Write", $"About to set TrackingRate property to {currentDriveRate}");
                                        telescopeDevice.TrackingRate = currentDriveRate;

                                        // Make sure that the set rate is returned
                                        LogCallToDriver("TrackingRate Write", $"About to get TrackingRate property");
                                        if (telescopeDevice.TrackingRate == currentDriveRate)
                                        {
                                            LogOk("TrackingRate Write", $"Successfully set drive rate: {currentDriveRate}");
                                        }
                                        else
                                        {
                                            LogIssue("TrackingRate Write", $"Unable to set drive rate: {currentDriveRate}");
                                        }

                                        // For ITelescopeV4 and later make sure that RightAscensionRate & DeclinationRate are only usable when tracking at Sidereal rate
                                        if (IsPlatform7OrLater)
                                        {
                                            CheckRateOffsetsRejectedForNonSiderealDriveRates(currentDriveRate, RateOffset.DeclinationRate);
                                            CheckRateOffsetsRejectedForNonSiderealDriveRates(currentDriveRate, RateOffset.RightAscensionRate);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "");
                                    }
                                }

                                //Now check that other drive rates are rejected by iterating over all possible values
                                foreach (DriveRate rate in Enum.GetValues(typeof(DriveRate)))
                                {
                                    LogDebug("TrackingRate Write", $"Checking drive rate {rate}");
                                    if (!driveRates.Contains(rate)) // This is not a supported drive rate so it should be rejected
                                    {
                                        LogDebug("TrackingRate Write", $"Making sure that drive rate {rate} is rejected because it is not supported.");
                                        try
                                        {
                                            telescopeDevice.TrackingRate = rate;
                                            LogIssue("TrackingRate Write", $"Tracking rate {rate} is not supported but did not throw an InvalidValueException exception.");
                                        }
                                        catch (Exception ex)
                                        {
                                            HandleInvalidValueExceptionAsOk("TrackingRate Write", MemberType.Property, Required.Optional, ex, $"", $"Drive rate {rate} is not supported but did not throw an InvalidValueException as expected.");
                                        }
                                    }
                                    else // This is a supported drive rate
                                    {
                                        // No action because it will have been tested above
                                        LogDebug("TrackingRate Write", $"No action for drive rate {rate} because it was tested above.");
                                    }
                                }
                            }
                            catch (NullReferenceException) // Catch issues in iterating over a new TrackingRates object after a previous TrackingRates object was disposed.
                            {
                                LogIssue("TrackingRate Write 1", "A NullReferenceException was thrown while iterating a new TrackingRates instance after a previous TrackingRates instance was disposed. TrackingRate.Write testing skipped");
                                LogInfo("TrackingRate Write 1", "This may indicate that the TrackingRates.Dispose method cleared a global variable shared by all TrackingRates instances.");
                            }
                            catch (Exception ex)
                            {
                                HandleException("TrackingRate Write 1", MemberType.Property, Required.Mandatory, ex, "");
                            }

                            // Attempt to write an invalid high tracking rate
                            try
                            {
                                LogCallToDriver("TrackingRate Write", "About to set TrackingRate property to invalid value (5)");
                                telescopeDevice.TrackingRate = (DriveRate)5;
                                LogIssue("TrackingRate Write", "No error generated when TrackingRate is set to an invalid value (5)");
                            }
                            catch (Exception ex)
                            {
                                HandleInvalidValueExceptionAsOk("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (5)");
                            }

                            // Attempt to write an invalid low tracking rate
                            try
                            {
                                LogCallToDriver("TrackingRate Write", "About to set TrackingRate property to invalid value (-1)");
                                telescopeDevice.TrackingRate = (DriveRate)(0 - 1); // Done this way to fool the compiler into allowing me to attempt to set a negative, invalid value
                                LogIssue("TrackingRate Write", "No error generated when TrackingRate is set to an invalid value (-1)");
                            }
                            catch (Exception ex)
                            {
                                HandleInvalidValueExceptionAsOk("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (-1)");
                            }

                            // Finally restore original TrackingRate
                            try
                            {
                                LogCallToDriver("TrackingRate Write", "About to set TrackingRate property to " + trackingRate.ToString());
                                telescopeDevice.TrackingRate = trackingRate;
                            }
                            catch (Exception ex)
                            {
                                HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "Unable to restore original tracking rate");
                            }
                        }
                    }
                    else // No TrackingRates object received after disposing of a previous instance
                    {
                        LogIssue("TrackingRate Write", "TrackingRates did not return an object after calling Disposed() on a previous instance, TrackingRate.Write testing skipped");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("TrackingRate Read", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("TrackingRate", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // UTCDate Read - Required
            try
            {
                LogCallToDriver("UTCDate Read", "About to get UTCDate property");
                utcDate = TimeFunc("UTCDate Read", () => telescopeDevice.UTCDate, TargetTime.Fast); // Save starting value
                LogOk("UTCDate Read", utcDate.ToString("dd-MMM-yyyy HH:mm:ss.fff"));

                try // UTCDate Write is optional since if you are using the PC time as UTCTime then you should not write to the PC clock!
                {
                    // Try to write a new UTCDate  1 hour in the future
                    LogCallToDriver("UTCDate Write", $"About to set UTCDate property to {utcDate.AddHours(1.0d)}");
                    TimeMethod("UTCDate Write", () => telescopeDevice.UTCDate = utcDate.AddHours(1.0d), TargetTime.Standard);
                    LogOk("UTCDate Write", $"New UTCDate written successfully: {utcDate.AddHours(1.0d)}");

                    // Restore original value
                    LogCallToDriver("UTCDate Write", $"About to set UTCDate property to {utcDate}");
                    telescopeDevice.UTCDate = utcDate;
                    LogOk("UTCDate Write", $"Original UTCDate restored successfully: {utcDate}");
                }
                catch (Exception ex)
                {
                    HandleException("UTCDate Write", MemberType.Property, Required.Optional, ex, "");
                }
            }
            catch (Exception ex)
            {
                HandleException("UTCDate Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;
        }

        public override void CheckMethods()
        {

            // CanMoveAxis - Required - This must be first test as Parked tests use its results
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[CAN_MOVE_AXIS] | telescopeTests[MOVE_AXIS] | telescopeTests[PARK_UNPARK])
                {
                    TelescopeRequiredMethodsTest(RequiredMethodType.TstCanMoveAxisPrimary, "CanMoveAxis:Primary");
                    if (cancellationToken.IsCancellationRequested) return;
                    TelescopeRequiredMethodsTest(RequiredMethodType.TstCanMoveAxisSecondary, "CanMoveAxis:Secondary");
                    if (cancellationToken.IsCancellationRequested) return;
                    TelescopeRequiredMethodsTest(RequiredMethodType.TstCanMoveAxisTertiary, "CanMoveAxis:Tertiary");
                    if (cancellationToken.IsCancellationRequested) return;
                }
                else
                {
                    LogInfo(CAN_MOVE_AXIS, "Tests skipped");
                }
            }
            else
            {
                LogInfo("CanMoveAxis",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // Test Park, Unpark - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[PARK_UNPARK])
                {
                    if (canPark) // Can Park
                    {
                        try
                        {
                            LogCallToDriver("Park", "About to get AtPark property");
                            if (!telescopeDevice.AtPark) // OK We are unparked so check that no error is generated
                            {
                                SetTest("Park");
                                try
                                {
                                    // Try to slew to an arbitrary start position that is not likely to be the Park position
                                    try
                                    {
                                        SetAction("Slewing to arbitrary start position");
                                        SlewToHa(-3.0);
                                    }
                                    catch (Exception ex)
                                    {
                                        LogDebug("Park", $"Exception when attempting to slew to start position: {ex.Message}");
                                        LogDebug("Park", ex.ToString());
                                    }

                                    SetAction("Parking scope...");
                                    LogTestAndMessage("Park", "Parking scope...");
                                    LogCallToDriver("Park", "About to call Park method");

                                    TimeMethod($"Park", () => telescopeDevice.Park(), IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

                                    // Wait for the park to complete using Platform 6 or 7 semantics as appropriate
                                    if (IsPlatform7OrLater) // Platform 7 or later device
                                    {
                                        LogCallToDriver("Park", "About to get Slewing property repeatedly (ITelescopeV4 interface)...");
                                        WaitWhile("Waiting for scope to park", () => telescopeDevice.Slewing, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                    }
                                    else // Platform 6 device
                                    {
                                        LogCallToDriver("Park", "About to get AtPark property repeatedly (ITelescopeV2/3 interface)...");
                                        WaitWhile("Waiting for scope to park", () => !telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                    }

                                    // Test outcome
                                    LogCallToDriver("Park", "About to get AtPark property");
                                    if (telescopeDevice.AtPark)
                                    {
                                        SetStatus("Scope parked");
                                        LogOk("Park", "Parked successfully (AtPark is true).");

                                        if (cancellationToken.IsCancellationRequested) return;

                                        // Scope Parked OK - Confirm that a second park is harmless
                                        try
                                        {
                                            telescopeDevice.Park();

                                            // Wait for the park to complete using Platform 6 or 7 semantics as appropriate
                                            if (IsPlatform7OrLater) // Platform 7 or later device
                                            {
                                                LogCallToDriver("Park", "About to get Slewing property repeatedly...");
                                                WaitWhile("Waiting for scope to park", () => telescopeDevice.Slewing, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                            }
                                            else // Platform 6 device
                                            {
                                                LogCallToDriver("Park", "About to get AtPark property repeatedly...");
                                                WaitWhile("Waiting for scope to park", () => !telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                            }

                                            // Test outcome
                                            LogCallToDriver("Park", "About to get AtPark property");
                                            if (telescopeDevice.AtPark)
                                            {
                                                SetStatus("Scope parked");
                                                LogOk("Park", "Calling Park twice is harmless and leaves the mount in the parked state.");
                                            }
                                            else
                                            {
                                                LogIssue("Park", $"Failed to park within {settings.TelescopeMaximumSlewTime} seconds.");
                                            }

                                            if (cancellationToken.IsCancellationRequested) return;
                                        }
                                        catch (COMException ex)
                                        {
                                            LogIssue("Park", $"Exception when calling Park two times in succession: {ex.Message} {((int)ex.ErrorCode):X8}");
                                        }
                                        catch (Exception ex)
                                        {
                                            LogIssue("Park", $"Exception when calling Park two times in succession: {ex.Message}");
                                        }

                                        // Confirm that methods do raise exceptions when scope is parked
                                        if (canSlew | canSlewAsync | canSlewAltAz | canSlewAltAzAsync)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepAbortSlew, "AbortSlew");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canFindHome)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepFindHome, "FindHome");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canMoveAxisPrimary)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepMoveAxisPrimary, "MoveAxis Primary");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canMoveAxisSecondary)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepMoveAxisSecondary, "MoveAxis Secondary");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canMoveAxisTertiary)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepMoveAxisTertiary, "MoveAxis Tertiary");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canPulseGuide)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepPulseGuide, "PulseGuide");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canSlew)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepSlewToCoordinates, "SlewToCoordinates");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canSlewAsync)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepSlewToCoordinatesAsync, "SlewToCoordinatesAsync");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canSlew)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepSlewToTarget, "SlewToTarget");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canSlewAsync)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepSlewToTargetAsync, "SlewToTargetAsync");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canSync)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepSyncToCoordinates, "SyncToCoordinates");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        if (canSync)
                                        {
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstPExcepSyncToTarget, "SyncToTarget");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        // Test Tracking behaviour when parked for ITelescopeV4 and later
                                        if (canSetTracking & IsPlatform7OrLater)
                                        {
                                            // Confirm that setting Tracking = true results in a ParkedException
                                            TelescopeParkedExceptionTest(ParkedExceptionType.TstExcepTracking, "Tracking = true");
                                            if (cancellationToken.IsCancellationRequested)
                                                return;

                                            // Confirm that setting Tracking = false works and doesn't result in an exception
                                            try
                                            {
                                                LogCallToDriver("Parked", "About to set Tracking to false");
                                                telescopeDevice.Tracking = false;
                                                LogOk("Parked:Tracking = false", $"Tracking = false was successful when parked.");
                                            }
                                            catch (Exception ex)
                                            {
                                                HandleException("Parked - Tracking=False", MemberType.Property, Required.MustBeImplemented, ex, "Setting Tracking false when parked must be successful for this interface version.");
                                            }
                                        }
                                    }
                                    else // Scope is not parked - first Park attempt failed
                                    {
                                        LogIssue("Park", $"Failed to park within {settings.TelescopeMaximumSlewTime} seconds (AtPark is false), skipping further Park tests.");
                                    }

                                    // Test unpark after park
                                    if (canUnpark)
                                    {
                                        try
                                        {
                                            SetStatus("Unparking scope");

                                            LogCallToDriver("Unpark", "About to call Unpark method");
                                            TimeMethod($"Unpark", () => telescopeDevice.Unpark(), IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

                                            // Wait for the un park to complete using Platform 6 or 7 semantics as appropriate
                                            if (IsPlatform7OrLater) // Platform 7 or later device
                                            {
                                                LogCallToDriver("Unpark", "About to get Slewing property repeatedly...");
                                                WaitWhile("Waiting for scope to unpark when parked", () => telescopeDevice.Slewing, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                            }
                                            else // Platform 6 device
                                            {
                                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly...");
                                                WaitWhile("Waiting for scope to unpark when parked", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                            }

                                            // Validate unparking
                                            LogCallToDriver("Unpark", "About to get AtPark property");
                                            if (!telescopeDevice.AtPark) // Scope unparked OK
                                            {
                                                LogOk("Unpark", "Unparked successfully.");
                                                SetStatus("Scope Unparked OK");

                                                // Enable tracking
                                                try // Make sure tracking doesn't generate an error if it is not implemented
                                                {
                                                    LogCallToDriver("Unpark", "About to set Tracking property true");
                                                    telescopeDevice.Tracking = true;
                                                }
                                                catch { }

                                                if (cancellationToken.IsCancellationRequested)
                                                    return;

                                                // Scope is unparked
                                                try // Confirm Unpark is harmless if already unparked
                                                {
                                                    LogCallToDriver("Unpark", "About to call Unpark method");
                                                    telescopeDevice.Unpark();

                                                    // Wait for the unpark to complete using Platform 6 or 7 semantics as appropriate
                                                    if (IsPlatform7OrLater) // Platform 7 or later device
                                                    {
                                                        LogCallToDriver("Unpark", "About to get Slewing property repeatedly...");
                                                        WaitWhile("Waiting for scope to unpark when parked", () => telescopeDevice.Slewing, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                                    }
                                                    else // Platform 6 device
                                                    {
                                                        LogCallToDriver("Unpark", "About to get AtPark property repeatedly...");
                                                        WaitWhile("Waiting for scope to unpark when parked", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                                    }

                                                    // Validate unparking
                                                    LogCallToDriver("Unpark", "About to get AtPark property");
                                                    if (!telescopeDevice.AtPark) // Scope remained unparked
                                                    {
                                                        LogOk("Unpark", "Calling Unpark twice is harmless and leaves the mount in the un-parked state.");
                                                        SetStatus("Scope Unparked OK");
                                                    }
                                                    else // Scope is still parked - it did not unpark
                                                    {
                                                        LogIssue("Unpark", $"Failed to unpark after {settings.TelescopeMaximumSlewTime} seconds when sending a second Unpark command.");
                                                    }
                                                }
                                                catch (COMException ex)
                                                {
                                                    LogIssue("Unpark",
                                                        $"Exception when calling Unpark two times in succession: {ex.Message} {((int)ex.ErrorCode):X8}");
                                                }
                                                catch (Exception ex)
                                                {
                                                    LogIssue("Unpark",
                                                        $"Exception when calling Unpark two times in succession: {ex.Message}");
                                                }
                                            }
                                            else // Scope is still parked - it did not unpark
                                            {
                                                LogIssue("Unpark", $"AtPark remains True - the scope failed to unpark.");
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            HandleException("Unpark", MemberType.Method, Required.MustBeImplemented, ex, "CanUnpark is true");
                                        }
                                    }
                                    else // Can't Unpark
                                    {
                                        // Confirm that Unpark generates an error
                                        try
                                        {
                                            LogCallToDriver("Unpark", "About to call Unpark method");
                                            telescopeDevice.Unpark();
                                            LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                            WaitWhile("Waiting for scope to unpark", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                            LogIssue("Unpark", "No exception thrown by Unpark when CanUnpark is false");
                                        }
                                        catch (Exception ex)
                                        {
                                            HandleException("Unpark", MemberType.Method, Required.MustNotBeImplemented, ex, "CanUnpark is false");
                                        }
                                        // Create user interface message asking for manual scope Unpark
                                        LogTestAndMessage("Unpark", "CanUnpark is false so you need to unpark manually");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Park", MemberType.Method, Required.MustBeImplemented, ex, "CanPark is true");
                                }
                            }
                            else // We are still in parked status despite a successful UnPark
                            {
                                LogIssue("Park", "AtPark still true despite an earlier successful unpark");
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleException("Park", MemberType.Method, Required.MustBeImplemented, ex, "CanPark is True");
                        }
                    }
                    else // Can't park so Park() should fail
                    {
                        try
                        {
                            LogCallToDriver("Park", "About to Park call method");
                            telescopeDevice.Park();
                            LogOk("Park", "Success if already parked");

                            LogCallToDriver("Park", "About to get AtPark property repeatedly...");

                            // Wait for the park to complete
                            WaitWhile("Waiting for scope to park", () => !telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                            if (cancellationToken.IsCancellationRequested) return;

                            SetStatus("Scope parked");
                            LogIssue("Park", "CanPark is false but no exception was generated on use");
                        }
                        catch (Exception ex)
                        {
                            HandleException("Park", MemberType.Method, Required.MustNotBeImplemented, ex, "CanPark is False");
                        }

                        // Now test unpark
                        if (canUnpark) // We should already be unparked so confirm that unpark works fine
                        {
                            try
                            {
                                LogCallToDriver("UnPark", "About to call UnPark method");
                                telescopeDevice.Unpark();
                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                WaitWhile("Waiting for scope to unpark", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                LogOk("UnPark", "CanPark is false and CanUnPark is true; no exception generated as expected");
                            }
                            catch (Exception ex)
                            {
                                HandleException("UnPark", MemberType.Method, Required.MustBeImplemented, ex, "CanUnPark is True");
                            }
                        }
                        else // Can't unpark so confirm an exception is raised
                        {
                            try
                            {
                                LogCallToDriver("UnPark", "About to call UnPark method");
                                telescopeDevice.Unpark();
                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                WaitWhile("Waiting for scope to unpark", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                LogIssue("UnPark", "CanPark and CanUnPark are false but no exception was generated on use");
                            }
                            catch (Exception ex)
                            {
                                HandleException("UnPark", MemberType.Method, Required.MustNotBeImplemented, ex, "CanUnPark is False");
                            }
                        }
                    }

                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
                else
                {
                    LogInfo(PARK_UNPARK, "Tests skipped");
                }
            }
            else
            {
                LogInfo("Park",
                    $"Skipping tests since behaviour of this method is not well defined in interface V{GetInterfaceVersion()}");
            }

            // AbortSlew - Optional
            if (telescopeTests[ABORT_SLEW])
            {
                // Check whether AbortSlew throws a not implemented error when called on its own
                try
                {
                    // Call AbortSlew
                    AbortSlew("AbortSlew");

                    // If we get here AbortSlew did not throw an exception so test its implementation
                    LogDebug("AbortSlew", $"The AbortSlew command did not throw an exception");
                    TelescopeOptionalMethodsTest(OptionalMethodType.AbortSlew, "AbortSlew", true);
                }
                catch (Exception ex)
                {
                    // AbortSlew threw an exception so process it accepting NotImplemnted exceptions
                    LogDebug("AbortSlew", $"The AbortSlew command did throw an exception");
                    HandleException("AbortSlew", MemberType.Method, Required.Optional, ex, "");
                }

                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo(ABORT_SLEW, "Tests skipped");
            }

            // AxisRates - Required
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[AXIS_RATE] | telescopeTests[MOVE_AXIS])
                {
                    TelescopeAxisRateTest("AxisRate:Primary", TelescopeAxis.Primary);
                    TelescopeAxisRateTest("AxisRate:Secondary", TelescopeAxis.Secondary);
                    TelescopeAxisRateTest("AxisRate:Tertiary", TelescopeAxis.Tertiary);
                }
                else
                {
                    LogInfo(AXIS_RATE, "Tests skipped");
                }
            }
            else
            {
                LogInfo("AxisRate",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // FindHome - Optional
            if (telescopeTests[FIND_HOME])
            {
                TelescopeOptionalMethodsTest(OptionalMethodType.FindHome, "FindHome", canFindHome);
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo(FIND_HOME, "Tests skipped");
            }

            // MoveAxis - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[MOVE_AXIS])
                {
                    TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisPrimary, "MoveAxis Primary", canMoveAxisPrimary);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisSecondary, "MoveAxis Secondary", canMoveAxisSecondary);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisTertiary, "MoveAxis Tertiary", canMoveAxisTertiary);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
                else
                {
                    LogInfo(MOVE_AXIS, "Tests skipped");
                }
            }
            else
            {
                LogInfo("MoveAxis",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // PulseGuide - Optional
            if (telescopeTests[PULSE_GUIDE])
            {
                TelescopeOptionalMethodsTest(OptionalMethodType.PulseGuide, "PulseGuide", canPulseGuide);
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Check whether the extended pulse guide tests are enabled
                if (settings.TelescopeExtendedPulseGuideTests)
                {
                    // Test whether this device can pulse guide
                    if (canPulseGuide)
                    {
                        // Test whether positive, non-zero, values have been provided for both RA and declination axes.
                        if ((guideRateRightAscension >= 0.0) & (guideRateDeclination >= 0.0))
                        {
                            TestPulseGuide(-9.0);
                            if (ApplicationCancellationToken.IsCancellationRequested)
                                return;

                            TestPulseGuide(+9.0);
                            if (ApplicationCancellationToken.IsCancellationRequested)
                                return;

                            TestPulseGuide(-3.0);
                            if (ApplicationCancellationToken.IsCancellationRequested)
                                return;

                            TestPulseGuide(+3.0);
                        }
                        else
                            LogIssue(PULSE_GUIDE, $"Extended pulse guide tests skipped because at least one of the GuideRateRightAscension or GuideRateDeclination properties was not implemented or returned a zero or negative value.");
                    }
                    else
                        LogInfo(PULSE_GUIDE, $"Extended pulse guide tests skipped because the CanPulseGuide property returned False.");
                }
            }
            else
                LogInfo(PULSE_GUIDE, "Tests skipped");

            // Test Equatorial slewing to coordinates - Optional
            if (telescopeTests[SLEW_TO_COORDINATES])
            {
                TelescopeSlewTest(SlewSyncType.SlewToCoordinates, "SlewToCoordinates", canSlew, "CanSlew");
                if (cancellationToken.IsCancellationRequested)
                    return;
                if (canSlew) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToCoordinates (Bad L)", SlewSyncType.SlewToCoordinates, BAD_RA_LOW, BAD_DEC_LOW);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeBadCoordinateTest("SlewToCoordinates (Bad H)", SlewSyncType.SlewToCoordinates, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
            {
                LogInfo(SLEW_TO_COORDINATES, "Tests skipped");
            }

            // Test Equatorial slewing to coordinates asynchronous - Optional
            if (telescopeTests[SLEW_TO_COORDINATES_ASYNC])
            {
                TelescopeSlewTest(SlewSyncType.SlewToCoordinatesAsync, "SlewToCoordinatesAsync", canSlewAsync, "CanSlewAsync");
                if (cancellationToken.IsCancellationRequested)
                    return;
                if (canSlewAsync) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToCoordinatesAsync (Bad L)", SlewSyncType.SlewToCoordinatesAsync, BAD_RA_LOW, BAD_DEC_LOW);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeBadCoordinateTest("SlewToCoordinatesAsync (Bad H)", SlewSyncType.SlewToCoordinatesAsync, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
            {
                LogInfo(SLEW_TO_COORDINATES_ASYNC, "Tests skipped");
            }

            // Equatorial Sync to Coordinates - Optional - Moved here so that it can be tested before any target coordinates are set - Peter 4th August 2018
            if (telescopeTests[SYNC_TO_COORDINATES])
            {
                TelescopeSyncTest(SlewSyncType.SyncToCoordinates, "SyncToCoordinates", canSync, "CanSync");
                if (cancellationToken.IsCancellationRequested)
                    return;
                if (canSync) // Test syncing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SyncToCoordinates (Bad L)", SlewSyncType.SyncToCoordinates, BAD_RA_LOW, BAD_DEC_LOW);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeBadCoordinateTest("SyncToCoordinates (Bad H)", SlewSyncType.SyncToCoordinates, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
            {
                LogInfo(SYNC_TO_COORDINATES, "Tests skipped");
            }

            // TargetRightAscension Write (bad negative value) - Optional
            // Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
            // Test Equatorial target slewing - Optional
            if (telescopeTests[SLEW_TO_TARGET] | telescopeTests[SLEW_TO_TARGET_ASYNC])
            {
                try
                {
                    LogCallToDriver("TargetRightAscension Write", "About to set TargetRightAscension property to -1.0");
                    telescopeDevice.TargetRightAscension = -1.0d;
                    LogIssue("TargetRightAscension Write", "No error generated on set TargetRightAscension < 0 hours");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension < 0 hours");
                }

                // TargetRightAscension Write (bad positive value) - Optional
                try
                {
                    LogCallToDriver("TargetRightAscension Write", "About to set TargetRightAscension property to 25.0");
                    telescopeDevice.TargetRightAscension = 25.0d;
                    LogIssue("TargetRightAscension Write", "No error generated on set TargetRightAscension > 24 hours");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension > 24 hours");
                }

                // TargetRightAscension Write (valid value) - Optional
                try
                {
                    targetRightAscension = TelescopeRaFromSiderealTime("TargetRightAscension Write", -4.0d);
                    LogCallToDriver("TargetRightAscension Write",
                                           $"About to set TargetRightAscension property to {targetRightAscension}");
                    telescopeDevice.TargetRightAscension = targetRightAscension; // Set a valid value
                    try
                    {
                        LogCallToDriver("TargetRightAscension Write", "About to get TargetRightAscension property");
                        switch (Math.Abs(telescopeDevice.TargetRightAscension - targetRightAscension))
                        {
                            case 0.0d:
                                {
                                    LogOk("TargetRightAscension Write",
                                        $"Legal value {targetRightAscension.ToHMS()} HH:MM:SS written successfully");
                                    break;
                                }

                            case var @case when @case <= 1.0d / 3600.0d: // 1 seconds
                                {
                                    LogOk("TargetRightAscension Write",
                                        $"Target RightAscension is within 1 second of the value set: {targetRightAscension.ToHMS()}");
                                    break;
                                }

                            case var case1 when case1 <= 2.0d / 3600.0d: // 2 seconds
                                {
                                    LogOk("TargetRightAscension Write",
                                        $"Target RightAscension is within 2 seconds of the value set: {targetRightAscension.ToHMS()}");
                                    break;
                                }

                            case var case2 when case2 <= 5.0d / 3600.0d: // 5 seconds
                                {
                                    LogOk("TargetRightAscension Write",
                                        $"Target RightAscension is within 5 seconds of the value set: {targetRightAscension.ToHMS()}");
                                    break;
                                }

                            default:
                                {
                                    LogInfo("TargetRightAscension Write",
                                        $"Target RightAscension: {telescopeDevice.TargetRightAscension.ToHMS()}");
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("TargetRightAscension Write", MemberType.Property, Required.MustBeImplemented, ex, "Unable to read TargetRightAscension before writing to it.");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "");
                }
            }
            else
            {
                LogInfo("TargetRightAscension Write", "Tests skipped");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetDeclination Write (bad negative value) - Optional
            // Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
            if (telescopeTests[SLEW_TO_TARGET] | telescopeTests[SLEW_TO_TARGET_ASYNC])
            {
                try
                {
                    LogCallToDriver("TargetDeclination Write", "About to set TargetDeclination property to -91.0");
                    telescopeDevice.TargetDeclination = -91.0d;
                    LogIssue("TargetDeclination Write", "No error generated on set TargetDeclination < -90 degrees");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
                }

                // TargetDeclination Write (bad positive value) - Optional
                try
                {
                    LogCallToDriver("TargetDeclination Write", "About to set TargetDeclination property to 91.0");
                    telescopeDevice.TargetDeclination = 91.0d;
                    LogIssue("TargetDeclination Write", "No error generated on set TargetDeclination > 90 degrees");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
                }

                // TargetDeclination Write (valid value) - Optional
                try
                {
                    targetDeclination = 1.0d;
                    LogCallToDriver("TargetDeclination Write",
                                           $"About to set TargetDeclination property to {targetDeclination}");
                    telescopeDevice.TargetDeclination = targetDeclination; // Set a valid value
                    try
                    {
                        LogCallToDriver("TargetDeclination Write", "About to get TargetDeclination property");
                        switch (Math.Abs(telescopeDevice.TargetDeclination - targetDeclination))
                        {
                            case 0.0d:
                                {
                                    LogOk("TargetDeclination Write",
                                        $"Legal value {targetDeclination.ToDMS()} DD:MM:SS written successfully");
                                    break;
                                }

                            case var case3 when case3 <= 1.0d / 3600.0d: // 1 seconds
                                {
                                    LogOk("TargetDeclination Write",
                                        $"Target Declination is within 1 second of the value set: {targetDeclination.ToDMS()}");
                                    break;
                                }

                            case var case4 when case4 <= 2.0d / 3600.0d: // 2 seconds
                                {
                                    LogOk("TargetDeclination Write",
                                        $"Target Declination is within 2 seconds of the value set: {targetDeclination.ToDMS()}");
                                    break;
                                }

                            case var case5 when case5 <= 5.0d / 3600.0d: // 5 seconds
                                {
                                    LogOk("TargetDeclination Write",
                                        $"Target Declination is within 5 seconds of the value set: {targetDeclination.ToDMS()}");
                                    break;
                                }

                            default:
                                {
                                    LogInfo("TargetDeclination Write", $"Target Declination: {targetDeclination.ToDMS()}");
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("TargetDeclination Write", MemberType.Property, Required.MustBeImplemented, ex, "Unable to read TargetDeclination before writing to it.");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "");
                }
            }
            else
            {
                LogInfo("TargetDeclination Write", "Tests skipped");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test Equatorial target slewing - Optional
            if (telescopeTests[SLEW_TO_TARGET])
            {
                TelescopeSlewTest(SlewSyncType.SlewToTarget, "SlewToTarget", canSlew, "CanSlew");
                if (cancellationToken.IsCancellationRequested)
                    return;
                if (canSlew) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToTarget (Bad L)", SlewSyncType.SlewToTarget, BAD_RA_LOW, BAD_DEC_LOW);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeBadCoordinateTest("SlewToTarget (Bad H)", SlewSyncType.SlewToTarget, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
            {
                LogInfo(SLEW_TO_TARGET, "Tests skipped");
            }

            // Test Equatorial target slewing asynchronous - Optional
            if (telescopeTests[SLEW_TO_TARGET_ASYNC])
            {
                TelescopeSlewTest(SlewSyncType.SlewToTargetAsync, "SlewToTargetAsync", canSlewAsync, "CanSlewAsync");
                if (cancellationToken.IsCancellationRequested)
                    return;
                if (canSlewAsync) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToTargetAsync (Bad L)", SlewSyncType.SlewToTargetAsync, BAD_RA_LOW, BAD_DEC_LOW);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeBadCoordinateTest("SlewToTargetAsync (Bad H)", SlewSyncType.SlewToTargetAsync, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
            {
                LogInfo(SLEW_TO_TARGET_ASYNC, "Tests skipped");
            }

            // DestinationSideOfPier - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[DESTINATION_SIDE_OF_PIER])
                {
                    if (alignmentMode == AlignmentMode.GermanPolar)
                    {
                        TelescopeOptionalMethodsTest(OptionalMethodType.DestinationSideOfPier, "DestinationSideOfPier", true);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                    else
                    {
                        LogTestAndMessage("DestinationSideOfPier", "Test skipped as AligmentMode is not German Polar");
                    }
                }
                else
                {
                    LogInfo(DESTINATION_SIDE_OF_PIER, "Tests skipped");
                }
            }
            else
            {
                LogInfo("DestinationSideOfPier",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // Test AltAz Slewing - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[SLEW_TO_ALTAZ])
                {
                    TelescopeSlewTest(SlewSyncType.SlewToAltAz, "SlewToAltAz", canSlewAltAz, "CanSlewAltAz");
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    if (canSlewAltAz) // Test slewing to bad co-ordinates
                    {
                        TelescopeBadCoordinateTest("SlewToAltAz (Bad L)", SlewSyncType.SlewToAltAz, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW);
                        if (cancellationToken.IsCancellationRequested)
                            return; // -100 is used for the Altitude limit to enable -90 to be used for parking the scope
                        TelescopeBadCoordinateTest("SlewToAltAz (Bad H)", SlewSyncType.SlewToAltAz, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                else
                {
                    LogInfo(SLEW_TO_ALTAZ, "Tests skipped");
                }
            }
            else
            {
                LogInfo("SlewToAltAz",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // Test AltAz Slewing asynchronous - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[SLEW_TO_ALTAZ_ASYNC])
                {
                    TelescopeSlewTest(SlewSyncType.SlewToAltAzAsync, "SlewToAltAzAsync", canSlewAltAzAsync, "CanSlewAltAzAsync");
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    if (canSlewAltAzAsync) // Test slewing to bad co-ordinates
                    {
                        TelescopeBadCoordinateTest("SlewToAltAzAsync (Bad L)", SlewSyncType.SlewToAltAzAsync, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        TelescopeBadCoordinateTest("SlewToAltAzAsync (Bad H)", SlewSyncType.SlewToAltAzAsync, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                else
                {
                    LogInfo(SLEW_TO_ALTAZ_ASYNC, "Tests skipped");
                }
            }
            else
            {
                LogInfo("SlewToAltAzAsync",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // Equatorial Sync to Target - Optional
            if (telescopeTests[SYNC_TO_TARGET])
            {
                TelescopeSyncTest(SlewSyncType.SyncToTarget, "SyncToTarget", canSync, "CanSync");
                if (cancellationToken.IsCancellationRequested)
                    return;
                if (canSync) // Test syncing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SyncToTarget (Bad L)", SlewSyncType.SyncToTarget, BAD_RA_LOW, BAD_DEC_LOW);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    TelescopeBadCoordinateTest("SyncToTarget (Bad H)", SlewSyncType.SyncToTarget, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
            {
                LogInfo(SYNC_TO_TARGET, "Tests skipped");
            }

            // AltAz Sync - Optional
            if (GetInterfaceVersion() > 1)
            {
                if (telescopeTests[SYNC_TO_ALTAZ])
                {
                    TelescopeSyncTest(SlewSyncType.SyncToAltAz, "SyncToAltAz", canSyncAltAz, "CanSyncAltAz");
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    if (canSyncAltAz) // Test syncing to bad co-ordinates
                    {
                        TelescopeBadCoordinateTest("SyncToAltAz (Bad L)", SlewSyncType.SyncToAltAz, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        TelescopeBadCoordinateTest("SyncToAltAz (Bad H)", SlewSyncType.SyncToAltAz, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                else
                {
                    LogInfo(SYNC_TO_ALTAZ, "Tests skipped");
                }
            }
            else
            {
                LogInfo("SyncToAltAz",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            // SideOfPier
            if (settings.TestSideOfPierRead)
            {
                LogNewLine();
                LogTestOnly("SideOfPier Model Tests"); LogDebug("SideOfPier Model Tests", "Starting tests");
                if (GetInterfaceVersion() > 1)
                {
                    // 3.0.0.14 - Skip these tests if unable to read SideOfPier
                    if (CanReadSideOfPier("SideOfPier Model Tests"))
                    {
                        // Only run tests for German mounts
                        LogCallToDriver("SideOfPier Model Tests", "About to get AlignmentMode property");
                        if (telescopeDevice.AlignmentMode == AlignmentMode.GermanPolar)
                        {
                            LogDebug("SideOfPier Model Tests", $"Site latitude: {siteLatitude.ToDMS()}");
                            switch (siteLatitude)
                            {
                                case >= -SIDE_OF_PIER_INVALID_LATITUDE and <= SIDE_OF_PIER_INVALID_LATITUDE: // Refuse to handle this value because the Conform targeting logic or the mount's SideofPier flip logic may fail when the poles are this close to the horizon
                                    LogInfo("SideOfPier Model Tests", $"Tests skipped because the site latitude is reported as {siteLatitude.ToDMS()}");
                                    LogInfo("SideOfPier Model Tests", "This places the celestial poles close to the horizon and the mount's flip logic may override Conform's expected behaviour.");
                                    LogInfo("SideOfPier Model Tests", $"Please set the site latitude to a value within the ranges {SIDE_OF_PIER_INVALID_LATITUDE:+0.0;-0.0} to +90.0 or {(-SIDE_OF_PIER_INVALID_LATITUDE):+0.0;-0.0} to -90.0 to obtain a reliable result.");
                                    break;

                                case >= -90.0d and <= 90.0d: // Normal case, just run the tests because latitude is outside the invalid range but within -90.0 to +90.0
                                    // SideOfPier write property test - Optional
                                    if (settings.TestSideOfPierWrite)
                                    {
                                        LogDebug("SideOfPier Model Tests", "Testing SideOfPier write...");
                                        TelescopeOptionalMethodsTest(OptionalMethodType.SideOfPierWrite, "SideOfPier Write", canSetPierside);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    SideOfPierTests();
                                    break;

                                // Values outside the range -90.0 to +90.0 are invalid
                                default:
                                    LogInfo("SideOfPier Model Tests", "Test skipped because the site latitude Is outside the range -90.0 to +90.0");
                                    break;
                            }
                        }
                        else
                            LogInfo("SideOfPier Model Tests", "Test skipped because this Is Not a German equatorial mount");
                    }
                    else
                        LogInfo("SideOfPier Model Tests", "Tests skipped because this driver does Not support SideOfPier Read");
                }
                else
                    LogInfo("SideOfPier Model Tests", $"Skipping test as this method Is Not supported in interface V{GetInterfaceVersion()}");
            }
        }

        public override void CheckPerformance()
        {
            SetTest("Performance"); // Clear status messages
            TelescopePerformanceTest(PerformanceType.TstPerfAltitude, "Altitude");
            if (cancellationToken.IsCancellationRequested)
                return;
            if (GetInterfaceVersion() > 1)
            {
                TelescopePerformanceTest(PerformanceType.TstPerfAtHome, "AtHome");
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("Performance: AtHome",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            if (GetInterfaceVersion() > 1)
            {
                TelescopePerformanceTest(PerformanceType.TstPerfAtPark, "AtPark");
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("Performance: AtPark",
                    $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
            }

            TelescopePerformanceTest(PerformanceType.TstPerfAzimuth, "Azimuth");
            if (cancellationToken.IsCancellationRequested)
                return;
            TelescopePerformanceTest(PerformanceType.TstPerfDeclination, "Declination");
            if (cancellationToken.IsCancellationRequested)
                return;
            if (GetInterfaceVersion() > 1)
            {
                if (canPulseGuide)
                {
                    TelescopePerformanceTest(PerformanceType.TstPerfIsPulseGuiding, "IsPulseGuiding");
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
                else
                {
                    LogInfo("Performance: IsPulseGuiding", "Test omitted since IsPulseGuiding is not implemented");
                }
            }
            else
            {
                LogInfo("Performance: IsPulseGuiding",
                    $"Skipping test as this method is not supported in interface v1{GetInterfaceVersion()}");
            }

            TelescopePerformanceTest(PerformanceType.TstPerfRightAscension, "RightAscension");
            if (cancellationToken.IsCancellationRequested)
                return;
            if (GetInterfaceVersion() > 1)
            {
                if (alignmentMode == AlignmentMode.GermanPolar)
                {
                    if (CanReadSideOfPier("Performance - SideOfPier"))
                    {
                        TelescopePerformanceTest(PerformanceType.TstPerfSideOfPier, "SideOfPier");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                    else
                    {
                        LogInfo("Performance: SideOfPier", "Test omitted since SideOfPier is not implemented");
                    }
                }
                else
                {
                    LogInfo("Performance: SideOfPier", "Test omitted since alignment mode is not German Polar");
                }
            }
            else
            {
                LogInfo("Performance: SideOfPier",
                    $"Skipping test as this method is not supported in interface v1{GetInterfaceVersion()}");
            }

            if (canReadSiderealTime)
            {
                TelescopePerformanceTest(PerformanceType.TstPerfSiderealTime, "SiderealTime");
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("Performance: SiderealTime", "Skipping test because the SiderealTime property throws an exception.");
            }

            TelescopePerformanceTest(PerformanceType.TstPerfSlewing, "Slewing");
            if (cancellationToken.IsCancellationRequested)
                return;
            TelescopePerformanceTest(PerformanceType.TstPerfUtcDate, "UTCDate");
            if (cancellationToken.IsCancellationRequested)
                return;
        }

        public override void PostRunCheck()
        {
            // Make things safe
            // LogMsg("", MessageLevel.Always, "") 'Blank line
            try
            {
                if (telescopeDevice.CanSetTracking)
                {
                    telescopeDevice.Tracking = false;
                    LogOk("Mount Safety", "Tracking stopped to protect your mount.");
                }
                else
                {
                    LogInfo("Mount Safety", "Tracking can't be turned off for this mount, please switch off manually.");
                }
            }
            catch (Exception ex)
            {
                LogIssue("Mount Safety", $"Exception when disabling tracking to protect mount: {ex}");
            }
        }

        public override void CheckConfiguration()
        {
            try
            {
                // Common configuration
                if (!settings.TestProperties)
                    LogConfigurationAlert("Property tests were omitted due to Conform configuration.");

                if (!settings.TestMethods)
                    LogConfigurationAlert("Method tests were omitted due to Conform configuration.");
                else
                {
                    // Method tests
                    if (!telescopeTests[CAN_MOVE_AXIS])
                        LogConfigurationAlert("Can move axis tests were omitted due to Conform configuration.");

                    if (!telescopeTests[PARK_UNPARK])
                        LogConfigurationAlert("Park and Unpark tests were omitted due to Conform configuration.");

                    if (!telescopeTests[ABORT_SLEW])
                        LogConfigurationAlert("Abort slew tests were omitted due to Conform configuration.");

                    if (!telescopeTests[AXIS_RATE])
                        LogConfigurationAlert("Axis rate tests were omitted due to Conform configuration.");

                    if (!telescopeTests[FIND_HOME])
                        LogConfigurationAlert("Find home tests were omitted due to Conform configuration.");

                    if (!telescopeTests[MOVE_AXIS])
                        LogConfigurationAlert("Move axis tests were omitted due to Conform configuration.");

                    if (!telescopeTests[PULSE_GUIDE])
                        LogConfigurationAlert("Pulse guide tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SLEW_TO_COORDINATES])
                        LogConfigurationAlert("Synchronous Slew to coordinates tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SLEW_TO_COORDINATES_ASYNC])
                        LogConfigurationAlert("Asynchronous slew to coordinates tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SLEW_TO_TARGET])
                        LogConfigurationAlert("Synchronous slew to target tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SLEW_TO_TARGET_ASYNC])
                        LogConfigurationAlert("Asynchronous slew to target tests were omitted due to Conform configuration.");

                    if (!telescopeTests[DESTINATION_SIDE_OF_PIER])
                        LogConfigurationAlert("Destination side of pier tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SLEW_TO_ALTAZ])
                        LogConfigurationAlert("Synchronous slew to altitude / azimuth tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SLEW_TO_ALTAZ_ASYNC])
                        LogConfigurationAlert("Asynchronous slew to altitude / azimuth tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SYNC_TO_COORDINATES])
                        LogConfigurationAlert("Sync to coordinates tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SYNC_TO_TARGET])
                        LogConfigurationAlert("Sync to target tests were omitted due to Conform configuration.");

                    if (!telescopeTests[SYNC_TO_ALTAZ])
                        LogConfigurationAlert("Sync to altitude / azimuth tests were omitted due to Conform configuration.");
                }

                // Miscellaneous configuration
                if (!settings.TelescopeFirstUseTests)
                    LogConfigurationAlert("First time use tests were omitted due to Conform configuration.");

                if (!settings.TestSideOfPierRead)
                    LogConfigurationAlert("Extended side of pier read tests were omitted due to Conform configuration.");

                if (!settings.TestSideOfPierWrite)
                    LogConfigurationAlert("Extended side of pier write tests were omitted due to Conform configuration.");

                if (!settings.TelescopeExtendedRateOffsetTests)
                    LogConfigurationAlert("Extended rate offset tests were omitted due to Conform configuration.");

                if (!settings.TelescopeExtendedPulseGuideTests)
                    LogConfigurationAlert("Extended pulse guide tests were omitted due to Conform configuration.");

                if (!settings.TelescopeExtendedSiteTests)
                    LogConfigurationAlert("Extended Site property tests were omitted due to Conform configuration.");

            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n:{ex}");
            }
        }

        #endregion

        #region Support Code

        /// <summary>
        /// Confirm that rate offsets are reported as 0.0 and that they cannot be set when tracking at drive rates other than sidereal.
        /// </summary>
        /// <param name="driveRate"></param>
        /// <param name="testRateOffset"></param>
        private void CheckRateOffsetsRejectedForNonSiderealDriveRates(DriveRate driveRate, RateOffset testRateOffset)
        {
            // Check whether or not we are tracking at sidereal rate
            if (driveRate != DriveRate.Sidereal) // We are tracking at a non-sidereal rate
            {
                // Get the test axis rate offset
                LogCallToDriver("", $"About to get {testRateOffset}");
                double currentRateOffset = (testRateOffset == RateOffset.DeclinationRate) ? telescopeDevice.DeclinationRate : telescopeDevice.RightAscensionRate;

                // Check whether the reported rate offset is 0.0
                if (currentRateOffset == 0.0) // Rate offset is 0.0
                {
                    LogOk("TrackingRate Write", $"{testRateOffset} is zero for {driveRate} drive rate.");
                }
                else // Rate offset is not 0.0
                {
                    LogIssue("TrackingRate Write", $"{testRateOffset} is not zero for drive rate: {driveRate}, it is: {currentRateOffset}");
                    LogInfo("TrackingRate Write", $"In ITelescopeV4 and later, rate offsets are only valid when tracking at Sidereal rate.");
                    LogInfo("TrackingRate Write", $"When the {driveRate} tracking rate is active, reading {testRateOffset} should return 0.0.");
                }

                // If the rate offset can be changed when tracking at sidereal rate, make sure that it cannot be set in this non-sidereal tracking rate
                if (testRateOffset == RateOffset.DeclinationRate ? canSetDeclinationRate : canSetRightAscensionRate) // RightAscensionRate can be set when tracking at sidereal rate
                {
                    try
                    {
                        // Try to set the appropriate offset rate
                        LogCallToDriver("", $"About to set {testRateOffset}");
                        if (testRateOffset == RateOffset.DeclinationRate) // Testing DeclinationRate
                            telescopeDevice.DeclinationRate = settings.TelescopeRateOffsetTestLowValue;
                        else // Testing RightAscensionRate
                            telescopeDevice.RightAscensionRate = settings.TelescopeRateOffsetTestLowValue / 15.0;

                        // Should never get here because an exception is expected
                        LogIssue("TrackingRate Write", $"Setting {testRateOffset} in {driveRate} drive rate did not result in an {(settings.DeviceTechnology == DeviceTechnology.COM ? "InvalidOperationException" : "InvalidOperation  error")} but should have.");
                        LogInfo("TrackingRate Write", $"In ITelescopeV4 and later, rate offsets are only valid when tracking at Sidereal rate.");
                        LogInfo("TrackingRate Write", $"When the {driveRate} tracking rate is active, writing to {testRateOffset} should result in an {(settings.DeviceTechnology == DeviceTechnology.COM ? "InvalidOperationException" : "InvalidOperation  error")}.");
                    }
                    catch (COMException ex)
                    {
                        // Check if this is an InvalidOperationExcception
                        if (ex.HResult == unchecked((int)0x8004040B)) // This is a COM InvalidOperationExcception
                        {
                            LogOk("TrackingRate Write", $"Setting {testRateOffset} in {driveRate} drive rate resulted in a COM InvalidOperationException.");
                        }
                        else // This is not a COM InvalidOperationExcception
                        {
                            LogIssue("TrackingRate Write", $"Setting {testRateOffset} in {driveRate} drive rate did not result in an InvalidOperationException. Instead it resulted in an {ex.GetType().Name} exception: {ex.Message}");
                            LogDebug("TrackingRate Write", ex.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex is ASCOM.InvalidOperationException) // This is an InvalidOperationExcception
                        {
                            LogOk("TrackingRate Write", $"Setting {testRateOffset} in {driveRate} drive rate resulted in an InvalidOperationException.");
                        }
                        else // This is not an InvalidOperationExcception
                        {
                            LogIssue("TrackingRate Write", $"Setting {testRateOffset} in {driveRate} drive rate did not result in an InvalidOperationException. Instead it resulted in an {ex.GetType().Name} exception: {ex.Message}");
                            LogDebug("TrackingRate Write", ex.ToString());
                        }
                    }
                }
            }
            else // Tracking at sidereal rate
            {
                // Nothing to do here because this test is only for non-sidereal drive rates
            }
        }

        private void TelescopeSyncTest(SlewSyncType testType, string testName, bool driverSupportsMethod, string canDoItName)
        {
            bool showOutcome = false;
            double difference, syncRa, syncDec, syncAlt = default, syncAz = default, newAlt, newAz, currentAz = default, currentAlt = default, startRa, startDec, currentRa, currentDec;

            SetTest(testName);
            SetAction("Running test...");

            // Basic test to make sure the method is either implemented OK or fails as expected if it is not supported in this driver.
            LogCallToDriver(testName, "About to get RightAscension property");
            syncRa = telescopeDevice.RightAscension;
            LogCallToDriver(testName, "About to get Declination property");
            syncDec = telescopeDevice.Declination;
            if (!driverSupportsMethod) // Call should fail
            {
                try
                {
                    switch (testType)
                    {
                        case SlewSyncType.SyncToCoordinates: // SyncToCoordinates
                            LogCallToDriver(testName, "About to get Tracking property");
                            if (canSetTracking & !telescopeDevice.Tracking)
                            {
                                LogCallToDriver(testName, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            LogDebug(testName, $"SyncToCoordinates: {syncRa.ToHMS()} {syncDec.ToDMS()}");
                            LogCallToDriver(testName,
                                                           $"About to call SyncToCoordinates method, RA: {syncRa.ToHMS()}, Declination: {syncDec.ToDMS()}");
                            telescopeDevice.SyncToCoordinates(syncRa, syncDec);
                            LogIssue(testName, "CanSyncToCoordinates is False but call to SyncToCoordinates did not throw an exception.");
                            break;

                        case SlewSyncType.SyncToTarget: // SyncToTarget
                            LogCallToDriver(testName, "About to get Tracking property");
                            if (canSetTracking & !telescopeDevice.Tracking)
                            {
                                LogCallToDriver(testName, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            try
                            {
                                LogDebug(testName, $"Setting TargetRightAscension: {syncRa.ToHMS()}");
                                LogCallToDriver(testName,
                                                                   $"About to set TargetRightAscension property to {syncRa.ToHMS()}");
                                telescopeDevice.TargetRightAscension = syncRa;
                                LogDebug(testName, "Completed Set TargetRightAscension");
                            }
                            catch (Exception)
                            {
                                // Ignore errors at this point as we aren't trying to test Telescope.TargetRightAscension
                            }

                            try
                            {
                                LogDebug(testName, $"Setting TargetDeclination: {syncDec.ToDMS()}");
                                LogCallToDriver(testName,
                                                                   $"About to set TargetDeclination property to {syncDec.ToDMS()}");
                                telescopeDevice.TargetDeclination = syncDec;
                                LogDebug(testName, "Completed Set TargetDeclination");
                            }
                            catch (Exception)
                            {
                                // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                            }
                            LogCallToDriver(testName, "About to call SyncToTarget method");
                            telescopeDevice.SyncToTarget(); // Sync to target coordinates
                            LogIssue(testName, "CanSyncToTarget is False but call to SyncToTarget did not throw an exception.");
                            break;

                        case SlewSyncType.SyncToAltAz:
                            if (canReadAltitide)
                            {
                                LogCallToDriver(testName, "About to get Altitude property");
                                syncAlt = telescopeDevice.Altitude;
                            }

                            if (canReadAzimuth)
                            {
                                LogCallToDriver(testName, "About to get Azimuth property");
                                syncAz = telescopeDevice.Azimuth;
                            }
                            LogCallToDriver(testName, "About to get Tracking property");
                            if (canSetTracking & telescopeDevice.Tracking)
                            {
                                LogCallToDriver(testName, "About to set Tracking property to false");
                                telescopeDevice.Tracking = false;
                            }
                            LogCallToDriver(testName,
                                $"About to call SyncToAltAz method, Altitude: {syncAlt.ToDMS()}, Azimuth: {syncAz.ToDMS()}");
                            telescopeDevice.SyncToAltAz(syncAz, syncAlt); // Sync to new Alt Az
                            LogIssue(testName, "CanSyncToAltAz is False but call to SyncToAltAz did not throw an exception.");
                            break;

                        default:
                            LogIssue(testName, $"Conform:SyncTest: Unknown test type {testType}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex,
                        $"{canDoItName} is False");
                }
            }
            else // Call should work
            {
                try
                {
                    switch (testType)
                    {
                        case SlewSyncType.SyncToCoordinates:
                        case SlewSyncType.SyncToTarget: // Only do this for equatorial syncs
                                                        // Calculate the Sync test RA position
                            startRa = TelescopeRaFromHourAngle(testName, +3.0d);
                            LogDebug(testName, string.Format("RA for sync tests: {0}", startRa.ToHMS()));

                            // Calculate the Sync test DEC position
                            if (siteLatitude > 0.0d) // We are in the northern hemisphere
                            {
                                startDec = 90.0d - (180.0d - siteLatitude) * 0.5d; // Calculate for northern hemisphere
                            }
                            else // We are in the southern hemisphere
                            {
                                startDec = -90.0d + (180.0d + siteLatitude) * 0.5d;
                            } // Calculate for southern hemisphere

                            LogDebug(testName, string.Format("Declination for sync tests: {0}", startDec.ToDMS()));
                            SlewScope(startRa, startDec, "start position");
                            if (cancellationToken.IsCancellationRequested)
                                return;
                            SetAction("Checking that scope slewed OK");
                            // Now test that we have actually arrived
                            CheckScopePosition(testName, "Slewed to start position", startRa, startDec);

                            // Calculate the sync test RA coordinate as a variation from the current RA coordinate
                            syncRa = startRa - SYNC_SIMULATED_ERROR / (15.0d * 60.0d); // Convert sync error in arc minutes to RA hours
                            if (syncRa < 0.0d)
                                syncRa += 24.0d; // Ensure legal RA

                            // Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                            syncDec = startDec - SYNC_SIMULATED_ERROR / 60.0d; // Convert sync error in arc minutes to degrees

                            SetAction("Syncing the scope");
                            // Sync the scope to the offset RA and DEC coordinates
                            SyncScope(testName, canDoItName, testType, syncRa, syncDec);

                            // Check that the scope's synchronised position is as expected
                            CheckScopePosition(testName, "Synced to sync position", syncRa, syncDec);

                            // Check that the TargetRA and TargetDec were 
                            SetAction("Checking that the scope synced OK");
                            if (testType == SlewSyncType.SyncToCoordinates)
                            {
                                // Check that target coordinates are present and set correctly per the ASCOM Telescope specification
                                try
                                {
                                    currentRa = telescopeDevice.TargetRightAscension;
                                    LogDebug(testName, string.Format("Current TargetRightAscension: {0}, Set TargetRightAscension: {1}", currentRa, syncRa));
                                    double raDifference;
                                    raDifference = RaDifferenceInArcSeconds(syncRa, currentRa);
                                    switch (raDifference)
                                    {
                                        case var @case when @case <= SLEW_SYNC_OK_TOLERANCE:  // Within specified tolerance
                                            {
                                                LogOk(testName, string.Format("The TargetRightAscension property {0} matches the expected RA OK. ", syncRa.ToHMS())); // Outside specified tolerance
                                                break;
                                            }

                                        default:
                                            {
                                                LogIssue(testName, string.Format("The TargetRightAscension property {0} does not match the expected RA {1}", currentRa.ToHMS(), syncRa.ToHMS()));
                                                break;
                                            }
                                    }
                                }
                                catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == ExNotSet1 | ex.ErrorCode == ExNotSet2)
                                {
                                    LogIssue(testName, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                                }
                                catch (ASCOM.InvalidOperationException)
                                {
                                    LogIssue(testName, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                }
                                catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == ExNotSet1 | ex.Number == ExNotSet2)
                                {
                                    LogIssue(testName, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
                                }

                                try
                                {
                                    currentDec = telescopeDevice.TargetDeclination;
                                    LogDebug(testName, string.Format("Current TargetDeclination: {0}, Set TargetDeclination: {1}", currentDec, syncDec));
                                    double decDifference;
                                    decDifference = Math.Round(Math.Abs(currentDec - syncDec) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // Dec difference is in arc seconds from degrees of Declination
                                    switch (decDifference)
                                    {
                                        case var case1 when case1 <= SLEW_SYNC_OK_TOLERANCE: // Within specified tolerance
                                            LogOk(testName, string.Format("The TargetDeclination property {0} matches the expected Declination OK. ", syncDec.ToDMS())); // Outside specified tolerance
                                            break;

                                        default:
                                            LogIssue(testName, string.Format("The TargetDeclination property {0} does not match the expected Declination {1}", currentDec.ToDMS(), syncDec.ToDMS()));
                                            break;
                                    }
                                }
                                catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == ExNotSet1 | ex.ErrorCode == ExNotSet2)
                                {
                                    LogIssue(testName, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                                }
                                catch (ASCOM.InvalidOperationException)
                                {
                                    LogIssue(testName, "The driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                }
                                catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == ExNotSet1 | ex.Number == ExNotSet2)
                                {
                                    LogIssue(testName, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
                                }
                            }

                            // Now slew to the scope's original position
                            SlewScope(startRa, startDec, "original position in post-sync coordinates");

                            // Check that the scope's position is the original position
                            CheckScopePosition(testName, "Slewed back to start position", startRa, startDec);

                            // Now "undo" the sync by reversing syncing in the opposition sense than originally made

                            // Calculate the sync test RA coordinate as a variation from the current RA coordinate
                            syncRa = startRa + SYNC_SIMULATED_ERROR / (15.0d * 60.0d); // Convert sync error in arc minutes to RA hours
                            if (syncRa >= 24.0d)
                                syncRa -= 24.0d; // Ensure legal RA

                            // Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                            syncDec = startDec + SYNC_SIMULATED_ERROR / 60.0d; // Convert sync error in arc minutes to degrees

                            // Sync back to the original coordinates
                            SetAction("Restoring original sync values");
                            SyncScope(testName, canDoItName, testType, syncRa, syncDec);

                            // Check that the scope's synchronised position is as expected
                            CheckScopePosition(testName, "Synced to reversed sync position", syncRa, syncDec);

                            // Now slew to the scope's original position
                            SlewScope(startRa, startDec, "original position in pre-sync coordinates");

                            // Check that the scope's position is the original position
                            CheckScopePosition(testName, "Slewed back to start position", startRa, startDec);
                            break;

                        case SlewSyncType.SyncToAltAz:
                            if (canReadAltitide)
                            {
                                LogCallToDriver(testName, "About to get Altitude property");
                                currentAlt = telescopeDevice.Altitude;
                            }

                            if (canReadAzimuth)
                            {
                                LogCallToDriver(testName, "About to get Azimuth property");
                                currentAz = telescopeDevice.Azimuth;
                            }

                            syncAlt = currentAlt - 1.0d;
                            syncAz = currentAz + 1.0d;
                            if (syncAlt < 0.0d)
                                syncAlt = 1.0d; // Ensure legal Alt
                            if (syncAz > 359.0d)
                                syncAz = 358.0d; // Ensure legal Az

                            LogCallToDriver(testName, "About to get Tracking property");
                            if (canSetTracking & telescopeDevice.Tracking)
                            {
                                LogCallToDriver(testName, "About to set Tracking property to false");
                                telescopeDevice.Tracking = false;
                            }
                            LogCallToDriver(testName, $"About to call SyncToAltAz method, Altitude: {syncAlt.ToDMS()}, Azimuth: {syncAz.ToDMS()}");
                            telescopeDevice.SyncToAltAz(syncAz, syncAlt); // Sync to new Alt Az

                            if (canReadAltitide & canReadAzimuth) // Can check effects of a sync
                            {
                                LogCallToDriver(testName, "About to get Altitude property");
                                newAlt = telescopeDevice.Altitude;
                                LogCallToDriver(testName, "About to get Azimuth property");
                                newAz = telescopeDevice.Azimuth;

                                // Compare old and new values
                                difference = Math.Abs(syncAlt - newAlt);
                                switch (difference)
                                {
                                    case var case2 when case2 <= 1.0d / (60 * 60): // Within 1 seconds
                                        LogOk(testName, "Synced Altitude OK");
                                        break;

                                    case var case3 when case3 <= 2.0d / (60 * 60): // Within 2 seconds
                                        LogOk(testName, "Synced within 2 seconds of Altitude");
                                        showOutcome = true;
                                        break;

                                    default:
                                        LogInfo(testName, $"Synced to within {difference.ToDMS()} DD:MM:SS of expected Altitude: {syncAlt.ToDMS()}");
                                        showOutcome = true;
                                        break;
                                }

                                difference = Math.Abs(syncAz - newAz);
                                switch (difference)
                                {
                                    case var case4 when case4 <= 1.0d / (60 * 60): // Within 1 seconds
                                        LogOk(testName, "Synced Azimuth OK");
                                        break;

                                    case var case5 when case5 <= 2.0d / (60 * 60): // Within 2 seconds
                                        LogOk(testName, "Synced within 2 seconds of Azimuth");
                                        showOutcome = true;
                                        break;

                                    default:
                                        LogInfo(testName,
                                            $"Synced to within {FormatAzimuth(difference)} DD:MM:SS of expected Azimuth: {FormatAzimuth(syncAz)}");
                                        showOutcome = true;
                                        break;
                                }

                                if (showOutcome)
                                {
                                    LogTestAndMessage(testName, "             Altitude       Azimuth");
                                    LogTestAndMessage(testName, $"Original:  {currentAlt.ToDMS()}   {FormatAzimuth(currentAz)}");
                                    LogTestAndMessage(testName, $"Sync to:   {syncAlt.ToDMS()}   {FormatAzimuth(syncAz)}");
                                    LogTestAndMessage(testName, $"New:       {newAlt.ToDMS()}   {FormatAzimuth(newAz)}");
                                }
                            }
                            else // Can't test effects of a sync
                            {
                                LogInfo(testName, "Can't test SyncToAltAz because Altitude or Azimuth values are not implemented");
                            } // Do nothing

                            break;

                        default:
                            break;
                    }
                }
                catch (Exception ex)
                {
                    HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, $"{canDoItName} is True");
                }
            }
        }

        private void TelescopeSlewTest(SlewSyncType testType, string testName, bool canPropertyIsTrue, string canPropertyName)
        {
            SetTest(testName);
            if (canSetTracking)
            {
                LogCallToDriver(testName, "About to set Tracking property to true");
                telescopeDevice.Tracking = true; // Enable tracking for these tests
            }

            try
            {
                switch (testType)
                {
                    case SlewSyncType.SlewToCoordinates:
                        LogCallToDriver(testName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(testName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        targetRightAscension = TelescopeRaFromSiderealTime(testName, -1.0d);
                        targetDeclination = 1.0d;
                        SetAction("Slewing synchronously...");

                        LogCallToDriver(testName, $"About to call SlewToCoordinates method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                        TimeMethod($"SlewToCoordinates", () => telescopeDevice.SlewToCoordinates(targetRightAscension, targetDeclination), TargetTime.Extended);
                        LogDebug(testName, "Returned from SlewToCoordinates method");
                        break;

                    case SlewSyncType.SlewToCoordinatesAsync:
                        LogCallToDriver(testName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(testName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        targetRightAscension = TelescopeRaFromSiderealTime(testName, -2.0d);
                        targetDeclination = 2.0d;

                        LogCallToDriver(testName, $"About to call SlewToCoordinatesAsync method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                        TimeMethod($"SlewToCoordinatesAsync", () => telescopeDevice.SlewToCoordinatesAsync(targetRightAscension, targetDeclination), TargetTime.Standard);

                        if (telescopeDevice.Slewing)
                            LogDebug(testName, "Slewing after SlewToCoordinatesAsync is True as expected, waiting for slew to complete.");
                        else
                            LogIssue(testName, "Slewing is FALSE after SlewToCoordinatesAsync returned, expected TRUE, continuing test.");

                        WaitForSlew(testName, $"Slewing to coordinates asynchronously");
                        break;

                    case SlewSyncType.SlewToTarget:
                        LogCallToDriver(testName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(testName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        targetRightAscension = TelescopeRaFromSiderealTime(testName, -3.0d);
                        targetDeclination = 3.0d;
                        try
                        {
                            LogCallToDriver(testName, $"About to set TargetRightAscension property to {targetRightAscension.ToHMS()}");
                            telescopeDevice.TargetRightAscension = targetRightAscension;
                        }
                        catch (Exception ex)
                        {
                            if (canPropertyIsTrue)
                            {
                                HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"{canPropertyName} is True but can't set TargetRightAscension");
                            }
                            else
                            {
                                // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                            }
                        }

                        try
                        {
                            LogCallToDriver(testName, $"About to set TargetDeclination property to {targetDeclination.ToDMS()}");
                            telescopeDevice.TargetDeclination = targetDeclination;
                        }
                        catch (Exception ex)
                        {
                            if (canPropertyIsTrue)
                            {
                                HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"{canPropertyName} is True but can't set TargetDeclination");
                            }
                            else
                            {
                                // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                            }
                        }

                        SetAction("Slewing synchronously...");
                        LogCallToDriver(testName, "About to call SlewToTarget method");
                        TimeMethod($"SlewToTarget", () => telescopeDevice.SlewToTarget(), TargetTime.Extended);
                        break;

                    case SlewSyncType.SlewToTargetAsync: // SlewToTargetAsync
                        LogCallToDriver(testName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(testName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        targetRightAscension = TelescopeRaFromSiderealTime(testName, -4.0d);
                        targetDeclination = 4.0d;
                        try
                        {
                            LogCallToDriver(testName, $"About to set TargetRightAscension property to {targetRightAscension.ToHMS()}");
                            telescopeDevice.TargetRightAscension = targetRightAscension;
                        }
                        catch (Exception ex)
                        {
                            if (canPropertyIsTrue)
                            {
                                HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"{canPropertyName} is True but can't set TargetRightAscension");
                            }
                            else
                            {
                                // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                            }
                        }

                        try
                        {
                            LogCallToDriver(testName, $"About to set TargetDeclination property to {targetDeclination.ToDMS()}");
                            telescopeDevice.TargetDeclination = targetDeclination;
                        }
                        catch (Exception ex)
                        {
                            if (canPropertyIsTrue)
                            {
                                HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"{canPropertyName} is True but can't set TargetDeclination");
                            }
                            else
                            {
                                // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                            }
                        }

                        LogCallToDriver(testName, "About to call SlewToTargetAsync method");
                        TimeMethod($"SlewToTargetAsync", () => telescopeDevice.SlewToTargetAsync(), TargetTime.Standard);

                        if (telescopeDevice.Slewing)
                            LogDebug(testName, "Slewing after SlewToTargetAsync is True as expected, waiting for slew to complete.");
                        else
                            LogIssue(testName, "Slewing is FALSE after SlewToTargetAsync returned, expected TRUE, continuing test.");

                        WaitForSlew(testName, $"Slewing to target asynchronously");
                        break;

                    case SlewSyncType.SlewToAltAz:
                        LogDebug(testName, $"Tracking 1: {telescopeDevice.Tracking}");
                        LogCallToDriver(testName, "About to get Tracking property");
                        if (canSetTracking & telescopeDevice.Tracking)
                        {
                            LogCallToDriver(testName, "About to set property Tracking to false");
                            telescopeDevice.Tracking = false;
                            LogDebug(testName, "Tracking turned off");
                        }

                        LogCallToDriver(testName, "About to get Tracking property");
                        LogDebug(testName, $"Tracking 2: {telescopeDevice.Tracking}");
                        targetAltitude = 50.0d;
                        targetAzimuth = 150.0d;
                        SetAction("Slewing to Alt/Az synchronously...");

                        LogCallToDriver(testName, $"About to call SlewToAltAz method, Altitude: {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                        TimeMethod($"SlewToAltAz", () => telescopeDevice.SlewToAltAz(targetAzimuth, targetAltitude), TargetTime.Extended);

                        LogCallToDriver(testName, "About to get Tracking property");
                        LogDebug(testName, $"Tracking 3: {telescopeDevice.Tracking}");
                        break;

                    case SlewSyncType.SlewToAltAzAsync:
                        LogCallToDriver(testName, "About to get Tracking property");
                        LogDebug(testName, $"Tracking 1: {telescopeDevice.Tracking}");
                        LogCallToDriver(testName, "About to get Tracking property");
                        if (canSetTracking & telescopeDevice.Tracking)
                        {
                            LogCallToDriver(testName, "About to set Tracking property false");
                            telescopeDevice.Tracking = false;
                            LogDebug(testName, "Tracking turned off");
                        }

                        LogCallToDriver(testName, "About to get Tracking property");
                        LogDebug(testName, $"Tracking 2: {telescopeDevice.Tracking}");
                        targetAltitude = 55.0d;
                        targetAzimuth = 155.0d;

                        LogCallToDriver(testName, $"About to call SlewToAltAzAsync method, Altitude: {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                        TimeMethod($"SlewToAltAzAsync", () => telescopeDevice.SlewToAltAzAsync(targetAzimuth, targetAltitude), TargetTime.Standard);

                        if (telescopeDevice.Slewing)
                            LogDebug(testName, "Slewing after SlewToAltAzAsync is True as expected, waiting for slew to complete.");
                        else
                            LogIssue(testName, "Slewing is FALSE after SlewToAltAzAsync returned, expected TRUE, continuing test.");

                        LogCallToDriver(testName, "About to get Tracking property");
                        LogDebug(testName, $"Tracking 3: {telescopeDevice.Tracking}");

                        WaitForSlew(testName, $"Slewing to Alt/Az asynchronously");
                        LogCallToDriver(testName, "About to get Tracking property");
                        LogDebug(testName, $"Tracking 4: {telescopeDevice.Tracking}");
                        break;

                    default:
                        LogError(testName, $"Conform:SlewTest: Unknown test type {testType}");
                        break;
                }

                if (cancellationToken.IsCancellationRequested) return;

                if (canPropertyIsTrue) // Should be able to do this so report what happened
                {
                    SetAction("Slew completed");
                    switch (testType)
                    {
                        case SlewSyncType.SlewToCoordinates:
                        case SlewSyncType.SlewToCoordinatesAsync:
                        case SlewSyncType.SlewToTarget:
                        case SlewSyncType.SlewToTargetAsync:
                            // Test how close the slew was to the required coordinates
                            CheckScopePosition(testName, "Slewed", targetRightAscension, targetDeclination);

                            // Check that the slews and syncs set the target RA coordinate correctly per the ASCOM Telescope specification
                            try
                            {
                                double actualTargetRa = telescopeDevice.TargetRightAscension;
                                LogDebug(testName, $"Current TargetRightAscension: {actualTargetRa}, Set TargetRightAscension: {targetRightAscension}");

                                if (RaDifferenceInArcSeconds(actualTargetRa, targetRightAscension) <= settings.TelescopeSlewTolerance) // Within specified tolerance
                                {
                                    LogOk(testName, $"The TargetRightAscension property: {actualTargetRa.ToHMS()} matches the expected RA {targetRightAscension.ToHMS()} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds."); // Outside specified tolerance
                                }
                                else
                                {
                                    LogIssue(testName, $"The TargetRightAscension property: {actualTargetRa.ToHMS()} does not match the expected RA {targetRightAscension.ToHMS()} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                                }
                            }
                            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == ExNotSet1 | ex.ErrorCode == ExNotSet2)
                            {
                                LogIssue(testName, "The Driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                            }
                            catch (ASCOM.InvalidOperationException)
                            {
                                LogIssue(testName, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                            }
                            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == ExNotSet1 | ex.Number == ExNotSet2)
                            {
                                LogIssue(testName, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                            }
                            catch (Exception ex)
                            {
                                HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
                            }

                            // Check that the slews and syncs set the target declination coordinate correctly per the ASCOM Telescope specification
                            try
                            {
                                double actualTargetDec = telescopeDevice.TargetDeclination;
                                LogDebug(testName, $"Current TargetDeclination: {actualTargetDec}, Set TargetDeclination: {targetDeclination}");

                                if (Math.Round(Math.Abs(actualTargetDec - targetDeclination) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero) <= settings.TelescopeSlewTolerance) // Within specified tolerance
                                {
                                    LogOk(testName, $"The TargetDeclination property {actualTargetDec.ToDMS()} matches the expected Declination {targetDeclination.ToDMS()} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds."); // Outside specified tolerance
                                }
                                else
                                {
                                    LogIssue(testName, $"The TargetDeclination property {actualTargetDec.ToDMS()} does not match the expected Declination {targetDeclination.ToDMS()} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                                }

                            }
                            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == ExNotSet1 | ex.ErrorCode == ExNotSet2)
                            {
                                LogIssue(testName, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                            }
                            catch (ASCOM.InvalidOperationException)
                            {
                                LogIssue(testName, "The Driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                            }
                            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == ExNotSet1 | ex.Number == ExNotSet2)
                            {
                                LogIssue(testName, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                            }
                            catch (Exception ex)
                            {
                                HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
                            }

                            break;

                        case SlewSyncType.SlewToAltAz:
                        case SlewSyncType.SlewToAltAzAsync:
                            // Test how close the slew was to the required coordinates
                            LogCallToDriver(testName, "About to get Azimuth property");
                            double actualAzimuth = telescopeDevice.Azimuth;

                            LogCallToDriver(testName, "About to get Altitude property");
                            double actualAltitude = telescopeDevice.Altitude;

                            double azimuthDifference = Math.Abs(actualAzimuth - targetAzimuth);
                            if (azimuthDifference > 350.0d) azimuthDifference = 360.0d - azimuthDifference; // Deal with the case where the two elements are on different sides of 360 degrees

                            if (azimuthDifference <= settings.TelescopeSlewTolerance)
                            {
                                LogOk(testName, $"Slewed to target Azimuth OK within tolerance: {settings.TelescopeSlewTolerance} arc seconds. Actual Azimuth: {FormatAzimuth(actualAzimuth)}, Target Azimuth: {FormatAzimuth(targetAzimuth)}");
                            }
                            else
                            {
                                LogIssue(testName, $"Slewed {azimuthDifference:0.0} arc seconds away from Azimuth target: {FormatAzimuth(targetAzimuth)} Actual Azimuth: {FormatAzimuth(actualAzimuth)}. Tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                            }

                            double altitudeDifference = Math.Abs(actualAltitude - targetAltitude);
                            if (altitudeDifference <= settings.DomeSlewTolerance)
                            {
                                LogOk(testName, $"Slewed to target Altitude OK within tolerance: {settings.TelescopeSlewTolerance} arc seconds. Actual Altitude: {actualAltitude.ToDMS()}, Target Altitude: {targetAltitude.ToDMS()}");
                            }
                            else
                            {
                                LogIssue(testName, $"Slewed {altitudeDifference:0.0} degree(s) away from Altitude target: {targetAltitude.ToDMS()} Actual Altitude: {actualAltitude.ToDMS()} Tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                            }

                            break;

                        default: // Do nothing
                            break;
                    }
                }
                else // Not supposed to be able to do this but no error generated so report an error
                {
                    LogIssue(testName, $"{canPropertyName} is false but no exception was generated on use");
                }
            }
            catch (Exception ex)
            {
                if (canPropertyIsTrue)
                {
                    HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, $"{canPropertyName} is True");
                }
                else
                {
                    HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex,
                        $"{canPropertyName} is False");
                }
            }

        }

        /// <summary>
        /// Confirm that InValidValueExceptions are thrown for invalid values
        /// </summary>
        /// <param name="pName"></param>
        /// <param name="pTest">The method to test</param>
        /// <param name="badCoordinate1">RA or Altitude</param>
        /// <param name="badCoordinate2">Dec or Azimuth</param>
        /// <remarks></remarks>
        private void TelescopeBadCoordinateTest(string pName, SlewSyncType pTest, double badCoordinate1, double badCoordinate2)
        {
            switch (pTest)
            {
                case SlewSyncType.SlewToCoordinates:
                case SlewSyncType.SlewToCoordinatesAsync:
                    {
                        LogCallToDriver(pName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(pName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = badCoordinate1;
                            targetDeclination = 0.0d;

                            if (pTest == SlewSyncType.SlewToCoordinates)
                            {
                                LogCallToDriver(pName,
                                    $"About to call SlewToCoordinates method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                                telescopeDevice.SlewToCoordinates(targetRightAscension, targetDeclination);
                            }
                            else
                            {
                                LogCallToDriver(pName,
                                    $"About to call SlewToCoordinatesAsync method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                                telescopeDevice.SlewToCoordinatesAsync(targetRightAscension, targetDeclination);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                AbortSlew(pName);
                            }
                            catch { } // Attempt to stop any motion that has actually started

                            LogIssue(pName, $"Failed to reject bad RA coordinate: {targetRightAscension.ToHMS()}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate",
                                $"Correctly rejected bad RA coordinate: {targetRightAscension.ToHMS()}");
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = TelescopeRaFromSiderealTime(pName, -2.0d);
                            targetDeclination = badCoordinate2;
                            if (pTest == SlewSyncType.SlewToCoordinates)
                            {
                                LogCallToDriver(pName,
                                    $"About to call SlewToCoordinates method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                                telescopeDevice.SlewToCoordinates(targetRightAscension, targetDeclination);
                            }
                            else
                            {
                                LogCallToDriver(pName,
                                    $"About to call SlewToCoordinatesAsync method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                                telescopeDevice.SlewToCoordinatesAsync(targetRightAscension, targetDeclination);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                AbortSlew(pName);
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogIssue(pName, $"Failed to reject bad Dec coordinate: {targetDeclination.ToDMS()}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate",
                                $"Correctly rejected bad Dec coordinate: {targetDeclination.ToDMS()}");
                        }

                        break;
                    }

                case SlewSyncType.SyncToCoordinates:
                    {
                        LogCallToDriver(pName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(pName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = badCoordinate1;
                            targetDeclination = 0.0d;
                            LogCallToDriver(pName,
                                                           $"About to call SyncToCoordinates method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                            telescopeDevice.SyncToCoordinates(targetRightAscension, targetDeclination);
                            LogIssue(pName, $"Failed to reject bad RA coordinate: {targetRightAscension.ToHMS()}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate",
                                $"Correctly rejected bad RA coordinate: {targetRightAscension.ToHMS()}");
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = TelescopeRaFromSiderealTime(pName, -3.0d);
                            targetDeclination = badCoordinate2;
                            LogCallToDriver(pName,
                                                           $"About to call SyncToCoordinates method, RA: {targetRightAscension.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                            telescopeDevice.SyncToCoordinates(targetRightAscension, targetDeclination);
                            LogIssue(pName, $"Failed to reject bad Dec coordinate: {targetDeclination.ToDMS()}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate",
                                $"Correctly rejected bad Dec coordinate: {targetDeclination.ToDMS()}");
                        }

                        break;
                    }

                case SlewSyncType.SlewToTarget:
                case SlewSyncType.SlewToTargetAsync:
                    {
                        LogCallToDriver(pName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(pName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = badCoordinate1;
                            targetDeclination = 0.0d;
                            LogCallToDriver(pName,
                                                           $"About to set TargetRightAscension property to {targetRightAscension.ToHMS()}");
                            telescopeDevice.TargetRightAscension = targetRightAscension;
                            // Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                LogCallToDriver(pName,
                                    $"About to set TargetDeclination property to {targetDeclination.ToDMS()}");
                                telescopeDevice.TargetDeclination = targetDeclination;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (pTest == SlewSyncType.SlewToTarget)
                                {
                                    LogCallToDriver(pName, "About to call SlewToTarget method");
                                    telescopeDevice.SlewToTarget();
                                }
                                else
                                {
                                    LogCallToDriver(pName, "About to call SlewToTargetAsync method");
                                    telescopeDevice.SlewToTargetAsync();
                                }

                                SetAction("Attempting to abort slew");
                                try
                                {
                                    AbortSlew(pName);
                                }
                                catch
                                {
                                } // Attempt to stop any motion that has actually started

                                LogIssue(pName, $"Failed to reject bad RA coordinate: {targetRightAscension.ToHMS()}");
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Slew rejected");
                                HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate",
                                    $"Correctly rejected bad RA coordinate: {targetRightAscension.ToHMS()}");
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate",
                                $"Telescope.TargetRA correctly rejected bad RA coordinate: {targetRightAscension.ToHMS()}");
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = TelescopeRaFromSiderealTime(pName, -2.0d);
                            targetDeclination = badCoordinate2;
                            LogCallToDriver(pName,
                                                           $"About to set TargetDeclination property to {targetDeclination.ToDMS()}");
                            telescopeDevice.TargetDeclination = targetDeclination;
                            // Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                LogCallToDriver(pName,
                                    $"About to set TargetRightAscension property to {targetRightAscension.ToHMS()}");
                                telescopeDevice.TargetRightAscension = targetRightAscension;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (pTest == SlewSyncType.SlewToTarget)
                                {
                                    LogCallToDriver(pName, "About to call SlewToTarget method");
                                    telescopeDevice.SlewToTarget();
                                }
                                else
                                {
                                    LogCallToDriver(pName, "About to call SlewToTargetAsync method");
                                    telescopeDevice.SlewToTargetAsync();
                                }

                                SetAction("Attempting to abort slew");
                                try
                                {
                                    AbortSlew(pName);
                                }
                                catch
                                {
                                } // Attempt to stop any motion that has actually started

                                LogIssue(pName, $"Failed to reject bad Dec coordinate: {targetDeclination.ToDMS()}");
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Slew rejected");
                                HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate",
                                    $"Correctly rejected bad Dec coordinate: {targetDeclination.ToDMS()}");
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate",
                                $"Telescope.TargetDeclination correctly rejected bad Dec coordinate: {targetDeclination.ToDMS()}");
                        }

                        break;
                    }

                case SlewSyncType.SyncToTarget:
                    {
                        LogCallToDriver(pName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            LogCallToDriver(pName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = badCoordinate1;
                            targetDeclination = 0.0d;
                            LogCallToDriver(pName,
                                                           $"About to set TargetRightAscension property to {targetRightAscension.ToHMS()}");
                            telescopeDevice.TargetRightAscension = targetRightAscension;
                            // Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                LogCallToDriver(pName,
                                    $"About to set TargetDeclination property to {targetDeclination.ToDMS()}");
                                telescopeDevice.TargetDeclination = targetDeclination;
                            }
                            catch
                            {
                            }

                            try
                            {
                                LogCallToDriver(pName, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                LogIssue(pName, $"Failed to reject bad RA coordinate: {targetRightAscension.ToHMS()}");
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Sync rejected");
                                HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate",
                                    $"Correctly rejected bad RA coordinate: {targetRightAscension.ToHMS()}");
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate",
                                $"Telescope.TargetRA correctly rejected bad RA coordinate: {targetRightAscension.ToHMS()}");
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = TelescopeRaFromSiderealTime(pName, -3.0d);
                            targetDeclination = badCoordinate2;
                            LogCallToDriver(pName,
                                                           $"About to set TargetDeclination property to {targetDeclination.ToDMS()}");
                            telescopeDevice.TargetDeclination = targetDeclination;
                            // Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                LogCallToDriver(pName,
                                    $"About to set TargetRightAscension property to {targetRightAscension.ToHMS()}");
                                telescopeDevice.TargetRightAscension = targetRightAscension;
                            }
                            catch
                            {
                            }

                            try
                            {
                                LogCallToDriver(pName, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                LogIssue(pName, $"Failed to reject bad Dec coordinate: {targetDeclination.ToDMS()}");
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Sync rejected");
                                HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate",
                                    $"Correctly rejected bad Dec coordinate: {targetDeclination.ToDMS()}");
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate",
                                $"Telescope.TargetDeclination correctly rejected bad Dec coordinate: {targetDeclination.ToDMS()}");
                        }

                        break;
                    }

                case SlewSyncType.SlewToAltAz:
                case SlewSyncType.SlewToAltAzAsync:
                    {
                        LogCallToDriver(pName, "About to get Tracking property");
                        if (canSetTracking & telescopeDevice.Tracking)
                        {
                            LogCallToDriver(pName, "About to set Tracking property to false");
                            telescopeDevice.Tracking = false;
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetAltitude = badCoordinate1;
                            targetAzimuth = 45.0d;
                            if (pTest == SlewSyncType.SlewToAltAz)
                            {
                                LogCallToDriver(pName,
                                    $"About to call SlewToAltAz method, Altitude: {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                                telescopeDevice.SlewToAltAz(targetAzimuth, targetAltitude);
                            }
                            else
                            {
                                LogCallToDriver(pName,
                                    $"About To call SlewToAltAzAsync method, Altitude:  {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                                telescopeDevice.SlewToAltAzAsync(targetAzimuth, targetAltitude);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                AbortSlew(pName);
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogIssue(pName, $"Failed to reject bad Altitude coordinate: {targetAltitude.ToDMS()}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "slewing to bad Altitude coordinate", $"Correctly rejected bad Altitude coordinate: {targetAltitude.ToDMS()}");
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetAltitude = 45.0d;
                            targetAzimuth = badCoordinate2;
                            if (pTest == SlewSyncType.SlewToAltAz)
                            {
                                LogCallToDriver(pName,
                                    $"About to call SlewToAltAz method, Altitude: {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                                telescopeDevice.SlewToAltAz(targetAzimuth, targetAltitude);
                            }
                            else
                            {
                                LogCallToDriver(pName,
                                    $"About to call SlewToAltAzAsync method, Altitude: {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                                telescopeDevice.SlewToAltAzAsync(targetAzimuth, targetAltitude);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                AbortSlew(pName);
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogIssue(pName, $"Failed to reject bad Azimuth coordinate: {FormatAzimuth(targetAzimuth)}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "slewing to bad Azimuth coordinate",
                                $"Correctly rejected bad Azimuth coordinate: {FormatAzimuth(targetAzimuth)}");
                        }

                        break;
                    }

                case SlewSyncType.SyncToAltAz:
                    {
                        LogCallToDriver(pName, "About to get Tracking property");
                        if (canSetTracking & telescopeDevice.Tracking)
                        {
                            LogCallToDriver(pName, "About to set Tracking property to false");
                            telescopeDevice.Tracking = false;
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetAltitude = badCoordinate1;
                            targetAzimuth = 45.0d;
                            LogCallToDriver(pName,
                                                           $"About to call SyncToAltAz method, Altitude: {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                            telescopeDevice.SyncToAltAz(targetAzimuth, targetAltitude);
                            LogIssue(pName, $"Failed to reject bad Altitude coordinate: {targetAltitude.ToDMS()}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "syncing to bad Altitude coordinate", $"Correctly rejected bad Altitude coordinate: {targetAltitude.ToDMS()}");
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetAltitude = 45.0d;
                            targetAzimuth = badCoordinate2;
                            LogCallToDriver(pName,
                                                           $"About to call SyncToAltAz method, Altitude: {targetAltitude.ToDMS()}, Azimuth: {targetAzimuth.ToDMS()}");
                            telescopeDevice.SyncToAltAz(targetAzimuth, targetAltitude);
                            LogIssue(pName, $"Failed to reject bad Azimuth coordinate: {FormatAzimuth(targetAzimuth)}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "syncing to bad Azimuth coordinate",
                                $"Correctly rejected bad Azimuth coordinate: {FormatAzimuth(targetAzimuth)}");
                        }

                        break;
                    }

                default:
                    {
                        LogIssue(pName, $"Conform:SlewTest: Unknown test type {pTest}");
                        break;
                    }
            }

            if (cancellationToken.IsCancellationRequested)
                return;
        }

        private void TelescopePerformanceTest(PerformanceType pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime, lRate;
            SetAction(pName);
            try
            {
                lStartTime = DateTime.Now;
                lCount = 0.0d;
                lLastElapsedTime = 0.0d;
                do
                {
                    lCount += 1.0d;
                    switch (pType)
                    {
                        case PerformanceType.TstPerfAltitude:
                            {
                                altitude = telescopeDevice.Altitude;
                                break;
                            }

                        case var @case when @case == PerformanceType.TstPerfAtHome:
                            {
                                atHome = telescopeDevice.AtHome;
                                break;
                            }

                        case PerformanceType.TstPerfAtPark:
                            {
                                atPark = telescopeDevice.AtPark;
                                break;
                            }

                        case PerformanceType.TstPerfAzimuth:
                            {
                                azimuth = telescopeDevice.Azimuth;
                                break;
                            }

                        case PerformanceType.TstPerfDeclination:
                            {
                                declination = telescopeDevice.Declination;
                                break;
                            }

                        case PerformanceType.TstPerfIsPulseGuiding:
                            {
                                isPulseGuiding = telescopeDevice.IsPulseGuiding;
                                break;
                            }

                        case PerformanceType.TstPerfRightAscension:
                            {
                                rightAscension = telescopeDevice.RightAscension;
                                break;
                            }

                        case PerformanceType.TstPerfSideOfPier:
                            {
                                sideOfPier = (PointingState)telescopeDevice.SideOfPier;
                                break;
                            }

                        case PerformanceType.TstPerfSiderealTime:
                            {
                                siderealTimeScope = telescopeDevice.SiderealTime;
                                break;
                            }

                        case PerformanceType.TstPerfSlewing:
                            {
                                slewing = telescopeDevice.Slewing;
                                break;
                            }

                        case PerformanceType.TstPerfUtcDate:
                            {
                                utcDate = telescopeDevice.UTCDate;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"Conform:PerformanceTest: Unknown test type {pType}");
                                break;
                            }
                    }

                    lElapsedTime = DateTime.Now.Subtract(lStartTime).TotalSeconds;
                    if (lElapsedTime > lLastElapsedTime + 1.0d)
                    {
                        SetStatus($"{lCount} transactions in {lElapsedTime:0} seconds");
                        lLastElapsedTime = lElapsedTime;
                        //Application.DoEvents();
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (lElapsedTime <= PERF_LOOP_TIME);
                lRate = lCount / lElapsedTime;
                switch (lRate)
                {
                    case var case1 when case1 > 10.0d:
                        {
                            LogInfo($"Performance: {pName}", $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case var case2 when 2.0d <= case2 && case2 <= 10.0d:
                        {
                            LogInfo($"Performance: {pName}", $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case var case3 when 1.0d <= case3 && case3 <= 2.0d:
                        {
                            LogInfo($"Performance: {pName}", $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    default:
                        {
                            LogInfo($"Performance: {pName}", $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogIssue($"Performance: {pName}", $"Exception {ex.Message}");
            }
        }

        private void TelescopeParkedExceptionTest(ParkedExceptionType pType, string methodName)
        {
            double targetRa;

            LogCallToDriver($"Parked:{methodName}", "About to get AtPark property");
            if (telescopeDevice.AtPark) // We are still parked so test AbortSlew
            {
                try
                {
                    switch (pType)
                    {
                        case ParkedExceptionType.TstPExcepAbortSlew:
                            AbortSlew("Parked");
                            break;

                        case ParkedExceptionType.TstPExcepFindHome:
                            LogCallToDriver($"Parked:{methodName}", "About to call FindHome method");
                            telescopeDevice.FindHome();
                            // Wait for mount to find home
                            WaitWhile("Waiting for mount to home...", () => !telescopeDevice.AtHome & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000), 200, settings.TelescopeMaximumSlewTime);
                            break;

                        case ParkedExceptionType.TstPExcepMoveAxisPrimary:
                            LogCallToDriver($"Parked:{methodName}", "About to call MoveAxis(Primary, 0.0) method");
                            telescopeDevice.MoveAxis(TelescopeAxis.Primary, 0.0d);
                            break;

                        case ParkedExceptionType.TstPExcepMoveAxisSecondary:
                            LogCallToDriver($"Parked:{methodName}", "About to call MoveAxis(Secondary, 0.0) method");
                            telescopeDevice.MoveAxis(TelescopeAxis.Secondary, 0.0d);
                            break;

                        case ParkedExceptionType.TstPExcepMoveAxisTertiary:
                            LogCallToDriver($"Parked:{methodName}", "About to call MoveAxis(Tertiary, 0.0) method");
                            telescopeDevice.MoveAxis(TelescopeAxis.Tertiary, 0.0d);
                            break;

                        case ParkedExceptionType.TstPExcepPulseGuide:
                            LogCallToDriver($"Parked:{methodName}", "About to call PulseGuide(East, 0.0) method");
                            telescopeDevice.PulseGuide(GuideDirection.East, 0);
                            break;

                        case ParkedExceptionType.TstPExcepSlewToCoordinates:
                            LogCallToDriver($"Parked:{methodName}", "About to call SlewToCoordinates method");
                            telescopeDevice.SlewToCoordinates(TelescopeRaFromSiderealTime($"Parked:{methodName}", 1.0d), 0.0d);
                            break;

                        case ParkedExceptionType.TstPExcepSlewToCoordinatesAsync:
                            LogCallToDriver($"Parked:{methodName}", "About to call SlewToCoordinatesAsync method");
                            telescopeDevice.SlewToCoordinatesAsync(TelescopeRaFromSiderealTime($"Parked:{methodName}", 1.0d), 0.0d);
                            WaitForSlew($"Parked:{methodName}", "Slewing to coordinates asynchronously");
                            break;

                        case ParkedExceptionType.TstPExcepSlewToTarget:
                            targetRa = TelescopeRaFromSiderealTime($"Parked:{methodName}", 1.0d);
                            LogCallToDriver($"Parked:{methodName}", $"About to set property TargetRightAscension to {targetRa.ToHMS()}");
                            telescopeDevice.TargetRightAscension = targetRa;
                            LogCallToDriver($"Parked:{methodName}", "About to set property TargetDeclination to 0.0");
                            telescopeDevice.TargetDeclination = 0.0d;
                            LogCallToDriver($"Parked:{methodName}", "About to call SlewToTarget method");
                            telescopeDevice.SlewToTarget();
                            break;

                        case ParkedExceptionType.TstPExcepSlewToTargetAsync:
                            targetRa = TelescopeRaFromSiderealTime($"Parked:{methodName}", 1.0d);
                            LogCallToDriver($"Parked:{methodName}", $"About to set property to {targetRa.ToHMS()}");
                            telescopeDevice.TargetRightAscension = targetRa;
                            LogCallToDriver($"Parked:{methodName}", "About to set property to 0.0");
                            telescopeDevice.TargetDeclination = 0.0d;
                            LogCallToDriver($"Parked:{methodName}", "About to call method");
                            telescopeDevice.SlewToTargetAsync();
                            WaitForSlew($"Parked:{methodName}", "Slewing to target asynchronously");
                            break;

                        case ParkedExceptionType.TstPExcepSyncToCoordinates:
                            targetRa = TelescopeRaFromSiderealTime($"Parked:{methodName}", 1.0d);
                            LogCallToDriver($"Parked:{methodName}", $"About to call method, RA: {targetRa.ToHMS()}, Declination: 0.0");
                            telescopeDevice.SyncToCoordinates(targetRa, 0.0d);
                            break;

                        case ParkedExceptionType.TstPExcepSyncToTarget:
                            targetRa = TelescopeRaFromSiderealTime($"Parked:{methodName}", 1.0d);
                            LogCallToDriver($"Parked:{methodName}", $"About to set property to {targetRa.ToHMS()}");
                            telescopeDevice.TargetRightAscension = targetRa;
                            LogCallToDriver($"Parked:{methodName}", "About to set property to 0.0");
                            telescopeDevice.TargetDeclination = 0.0d;
                            LogCallToDriver($"Parked:{methodName}", "About to call SyncToTarget method");
                            telescopeDevice.SyncToTarget();
                            break;

                        case ParkedExceptionType.TstExcepTracking:
                            LogCallToDriver($"Parked:{methodName}", "About to get Tracking");
                            bool orginalTrackingState = telescopeDevice.Tracking;
                            LogCallToDriver($"Parked:{methodName}", "About to set Tracking True");
                            telescopeDevice.Tracking = true;

                            // If we get here setting Tracking to true did not throw an exception when parked so try to rest the original Tracking state if necessary
                            if (orginalTrackingState == false) // Tracking was originally false so try to reset it to this state ignoring any errors
                            {
                                try
                                {
                                    LogCallToDriver($"Parked:{methodName}", "About to set Tracking False");
                                    telescopeDevice.Tracking = false;
                                }
                                catch (Exception)
                                {
                                    // Ignore any errors here                                }
                                }
                            }
                            break;

                        default:
                            LogError($"Parked:{methodName}", $"Conform:ParkedExceptionTest: Unknown test type {pType}");
                            break;
                    }

                    LogIssue($"Parked:{methodName}", $"{methodName} didn't raise an error when Parked as required");
                }
                catch (Exception ex)
                {
                    // Check whether this is a Platform 7 or later device
                    if (IsPlatform7OrLater) // This is a Platform 7 or later device so expect a ParkedException
                    {
                        // Check whether this is a ParkedException or its COM equivalent
                        if (IsParkedException(ex))
                            LogOk($"Parked:{methodName}", $"{methodName} threw a ParkedException when Parked, as required. ({ex.GetType().Name} - {ex.Message})");
                        else // Not a parked exception so report an issue
                        {
                            // Handle COM and .NET exceptions
                            if (ex is COMException comException) // COM exception
                                LogIssue($"Parked:{methodName}", $"{methodName} threw a COMException with error number 0x{comException.HResult:X8} when a ParkedException (0x80040408) was expected. Exception message: {ex.Message}");
                            else // .NET exception
                                LogIssue($"Parked:{methodName}", $"{methodName} threw a {ex.GetType().Name} exception when a ParkedException was expected. Exception message: {ex.Message}");
                        }
                    }
                    else // Platform 6 or earlier device so accept any exception
                        LogOk($"Parked:{methodName}", $"{methodName} threw an exception when Parked, as required. The {ex.GetType().Name} exception message was: {ex.Message}");
                }
                // Check that Telescope is still parked after issuing the command!
                LogCallToDriver($"Parked:{methodName}", "About to get AtPark property");
                if (!telescopeDevice.AtPark)
                    LogIssue($"Parked:{methodName}",
                        $"Telescope was unparked by the {methodName} command. This should not happen!");
            }
            else
            {
                LogIssue($"Parked:{methodName}",
                    $"Not parked after Telescope.Park command, {methodName} when parked test skipped");
            }

        }

        protected bool IsParkedException(Exception deviceException)
        {
            bool isParkedException = false; // Set false default value

            try
            {
                switch (deviceException)
                {
                    // This is a COM exception so test whether the error code indicates that it is a parked exception
                    case COMException exception:
                        if (exception.ErrorCode == ErrorCodes.InvalidWhileParked) // This is a parked exception
                            isParkedException = true;
                        break;

                    case ParkedException: // This is a parked exception
                        isParkedException = true;
                        break;

                    default: // Unexpected type of exception so report it as a Conform error
                        LogError("IsParkedException", $"Unsupported exception: {deviceException.GetType().Name} - {deviceException.Message}\r\n{deviceException}");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError("IsParkedException", $"Unexpected exception: {ex}");
            }

            return isParkedException;
        }

        private void TelescopeAxisRateTest(string pName, TelescopeAxis pAxis)
        {
            int lNAxisRates, lI, lJ;
            bool lAxisRateOverlap = default, lAxisRateDuplicate, lCanGetAxisRates = default, lHasRates = default;
            int lCount;

            IAxisRates lAxisRatesIRates;
            IAxisRates lAxisRates = null;
            IRate lRate = null;

            try
            {
                lNAxisRates = 0;
                lAxisRates = null;
                switch (pAxis)
                {
                    case TelescopeAxis.Primary:
                        LogCallToDriver(pName, $"About to call AxisRates method, Axis: {((int)TelescopeAxis.Primary)}");
                        lAxisRates = telescopeDevice.AxisRates(TelescopeAxis.Primary); // Get primary axis rates
                        break;

                    case TelescopeAxis.Secondary:
                        LogCallToDriver(pName, $"About to call AxisRates method, Axis: {((int)TelescopeAxis.Secondary)}");
                        lAxisRates = telescopeDevice.AxisRates(TelescopeAxis.Secondary); // Get secondary axis rates
                        break;

                    case TelescopeAxis.Tertiary:
                        LogCallToDriver(pName, $"About to call AxisRates method, Axis: {((int)TelescopeAxis.Tertiary)}");
                        lAxisRates = telescopeDevice.AxisRates(TelescopeAxis.Tertiary); // Get tertiary axis rates
                        break;

                    default:
                        LogIssue("TelescopeAxisRateTest", $"Unknown telescope axis: {pAxis}");
                        break;
                }

                try
                {
                    if (lAxisRates is null)
                    {
                        LogDebug(pName, "ERROR: The driver did NOT return an AxisRates object!");
                    }
                    else
                    {
                        LogDebug(pName, "OK - the driver returned an AxisRates object");
                    }

                    lCount = lAxisRates.Count; // Save count for use later if no members are returned in the for each loop test
                    LogDebug($"{pName} Count", $"The driver returned {lCount} rates");
                    int i;
                    var loopTo = lCount;
                    for (i = 1; i <= loopTo; i++)
                    {
                        IRate axisRateItem;

                        axisRateItem = lAxisRates[i];
                        LogDebug($"{pName} Count",
                            $"Rate {i} - Minimum: {axisRateItem.Minimum}, Maximum: {axisRateItem.Maximum}");
                    }
                }
                catch (Exception ex)
                {
                    LogIssue($"{pName} Count", $"Unexpected exception: {ex.Message}");
                }

                try
                {
                    IEnumerator lEnum;
                    dynamic lObj;
                    IRate axisRateItem = null;

                    lEnum = (IEnumerator)lAxisRates.GetEnumerator();
                    if (lEnum is null)
                    {
                        LogDebug($"{pName} Enum", "ERROR: The driver did NOT return an Enumerator object!");
                    }
                    else
                    {
                        LogDebug($"{pName} Enum", "OK - the driver returned an Enumerator object");
                    }

                    lEnum.Reset();
                    LogDebug($"{pName} Enum", "Reset Enumerator");
                    while (lEnum.MoveNext())
                    {
                        LogDebug($"{pName} Enum", "Reading Current");
                        lObj = lEnum.Current;
                        LogDebug($"{pName} Enum", "Read Current OK, Type: " + lObj.GetType().Name);
                        axisRateItem = lObj;

                        LogDebug($"{pName} Enum",
                            $"Found axis rate - Minimum: {axisRateItem.Minimum}, Maximum: {axisRateItem.Maximum}");
                    }

                    lEnum.Reset();
                    lEnum = null;
                    axisRateItem = null;
                }
                catch (Exception ex)
                {
                    LogIssue($"{pName} Enum", $"Exception: {ex}");
                }

                if (lAxisRates.Count > 0)
                {
                    try
                    {
                        lAxisRatesIRates = lAxisRates;
                        foreach (IRate currentLRate in lAxisRatesIRates)
                        {
                            lRate = currentLRate;
                            if ((lRate.Minimum < 0.0d) | (lRate.Maximum < 0.0d)) // Error because negative values are not allowed
                            {
                                LogIssue(pName,
                                    $"Minimum or maximum rate is negative: {lRate.Minimum}, {lRate.Maximum}");
                            }
                            else  // All positive values so continue tests
                            {
                                if (lRate.Minimum <= lRate.Maximum) // Minimum <= Maximum so OK
                                {
                                    LogOk(pName,
                                        $"Axis rate minimum: {lRate.Minimum} Axis rate maximum: {lRate.Maximum}");
                                }
                                else // Minimum > Maximum so error!
                                {
                                    LogIssue(pName,
                                        $"Maximum rate is less than minimum rate - minimum: {lRate.Minimum} maximum: {lRate.Maximum}");
                                }
                            }
                            // Save rates for overlap testing
                            lNAxisRates += 1;
                            axisRatesArray[lNAxisRates, AXIS_RATE_MINIMUM] = lRate.Minimum;
                            axisRatesArray[lNAxisRates, AXIS_RATE_MAXIMUM] = lRate.Maximum;
                            lHasRates = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogIssue(pName, $"Unable to read AxisRates object - Exception: {ex.Message}");
                        LogDebug(pName, $"Unable to read AxisRates object - Exception: {ex}");
                    }

                    // Overlap testing
                    if (lNAxisRates > 1) // Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    {
                        int loopTo1 = lNAxisRates;
                        for (lI = 1; lI <= loopTo1; lI++)
                        {
                            int loopTo2 = lNAxisRates;
                            for (lJ = 1; lJ <= loopTo2; lJ++)
                            {
                                if (lI != lJ) // Only test different lines, shouldn't compare same lines!
                                {
                                    if (axisRatesArray[lI, AXIS_RATE_MINIMUM] >= axisRatesArray[lJ, AXIS_RATE_MINIMUM] & axisRatesArray[lI, AXIS_RATE_MINIMUM] <= axisRatesArray[lJ, AXIS_RATE_MAXIMUM])
                                        lAxisRateOverlap = true;
                                }
                            }
                        }
                    }

                    if (lAxisRateOverlap)
                    {
                        LogIssue(pName, "Overlapping axis rates found, suggest these be rationalised to remove overlaps");
                    }
                    else
                    {
                        LogOk(pName, "No overlapping axis rates found");
                    }

                    // Duplicate testing
                    lAxisRateDuplicate = false;
                    if (lNAxisRates > 1) // Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    {
                        int loopTo3 = lNAxisRates;
                        for (lI = 1; lI <= loopTo3; lI++)
                        {
                            int loopTo4 = lNAxisRates;
                            for (lJ = 1; lJ <= loopTo4; lJ++)
                            {
                                if (lI != lJ) // Only test different lines, shouldn't compare same lines!
                                {
                                    if (axisRatesArray[lI, AXIS_RATE_MINIMUM] == axisRatesArray[lJ, AXIS_RATE_MINIMUM] & axisRatesArray[lI, AXIS_RATE_MAXIMUM] == axisRatesArray[lJ, AXIS_RATE_MAXIMUM])
                                        lAxisRateDuplicate = true;
                                }
                            }
                        }
                    }

                    if (lAxisRateDuplicate)
                    {
                        LogIssue(pName, "Duplicate axis rates found, suggest these be removed");
                    }
                    else
                    {
                        LogOk(pName, "No duplicate axis rates found");
                    }
                }
                else
                {
                    LogOk(pName, "Empty axis rate returned");
                }

                lCanGetAxisRates = true; // Record that this driver can deliver a viable AxisRates object that can be tested for AxisRates.Dispose() later
            }
            catch (Exception ex)
            {
                LogIssue(pName, $"Unable to get or unable to use an AxisRates object - Exception: {ex}");
            }

            // Clean up AxisRate object if used
            if (lAxisRates is object)
            {
                try
                {
                    lAxisRates.Dispose();
                }
                catch
                {
                }

                if (OperatingSystem.IsWindows())
                {
                    try
                    {
                        Marshal.ReleaseComObject(lAxisRates);
                    }
                    catch
                    {
                    }
                }

                lAxisRates = null;
            }

            // Clean up and release rate object if used
            if (lRate is object)
            {
                try
                {
                    lRate.Dispose();
                }
                catch
                {
                }

                if (OperatingSystem.IsWindows())
                {
                    try
                    {
                        Marshal.ReleaseComObject(lRate);
                    }
                    catch { }
                }
            }

            if (lCanGetAxisRates) // The driver does return a viable AxisRates object that can be tested for correct AxisRates.Dispose() and Rate.Dispose() operation
            {
                try
                {
                    // Test Rate.Dispose()
                    switch (pAxis) // Get the relevant axis rates object for this axis
                    {
                        case TelescopeAxis.Primary:
                            {
                                lAxisRates = telescopeDevice.AxisRates(TelescopeAxis.Primary);
                                break;
                            }

                        case TelescopeAxis.Secondary:
                            {
                                lAxisRates = telescopeDevice.AxisRates(TelescopeAxis.Secondary);
                                break;
                            }

                        case TelescopeAxis.Tertiary:
                            {
                                lAxisRates = telescopeDevice.AxisRates(TelescopeAxis.Tertiary);
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"AxisRate.Dispose() - Unknown axis: {pAxis}");
                                break;
                            }
                    }

                    if (lHasRates) // This axis does have one or more rates that can be accessed through ForEach so test these for correct Rate.Dispose() action
                    {
                        foreach (IRate currentLRate2 in (IEnumerable)lAxisRates)
                        {
                            lRate = currentLRate2;
                            try
                            {
                                lRate.Dispose();
                                LogOk(pName, string.Format("Successfully disposed of rate {0} - {1}", lRate.Minimum, lRate.Maximum));
                            }
                            catch (MissingMemberException)
                            {
                                LogOk(pName, string.Format("Rate.Dispose() member not present for rate {0} - {1}", lRate.Minimum, lRate.Maximum));
                            }
                            catch (Exception ex1)
                            {
                                LogIssue(pName, string.Format("Rate.Dispose() for rate {0} - {1} threw an exception but it is poor practice to throw exceptions in Dispose methods: {2}", lRate.Minimum, lRate.Maximum, ex1.Message));
                                LogDebug("TrackingRates.Dispose", $"Exception: {ex1}");
                            }
                        }
                    }

                    // Test AxisRates.Dispose()
                    try
                    {
                        LogDebug(pName, "Disposing axis rates");
                        lAxisRates.Dispose();
                        LogOk(pName, "Disposed axis rates OK");
                    }
                    catch (MissingMemberException)
                    {
                        LogOk(pName, $"AxisRates.Dispose() member not present for axis {pAxis}");
                    }
                    catch (Exception ex1)
                    {
                        LogIssue(pName,
                            $"AxisRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: {ex1.Message}");
                        LogDebug("AxisRates.Dispose", $"Exception: {ex1}");
                    }
                }
                catch (Exception ex)
                {
                    LogIssue(pName,
                        $"AxisRate.Dispose() - Unable to get or unable to use an AxisRates object - Exception: {ex}");
                }
            }
            else
            {
                LogInfo(pName, "AxisRates.Dispose() testing skipped because of earlier issues in obtaining a viable AxisRates object.");
            }

        }

        private void TelescopeRequiredMethodsTest(RequiredMethodType pType, string pName)
        {
            try
            {
                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case RequiredMethodType.TstAxisrates:
                            {
                                break;
                            }
                        // This is now done by TelescopeAxisRateTest subroutine 
                        case RequiredMethodType.TstCanMoveAxisPrimary:
                            {
                                LogCallToDriver(pName, $"About to call CanMoveAxis method {((int)TelescopeAxis.Primary)}");
                                canMoveAxisPrimary = telescopeDevice.CanMoveAxis(TelescopeAxis.Primary);
                                LogOk(pName, $"{pName} {canMoveAxisPrimary}");
                                break;
                            }

                        case RequiredMethodType.TstCanMoveAxisSecondary:
                            {
                                LogCallToDriver(pName, $"About to call CanMoveAxis method {((int)TelescopeAxis.Secondary)}");
                                canMoveAxisSecondary = telescopeDevice.CanMoveAxis(TelescopeAxis.Secondary);
                                LogOk(pName, $"{pName} {canMoveAxisSecondary}");
                                break;
                            }

                        case RequiredMethodType.TstCanMoveAxisTertiary:
                            {
                                LogCallToDriver(pName, $"About to call CanMoveAxis method {((int)TelescopeAxis.Tertiary)}");
                                canMoveAxisTertiary = telescopeDevice.CanMoveAxis(TelescopeAxis.Tertiary);
                                LogOk(pName, $"{pName} {canMoveAxisTertiary}");
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"Conform:RequiredMethodsTest: Unknown test type {pType}");
                                break;
                            }
                    }
                }, TargetTime.Fast);
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Method, Required.Mandatory, ex, "");
            }

            // Clean up and release each object after use
            // If Not (m_Rate Is Nothing) Then Try : Marshal.ReleaseComObject(m_Rate) : Catch : End Try
            // m_Rate = Nothing
        }

        private void TelescopeOptionalMethodsTest(OptionalMethodType testType, string testName, bool canTest)
        {
            IAxisRates axisRates = null;

            SetTest(testName);
            LogDebug("TelescopeOptionalMethodsTest", $"Test type: {testType}, Test name: {testName}, Can test: {canTest}");

            // Check whether this method is supported by the driver
            if (canTest) // This method is supported by the driver
            {
                try
                {
                    // Process each type of test
                    switch (testType)
                    {
                        case OptionalMethodType.AbortSlew:
                            AbortSlew(testName);
                            LogOk(testName, "AbortSlew OK when not slewing");

                            // If we get here the AbortSlew method is available so start a slew and then try to halt it.
                            // The following test relies on SlewToCoordinatesAsync so only undertake it if this method is supported
                            if (canSlewAsync) // Async slew is supported
                            {

                                // Start a slew to HA -3.0
                                double targetHa = -3.0;

                                if (canSetTracking)
                                {
                                    LogCallToDriver(testName, "About to set Tracking property to true");
                                    telescopeDevice.Tracking = true; // Enable tracking for these tests
                                }

                                // Calculate the target RA
                                LogCallToDriver(testName, "About to get SiderealTime property");
                                double targetRa = Utilities.ConditionRA(telescopeDevice.SiderealTime - targetHa);

                                // Calculate the target declination
                                targetDeclination = GetTestDeclinationLessThan65(testName, targetRa);
                                LogDebug(testName, $"Slewing to HA: {targetHa.ToHMS()} (RA: {targetRa.ToHMS()}), Dec: {targetDeclination.ToDMS()}");

                                // Slew to the target coordinates
                                LogCallToDriver(testName, $"About to call SlewToCoordinatesAsync. RA: {targetRa.ToHMS()}, Declination: {targetDeclination.ToDMS()}");

                                LogCallToDriver(testName, $"About to call SlewToCoordinatesAsync method, RA: {targetRa.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                                telescopeDevice.SlewToCoordinatesAsync(targetRa, targetDeclination);

                                // Wait for 1.5 seconds
                                WaitFor(1500, 100);

                                // Validate that the slew is still going
                                ValidateSlewing(testName, true);

                                // Now try to end the slew, waiting up to 30 seconds for this to happen
                                Stopwatch sw = Stopwatch.StartNew();
                                TimeMethod(testName, () => AbortSlew(testName), TargetTime.Standard);
                                try
                                {
                                    // Wait for the mount to report that it is no longer slewing or for the wait to time out
                                    LogCallToDriver(testName, $"About to call Slewing repeatedly...");
                                    WaitWhile("Waiting for slew to stop", () => telescopeDevice.Slewing == true, 500, settings.TelescopeTimeForSlewingToBecomeFalse);
                                    LogOk(testName, $"AbortSlew stopped the mount from slewing in {sw.Elapsed.TotalSeconds:0.0} seconds.");
                                }
                                catch (TimeoutException)
                                {
                                    LogIssue(testName, $"The mount still reports Slewing as TRUE {settings.TelescopeTimeForSlewingToBecomeFalse} seconds after AbortSlew returned.");
                                }
                                catch (Exception ex)
                                {
                                    LogIssue(testName, $"The mount reported an exception while waiting for Slewing to become false after AbortSlew: {ex.Message}");
                                    LogDebug(testName, ex.ToString());
                                }
                                finally
                                {
                                    sw.Stop();
                                }
                            }
                            else // Async slew is not supported
                            {
                                LogInfo(testName, $"Skipping the abort SlewToCoordinatesAsync test because CanSlewAsync is {canSlewAsync}.");
                            }
                            break;

                        case OptionalMethodType.DestinationSideOfPier:
                            // This test only applies to German Polar (German Equatorial) mounts

                            // Set the test declination value depending on whether the scope is in the northern or southern hemisphere
                            if (siteLatitude > 0.0d)
                                targetDeclination = 45.0d; // Positive for the northern hemisphere
                            else
                                targetDeclination = -45.0d; // Negative for the southern hemisphere

                            // Set the hour angle offset as 2 hours from local sider5eal time
                            double hourAngleOffset = 2.0d;
                            LogDebug(testName, $"Test hour angle: {hourAngleOffset}, Test declination: {targetDeclination}");

                            // Slew to HA +3.0 if possible. This puts the tube on the East side of the pier looking West which and, by ASCOM convention, SideofPier should report the pierEast / Normal pointing state
                            if (canSlewAsync)
                                SlewToHa(+3.0);

                            PointingState currentSideOfPier = PointingState.Unknown;
                            if (CanReadSideOfPier(testName))
                            {
                                LogCallToDriver(testName, "About to get SideOfPier");
                                currentSideOfPier = telescopeDevice.SideOfPier;
                            }

                            // Get the DestinationSideOfPier for a target in the West (positive HA offset) i.e. for a German mount when the tube is on the East side of the pier looking West
                            double targetRightAscensionWest = TelescopeRaFromHourAngle(testName, +hourAngleOffset);

                            LogCallToDriver(testName, $"About to call DestinationSideOfPier method, RA: {targetRightAscensionWest.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                            destinationSideOfPierWest = TimeFunc(testName, () => telescopeDevice.DestinationSideOfPier(targetRightAscensionWest, targetDeclination), TargetTime.Fast);
                            LogDebug(testName, $"Current pointing state: {currentSideOfPier} - Western target pointing state: {destinationSideOfPierWest} at RA: {targetRightAscensionWest.ToHMS()}, Declination {targetDeclination.ToDMS()} ");

                            // Get the DestinationSideOfPier for a target in the East (negative HA offset) i.e. for a German mount when the tube is on the West side of the pier looking East
                            double targetRightAscensionEast = TelescopeRaFromHourAngle(testName, -hourAngleOffset);
                            LogCallToDriver(testName, $"About to call DestinationSideOfPier method, Target RA: {targetRightAscensionEast.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
                            destinationSideOfPierEast = TimeFunc(testName, () => telescopeDevice.DestinationSideOfPier(targetRightAscensionEast, targetDeclination), TargetTime.Fast);
                            LogDebug(testName, $"Current pointing state: {currentSideOfPier} - Eastern target pointing state: {destinationSideOfPierEast} {targetRightAscensionEast.ToHMS()} {targetDeclination.ToDMS()} ");

                            // Make sure that we received two valid values i.e. that neither side returned PierSide.Unknown and that the two valid returned values are not the same i.e. we got one PointingState.Normal and one PointingState.ThroughThePole
                            if (destinationSideOfPierEast == PointingState.Unknown | destinationSideOfPierWest == PointingState.Unknown)
                            {
                                LogIssue(testName, $"Unexpected SideOfPier value received (Unknown), Target in East: {destinationSideOfPierEast}, Target in West: {destinationSideOfPierWest}");
                            }
                            else if (destinationSideOfPierEast == destinationSideOfPierWest)
                            {
                                LogIssue(testName, $"Same value for DestinationSideOfPier received on both sides of the meridian - East: {destinationSideOfPierEast}, West: {destinationSideOfPierWest}");
                            }
                            else
                            {
                                LogOk(testName, $"DestinationSideOfPier is different on either side of the meridian - East: {destinationSideOfPierEast}, West: {destinationSideOfPierWest}");
                            }
                            break;

                        case OptionalMethodType.FindHome:
                            SetAction("Homing mount...");
                            if (GetInterfaceVersion() > 1)
                            {
                                bool currentTrackngState = false;

                                // Report the current Tracking state
                                if (canReadTracking)
                                {
                                    LogCallToDriver(testName, "About to get Tracking property");
                                    currentTrackngState = telescopeDevice.Tracking; // Save the current tracking state
                                    LogDebug(testName, $"Current tracking state: {currentTrackngState}");
                                }
                                else
                                    LogDebug(testName, "Cannot get tracking property");

                                // Set tracking False if possible
                                if (canSetTracking)
                                {
                                    LogCallToDriver(testName, $"About to set Tracking property to false (currently {currentTrackngState})");
                                    telescopeDevice.Tracking = false; // Disable tracking for this test
                                }
                                else
                                    LogDebug(testName, "Cannot set tracking to false");

                                // Test FindHome()
                                LogCallToDriver(testName, "About to call FindHome method");
                                TimeMethod(testName, () => telescopeDevice.FindHome(), IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

                                // Wait for the home to complete using Platform 6 or 7 semantics as appropriate
                                if (IsPlatform7OrLater) // Platform 7 or later device
                                {
                                    LogCallToDriver(testName, "About to get Slewing property repeatedly...");
                                    WaitWhile("Waiting for scope to home...", () => telescopeDevice.Slewing, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                }
                                else // Platform 6 device
                                {
                                    LogCallToDriver(testName, "About to get AtHome property repeatedly...");
                                    WaitWhile("Waiting for scope to home...", () => !telescopeDevice.AtHome, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                }

                                // Ensure Tracking is false after FindHome if possible
                                if (canReadTracking) // Can read the tracking state
                                {
                                    LogCallToDriver(testName, "About to get Tracking property");
                                    currentTrackngState = telescopeDevice.Tracking; // Save the current tracking state
                                    LogDebug(testName, $"Post FindHome tracking state: {currentTrackngState}");

                                    // Reset tracking to false if it has been enabled by FindHome
                                    if (currentTrackngState)
                                    {
                                        LogCallToDriver(testName, $"About to set Tracking property to false (currently {currentTrackngState})");
                                        telescopeDevice.Tracking = false; // DIsable tracking for this test
                                    }
                                    else
                                        LogDebug(testName, "Tracking is already false so no need to change it");
                                }
                                else // Cannot read the tracking state
                                {
                                    // Set tracking False if possible
                                    if (canSetTracking)
                                    {
                                        LogCallToDriver(testName, $"About to set Tracking property to false (state currently unknown)");
                                        telescopeDevice.Tracking = false; // Disable tracking for this test
                                    }
                                    else
                                        LogDebug(testName, "Cannot set tracking to false");
                                }

                                // Validate FindHome outcome                                
                                LogCallToDriver(testName, "About to get AtHome property");
                                if (telescopeDevice.AtHome)
                                {
                                    LogOk(testName, "Found home OK, AtHome reports TRUE as expected.");
                                }
                                else
                                {
                                    LogIssue(testName, $"AtHome reports false after FindHome.");
                                    LogInfo(testName, $"This could be because of an issue in implementing AtHome, or because Tracking has been automatically enabled by FindHome or because FindHome did not complete within the configured {settings.TelescopeMaximumSlewTime} second timeout.");
                                }

                                // Unpark the mount in case it was parked by the Home command
                                LogCallToDriver(testName, "About to get AtPark property");
                                if (telescopeDevice.AtPark)
                                {
                                    LogIssue(testName, "FindHome has parked the scope as well as finding home");

                                    LogCallToDriver(testName, "About to call UnPark method");
                                    telescopeDevice.Unpark(); // Unpark it ready for further tests

                                    LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                    WaitWhile("Waiting for scope to unpark", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                }

                                // Re-enable tracking
                                if (canSetTracking)
                                {
                                    LogCallToDriver(testName, $"About to set Tracking property to true after FindHome test");
                                    telescopeDevice.Tracking = true;
                                }
                                else
                                    LogDebug(testName, "Cannot set tracking to true after FindHome test.");
                            }
                            else // Interface version 1
                            {
                                SetAction("Waiting for mount to home");
                                LogCallToDriver(testName, "About to call FindHome method");
                                TimeMethod(testName, () => telescopeDevice.FindHome(), TargetTime.Extended);

                                // Wait for mount to find home
                                WaitWhile("Waiting for mount to home...", () => !telescopeDevice.AtHome, 200, settings.TelescopeMaximumSlewTime);

                                // Now make sure that the mount is unparked
                                LogOk(testName, "Found home OK.");
                                LogCallToDriver(testName, "About to call Unpark method");
                                telescopeDevice.Unpark();
                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                WaitWhile("Waiting for scope to unpark", () => telescopeDevice.AtPark, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                            }
                            break;

                        case OptionalMethodType.MoveAxisPrimary:
                            // Get axis rates for primary axis
                            LogCallToDriver(testName, $"About to call AxisRates method for {TelescopeAxis.Primary} axis");
                            axisRates = telescopeDevice.AxisRates(TelescopeAxis.Primary);
                            TelescopeMoveAxisTest(testName, TelescopeAxis.Primary, axisRates);
                            break;

                        case OptionalMethodType.MoveAxisSecondary:
                            // Get axis rates for secondary axis
                            LogCallToDriver(testName, $"About to call AxisRates method for {TelescopeAxis.Secondary} axis");
                            axisRates = telescopeDevice.AxisRates(TelescopeAxis.Secondary);
                            TelescopeMoveAxisTest(testName, TelescopeAxis.Secondary, axisRates);
                            break;

                        case OptionalMethodType.MoveAxisTertiary:
                            // Get axis rates for tertiary axis
                            LogCallToDriver(testName, $"About to call AxisRates method for {TelescopeAxis.Tertiary} axis");
                            axisRates = telescopeDevice.AxisRates(TelescopeAxis.Tertiary);
                            TelescopeMoveAxisTest(testName, TelescopeAxis.Tertiary, axisRates);
                            break;

                        case OptionalMethodType.PulseGuide:
                            // Ensure that tracking is true for this test
                            if (canSetTracking)
                            {
                                LogCallToDriver(testName, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            LogCallToDriver(testName, "About to get IsPulseGuiding property");
                            if (telescopeDevice.IsPulseGuiding) // IsPulseGuiding is true before we've started so this is an error and voids a real test
                            {
                                LogIssue(testName, "IsPulseGuiding is True when not pulse guiding - PulseGuide test omitted");
                            }
                            else // OK to test pulse guiding
                            {
                                SetAction("Calling PulseGuide east");
                                startTime = DateTime.Now;
                                LogCallToDriver(testName, $"About to call PulseGuide method, Direction: {((int)GuideDirection.East)}, Duration: {PULSEGUIDE_MOVEMENT_TIME * 1000}ms");
                                TimeMethod($"{testName} {GuideDirection.East} {PULSEGUIDE_MOVEMENT_TIME}s", () => telescopeDevice.PulseGuide(GuideDirection.East, PULSEGUIDE_MOVEMENT_TIME * 1000), TargetTime.Standard); // Start a 2 second pulse
                                endTime = DateTime.Now;
                                double pulseGuideTime = endTime.Subtract(startTime).TotalSeconds; // Seconds
                                LogDebug(testName, $"PulseGuide command time: {PULSEGUIDE_MOVEMENT_TIME:0.0} seconds, PulseGuide call duration: {pulseGuideTime:0.0} seconds");

                                // Check whether the pulse guide completed before the pulse duration. Target time depends on interface version. Standard response time for Platform 7, 75% of pulse guide time for Platform 6
                                if (pulseGuideTime < (IsPlatform7OrLater ? Globals.STANDARD_TARGET_RESPONSE_TIME : PULSEGUIDE_MOVEMENT_TIME * 0.75)) // Returned in less than the synchronous limit time so treat as asynchronous
                                {
                                    LogCallToDriver(testName, "About to get IsPulseGuiding property");
                                    if (telescopeDevice.IsPulseGuiding)
                                    {
                                        LogCallToDriver(testName, "About to get IsPulseGuiding property multiple times");
                                        WaitWhile("Pulse guiding East", () => telescopeDevice.IsPulseGuiding, SLEEP_TIME, PULSEGUIDE_TIMEOUT_TIME);

                                        LogCallToDriver(testName, "About to get IsPulseGuiding property");
                                        if (!telescopeDevice.IsPulseGuiding)
                                        {
                                            LogOk(testName, "Asynchronous single axis pulse guide East found OK");
                                            LogDebug(testName, $"IsPulseGuiding = True duration: {DateTime.Now.Subtract(startTime).TotalMilliseconds} milliseconds");

                                            // Successfully tested an asynchronous guide to the East, now test to the North.
                                            TimeMethod($"{testName} {GuideDirection.North} {PULSEGUIDE_MOVEMENT_TIME}s", () => telescopeDevice.PulseGuide(GuideDirection.North, PULSEGUIDE_MOVEMENT_TIME * 1000), TargetTime.Standard); // Start a 2 second pulse
                                            LogCallToDriver(testName, "About to get IsPulseGuiding property multiple times");
                                            WaitWhile("Pulse guiding North", () => telescopeDevice.IsPulseGuiding, SLEEP_TIME, PULSEGUIDE_TIMEOUT_TIME);

                                            // Check outcome of pulse guide north
                                            if (!telescopeDevice.IsPulseGuiding)
                                            {
                                                LogOk(testName, "Asynchronous single axis pulse guide North found OK");
                                                LogDebug(testName, $"IsPulseGuiding = True duration: {DateTime.Now.Subtract(startTime).TotalMilliseconds} milliseconds");

                                                // Successfully tested asynchronous guides to the East and North individually, now test whether they can run concurrently

                                                LogCallToDriver(testName, "About to call PulseGuide East");
                                                telescopeDevice.PulseGuide(GuideDirection.East, PULSEGUIDE_MOVEMENT_TIME * 1000); // Start a 2 second pulse guide East

                                                bool dualAxisGuiding = false; // Flag indicating whether the mount supports dual axis guiding
                                                try
                                                {
                                                    LogCallToDriver(testName, "About to call PulseGuide North");
                                                    telescopeDevice.PulseGuide(GuideDirection.North, PULSEGUIDE_MOVEMENT_TIME * 1000); // Start a 2 second pulse guide North
                                                    dualAxisGuiding = true;
                                                }
                                                catch (ASCOM.InvalidOperationException ex)
                                                {
                                                    LogOk(testName, $"The mount is not capable of concurrent guiding in two directions, it correctly threw an InvalidOperationException: {ex.Message}");
                                                    LogDebug(testName, ex.ToString());
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleException("PulseGuide - Dual Axis", MemberType.Method, Required.MustBeImplemented, ex, $"received a {ex.GetType().Name} exception: {ex.Message}");
                                                    LogInfo("PulseGuide - Dual Axis", $"If the mount is not capable of dual axis auto-guiding it should throw " +
                                                        $"an InvalidOperationException when the second PulseGuide is started.");
                                                    LogDebug("PulseGuide - Dual Axis", ex.ToString());
                                                }

                                                LogCallToDriver(testName, "About to get IsPulseGuiding property multiple times");
                                                WaitWhile("Pulse guiding East and North", () => telescopeDevice.IsPulseGuiding, SLEEP_TIME, PULSEGUIDE_TIMEOUT_TIME);

                                                // Check outcome of pulse guide East and North
                                                if (!telescopeDevice.IsPulseGuiding)
                                                {
                                                    LogOk(testName, dualAxisGuiding ? "Asynchronous dual axis pulse guide East and North found OK" :
                                                        "Single axis guide completed OK because the mount can not guide both axes simultaneously.");
                                                }
                                                else
                                                {
                                                    LogIssue(testName, $"Asynchronous dual axis pulse guide East and North expected, but IsPulseGuiding is still TRUE {PULSEGUIDE_TIMEOUT_TIME} seconds beyond expected time");
                                                }
                                            }
                                            else
                                            {
                                                LogIssue(testName, $"Asynchronous pulse guide North expected, but IsPulseGuiding is still TRUE {PULSEGUIDE_TIMEOUT_TIME} seconds beyond expected time");
                                            }
                                        }
                                        else
                                        {
                                            LogIssue(testName, $"Asynchronous pulse guide East expected, but IsPulseGuiding is still TRUE {PULSEGUIDE_TIMEOUT_TIME} seconds beyond expected time");
                                        }
                                    }
                                    else
                                    {
                                        LogIssue(testName, "Asynchronous pulse guide expected, but IsPulseGuiding returned FALSE");
                                    }
                                }
                                else // Took longer than the synchronous limit time so treat as synchronous
                                {
                                    // Check whether this is a Platform 7 or later interface
                                    if (IsPlatform7OrLater) // Platform 7 or later is expected
                                    {
                                        LogIssue(testName, $"Synchronous pulse guide found. The guide took {pulseGuideTime} seconds to complete, the expected maximum time is {Globals.STANDARD_TARGET_RESPONSE_TIME} seconds.");
                                        LogInfo(testName, $"In ITelescopeV4 and later interfaces, PulseGuide should be implemented asynchronously: returning quickly with IsPulseGuiding returning TRUE.");
                                        LogInfo(testName, $"When guide movement is complete, IsPulseGuiding should return FALSE.");
                                    }
                                    else // Platform 6 or earlier behaviour is expected
                                    {
                                        LogCallToDriver(testName, "About to get IsPulseGuiding property");
                                        if (!telescopeDevice.IsPulseGuiding)
                                        {
                                            LogOk(testName, "Synchronous pulse guide found OK");
                                        }
                                        else
                                        {
                                            LogIssue(testName, "Synchronous pulse guide expected but IsPulseGuiding has returned TRUE");
                                        }
                                    }
                                }
                            }
                            break;

                        case OptionalMethodType.SideOfPierWrite:
                            // SideOfPier Write
                            if (canSetPierside) // Can set pier side so test if we can
                            {
                                SlewScope(TelescopeRaFromHourAngle(testName, -3.0d), 0.0d, "far start point");
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                SlewScope(TelescopeRaFromHourAngle(testName, -0.03d), 0.0d, "near start point"); // 2 minutes from zenith
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                // We are now 2 minutes from the meridian looking east so allow the mount to track for 7 minutes 
                                // so it passes through the meridian and ends up 5 minutes past the meridian
                                LogInfo(testName, "This test will now wait for 7 minutes while the mount tracks through the Meridian");

                                // Wait for mount to move
                                startTime = DateTime.Now;
                                do
                                {
                                    Thread.Sleep(1000);
                                    SetFullStatus(testName, "Waiting for transit through Meridian", $"{Convert.ToInt32(DateTime.Now.Subtract(startTime).TotalSeconds)}/{SIDEOFPIER_MERIDIAN_TRACKING_PERIOD / 1000d} seconds");
                                }
                                while ((DateTime.Now.Subtract(startTime).TotalMilliseconds <= SIDEOFPIER_MERIDIAN_TRACKING_PERIOD) & !cancellationToken.IsCancellationRequested);

                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                LogCallToDriver(testName, "About to get SideOfPier property");

                                switch (telescopeDevice.SideOfPier)
                                {
                                    case PointingState.Normal: // We are on pierEast so try pierWest
                                        {
                                            try
                                            {
                                                LogDebug(testName, "Scope is pierEast so flipping West");
                                                SetAction("Flipping mount to pointing state pierWest");
                                                LogCallToDriver(testName, $"About to set SideOfPier property to {PointingState.ThroughThePole}");
                                                TimeMethod($"{testName} {PointingState.ThroughThePole}", () => telescopeDevice.SideOfPier = PointingState.ThroughThePole, TargetTime.Standard);

                                                WaitForSlew(testName, $"Moving to the pierEast pointing state asynchronously");

                                                LogCallToDriver(testName, "About to get SideOfPier property");
                                                sideOfPier = telescopeDevice.SideOfPier;

                                                if (sideOfPier == PointingState.ThroughThePole)
                                                {
                                                    LogOk(testName, "Successfully flipped pierEast to pierWest");
                                                }
                                                else
                                                {
                                                    LogIssue(testName, $"Failed to set SideOfPier to pierWest, got: {sideOfPier}");
                                                }

                                                if (cancellationToken.IsCancellationRequested)
                                                    return;
                                            }
                                            catch (Exception ex)
                                            {
                                                HandleException("SideOfPier Write pierWest", MemberType.Method, Required.MustBeImplemented, ex, "CanSetPierSide is True");
                                            }

                                            break;
                                        }

                                    case PointingState.ThroughThePole: // We are on pierWest so try pierEast
                                        {
                                            try
                                            {
                                                LogDebug(testName, "Scope is pierWest so flipping East");
                                                SetAction("Flipping mount to pointing state pierEast");

                                                LogCallToDriver(testName, $"About to set SideOfPier property to {PointingState.Normal}");
                                                TimeMethod($"{testName} {PointingState.Normal}", () => telescopeDevice.SideOfPier = PointingState.Normal, TargetTime.Standard);

                                                WaitForSlew(testName, $"Moving to the pierWest pointing state asynchronously");

                                                if (cancellationToken.IsCancellationRequested)
                                                    return;

                                                LogCallToDriver(testName, "About to get SideOfPier property");
                                                sideOfPier = telescopeDevice.SideOfPier;
                                                if (sideOfPier == PointingState.Normal)
                                                    LogOk(testName, "Successfully flipped pierWest to pierEast");
                                                else
                                                    LogIssue(testName, $"Failed to set SideOfPier to pierEast, got: {sideOfPier}");
                                            }
                                            catch (Exception ex)
                                            {
                                                HandleException("SideOfPier Write pierEast", MemberType.Method, Required.MustBeImplemented, ex, "CanSetPierSide is True");
                                            }
                                            break;
                                        }

                                    default:
                                        LogIssue(testName, $"Unknown PierSide: {sideOfPier}");
                                        break;
                                }
                            }
                            else // Can't set pier side so it should generate an error
                            {
                                try
                                {
                                    LogDebug(testName, "Attempting to set SideOfPier");
                                    LogCallToDriver(testName, $"About to set SideOfPier property to {((int)PointingState.Normal)}");
                                    telescopeDevice.SideOfPier = PointingState.Normal;

                                    LogDebug(testName, "SideOfPier set OK to pierEast but should have thrown an error");
                                    WaitForSlew(testName, $"Moving to the pierWest pointing state asynchronously");
                                    LogIssue(testName, "CanSetPierSide is false but no exception was generated when set was attempted");
                                }
                                catch (Exception ex)
                                {
                                    HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, "CanSetPierSide is False");
                                }
                                finally
                                {
                                    WaitForSlew(testName, $"Moving to the pierWest pointing state asynchronously");
                                } // Make sure slewing is stopped if an exception was thrown
                            }

                            LogCallToDriver(testName, "About to set Tracking property to false");
                            telescopeDevice.Tracking = false;

                            if (cancellationToken.IsCancellationRequested)
                                return;
                            break;

                        default:
                            LogIssue(testName, $"Conform:OptionalMethodsTest: Unknown test type {testType}");
                            break;
                    }

                    // Clean up AxisRate object, if used
                    if (axisRates is object)
                    {
                        try
                        {
                            LogCallToDriver(testName, "About to dispose of AxisRates object");
                            axisRates.Dispose();

                            LogOk(testName, "AxisRates object successfully disposed");
                        }
                        catch (Exception ex)
                        {
                            LogIssue(testName, $"AxisRates.Dispose threw an exception but must not: {ex.Message}");
                            LogDebug(testName, $"Exception: {ex}");
                        }

                        try
                        {
#if WINDOWS
                            Marshal.ReleaseComObject(axisRates);
#endif
                        }
                        catch
                        {
                        }

                        axisRates = null;
                    }
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
                catch (Exception ex)
                {
                    LogIssue(testName, $"{testType} Exception\r\n{ex}");
                    HandleException(testName, MemberType.Method, Required.Optional, ex, "");
                }
            }
            else // Can property is false so confirm that an error is generated
            {
                try
                {
                    switch (testType)
                    {
                        case OptionalMethodType.AbortSlew:
                            {
                                AbortSlew(testName);
                                break;
                            }

                        case OptionalMethodType.DestinationSideOfPier:
                            {
                                targetRightAscension = TelescopeRaFromSiderealTime(testName, -1.0d);
                                LogCallToDriver(testName, $"About to call DestinationSideOfPier method, RA: {targetRightAscension.ToHMS()}, Declination: {0.0d.ToDMS()}");
                                destinationSideOfPier = (PointingState)telescopeDevice.DestinationSideOfPier(targetRightAscension, 0.0d);
                                break;
                            }

                        case OptionalMethodType.FindHome:
                            {
                                LogCallToDriver(testName, "About to call FindHome method");
                                telescopeDevice.FindHome();
                                // Wait for mount to find home
                                WaitWhile("Waiting for mount to home...", () => !telescopeDevice.AtHome & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000), 200, settings.TelescopeMaximumSlewTime);
                                break;
                            }

                        case OptionalMethodType.MoveAxisPrimary:
                            {
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {((int)TelescopeAxis.Primary)} at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxis.Primary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.MoveAxisSecondary:
                            {
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {((int)TelescopeAxis.Secondary)} at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxis.Secondary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.MoveAxisTertiary:
                            {
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {((int)TelescopeAxis.Tertiary)} at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxis.Tertiary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.PulseGuide:
                            {
                                LogCallToDriver(testName, $"About to call PulseGuide method, Direction: {((int)GuideDirection.East)}, Duration: 0ms");
                                telescopeDevice.PulseGuide(GuideDirection.East, 0);
                                break;
                            }

                        case OptionalMethodType.SideOfPierWrite:
                            {
                                LogCallToDriver(testName, $"About to set SideOfPier property to {((int)PointingState.Normal)}");
                                telescopeDevice.SideOfPier = PointingState.Normal;
                                break;
                            }

                        default:
                            {
                                LogIssue(testName, $"Conform:OptionalMethodsTest: Unknown test type {testType}");
                                break;
                            }
                    }

                    LogIssue(testName, $"Can{testName} is false but no exception was generated on use");
                }
                catch (Exception ex)
                {
                    if (IsInvalidValueException(testName, ex))
                    {
                        LogOk(testName, "Received an invalid value exception");
                    }
                    else if (testType == OptionalMethodType.SideOfPierWrite) // PierSide is actually a property even though I have it in the methods section!!
                    {
                        HandleException(testName, MemberType.Property, Required.MustNotBeImplemented, ex, $"Can{testName} is False");
                    }
                    else
                    {
                        HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, $"Can{testName} is False");
                    }
                }
            }

            ClearStatus();
        }

        private void TelescopeCanTest(CanType pType, string pName)
        {
            try
            {
                LogCallToDriver(pName, string.Format("About to get {0} property", pType.ToString()));

                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case CanType.CanFindHome:
                            canFindHome = telescopeDevice.CanFindHome;
                            LogOk(pName, canFindHome.ToString());
                            break;

                        case CanType.CanPark:
                            canPark = telescopeDevice.CanPark;
                            LogOk(pName, canPark.ToString());
                            break;

                        case CanType.CanPulseGuide:
                            canPulseGuide = telescopeDevice.CanPulseGuide;
                            LogOk(pName, canPulseGuide.ToString());
                            break;

                        case CanType.CanSetDeclinationRate:
                            if (GetInterfaceVersion() > 1)
                            {
                                canSetDeclinationRate = telescopeDevice.CanSetDeclinationRate;
                                LogOk(pName, canSetDeclinationRate.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetDeclinationRate", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
                            }
                            break;

                        case CanType.CanSetGuideRates:
                            if (GetInterfaceVersion() > 1)
                            {
                                canSetGuideRates = telescopeDevice.CanSetGuideRates;
                                LogOk(pName, canSetGuideRates.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetGuideRates", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
                            }
                            break;

                        case CanType.CanSetPark:
                            canSetPark = telescopeDevice.CanSetPark;
                            LogOk(pName, canSetPark.ToString());
                            break;

                        case CanType.CanSetPierSide:
                            if (GetInterfaceVersion() > 1)
                            {
                                canSetPierside = telescopeDevice.CanSetPierSide;
                                LogOk(pName, canSetPierside.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetPierSide", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
                            }
                            break;

                        case CanType.CanSetRightAscensionRate:
                            if (GetInterfaceVersion() > 1)
                            {
                                canSetRightAscensionRate = telescopeDevice.CanSetRightAscensionRate;
                                LogOk(pName, canSetRightAscensionRate.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetRightAscensionRate", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
                            }
                            break;

                        case CanType.CanSetTracking:
                            canSetTracking = telescopeDevice.CanSetTracking;
                            LogOk(pName, canSetTracking.ToString());
                            break;

                        case CanType.CanSlew:
                            canSlew = telescopeDevice.CanSlew;
                            LogOk(pName, canSlew.ToString());
                            break;

                        case CanType.CanSlewAltAz:
                            if (GetInterfaceVersion() > 1)
                            {
                                canSlewAltAz = telescopeDevice.CanSlewAltAz;
                                LogOk(pName, canSlewAltAz.ToString());
                            }
                            else
                            {
                                LogInfo("CanSlewAltAz", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
                            }
                            break;

                        case CanType.CanSlewAltAzAsync:
                            if (GetInterfaceVersion() > 1)
                            {
                                canSlewAltAzAsync = telescopeDevice.CanSlewAltAzAsync;
                                LogOk(pName, canSlewAltAzAsync.ToString());
                            }
                            else
                            {
                                LogInfo("CanSlewAltAzAsync", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
                            }
                            break;

                        case CanType.CanSlewAsync:
                            canSlewAsync = telescopeDevice.CanSlewAsync;
                            LogOk(pName, canSlewAsync.ToString());
                            break;

                        case CanType.CanSync:
                            canSync = telescopeDevice.CanSync;
                            LogOk(pName, canSync.ToString());
                            break;

                        case CanType.CanSyncAltAz:
                            if (GetInterfaceVersion() > 1)
                            {
                                canSyncAltAz = telescopeDevice.CanSyncAltAz;
                                LogOk(pName, canSyncAltAz.ToString());
                            }
                            else
                            {
                                LogInfo("CanSyncAltAz", $"Skipping test as this method is not supported in interface V{GetInterfaceVersion()}");
                            }
                            break;

                        case CanType.CanUnPark:
                            canUnpark = telescopeDevice.CanUnpark;
                            LogOk(pName, canUnpark.ToString());
                            break;

                        default:
                            LogIssue(pName, $"Conform:CanTest: Unknown test type {pType}");
                            break;
                    }
                }, TargetTime.Fast);
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.Mandatory, ex, "");
            }
        }

        private void TelescopeMoveAxisTest(string testName, TelescopeAxis testAxis, IAxisRates axisRates)
        {
            IRate rate = null;

            double moveRate = default, rateMinimum, rateMaximum;
            bool trackingStart, trackingEnd, canSetZero;
            int rateCount;

            SetTest(testName);

            // Determine lowest and highest tracking rates
            rateMinimum = double.PositiveInfinity; // Set to invalid values
            rateMaximum = double.NegativeInfinity;
            LogDebug(testName, $"Number of rates found: {axisRates.Count}");

            // Make sure that some axis rates are available
            if (axisRates.Count > 0) // Some rates are available
            {
                IAxisRates lAxisRatesIRates = axisRates;
                rateCount = 0;
                foreach (IRate currentLRate in lAxisRatesIRates)
                {
                    rate = currentLRate;
                    if (rate.Minimum < rateMinimum) rateMinimum = rate.Minimum;
                    if (rate.Maximum > rateMaximum) rateMaximum = rate.Maximum;
                    LogDebug(testName, $"Checking rates: {rate.Minimum} {rate.Maximum} Current rates: {rateMinimum} {rateMaximum}");
                    rateCount += 1;
                }

                if (rateMinimum != double.PositiveInfinity & rateMaximum != double.NegativeInfinity) // Found valid rates
                {
                    LogDebug(testName, $"Found minimum rate: {rateMinimum} found maximum rate: {rateMaximum}");

                    // Confirm setting a zero rate works
                    SetAction("Set zero rate");
                    canSetZero = false;
                    try
                    {
                        LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                        TimeMethod($"{testName} {testAxis} 0", () => telescopeDevice.MoveAxis(testAxis, 0.0d), TargetTime.Standard); // Set a value of zero

                        if (WaitForSlewingTobecomeFalse(testName))
                        {
                            LogOk(testName, "Can successfully set a movement rate of zero");
                            canSetZero = true;
                        }
                        else
                        {
                            LogIssue(testName, "Slewing was not false after setting a rate of 0..");
                            canSetZero = false;
                        }
                    }
                    catch (COMException ex)
                    {
                        LogIssue(testName, $"Unable to set a movement rate of zero - {ex.Message} {((int)ex.ErrorCode):X8}");
                    }
                    catch (DriverException ex)
                    {
                        LogIssue(testName, $"Unable to set a movement rate of zero - {ex.Message} {ex.Number:X8}");
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unable to set a movement rate of zero - {ex.Message}");
                    }

                    SetAction("Set lower rate");

                    // Test that error is generated on attempt to set rate lower than minimum
                    try
                    {
                        if (rateMinimum > 0d) // choose a value between the minimum and zero
                            moveRate = rateMinimum / 2.0d;
                        else // Choose a large negative value
                            moveRate = -rateMaximum - 1.0d;

                        LogDebug(testName, $"Using minimum rate: {moveRate}");
                        LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {moveRate}");
                        telescopeDevice.MoveAxis(testAxis, moveRate); // Set a value lower than the minimum
                        LogIssue(testName, $"No exception raised when move axis value < minimum rate: {moveRate}");
                        // Clean up and release each object after use
#if WINDOWS
                        try { Marshal.ReleaseComObject(rate); } catch { }
#endif

                        rate = null;
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsOk(testName, MemberType.Method, Required.MustBeImplemented, ex,
                            $"when move axis is set below lowest rate ({moveRate})",
                            $"Exception correctly generated when move axis is set below lowest rate ({moveRate})");
                    }

                    // Clean up and release each object after use
                    try
                    {
#if WINDOWS
                        Marshal.ReleaseComObject(rate);
#endif
                    }
                    catch
                    {
                    }

                    rate = null;
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    // Test that error is generated when rate is above maximum set
                    SetAction("Set upper rate");
                    try
                    {
                        moveRate = rateMaximum + 1.0d;
                        LogDebug(testName, $"Using maximum rate: {moveRate}");
                        LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {moveRate}");
                        telescopeDevice.MoveAxis(testAxis, moveRate); // Set a value higher than the maximum
                        LogIssue(testName, $"No exception raised when move axis value > maximum rate: {moveRate}");
                        // Clean up and release each object after use
#if WINDOWS
                        try { Marshal.ReleaseComObject(rate); } catch { }
#endif

                        rate = null;
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsOk(testName, MemberType.Method, Required.MustBeImplemented, ex,
                            $"when move axis is set above highest rate ({moveRate})",
                            $"Exception correctly generated when move axis is set above highest rate ({moveRate})");
                    }
                    // Clean up and release each object after use
                    try
                    {
#if WINDOWS
                        if (rate is not null) Marshal.ReleaseComObject(rate);
#endif
                    }
                    catch
                    {
                    }

                    rate = null;
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    if (canSetZero) // Can set a rate of zero so undertake these tests
                    {
                        // Confirm that lowest tracking rate can be set
                        if (rateMinimum != double.PositiveInfinity) // Valid value found so try and set it
                        {
                            try
                            {
                                SetAction("Moving forward at minimum rate");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {rateMinimum}");
                                TimeMethod($"{testName} {testAxis} {rateMinimum}", () => telescopeDevice.MoveAxis(testAxis, rateMinimum), TargetTime.Standard); // Set the minimum rate

                                // Assess outcome depending on whether or not the minimum rate was 0.0
                                if (rateMinimum == 0.0)
                                {
                                    if (ValidateSlewing(testName, false))
                                        LogOk(testName, $"Can successfully set a movement rate of {rateMinimum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {rateMinimum}");
                                }
                                else // rateMinimum != 0.0
                                {
                                    if (ValidateSlewing(testName, true))
                                        LogOk(testName, $"Can successfully set a movement rate of {rateMinimum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {rateMinimum}");
                                }

                                WaitFor(MOVE_AXIS_TIME);

                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                // Stop movement at minimum rate
                                SetAction("Stopping movement");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis

                                // Wait up to the configured time for slewing to become false
                                if (WaitForSlewingTobecomeFalse(testName)) // Wait was successful
                                    LogOk(testName, $"Successfully stopped movement.");
                                else // Something went wrong
                                    LogIssue(testName, $"Slewing was not false after setting a rate of 0.0");

                                // Move back at minimum rate
                                SetAction("Moving back at minimum rate");

                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {testAxis} at speed {-rateMinimum}");
                                telescopeDevice.MoveAxis(testAxis, -rateMinimum); // Set the minimum rate

                                // Assess outcome depending on whether or not the minimum rate was 0.0
                                if (rateMinimum == 0.0)
                                {
                                    if (ValidateSlewing(testName, false))
                                        LogOk(testName, $"Can successfully set a movement rate of {rateMinimum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {rateMinimum}");
                                }
                                else // rateMinimum != 0.0
                                {
                                    if (ValidateSlewing(testName, true))
                                        LogOk(testName, $"Can successfully set a movement rate of {-rateMinimum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {-rateMinimum}");
                                }

                                WaitFor(MOVE_AXIS_TIME);

                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                // Stop movement at -minimum rate
                                SetAction("Stopping movement");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis

                                // Wait up to the configured time for slewing to become false
                                if (WaitForSlewingTobecomeFalse(testName)) // Wait was successful
                                    LogOk(testName, $"Successfully stopped movement.");
                                else // Something went wrong
                                    LogIssue(testName, $"Slewing was not false after setting a rate of 0.0");
                            }
                            catch (Exception ex)
                            {
                                HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, $"when setting rate: {rateMinimum}");
                            }

                            SetStatus(""); // Clear status flag
                        }
                        else // No valid rate was found so print an error
                        {
                            LogIssue(testName, "Minimum rate test - unable to find lowest axis rate");
                        }

                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Confirm that highest tracking rate can be set
                        if (rateMaximum != double.NegativeInfinity) // Valid value found so try and set it
                        {
                            try
                            {
                                SetAction("Moving forward at maximum rate");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {rateMaximum}");
                                TimeMethod($"{testName} {testAxis} {rateMaximum}", () => telescopeDevice.MoveAxis(testAxis, rateMaximum), TargetTime.Standard); // Set the maximum rate

                                // Assess outcome depending on whether or not the maximum rate was 0.0
                                if (rateMaximum == 0.0)
                                {
                                    if (ValidateSlewing(testName, false))
                                        LogOk(testName, $"Can successfully set a movement rate of {rateMaximum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {rateMaximum}");

                                }
                                else // rateMaximum != 0.0
                                {
                                    if (ValidateSlewing(testName, true))
                                        LogOk(testName, $"Can successfully set a movement rate of {rateMaximum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {rateMaximum}");
                                }

                                WaitFor(MOVE_AXIS_TIME);

                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                // Stop movement at maximum rate
                                SetAction("Stopping movement");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis

                                // Wait up to the configured time for slewing to become false
                                if (WaitForSlewingTobecomeFalse(testName)) // Wait was successful
                                    LogOk(testName, $"Successfully stopped movement.");
                                else // Something went wrong
                                    LogIssue(testName, $"Slewing was not false after setting a rate of 0.0");

                                // Move back at maximum rate
                                SetAction("Moving back at maximum rate");

                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {testAxis} at speed {-rateMaximum}");
                                telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate

                                // Assess outcome depending on whether or not the maximum rate was 0.0
                                if (rateMaximum == 0.0)
                                {
                                    if (ValidateSlewing(testName, false))
                                        LogOk(testName, $"Can successfully set a movement rate of {rateMaximum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {rateMaximum}");
                                }
                                else // rateMaximum != 0.0
                                {
                                    if (ValidateSlewing(testName, true))
                                        LogOk(testName, $"Can successfully set a movement rate of {-rateMaximum}");
                                    else
                                        LogIssue(testName, $"Slewing was not false after setting a rate of {-rateMaximum}");
                                }

                                WaitFor(MOVE_AXIS_TIME);

                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                // Stop movement at -maximum rate
                                SetAction("Stopping movement");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis

                                // Wait up to the configured time for slewing to become false
                                if (WaitForSlewingTobecomeFalse(testName)) // Wait was successful
                                    LogOk(testName, $"Successfully stopped movement.");
                                else // Something went wrong
                                    LogIssue(testName, $"Slewing was not false after setting a rate of 0.0");
                            }
                            catch (Exception ex)
                            {
                                HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, $"when setting rate: {rateMaximum}");
                            }

                            SetStatus(""); // Clear status flag
                        }
                        else // No valid rate was found so print an error
                        {
                            LogIssue(testName, "Maximum rate test - unable to find lowest axis rate");
                        }

                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Confirm that tracking state is correctly restored after a move axis command
                        try
                        {
                            SetAction("Tracking state restore");
                            if (canSetTracking)
                            {
                                LogCallToDriver(testName, "About to get Tracking property");
                                trackingStart = telescopeDevice.Tracking; // Save the start tracking state
                                SetStatus("Moving forward");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {rateMaximum}");
                                telescopeDevice.MoveAxis(testAxis, rateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                SetStatus("Stop movement");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis                                
                                WaitForSlewingTobecomeFalse(testName); // Wait up to the configured time for slewing to become false

                                LogCallToDriver(testName, "About to get Tracking property");
                                trackingEnd = telescopeDevice.Tracking; // Save the final tracking state
                                if (trackingStart == trackingEnd) // Successfully retained tracking state
                                {
                                    if (trackingStart) // Tracking is true so switch to false for return movement
                                    {
                                        SetStatus("Set tracking off");
                                        LogCallToDriver(testName, "About to set Tracking property false");
                                        telescopeDevice.Tracking = false;
                                        SetStatus("Move back");
                                        LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {-rateMaximum}");
                                        telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                        WaitFor(MOVE_AXIS_TIME);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;

                                        LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                        telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                        WaitForSlewingTobecomeFalse(testName); // Wait up to the configured time for slewing to become false

                                        SetStatus("");
                                        LogCallToDriver(testName, "About to get Tracking property");
                                        if (telescopeDevice.Tracking == false) // tracking correctly retained in both states
                                            LogOk(testName, "Tracking state correctly retained for both tracking states");
                                        else
                                            LogIssue(testName, $"Tracking state correctly retained when tracking is {trackingStart}, but not when tracking is false");
                                    }
                                    else // Tracking false so switch to true for return movement
                                    {
                                        SetStatus("Set tracking on");
                                        LogCallToDriver(testName, "About to set Tracking property true");
                                        telescopeDevice.Tracking = true;
                                        SetStatus("Move back");
                                        LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {-rateMaximum}");
                                        telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                        WaitFor(MOVE_AXIS_TIME);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;

                                        LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                        telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                        WaitForSlewingTobecomeFalse(testName); // Wait up to the configured time for slewing to become false

                                        SetStatus("");
                                        LogCallToDriver(testName, "About to get Tracking property");
                                        if (telescopeDevice.Tracking == true) // tracking correctly retained in both states
                                            LogOk(testName, "Tracking state correctly retained for both tracking states");
                                        else
                                            LogIssue(testName, $"Tracking state correctly retained when tracking is {trackingStart}, but not when tracking is true");
                                    }

                                    SetStatus(""); // Clear status flag
                                }
                                else // Tracking state not correctly restored
                                {
                                    SetStatus("Move back");
                                    LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {-rateMaximum}");
                                    telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                    WaitFor(MOVE_AXIS_TIME);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    SetStatus("");
                                    LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                    telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                    WaitForSlewingTobecomeFalse(testName); // Wait up to the configured time for slewing to become false

                                    LogCallToDriver(testName, $"About to set Tracking property {trackingStart}");
                                    telescopeDevice.Tracking = trackingStart; // Restore original value
                                    LogIssue(testName, "Tracking state not correctly restored after MoveAxis when CanSetTracking is true");
                                }
                            }
                            else // Can't set tracking so just test the current state
                            {
                                LogCallToDriver(testName, "About to get Tracking property");
                                trackingStart = telescopeDevice.Tracking; SetStatus("Moving forward");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed {rateMaximum}");
                                telescopeDevice.MoveAxis(testAxis, rateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                SetStatus("Stop movement");
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                WaitForSlewingTobecomeFalse(testName); // Wait up to the configured time for slewing to become false

                                LogCallToDriver(testName, "About to get Tracking property");
                                trackingEnd = telescopeDevice.Tracking; // Save tracking state
                                SetStatus("Move back");
                                LogCallToDriver(testName, $"About to call method MoveAxis for axis {(int)testAxis} at speed {-rateMaximum}");
                                telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                // v1.0.12 next line added because movement wasn't stopped
                                LogCallToDriver(testName, $"About to call MoveAxis method for axis {(int)testAxis} at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                WaitForSlewingTobecomeFalse(testName); // Wait up to the configured time for slewing to become false

                                if (trackingStart == trackingEnd)
                                {
                                    LogOk(testName, "Tracking state correctly restored after MoveAxis when CanSetTracking is false");
                                }
                                else
                                {
                                    LogCallToDriver(testName, $"About to set Tracking property to {trackingStart}");
                                    telescopeDevice.Tracking = trackingStart; // Restore correct value
                                    LogIssue(testName, "Tracking state not correctly restored after MoveAxis when CanSetTracking is false");
                                }

                                SetStatus("");
                            } // Clear status flag
                        }
                        catch (Exception ex)
                        {
                            HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, "");
                        }
                    }
                    else // Cant set zero so tests skipped 
                    {
                        LogInfo(testName, "Remaining MoveAxis tests skipped because unable to set a movement rate of zero");
                    }

                    ClearStatus(); // Clear status
                }
                else // Some problem in finding rates inside the AxisRates object
                {
                    LogInfo(testName, $"Found minimum rate: {rateMinimum} found maximum rate: {rateMaximum}");
                    LogIssue(testName, $"Unable to determine lowest or highest rates, expected {axisRates.Count} rates, found {rateCount}");
                }
            }
            else
            {
                LogIssue(testName, "MoveAxis tests skipped because there are no AxisRate values");
            }
        }

        private void SideOfPierTests()
        {
            SideOfPierResults pierSideMinus3, pierSideMinus9, pierSidePlus3, pierSidePlus9;
            double declination3, declination9, startRa, startDeclination;

            LogDebug("SideofPier", "Starting Side of Pier tests");
            SetTest("Side of pier tests");

            // Calculate the test RA form hour angle -3.0
            startRa = TelescopeRaFromHourAngle("SideofPier", -3.0d);
            startDeclination = GetTestDeclinationLessThan65("SideofPier", startRa);

            // Calculate the acceptable declinations for hour angle 
            declination3 = GetTestDeclinationHalfwayToHorizon("SideOfPierTests", startRa); // 90.0d - (180.0d - siteLatitude) * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR; // Calculate for northern hemisphere
            declination9 = GetTestDeclinationHalfwayToHorizon("SideOfPierTests", TelescopeRaFromHourAngle("SideofPier", -9.0d)); // 90.0d - siteLatitude * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR;
            LogDebug("SideofPier", $"Declination for hour angle = +-3.0 tests: {declination3.ToDMS()}, Declination for hour angle = +-9.0 tests: {declination9.ToDMS()}");

            // Run tests
            SetAction($"Test hour angle -3.0 at declination: {declination3.ToDMS()}");
            pierSideMinus3 = SopPierTest(TelescopeRaFromHourAngle("SideofPier", -3.0d), declination3, startRa, startDeclination, "hour angle -3.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            SetAction($"Test hour angle -9.0 at declination: {declination9.ToDMS()}");
            pierSideMinus9 = SopPierTest(TelescopeRaFromHourAngle("SideofPier", -9.0d), declination9, startRa, startDeclination, "hour angle -9.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            SetAction($"Test hour angle +3.0 at declination: {declination3.ToDMS()}");
            pierSidePlus3 = SopPierTest(TelescopeRaFromHourAngle("SideofPier", +3.0d), declination3, startRa, startDeclination, "hour angle +3.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            SetAction($"Test hour angle +9.0 at declination: {declination9.ToDMS()}");
            pierSidePlus9 = SopPierTest(TelescopeRaFromHourAngle("SideofPier", +9.0d), declination9, startRa, startDeclination, "hour angle +9.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            LogDebug(" ", " ");

            if ((pierSideMinus3.SideOfPier == pierSidePlus9.SideOfPier) & (pierSidePlus3.SideOfPier == pierSideMinus9.SideOfPier))// Reporting physical pier side
            {
                LogIssue("SideofPier", "SideofPier reports physical pier side rather than pointing state");
            }
            else if ((pierSideMinus3.SideOfPier == pierSideMinus9.SideOfPier) & (pierSidePlus3.SideOfPier == pierSidePlus9.SideOfPier)) // Make other tests
            {
                LogOk("SideofPier", "Reports the pointing state of the mount as expected");
            }
            else // Don't know what this means!
            {
                LogInfo("SideofPier", $"Unknown SideofPier reporting model: HA-3: {pierSideMinus3.SideOfPier} HA-9: {pierSideMinus9.SideOfPier} HA+3: {pierSidePlus3.SideOfPier} HA+9: {pierSidePlus9.SideOfPier}");
            }

            LogInfo("SideofPier", $"Reported SideofPier at HA -9, +9: {TranslatePierSide(pierSideMinus9.SideOfPier, false)}{TranslatePierSide(pierSidePlus9.SideOfPier, false)}");
            LogInfo("SideofPier", $"Reported SideofPier at HA -3, +3: {TranslatePierSide(pierSideMinus3.SideOfPier, false)}{TranslatePierSide(pierSidePlus3.SideOfPier, false)}");

            // Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east observing at hour angle -3
            if (pierSideMinus3.SideOfPier == PointingState.ThroughThePole)
            {
                LogOk("SideofPier", "pierWest is returned when the mount is observing at an hour angle between -6.0 and 0.0");
            }
            else
            {
                LogIssue("SideofPier", "pierEast is returned when the mount is observing at an hour angle between -6.0 and 0.0");
                LogInfo("SideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.");
            }

            if (pierSidePlus3.SideOfPier == (int)PointingState.Normal)
            {
                LogOk("SideofPier", "pierEast is returned when the mount is observing at an hour angle between 0.0 and +6.0");
            }
            else
            {
                LogIssue("SideofPier", "pierWest is returned when the mount is observing at an hour angle between 0.0 and +6.0");
                LogInfo("SideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.");
            }

            // Test whether DestinationSideOfPier is implemented
            if (pierSideMinus3.DestinationSideOfPier == PointingState.Unknown & pierSideMinus9.DestinationSideOfPier == PointingState.Unknown & pierSidePlus3.DestinationSideOfPier == PointingState.Unknown & pierSidePlus9.DestinationSideOfPier == PointingState.Unknown)
            {
                LogInfo("DestinationSideofPier", "Analysis skipped as this method is not implemented"); // Not implemented
            }
            else // It is implemented so assess the results
            {
                if (pierSideMinus3.DestinationSideOfPier == pierSidePlus9.DestinationSideOfPier & pierSidePlus3.DestinationSideOfPier == pierSideMinus9.DestinationSideOfPier) // Reporting physical pier side
                    LogIssue("DestinationSideofPier", "DestinationSideofPier reports physical pier side rather than pointing state");
                else if (pierSideMinus3.DestinationSideOfPier == pierSideMinus9.DestinationSideOfPier & pierSidePlus3.DestinationSideOfPier == pierSidePlus9.DestinationSideOfPier) // Make other tests
                    LogOk("DestinationSideofPier", "Reports the pointing state of the mount as expected");
                else // Don't know what this means!
                    LogInfo("DestinationSideofPier", $"Unknown DestinationSideofPier reporting model: HA-3: {pierSideMinus3.SideOfPier} HA-9: {pierSideMinus9.SideOfPier} HA+3: {pierSidePlus3.SideOfPier} HA+9: {pierSidePlus9.SideOfPier}");

                // Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east at hour angle -3
                if ((int)pierSideMinus3.DestinationSideOfPier == (int)PointingState.ThroughThePole)
                    LogOk("DestinationSideofPier", "pierWest is returned when the mount will observe at an hour angle between -6.0 and 0.0");
                else
                {
                    LogIssue("DestinationSideofPier", "pierEast is returned when the mount will observe at an hour angle between -6.0 and 0.0");
                    LogInfo("DestinationSideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.");
                }

                if (pierSidePlus3.DestinationSideOfPier == (int)PointingState.Normal)
                    LogOk("DestinationSideofPier", "pierEast is returned when the mount will observe at an hour angle between 0.0 and +6.0");
                else
                {
                    LogIssue("DestinationSideofPier", "pierWest is returned when the mount will observe at an hour angle between 0.0 and +6.0");
                    LogInfo("DestinationSideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.");
                }
            }

            LogInfo("DestinationSideofPier", $"Reported DesintationSideofPier at HA -9, +9: {TranslatePierSide((PointingState)pierSideMinus9.DestinationSideOfPier, false)}{TranslatePierSide((PointingState)pierSidePlus9.DestinationSideOfPier, false)}");
            LogInfo("DestinationSideofPier", $"Reported DesintationSideofPier at HA -3, +3: {TranslatePierSide((PointingState)pierSideMinus3.DestinationSideOfPier, false)}{TranslatePierSide((PointingState)pierSidePlus3.DestinationSideOfPier, false)}");

            // Clean up
            // 3.0.0.12 added conditional test to next line
            if (canSetTracking)
                telescopeDevice.Tracking = false;
            ClearStatus();
        }

        public SideOfPierResults SopPierTest(double testRa, double testDec, double startRa, double startDeclination, string message)
        {
            // Determine side of pier and destination side of pier results for a particular RA and DEC
            SideOfPierResults results = new SideOfPierResults(); // Create result set object

            try
            {
                LogDebug("", "");

                // Slew to start RA
                LogDebug("SOPPierTest", $"Starting test for {message}");
                //LogDebug("SOPPierTest", $"Slewing to start position: {startRa.ToHMS()} {startDeclination.ToDMS()}");
                SlewScope(startRa, startDeclination, "start position");
                //LogDebug("SOPPierTest", $"Slewed to start position:  {startRa.ToHMS()} {startDeclination.ToDMS()}");

                LogCallToDriver("SopPierTest", "About to get SideofPier");
                PointingState currentPointingState = telescopeDevice.SideOfPier;
                LogDebug("SOPPierTest", $"Initial pointing state before slewing to test RA/Dec: {TranslatePierSide(currentPointingState, true)} ({currentPointingState})");
                LogDebug("SOPPierTest", $"Test RA: {testRa.ToHMS()},Test declination: {testDec.ToDMS()}");

                // Do destination side of pier test to see what side of pier we should end up on if we slew to the test coordinates
                try
                {
                    results.DestinationSideOfPier = TimeFunc($"DestinatitonSideOfPier", () => telescopeDevice.DestinationSideOfPier(testRa, testDec), TargetTime.Fast);
                    LogDebug("SOPPierTest", $"==> Pointing state predicted by DestinationSideOfPier at test RA/Dec: {TranslatePierSide(results.DestinationSideOfPier, true)} ({results.DestinationSideOfPier})");
                }
                catch (COMException ex)
                {
                    switch (ex.ErrorCode)
                    {
                        case var @case when @case == ErrorCodes.NotImplemented:
                            results.DestinationSideOfPier = PointingState.Unknown;
                            LogDebug("SOPPierTest", $"DestinationSideOfPier is not implemented setting result to: {TranslatePierSide(results.DestinationSideOfPier, true)} ({results.DestinationSideOfPier})");
                            break;

                        default:
                            LogIssue("SOPPierTest", $"DestinationSideOfPier Exception: {ex}");
                            break;
                    }
                }
                catch (MethodNotImplementedException) // DestinationSideOfPier not available so mark as unknown
                {
                    results.DestinationSideOfPier = PointingState.Unknown;
                    LogDebug("SOPPierTest", $"DestinationSideOfPier is not implemented setting result to: {TranslatePierSide(results.DestinationSideOfPier, true)} ({results.DestinationSideOfPier})");
                }
                catch (Exception ex)
                {
                    LogIssue("SOPPierTest", $"DestinationSideOfPier Exception: {ex}");
                }

                // Now do an actual slew and record side of pier we actually get
                SlewScope(testRa, testDec, $"test position {message}");
                results.SideOfPier = telescopeDevice.SideOfPier;
                LogDebug("SOPPierTest", $"==> Actual pointing state after slewing to test RA/Dec: {TranslatePierSide(results.SideOfPier, true)} ({results.SideOfPier})");
            }
            catch (Exception ex)
            {
                LogIssue("SOPPierTest", $"SideofPierException: {ex}");
            }

            return results;
        }

        private void DestinationSideOfPierTests()
        {
            PointingState lPierSideMinus3, lPierSideMinus9, lPierSidePlus3, lPierSidePlus9;

            // Slew to one position, then call destination side of pier 4 times and report the pattern
            SlewScope(TelescopeRaFromHourAngle("DestinationSideofPier", -3.0d), 0.0d, "start position");
            lPierSideMinus3 = telescopeDevice.DestinationSideOfPier(-3.0d, 0.0d);
            lPierSidePlus3 = telescopeDevice.DestinationSideOfPier(3.0d, 0.0d);
            lPierSideMinus9 = telescopeDevice.DestinationSideOfPier(-9.0d, 90.0d - siteLatitude);
            lPierSidePlus9 = telescopeDevice.DestinationSideOfPier(9.0d, 90.0d - siteLatitude);
            if (lPierSideMinus3 == lPierSidePlus9 & lPierSidePlus3 == lPierSideMinus9) // Reporting physical pier side
            {
                LogIssue("DestinationSideofPier", "The driver appears to be reporting physical pier side rather than pointing state");
            }
            else if (lPierSideMinus3 == lPierSideMinus9 & lPierSidePlus3 == lPierSidePlus9) // Make other tests
            {
                LogOk("DestinationSideofPier", "The driver reports the pointing state of the mount");
            }
            else // Don't know what this means!
            {
                LogInfo("DestinationSideofPier",
                    $"Unknown pier side reporting model: HA-3: {lPierSideMinus3} HA-9: {lPierSideMinus9} HA+3: {lPierSidePlus3} HA+9: {lPierSidePlus9}");
            }

            telescopeDevice.Tracking = false;
            LogInfo("DestinationSideofPier", TranslatePierSide(lPierSideMinus9, false) + TranslatePierSide(lPierSidePlus9, false));
            LogInfo("DestinationSideofPier", TranslatePierSide(lPierSideMinus3, false) + TranslatePierSide(lPierSidePlus3, false));
        }

        private void CheckScopePosition(string testName, string functionName, double expectedRa, double expectedDec)
        {
            double actualRa, actualDec, difference;

            LogCallToDriver(testName, "About to get RightAscension property");
            actualRa = telescopeDevice.RightAscension;
            LogDebug(testName, $"Read RightAscension: {actualRa.ToHMS()}");

            LogCallToDriver(testName, "About to get Declination property");
            actualDec = telescopeDevice.Declination;
            LogDebug(testName, $"Read Declination: {actualDec.ToDMS()}");

            // Check that we have actually arrived where we are expected to be
            difference = RaDifferenceInArcSeconds(actualRa, expectedRa); // Convert RA difference to arc seconds

            if (difference <= settings.TelescopeSlewTolerance)
            {
                LogOk(testName, $"{functionName} OK within tolerance: ±{settings.TelescopeSlewTolerance} arc seconds. Actual RA: {actualRa.ToHMS()}, Target RA: {expectedRa.ToHMS()}");

            }
            else
            {
                LogIssue(testName, $"{functionName} {difference:0.0} arc seconds away from RA target: {expectedRa.ToHMS()} Actual RA: {actualRa.ToHMS()} Tolerance: ±{settings.TelescopeSlewTolerance} arc seconds");
            }

            difference = Math.Round(Math.Abs(actualDec - expectedDec) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero);
            if (difference <= settings.TelescopeSlewTolerance) // Dec difference is in arc seconds from degrees of Declination
            {
                LogOk(testName, $"{functionName} OK within tolerance: ±{settings.TelescopeSlewTolerance} arc seconds. Actual DEC: {actualDec.ToDMS()}, Target DEC: {expectedDec.ToDMS()}");
            }
            else
            {
                LogIssue(testName, $"{functionName} {difference:0.0} arc seconds from the expected DEC: {expectedDec.ToDMS()} Actual DEC: {actualDec.ToDMS()} Tolerance: ±{settings.TelescopeSlewTolerance} degrees.");
            }
        }

        /// <summary>
        /// Return the difference between two RAs (in hours) as seconds
        /// </summary>
        /// <param name="firstRa">First RA (hours)</param>
        /// <param name="secondRa">Second RA (hours)</param>
        /// <returns>Difference (seconds) between the supplied RAs</returns>
        private static double RaDifferenceInArcSeconds(double firstRa, double secondRa)
        {
            double raDifferenceInSecondsRet = Math.Abs(firstRa - secondRa); // Calculate the difference allowing for negative outcomes
            if (raDifferenceInSecondsRet > 12.0d) raDifferenceInSecondsRet = 24.0d - raDifferenceInSecondsRet; // Deal with the cases where the two elements are more than 12 hours apart going in the initial direction

            raDifferenceInSecondsRet = Math.Round(raDifferenceInSecondsRet * 15.0d * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // RA difference is in arc seconds from hours of RA
            return raDifferenceInSecondsRet;
        }

        private void SyncScope(string testName, string canDoItName, SlewSyncType testType, double syncRa, double syncDec)
        {
            switch (testType)
            {
                case SlewSyncType.SyncToCoordinates: // SyncToCoordinates
                    LogCallToDriver(testName, "About to get Tracking property");
                    if (canSetTracking & !telescopeDevice.Tracking)
                    {
                        LogCallToDriver(testName, "About to set Tracking property to true");
                        telescopeDevice.Tracking = true;
                    }
                    LogCallToDriver(testName, $"About to call SyncToCoordinates method, RA: {syncRa.ToHMS()}, Declination: {syncDec.ToDMS()}");
                    TimeMethod(testName, () => telescopeDevice.SyncToCoordinates(syncRa, syncDec), TargetTime.Standard); // Sync to slightly different coordinates
                    LogDebug(testName, "Completed SyncToCoordinates");
                    break;

                case SlewSyncType.SyncToTarget: // SyncToTarget
                    LogCallToDriver(testName, "About to get Tracking property");
                    if (canSetTracking & !telescopeDevice.Tracking)
                    {
                        LogCallToDriver(testName, "About to set Tracking property to true");
                        telescopeDevice.Tracking = true;
                    }

                    try
                    {
                        LogCallToDriver(testName, $"About to set TargetRightAscension property to {syncRa.ToHMS()}");
                        telescopeDevice.TargetRightAscension = syncRa;
                        LogDebug(testName, "Completed Set TargetRightAscension");
                    }
                    catch (Exception ex)
                    {
                        HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"{canDoItName} is True but can't set TargetRightAscension");
                    }

                    try
                    {
                        LogCallToDriver(testName, $"About to set TargetDeclination property to {syncDec.ToDMS()}");
                        telescopeDevice.TargetDeclination = syncDec;
                        LogDebug(testName, "Completed Set TargetDeclination");
                    }
                    catch (Exception ex)
                    {
                        HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"{canDoItName} is True but can't set TargetDeclination");
                    }
                    LogCallToDriver(testName, "About to call SyncToTarget method");
                    TimeMethod(testName, () => telescopeDevice.SyncToTarget(), TargetTime.Standard); // Sync to slightly different coordinates
                    LogDebug(testName, "Completed SyncToTarget");
                    break;

                default:
                    LogError(testName, $"Conform:SyncTest: Unknown test type {testType}");
                    break;
            }
        }

        public void SlewScope(double pRa, double pDec, string pMsg)
        {
            if (canSetTracking)
            {
                LogCallToDriver("SlewScope", "About to set Tracking property to true");
                telescopeDevice.Tracking = true;
            }

            if (canSlewAsync)
            {
                LogDebug("SlewScope", $"Slewing asynchronously to {pMsg} {pRa.ToHMS()} {pDec.ToDMS()}");

                LogCallToDriver("SlewScope", $"About to call SlewToCoordinatesAsync method, RA: {pRa.ToHMS()}, Declination: {pDec.ToDMS()}");
                telescopeDevice.SlewToCoordinatesAsync(pRa, pDec);

                WaitForSlew(pMsg, $"Slewing asynchronously to {pMsg}");
            }
            else
            {
                if (canSlew)
                {
                    SetStatus($"Slewing synchronously to {pMsg}: {pRa.ToHMS()} {pDec.ToDMS()}");
                    LogDebug("SlewScope", $"Slewing synchronously to {pMsg} {pRa.ToHMS()} {pDec.ToDMS()}");
                    LogCallToDriver("SlewScope", $"About to call SlewToCoordinates method, RA: {pRa.ToHMS()}, Declination: {pDec.ToDMS()}");
                    telescopeDevice.SlewToCoordinates(pRa, pDec);
                }
                else
                {
                    LogInfo("SlewScope", "Unable to slew this scope as both CanSlew and CanSlewAsync are false, slew omitted");
                }
            }

            SetAction("");
        }

        private void WaitForSlew(string testName, string actionMessage)
        {
            Stopwatch sw = Stopwatch.StartNew();

            LogDebug(testName, $"Starting wait for slew.");

            LogCallToDriver(testName, "About to get Slewing property multiple times");
            WaitWhile(actionMessage, () => telescopeDevice.Slewing | (sw.Elapsed.TotalSeconds <= WAIT_FOR_SLEW_MINIMUM_DURATION), SLEEP_TIME, settings.TelescopeMaximumSlewTime);

            LogDebug(testName, $"Completed wait for slew.");
        }

        private double TelescopeRaFromHourAngle(string testName, double hourAngle)
        {
            double telescopeRa;

            // Handle the possibility that the mandatory SideealTime property has not been implemented
            if (canReadSiderealTime)
            {
                // Create a legal RA based on an offset from Sidereal time
                LogCallToDriver(testName, "About to get SiderealTime property");
                telescopeRa = telescopeDevice.SiderealTime - hourAngle;
                switch (telescopeRa)
                {
                    case var @case when @case < 0.0d: // Illegal if < 0 hours
                        {
                            telescopeRa += 24.0d;
                            break;
                        }

                    case var case1 when case1 >= 24.0d: // Illegal if > 24 hours
                        {
                            telescopeRa -= 24.0d;
                            break;
                        }
                }
            }
            else
            {
                telescopeRa = 0.0d - hourAngle;
            }

            return telescopeRa;
        }

        private double TelescopeRaFromSiderealTime(string testName, double raOffset)
        {
            double telescopeRa;
            double currentSiderealTime;

            // Handle the possibility that the mandatory SideealTime property has not been implemented
            if (canReadSiderealTime)
            {
                // Create a legal RA based on an offset from Sidereal time
                LogCallToDriver(testName, "About to get SiderealTime property");
                currentSiderealTime = telescopeDevice.SiderealTime;

                // Deal with possibility that sidereal time from the driver is bad
                switch (currentSiderealTime)
                {
                    case var @case when @case < 0.0d: // Illegal if < 0 hours
                        currentSiderealTime = 0d;
                        break;

                    case var case1 when case1 >= 24.0d: // Illegal if > 24 hours
                        currentSiderealTime = 0d;
                        break;
                }

                telescopeRa = currentSiderealTime + raOffset;
                switch (telescopeRa)
                {
                    case var case2 when case2 < 0.0d: // Illegal if < 0 hours
                        telescopeRa += 24.0d;
                        break;

                    case var case3 when case3 >= 24.0d: // Illegal if > 24 hours
                        telescopeRa -= 24.0d;
                        break;
                }

                LogDebug("TelescopeRaFromSiderealTime", $"Current sidereal time: {currentSiderealTime.ToHMS()}, Target RA: {telescopeRa.ToHMS()}");
            }
            else
            {
                telescopeRa = 0.0d + raOffset;
            }

            return telescopeRa;
        }
#if WINDOWS
        private void TestEarlyBinding(InterfaceType testType)
        {
            dynamic lITelescope;
            dynamic lDeviceObject = null;
            string lErrMsg;
            int lTryCount = 0;
            try
            {
                // Try early binding
                lITelescope = null;
                do
                {
                    lTryCount += 1;
                    try
                    {
                        LogCallToDriver("AccessChecks", "About to create driver object with CreateObject");
                        LogDebug("AccessChecks", "Creating late bound object for interface test");
                        Type driverType = Type.GetTypeFromProgID(settings.ComDevice.ProgId);
                        lDeviceObject = Activator.CreateInstance(driverType);
                        LogDebug("AccessChecks", "Created late bound object OK");
                        switch (testType)
                        {
                            case InterfaceType.TelescopeV2:
                                {
                                    //l_ITelescope = (ASCOM.Interface.ITelescope)l_DeviceObject;
                                    break;
                                }

                            case InterfaceType.TelescopeV3:
                                {
                                    lITelescope = (ITelescopeV3)lDeviceObject;
                                    break;
                                }

                            default:
                                {
                                    LogIssue("TestEarlyBinding", $"Unknown interface type: {testType}");
                                    break;
                                }
                        }

                        LogDebug("AccessChecks", $"Successfully created driver with interface {testType}");
                        try
                        {
                            LogCallToDriver("AccessChecks", "About to set Connected property true");
                            lITelescope.Connected = true;
                            LogInfo("AccessChecks", $"Device exposes interface {testType}");
                            LogCallToDriver("AccessChecks", "About to set Connected property false");
                            lITelescope.Connected = false;
                        }
                        catch (Exception)
                        {
                            LogInfo("AccessChecks", $"Device does not expose interface {testType}");
                            LogNewLine();
                        }
                    }
                    catch (Exception ex)
                    {
                        lErrMsg = ex.ToString();
                        LogDebug("AccessChecks", $"Exception: {ex.Message}");
                    }

                    if (lDeviceObject is null)
                        WaitFor(200);
                }
                while (lTryCount < 3 & lITelescope is not object); // Exit if created OK
                if (lITelescope is null)
                {
                    LogInfo("AccessChecks", $"Device does not expose interface {testType}");
                }
                else
                {
                    LogDebug("AccessChecks", $"Created telescope on attempt: {lTryCount}");
                }

                // Clean up
                try
                {
                    //DisposeAndReleaseObject("Telescope V1", l_ITelescope);
                }
                catch
                {
                }

                try
                {
                    //DisposeAndReleaseObject("Telescope V3", l_DeviceObject);
                }
                catch
                {
                }

                lDeviceObject = null;
                lITelescope = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                SetFullStatus("Early Binding", "Waiting for driver to be destroyed", "");
                WaitFor(1000, 100);
            }
            catch (Exception ex)
            {
                LogIssue("Telescope:TestEarlyBinding.EX1", ex.ToString());
            }
        }
#endif

        private static string FormatAzimuth(double az)
        {
            return az.ToDMS().PadLeft(11);
        }

        public static string TranslatePierSide(PointingState pPierSide, bool pLong)
        {
            string lPierSide;
            switch (pPierSide)
            {
                case PointingState.Normal:
                    {
                        if (pLong)
                        {
                            lPierSide = "pierEast";
                        }
                        else
                        {
                            lPierSide = "E";
                        }

                        break;
                    }

                case PointingState.ThroughThePole:
                    {
                        if (pLong)
                        {
                            lPierSide = "pierWest";
                        }
                        else
                        {
                            lPierSide = "W";
                        }

                        break;
                    }

                default:
                    {
                        if (pLong)
                        {
                            lPierSide = "pierUnknown";
                        }
                        else
                        {
                            lPierSide = "U";
                        }

                        break;
                    }
            }

            return lPierSide;
        }

        private enum Axis
        {
            Ra,
            Dec
        }

        /// <summary>
        /// Test whether the offset rate can be set successfully
        /// </summary>
        /// <param name="testName">Test name</param>
        /// <param name="description">Test description</param>
        /// <param name="axis">RA or Dec axis</param>
        /// <param name="testRate">Offset rate to be set</param>
        /// <param name="includeSlewiingTest">Flag indicating whether the Slewing property should be tested to confirm that it is false.</param>
        /// <returns>True if the rate could be set</returns>
        private bool TestRaDecRate(string testName, string description, Axis axis, double testRate, bool includeSlewiingTest)
        {
            // Initialise the outcome to false
            bool success = false;
            double offsetRate;

            try
            {
                // Tracking must be enabled for this test so make sure that it is enabled
                LogCallToDriver(testName,
                    $"{description} - About to get Tracking property");
                tracking = telescopeDevice.Tracking;

                // Test whether we are tracking and if not enable this if possible, abort the test if tracking cannot be set True
                if (!tracking)
                {
                    // We are not tracking so test whether Tracking can be set and abandon test if this is not possible.
                    // Tracking will be enabled by the SlewToCoordinates() method when called.
                    if (!canSetTracking)
                    {
                        LogIssue(testName, $"{description} - Test abandoned because {RateOffsetName(axis)} is true but Tracking is false and CanSetTracking is also false.");
                        return false;
                    }
                }

                // Include the slewing test if required
                if (includeSlewiingTest)
                {
                    // Test whether the mount is slewing, if so, abort the test
                    LogCallToDriver(testName, $"{description} - About to get Slewing property");
                    slewing = telescopeDevice.Slewing;

                    if (slewing) // Slewing is true
                    {
                        LogIssue(testName, $"{description} - Telescope.Slewing should be False at the start of this test but is returning True, test abandoned");
                        return false;
                    }
                }

                // Set the action name
                SetAction(testName);

                // Check that we can set the rate
                try
                {
                    LogCallToDriver(testName, $"{description} - About to set {RateOffsetName(axis)} property to {testRate}");

                    // Set the appropriate offset rate
                    if (axis == Axis.Ra)
                        TimeMethod($"{testName} {testRate:#0.000}", () => telescopeDevice.RightAscensionRate = testRate, TargetTime.Standard);
                    else
                        TimeMethod($"{testName} {testRate:#0.000}", () => telescopeDevice.DeclinationRate = testRate, TargetTime.Standard);

                    SetAction("Waiting for mount to settle");
                    WaitFor(1000); // Give a short wait to allow the mount to settle at the new rate

                    // If we get here the value was set OK, now check that the new rate is returned by RightAscensionRate Get and that Slewing is false
                    LogCallToDriver(testName, $"{description} - About to get {RateOffsetName(axis)} property");

                    // Get the appropriate offset rate
                    if (axis == Axis.Ra)
                        offsetRate = telescopeDevice.RightAscensionRate;
                    else
                        offsetRate = telescopeDevice.DeclinationRate;

                    LogCallToDriver(testName, $"{description} - About to get the Slewing property");
                    slewing = telescopeDevice.Slewing;

                    // Check the rate assignment outcome to within a small tolerance
                    bool ratesMatch = CompareDouble(offsetRate, testRate, 0.00001);

                    if (ratesMatch & !slewing) // Success - The rate has been correctly set and the mount reports that Slewing is False
                    {
                        LogOk(testName, $"{description} - successfully set rate to {offsetRate,7:+0.000;-0.000;+0.000}");
                        success = true;
                    }
                    else // Failed - Report what went wrong
                    {
                        if (slewing & ratesMatch) // The correct rate was returned but Slewing is True
                            LogIssue(testName, $"{RateOffsetName(axis)} was successfully set to {testRate} but Slewing is returning True, it should return False.");

                        if (slewing & !ratesMatch) // An incorrect rate was returned and Slewing is True
                            LogIssue(testName, $"{RateOffsetName(axis)} Read does not return {testRate} as set, instead it returns {offsetRate}. Slewing is also returning True, it should return False.");

                        if (!slewing & !ratesMatch) // An incorrect rate was returned and Slewing is False
                            LogIssue(testName, $"{RateOffsetName(axis)} Read does not return {testRate} as set, instead it returns {offsetRate}.");
                    }
                }
                catch (Exception ex)
                {
                    if (IsInvalidOperationException(testName, ex)) // We can't know what the valid range for this telescope is in advance so its possible that our test value will be rejected, if so just report this.
                    {
                        LogInfo(testName, $"Unable to set test rate {testRate}, it was rejected as an invalid value.");
                    }
                    else
                    {
                        HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"CanSet{RateOffsetName(axis)} is True");
                    }
                }
            }
            catch (Exception ex)
            {
                HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, "Tried to read Slewing property");
            }

            ClearStatus();
            return success;
        }

        /// <summary>
        /// Return the rate offset name, RightAscensionRate or DeclinationRate, corresponding to the selected axis
        /// </summary>
        /// <param name="axis">RA or Declination axis</param>
        /// <returns>RightAscensionRate or DeclinationRate string</returns>
        private static string RateOffsetName(Axis axis)
        {
            if (axis == Axis.Ra)
                return "RightAscensionRate";
            return "DeclinationRate";
        }

        /// <summary>
        /// Return the declination which provides the highest elevation that is less than 65 degrees at the given RA.
        /// </summary>
        /// <param name="testRa">RA to use when determining the optimum declination.</param>
        /// <returns>Declination in the range -80.0 to +80.0</returns>
        /// <remarks>The returned declination will correspond to an elevation in the range 0 to 65 degrees.</remarks>
        internal double GetTestDeclinationLessThan65(string testName, double testRa)
        {
            double testDeclination = 0.0;
            double testElevation = double.MinValue;

            // Find a test declination that yields an elevation that is as high as possible but under 65 degrees
            // Initialise transform with site parameters
            LogCallToDriver(testName, $"About to get SiteLatitude property");
            transform.SiteLatitude = telescopeDevice.SiteLatitude;
            LogCallToDriver(testName, $"About to get SiteLongitude property");
            transform.SiteLongitude = telescopeDevice.SiteLongitude;
            LogCallToDriver(testName, $"About to get SiteElevation property");
            transform.SiteElevation = telescopeDevice.SiteElevation;

            // Set remaining transform parameters
            transform.SitePressure = 1010.0;
            transform.Refraction = false;

            LogDebug("GetTestDeclinationLessThan65", $"TRANSFORM: Site: Latitude: {transform.SiteLatitude.ToDMS()}, Longitude: {transform.SiteLongitude.ToDMS()}, Elevation: {transform.SiteElevation}, Pressure: {transform.SitePressure}, Temperature: {transform.SiteTemperature}, Refraction: {transform.Refraction}");

            // Iterate from declination -85 to +85 in steps of 10 degrees
            Stopwatch sw = Stopwatch.StartNew();
            for (double declination = -85.0; declination < 90.0; declination += 10.0)
            {
                // Set transform's topocentric coordinates
                transform.SetTopocentric(testRa, declination);

                // Retrieve the corresponding elevation
                double elevation = transform.ElevationTopocentric;

                //LogDebug("GetTestDeclinationLessThan65", $"TRANSFORM: RA: {testRa.ToHMS()}, Declination: {declination.ToDMS()}, Azimuth: {transform.AzimuthTopocentric.ToDMS()}, Elevation: {elevation.ToDMS()}");

                // Update the test declination if the new elevation is less that 65 degrees and also greater than the current highest elevation 
                if ((elevation < 65.0) & (elevation > testElevation))
                {
                    testDeclination = declination;
                    testElevation = elevation;
                    //LogDebug("GetTestDeclinationLessThan65", $"Saved declination {testDeclination.ToDMS()} as having highest elevation: {testElevation.ToDMS()}");
                }
            }
            sw.Stop();
            LogDebug("GetTestDeclinationLessThan65", $"Test RightAscension: {testRa.ToHMS()}, Test Declination: {testDeclination.ToDMS()} at Elevation: {testElevation.ToDMS()} found in {sw.Elapsed.TotalMilliseconds:0.0}ms.");

            // Throw an exception if the test elevation is below 0 degrees
            if (testElevation < 0.0)
            {
                throw new ASCOM.InvalidOperationException($"The highest elevation available: {testElevation.ToDMS()} is below the horizon");
            }
            // Return the test declination
            return testDeclination;
        }

        /// <summary>
        /// Return the declination which provides the declination that is half way to the horizon at the give RA.
        /// </summary>
        /// <param name="testRa">RA to use when determining the optimum declination.</param>
        /// <returns>Declination in the range -80.0 to +80.0</returns>
        /// <remarks>The returned declination will correspond to an elevation in the range 0 to 65 degrees.</remarks>
        internal double GetTestDeclinationHalfwayToHorizon(string testName, double testRa)
        {
            double testDeclination = 0.0;
            double lowestDeclinationAboveHorizon;
            double siteLatitude;

            // Find a test declination that yields an elevation that is as high as possible but under 65 degrees
            // Initialise transform with site parameters
            LogCallToDriver(testName, $"About to get SiteLatitude property");
            siteLatitude = telescopeDevice.SiteLatitude;
            transform.SiteLatitude = siteLatitude;

            LogCallToDriver(testName, $"About to get SiteLongitude property");
            transform.SiteLongitude = telescopeDevice.SiteLongitude;

            LogCallToDriver(testName, $"About to get SiteElevation property");
            transform.SiteElevation = telescopeDevice.SiteElevation;

            // Set remaining transform parameters
            transform.SitePressure = 1010.0;
            transform.Refraction = false;

            LogDebug("GetTestDeclinationHalfwayToHorizon", $"TRANSFORM: Site: Latitude: {transform.SiteLatitude.ToDMS()}, Longitude: {transform.SiteLongitude.ToDMS()}, Elevation: {transform.SiteElevation}, Pressure: {transform.SitePressure}, Temperature: {transform.SiteTemperature}, Refraction: {transform.Refraction}");

            if (siteLatitude >= 0)
                lowestDeclinationAboveHorizon = double.MaxValue;
            else
                lowestDeclinationAboveHorizon = double.MinValue;

            // Iterate from declination -85 to +85 in steps of 10 degrees
            Stopwatch sw = Stopwatch.StartNew();
            for (double declination = -85.0; declination < 90.0; declination += 10.0)
            {
                // Set transform's topocentric coordinates
                transform.SetTopocentric(testRa, declination);

                // Retrieve the corresponding elevation
                double elevation = transform.ElevationTopocentric;

                //LogDebug("GetTestDeclinationHalfwayToHorizon", $"TRANSFORM: RA: {testRa.ToHMS()}, Declination: {declination.ToDMS()}, Azimuth: {transform.AzimuthTopocentric.ToDMS()}, Elevation: {elevation.ToDMS()}");

                // Update the lowest declination above the horizon if the sky position is above the horizon and also has a declination closer to the horizon than the current value.
                if (elevation > 0.0)
                {
                    if (siteLatitude >= 0.0) // Northern hemisphere
                    {
                        if (declination < lowestDeclinationAboveHorizon)
                        {
                            //LogDebug("GetTestDeclinationHalfwayToHorizon", $"Northern hemisphere - Saving this as the lowest declination above horizon:: {declination.ToDMS()}");
                            lowestDeclinationAboveHorizon = declination;
                        }
                    }
                    else // Southern hemisphere
                    {
                        if (declination > lowestDeclinationAboveHorizon)
                        {
                            //LogDebug("GetTestDeclinationHalfwayToHorizon", $"Southern hemisphere - Saving this as the lowest declination above horizon:: {declination.ToDMS()}");
                            lowestDeclinationAboveHorizon = declination;
                        }
                    }
                }
            }
            sw.Stop();

            // Check whether a valid declination has been found
            if ((lowestDeclinationAboveHorizon < 90.0) & (lowestDeclinationAboveHorizon > -90.0)) // A valid declination has been found
            {
                // Calculate the test declination as half way between the relevant celestial pole and the lowest declination above the horizon
                if (siteLatitude >= 0.0) // Northern hemisphere
                {
                    testDeclination = (lowestDeclinationAboveHorizon + 90.0) / 2.0;
                    LogDebug("GetTestDeclinationHalfwayToHorizon", $"Northern hemisphere or equator - Test RightAscension: {testRa.ToHMS()}, Lowest declination above horizon: {lowestDeclinationAboveHorizon.ToDMS()}, Test Declination: {testDeclination.ToDMS()} found in {sw.Elapsed.TotalMilliseconds:0.0}ms.");
                }
                else // Southern hemisphere
                {
                    testDeclination = (lowestDeclinationAboveHorizon - 90.0) / 2.0;
                    LogDebug("GetTestDeclinationHalfwayToHorizon", $"Southern hemisphere - Test RightAscension: {testRa.ToHMS()}, Lowest declination above horizon: {lowestDeclinationAboveHorizon.ToDMS()}, Test Declination: {testDeclination.ToDMS()} found in {sw.Elapsed.TotalMilliseconds:0.0}ms.");
                }

                // Return the test declination
                return testDeclination;
            }

            // A valid declination was not found so throw an error
            throw new ASCOM.InvalidOperationException($"GetTestDeclinationHalfwayToHorizon was unable to find a declination above the horizon at RA: {testRa.ToHMS()}");
        }

        internal void TestRAOffsetRates(string testName, double hourAngle)
        {
            // Test positive low value
            TestOffsetRate(testName, hourAngle, Axis.Ra, settings.TelescopeRateOffsetTestLowValue * ARC_SECONDS_TO_RA_SECONDS);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test negative low value
            TestOffsetRate(testName, hourAngle, Axis.Ra, -settings.TelescopeRateOffsetTestLowValue * ARC_SECONDS_TO_RA_SECONDS);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test positive high value
            TestOffsetRate(testName, hourAngle, Axis.Ra, settings.TelescopeRateOffsetTestHighValue * ARC_SECONDS_TO_RA_SECONDS);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test negative hight value
            TestOffsetRate(testName, hourAngle, Axis.Ra, -settings.TelescopeRateOffsetTestHighValue * ARC_SECONDS_TO_RA_SECONDS);
            if (cancellationToken.IsCancellationRequested)
                return;
        }

        internal void TestDeclinationOffsetRates(string testName, double hourAngle)
        {
            // Test positive low value
            TestOffsetRate(testName, hourAngle, Axis.Dec, settings.TelescopeRateOffsetTestLowValue);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test negative low value
            TestOffsetRate(testName, hourAngle, Axis.Dec, -settings.TelescopeRateOffsetTestLowValue);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test positive high value
            TestOffsetRate(testName, hourAngle, Axis.Dec, settings.TelescopeRateOffsetTestHighValue);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test negative hight value
            TestOffsetRate(testName, hourAngle, Axis.Dec, -settings.TelescopeRateOffsetTestHighValue);
            if (cancellationToken.IsCancellationRequested)
                return;
        }

        private void TestOffsetRate(string testName, double testHa, Axis testAxis, double expectedRate)
        {
            double expectedRaRate = 0.0, expectedDeclinationRate = 0.0;

            double testDuration = settings.TelescopeRateOffsetTestDuration; // Seconds
            double testDeclination;

            // Update the test name with the test hour angle
            testName = $"{testName} {testHa:+0.0;-0.0;+0.0}";

            // Set the expected rates
            switch (testAxis)
            {
                case Axis.Ra:
                    expectedRaRate = expectedRate;
                    break;

                case Axis.Dec:
                    expectedDeclinationRate = expectedRate;
                    break;
            }
            LogDebug(testName, $" ");
            LogDebug(testName, $"Starting test");

            // Create the test RA and declination
            double testRa = Utilities.ConditionRA(telescopeDevice.SiderealTime - testHa);
            try
            {
                testDeclination = GetTestDeclinationHalfwayToHorizon(testName, testRa); // Get the test declination for the test RA
            }
            catch (ASCOM.InvalidOperationException)
            {
                LogInfo(testName, $"Test omitted because it was not possible to find a declination above the horizon at {testRa.ToHMS()} - this is an expected condition at latitudes close to the equator.");
                return;
            }

            /// Slew the scope to the test position
            SlewScope(testRa, testDeclination, $"Hour angle {testHa:0.00}");

            if ((telescopeDevice.InterfaceVersion <= 2) | !CanReadSideOfPier(testName))
            {
                LogDebug(testName, $"Testing Primary rate: {expectedRaRate}, Secondary rate: {expectedDeclinationRate}, SideofPier: {PointingState.Unknown}");
            }
            else
            {
                LogCallToDriver(testName, "About to get SideOfPier property");
                LogDebug(testName, $"Testing Primary rate: {expectedRaRate}, Secondary rate: {expectedDeclinationRate}, SideofPier: {telescopeDevice.SideOfPier}");
            }

            // Get the telescope's start of test state
            LogCallToDriver(testName, "About to get RightAscension property");
            double priStart = telescopeDevice.RightAscension;

            LogCallToDriver(testName, "About to get Declination property");
            double secStart = telescopeDevice.Declination;

            // Set the rate offsets to the supplied test values
            LogCallToDriver(testName, "About to set RightAscensionRate property");
            telescopeDevice.RightAscensionRate = expectedRaRate * SIDEREAL_SECONDS_TO_SI_SECONDS;

            LogCallToDriver(testName, "About to set DeclinationRate property");
            telescopeDevice.DeclinationRate = expectedDeclinationRate;

            WaitFor(Convert.ToInt32(testDuration * 1000));

            // Get the telescope's end of test state
            LogCallToDriver(testName, "About to get RightAscension property");
            double priEnd = telescopeDevice.RightAscension;

            LogCallToDriver(testName, "About to get Declination property");
            double secEnd = telescopeDevice.Declination;

            // Restore previous state
            LogCallToDriver(testName, "About to set RightAscensionRate property");
            telescopeDevice.RightAscensionRate = 0.0;

            LogCallToDriver(testName, "About to set DeclinationRate property");
            telescopeDevice.DeclinationRate = 0.0;

            LogDebug(testName, $"Start      - : {priStart.ToHMS()}, {secStart.ToDMS()}");
            LogDebug(testName, $"Finish     - : {priEnd.ToHMS()}, {secEnd.ToDMS()}");
            LogDebug(testName, $"Difference - : {(priEnd - priStart).ToHMS()}, {(secEnd - secStart).ToDMS()}, {priEnd - priStart:N10}, {secEnd - secStart:N10}");

            // Condition results
            double actualPriRate = (priEnd - priStart) / testDuration; // Calculate offset rate in RA hours per SI second
            actualPriRate = actualPriRate * 60.0 * 60.0; // Convert rate in RA hours per SI second to RA seconds per SI second

            double actualSecRate = (secEnd - secStart) / testDuration * 60.0 * 60.0;

            LogDebug(testName, $"Actual primary rate: {actualPriRate}, Expected rate: {expectedRaRate}, Ratio: {actualPriRate / expectedRaRate}, Actual secondary rate: {actualSecRate}, Expected rate: {expectedDeclinationRate}, Ratio: {actualSecRate / expectedDeclinationRate}");

            TestDouble(testName, $"Rate: {expectedRate,7:+0.000;-0.000;+0.000} - RightAscensionRate", actualPriRate, expectedRaRate);
            TestDouble(testName, $"Rate: {expectedRate,7:+0.000;-0.000;+0.000} - DeclinationRate   ", actualSecRate, expectedDeclinationRate);

            LogDebug("", "");
        }

        private void TestDouble(string testName, string name, double actualValue, double expectedValue, double tolerance = 0.0)
        {
            // Tolerance 0 = 2%
            const double toleranceDefault = 0.05; // 5%

            if (tolerance == 0.0)
            {
                tolerance = toleranceDefault;
            }

            if (expectedValue == 0.0)
            {
                if (Math.Abs(actualValue - expectedValue) <= tolerance)
                {
                    LogOk(testName, $"{name} is within expected tolerance. Expected: {expectedValue,8:+0.0000;-0.0000;+0.0000}, Actual: {actualValue,8:+0.0000;-0.0000;+0.0000}, " +
                        $"Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / tolerance):N2}%.");
                }
                else
                {
                    LogIssue(testName, $"{name} is outside the expected tolerance. Expected: {expectedValue,8:+0.0000;-0.0000;+0.0000}, Actual: {actualValue,8:+0.0000;-0.0000;+0.0000}, " +
                        $"Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / tolerance),5:N2}%, Tolerance:{tolerance * 100:N0}%, Test duration: {settings.TelescopeRateOffsetTestDuration} seconds.");
                }
            }
            else
            {
                if (Math.Abs(Math.Abs(actualValue - expectedValue) / expectedValue) <= tolerance)
                {
                    LogOk(testName, $"{name} is within expected tolerance. Expected: {expectedValue,8:+0.0000;-0.0000;+0.0000}, Actual: {actualValue,8:+0.0000;-0.0000;+0.0000}, " +
                        $"Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / expectedValue):N2}%.");
                }
                else
                {
                    LogIssue(testName, $"{name} is outside the expected tolerance. Expected: {expectedValue,8:+0.0000;-0.0000;+0.0000}, Actual: {actualValue,8:+0.0000;-0.0000;+0.0000}, " +
                        $"Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / expectedValue),5:N2}%, Tolerance:{tolerance * 100:N0}%, Test duration: {settings.TelescopeRateOffsetTestDuration} seconds.");
                }
            }
        }

        /// <summary>
        /// Compare two numbers and return true if they match within a tolerance
        /// </summary>
        /// <param name="value1"></param>
        /// <param name="value2"></param>
        /// <returns></returns>
        private bool CompareDouble(double actualValue, double expectedValue, double tolerance = 0.0001)
        {
            if (expectedValue == 0.0)
            {
                if (Math.Abs(actualValue - expectedValue) <= tolerance)
                    return true;
                else
                    return false;
            }
            else
            {
                if (Math.Abs(Math.Abs(actualValue - expectedValue) / expectedValue) <= tolerance)
                    return true;
                else
                    return false;
            }
        }

        /// <summary>
        /// Slew to test coordinates and test all four guide directions
        /// </summary>
        /// <param name="ha">Hour angle at which to conduct test.</param>
        private void TestPulseGuide(double ha)
        {
            // Slew to the test HA/RA and a sensible Declination
            SlewToHa(ha);
            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            // Test each of the four guide directions
            TestPulseGuideDirection(ha, GuideDirection.North);
            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            TestPulseGuideDirection(ha, GuideDirection.South);
            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            TestPulseGuideDirection(ha, GuideDirection.East);
            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            TestPulseGuideDirection(ha, GuideDirection.West);
        }

        /// <summary>
        /// Ensure that the scope moves in the specified guide direction and for the expected distance.
        /// </summary>
        /// <param name="direction">Guide direction: North, South, East or West</param>
        private void TestPulseGuideDirection(double ha, GuideDirection direction)
        {
            const int PULSE_GUIDE_DURATION = 5; // Pulse guide test duration (Seconds)
            const int PULSE_GUIDE_TIMEOUT = PULSE_GUIDE_DURATION + 5; // Add 5 seconds of timeout allowance

            double expectedDecChange = guideRateDeclination * PULSE_GUIDE_DURATION; // Degrees
            double expectedRaChange = guideRateRightAscension * PULSE_GUIDE_DURATION / (15.0 * SIDEREAL_SECONDS_TO_SI_SECONDS); // Hours - 15 converts from 360 degrees = 24 hours, SIDEREAL_RATE converts from change in SI hours to change in sidereal hours

            // Handle direction reverse for south and west vs north and east
            if (direction == GuideDirection.South)
                expectedDecChange = -expectedDecChange;
            if (direction == GuideDirection.West)
                expectedRaChange = -expectedRaChange;

            double changeToleranceDec = settings.TelescopePulseGuideTolerance / 3600.0; // Degrees
            double changeToleranceRa = changeToleranceDec / (15.0 * SIDEREAL_SECONDS_TO_SI_SECONDS); // Sidereal hours

            string actionName = $"Pulse guiding {direction} at {ha}";
            string logName = $"PulseGuide {ha:+0.0;-0.0;+0.0} {direction}";

            LogDebug(logName, $"Test guiding direction: {direction}, RA change tolerance: {changeToleranceRa} ({changeToleranceRa.ToHMS()}, {changeToleranceRa.ToDMS()}), Declination change tolerance: {changeToleranceDec} ({changeToleranceDec.ToDMS()}, {changeToleranceDec.ToHMS()})");
            LogCallToDriver(logName, "About to get RightAscension property");
            double initialRaCoordinate = telescopeDevice.RightAscension;
            LogCallToDriver(logName, "About to get Declination property");
            double initialDecCoordinate = telescopeDevice.Declination;
            LogCallToDriver(logName, $"About to call PulseGuide. Direction: {direction}, Duration: {PULSE_GUIDE_DURATION * 1000}ms.");
            TimeMethod($"PulseGuide {direction} {PULSE_GUIDE_DURATION}s", () => telescopeDevice.PulseGuide(direction, PULSE_GUIDE_DURATION * 1000), TargetTime.Standard);

            WaitWhile($"Pulse guiding {direction} at HA {ha:+0.0;-0.0;+0.0}", () => telescopeDevice.IsPulseGuiding, SLEEP_TIME, PULSE_GUIDE_TIMEOUT);
            LogCallToDriver(logName, "About to get RightAscension property");
            double finalRaCoordinate = telescopeDevice.RightAscension;
            LogCallToDriver(logName, "About to get Declination property");
            double finalDecCoordinate = telescopeDevice.Declination;

            double raChange = finalRaCoordinate - initialRaCoordinate;
            double decChange = finalDecCoordinate - initialDecCoordinate;

            double raDifference = Math.Abs(raChange - expectedRaChange);
            double decDifference = Math.Abs(decChange - expectedDecChange);

            switch (direction)
            {
                case GuideDirection.North:
                    if ((decChange > 0.0) & (decDifference <= changeToleranceDec) & (Math.Abs(raChange) <= changeToleranceRa)) // Moved north by expected amount with no east-west movement
                        LogOk(logName, $"Moved north with no east-west movement  - Declination change (DMS): {decChange.ToDMS()},  Expected: {expectedDecChange.ToDMS()},  Difference: {decDifference.ToDMS()},  Test tolerance: {changeToleranceDec.ToDMS()}");
                    else
                    {
                        switch (decChange)
                        {
                            case < 0.0: //Moved south
                                LogIssue(logName, $"Moved south instead of north - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The declination axis did not move.");
                                break;

                            default: // Moved north
                                if (decDifference <= changeToleranceDec) // Moved north OK.
                                    LogOk(logName, $"Moved north - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                                else // Moved north but by wrong amount.
                                    LogIssue(logName, $"Moved north but outside test tolerance - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                                break;
                        }
                        if (Math.Abs(raChange) <= changeToleranceRa)
                            LogOk(logName, $"No significant east-west movement as expected. RA change (HMS): {raChange.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                        else
                            LogIssue(logName, $"East-West movement was outside test tolerance: RA change (HMS): {raChange.ToHMS()}, Expected: {0.0.ToHMS()}, Tolerance: {changeToleranceRa.ToHMS()}");
                    }
                    break;

                case GuideDirection.South:
                    if ((decChange < 0.0) & (decDifference <= changeToleranceDec) & (Math.Abs(raChange) <= changeToleranceRa)) // Moved south by expected amount with no east-west movement
                        LogOk(logName, $"Moved south with no east-west movement  - Declination change (DMS): {decChange.ToDMS()},  Expected: {expectedDecChange.ToDMS()},  Difference: {decDifference.ToDMS()},  Test tolerance: {changeToleranceDec.ToDMS()}");
                    else
                    {
                        switch (decChange)
                        {
                            case > 0.0: //Moved north
                                LogIssue(logName, $"Moved north instead of south - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The declination axis did not move.");
                                break;

                            default: // Moved south
                                if (decDifference <= changeToleranceDec)// Moved south OK
                                    LogOk(logName, $"Moved south - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                                else// Moved south but by wrong amount.
                                    LogIssue(logName, $"Moved south but outside test tolerance - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                                break;
                        }
                        if (Math.Abs(raChange) <= changeToleranceRa)
                            LogOk(logName, $"No significant east-west movement as expected. RA change (HMS): {raChange.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                        else
                            LogIssue(logName, $"East-West movement was outside test tolerance: RA change (HMS): {raChange.ToHMS()}, Expected: {0.0.ToHMS()}, Tolerance: {changeToleranceRa.ToHMS()}");
                    }
                    break;

                case GuideDirection.East:
                    if ((raChange > 0.0) & (raDifference <= changeToleranceRa) & (Math.Abs(decChange) <= changeToleranceDec)) // Moved east by expected amount with no north-south movement
                        LogOk(logName, $"Moved east with no north-south movement - RA change (HMS):          {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                    else
                    {
                        switch (raChange)
                        {
                            case < 0.0: //Moved west
                                LogIssue(logName, $"Moved west instead of east - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The RA axis did not move.");
                                break;

                            default: // Moved east
                                if (raDifference <= changeToleranceRa)  // Moved east OK
                                    LogOk(logName, $"Moved east - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                                else // Moved east but by wrong amount.
                                    LogIssue(logName, $"Moved east but outside test tolerance - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                                break;
                        }
                        if (Math.Abs(decChange) <= changeToleranceDec)
                            LogOk(logName, $"No significant north-south movement as expected. Declination change (DMS): {decChange.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                        else
                            LogIssue(logName, $"North-south movement was outside test tolerance: Declination change (DMS): {decChange.ToDMS()}, Expected: {0.0.ToDMS()}, Tolerance: {changeToleranceDec.ToDMS()}");
                    }
                    break;

                case GuideDirection.West:
                    if ((raChange < 0.0) & (raDifference <= changeToleranceRa) & (Math.Abs(decChange) <= changeToleranceDec)) // Moved west by expected amount with no north-south movement
                        LogOk(logName, $"Moved west with no north-south movement - RA change (HMS):          {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                    else
                    {
                        switch (raChange)
                        {
                            case > 0.0: //Moved east
                                LogIssue(logName, $"Moved east instead of west - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The RA axis did not move.");
                                break;

                            default: // Moved west
                                if (raDifference <= changeToleranceRa) // Moved west OK
                                    LogOk(logName, $"Moved west - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                                else// Moved west but by wrong amount.
                                    LogIssue(logName, $"Moved west but outside test tolerance - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}");
                                break;
                        }
                        if (Math.Abs(decChange) <= changeToleranceDec)
                            LogOk(logName, $"No significant north-south movement as expected. Declination change (DMS): {decChange.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}");
                        else
                            LogIssue(logName, $"North-south movement was outside test tolerance: Declination change (DMS): {decChange.ToDMS()}, Expected: {0.0.ToDMS()}, Tolerance: {changeToleranceDec.ToDMS()}");
                    }
                    break;

                default:
                    break;
            }
        }

        /// <summary>
        /// Slew to the target hour angle and to a declination that results in the highest elevation under 65 degrees.
        /// </summary>
        /// <param name="targetHa">Target hour angle (hours)</param>
        private void SlewToHa(double targetHa)
        {
            if (canSetTracking)
            {
                LogCallToDriver("SlewToHA", "About to set Tracking property to true");
                telescopeDevice.Tracking = true; // Enable tracking for these tests
            }

            // Calculate the target RA
            LogCallToDriver("SlewToHa", "About to get SiderealTime property");
            double targetRa = Utilities.ConditionRA(telescopeDevice.SiderealTime - targetHa);

            // Calculate the target declination
            double targetDeclination = GetTestDeclinationLessThan65("SlewToHa", targetRa);
            LogDebug("SlewToHa", $"Slewing to HA: {targetHa.ToHMS()} (RA: {targetRa.ToHMS()}), Dec: {targetDeclination.ToDMS()}");

            // Slew to the target coordinates
            LogCallToDriver("SlewToHa", $"About to call SlewToCoordinatesAsync. RA: {targetRa.ToHMS()}, Declination: {targetDeclination.ToDMS()}");

            LogCallToDriver("SlewToHa",
                $"About to call SlewToCoordinatesAsync method, RA: {targetRa.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
            telescopeDevice.SlewToCoordinatesAsync(targetRa, targetDeclination);

            WaitForSlew("SlewToHa", $"Slewing to HA {targetHa:+0.0;-0.0;+0.0}");

            // Report the outcome of the slew
            if ((telescopeDevice.InterfaceVersion <= 2) | !CanReadSideOfPier("SlewToHa"))
            {
                LogDebug("SlewToHa", $"Slewed to RA:  {telescopeDevice.RightAscension.ToHMS()}, Dec: {telescopeDevice.Declination.ToDMS()}, Pointing state: {PointingState.Unknown}");
            }
            else
            {
                LogDebug("SlewToHa", $"Slewed to RA:  {telescopeDevice.RightAscension.ToHMS()}, Dec: {telescopeDevice.Declination.ToDMS()}, Pointing state: {telescopeDevice.SideOfPier}");
            }
        }

        /// <summary>
        /// Determine whether the SideofPier property can be successfully read.
        /// </summary>
        /// <param name="testName">Name of the current test</param>
        /// <returns>True if SideOfPier can be read successfully, otherwise returns False.</returns>
        /// <remarks>If the test has been run before, the previous answer is returned without calling SideOfPier again.</remarks>
        private bool CanReadSideOfPier(string testName)
        {
            // test whether a value has already been determined
            if (!canReadSideOfPier.HasValue) // This is the first time the function has been called so test whther reading is possible
            {
                try
                {
                    LogCallToDriver(testName, "About to get SideOfPier");
                    PointingState pointingState = telescopeDevice.SideOfPier;

                    // If we get here the read was successful so flag this
                    canReadSideOfPier = true;
                }
                catch
                {
                    // If we get here the read was unsuccessful so flag this.
                    // We ignore exceptions because we are only interested in whenther SideOfPier can be read successfully
                    canReadSideOfPier = false;
                }
            }

            // Return the outcome
            return canReadSideOfPier.Value;
        }

        private bool ValidateSlewing(string test, bool expectedState)
        {
            try
            {
                LogCallToDriver(test, "About to call Slewing property");
                bool slewing = telescopeDevice.Slewing;

                if (slewing == expectedState)
                {
                    return true; // Got expected outcome so no action
                }
                else
                {
                    LogIssue(test, $"Slewing did not have the expected state: {expectedState}, it was: {slewing}.");
                }
            }
            catch (Exception ex)
            {
                LogIssue(test, $"Unexpected exception from Slewing: {ex.Message}");
                LogDebug(test, ex.ToString());
            }
            return false;
        }

        /// <summary>
        /// Wait up to the configured time for Slewing to become false.
        /// </summary>
        /// <param name="testName">Test name</param>
        /// <returns>TRUE if wait completed successfully within the timeout, otherwise FALSE</returns>
        private bool WaitForSlewingTobecomeFalse(string testName)
        {
            try
            {
                // Wait for the mount to report that it is no longer slewing or for the wait to time out
                LogCallToDriver(testName, $"About to call Slewing repeatedly...");
                WaitWhile("Waiting for slew to stop", () => telescopeDevice.Slewing == true, 500, settings.TelescopeTimeForSlewingToBecomeFalse);

                // Indicate that the wait was successful
                return true;
            }
            catch (TimeoutException)
            {
                LogIssue(testName, $"Timed out after {settings.TelescopeTimeForSlewingToBecomeFalse} seconds while waiting for Slewing to become FALSE.");
            }
            catch (Exception ex)
            {
                LogIssue(testName, $"The mount reported an exception while waiting for Slewing to become false: {ex.Message}");
                LogDebug(testName, ex.ToString());
            }

            // Indicate that the wait failed
            return false;
        }

        private void AbortSlew(string testName)
        {
            LogCallToDriver(testName, "About to call AbortSlew method");
            telescopeDevice.AbortSlew();
        }

        #endregion

    }
}

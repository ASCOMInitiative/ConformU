using ASCOM;
using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.Tools;
using System;
using System.Collections;
using System.Collections.Generic;
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
        internal const string TELTEST_ABORT_SLEW = "AbortSlew";
        internal const string TELTEST_AXIS_RATE = "AxisRate";
        internal const string TELTEST_CAN_MOVE_AXIS = "CanMoveAxis";
        internal const string TELTEST_COMMANDXXX = "CommandXXX";
        internal const string TELTEST_DESTINATION_SIDE_OF_PIER = "DestinationSideOfPier";
        internal const string TELTEST_FIND_HOME = "FindHome";
        internal const string TELTEST_MOVE_AXIS = "MoveAxis";
        internal const string TELTEST_PARK_UNPARK = "Park/Unpark";
        internal const string TELTEST_PULSE_GUIDE = "PulseGuide";
        internal const string TELTEST_SLEW_TO_ALTAZ = "SlewToAltAz";
        internal const string TELTEST_SLEW_TO_ALTAZ_ASYNC = "SlewToAltAzAsync";
        internal const string TELTEST_SLEW_TO_TARGET = "SlewToTarget";
        internal const string TELTEST_SLEW_TO_TARGET_ASYNC = "SlewToTargetAsync";
        internal const string TELTEST_SYNC_TO_ALTAZ = "SyncToAltAz";
        internal const string TELTEST_SLEW_TO_COORDINATES = "SlewToCoordinates";
        internal const string TELTEST_SLEW_TO_COORDINATES_ASYNC = "SlewToCoordinatesAsync";
        internal const string TELTEST_SYNC_TO_COORDINATES = "SyncToCoordinates";
        internal const string TELTEST_SYNC_TO_TARGET = "SyncToTarget";

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

        private bool canFindHome, canPark, canPulseGuide, canSetDeclinationRate, canSetGuideRates, canSetPark, canSetPierside, canSetRightAscensionRate;
        private bool canSetTracking, canSlew, canSlewAltAz, canSlewAltAzAsync, canSlewAsync, canSync, canSyncAltAz, canUnpark;
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
        private double siderealTimeASCOM;
        private DateTime startTime, endTime;
        private bool? canReadSideOfPier = null; // Start out in the "not read" state
        private double targetAltitude, targetAzimuth;
        private bool canReadAltitide, canReadAzimuth, canReadSiderealTime;

        private double operationInitiationTime = 1.0; // Time within which an operation initiation must complete (seconds)

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
            CanSetCCDTemperature = 20,
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
            tstAxisrates = 1,
            tstCanMoveAxisPrimary = 2,
            tstCanMoveAxisSecondary = 3,
            tstCanMoveAxisTertiary = 4
        }

        private enum ParkedExceptionType
        {
            tstPExcepAbortSlew = 1,
            tstPExcepFindHome = 2,
            tstPExcepMoveAxisPrimary = 3,
            tstPExcepMoveAxisSecondary = 4,
            tstPExcepMoveAxisTertiary = 5,
            tstPExcepSlewToCoordinates = 6,
            tstPExcepSlewToCoordinatesAsync = 7,
            tstPExcepSlewToTarget = 8,
            tstPExcepSlewToTargetAsync = 9,
            tstPExcepSyncToCoordinates = 10,
            tstPExcepSyncToTarget = 11,
            tstPExcepPulseGuide = 12
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
            tstPerfAltitude = 0,
            tstPerfAtHome = 1,
            tstPerfAtPark = 2,
            tstPerfAzimuth = 3,
            tstPerfDeclination = 4,
            tstPerfIsPulseGuiding = 5,
            tstPerfRightAscension = 6,
            tstPerfSideOfPier = 7,
            tstPerfSiderealTime = 8,
            tstPerfSlewing = 9,
            tstPerfUTCDate = 10
        }

        public enum FlipTestType
        {
            DestinationSideOfPier,
            SideOfPier
        }

        private enum InterfaceType
        {
            ITelescope,
            ITelescopeV2,
            ITelescopeV3
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
            LogDebug("Dispose", "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
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

        #region Code
        public override void CheckCommonMethods()
        {
            CheckCommonMethods(telescopeDevice, DeviceTypes.Telescope);
        }

        public new void CheckInitialise()
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
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040401;
                            g_ExInvalidValue2 = (int)0x80040402;
                            g_ExInvalidValue3 = (int)0x80040405;
                            g_ExInvalidValue4 = (int)0x80040402;
                            g_ExInvalidValue5 = (int)0x80040402;
                            g_ExInvalidValue6 = (int)0x80040402;
                            g_ExNotSet1 = (int)0x80040403;
                            break;
                        }

                    case "ASCOM.MI250SA.Telescope":
                    case "Celestron.Telescope":
                    case "ASCOM.MI250.Telescope":
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040401;
                            g_ExInvalidValue2 = (int)0x80040402;
                            g_ExInvalidValue3 = (int)0x80040402;
                            g_ExInvalidValue4 = (int)0x80040402;
                            g_ExInvalidValue5 = (int)0x80040402;
                            g_ExInvalidValue6 = (int)0x80040402;
                            g_ExNotSet1 = (int)0x80040403;
                            break;
                        }

                    case "TemmaLite.Telescope":
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040410;
                            g_ExInvalidValue2 = (int)0x80040418;
                            g_ExInvalidValue3 = (int)0x80040418;
                            g_ExInvalidValue4 = (int)0x80040418;
                            g_ExInvalidValue5 = (int)0x80040418;
                            g_ExInvalidValue6 = (int)0x80040418;
                            g_ExNotSet1 = (int)0x80040417;
                            break;
                        }

                    case "Gemini.Telescope":
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040410;
                            g_ExInvalidValue2 = (int)0x80040418;
                            g_ExInvalidValue3 = (int)0x80040419;
                            g_ExInvalidValue4 = (int)0x80040420;
                            g_ExInvalidValue5 = (int)0x80040420;
                            g_ExInvalidValue6 = (int)0x80040420;
                            g_ExNotSet1 = (int)0x80040417;
                            break;
                        }

                    case "POTH.Telescope":
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040405;
                            g_ExInvalidValue2 = (int)0x80040406;
                            g_ExInvalidValue3 = (int)0x80040406;
                            g_ExInvalidValue4 = (int)0x80040406;
                            g_ExInvalidValue5 = (int)0x80040406;
                            g_ExInvalidValue6 = (int)0x80040406;
                            g_ExNotSet1 = (int)0x80040403;
                            break;
                        }

                    case "ServoCAT.Telescope":
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = ErrorCodes.InvalidValue;
                            g_ExInvalidValue2 = (int)0x80040405;
                            g_ExInvalidValue3 = (int)0x80040405;
                            g_ExInvalidValue4 = (int)0x80040405;
                            g_ExInvalidValue5 = (int)0x80040405;
                            g_ExInvalidValue6 = (int)0x80040405;
                            g_ExNotSet1 = (int)0x80040403;
                            g_ExNotSet2 = (int)0x80040404; // I'm using the simulator values as the defaults since it is the reference platform
                            break;
                        }

                    default:
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = ErrorCodes.InvalidValue;
                            g_ExInvalidValue2 = (int)0x80040405;
                            g_ExInvalidValue3 = (int)0x80040405;
                            g_ExInvalidValue4 = (int)0x80040405;
                            g_ExInvalidValue5 = (int)0x80040405;
                            g_ExInvalidValue6 = (int)0x80040405;
                            g_ExNotSet1 = (int)0x80040403;
                            break;
                        }
                }
            }

            CheckInitialise();
        }

        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        LogInfo("CreateDevice", $"Creating Alpaca device: Access service: {settings.AlpacaConfiguration.AccessServiceType}, IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber} " +
                            $"({(settings.AlpacaConfiguration.TrustUserGeneratedSslCertificates ? "A" : "Not a")}ccepting user generated SSL certificates)");
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
                baseClassDevice = telescopeDevice; // Assign the driver to the base class
                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);
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

        public override bool Connected
        {
            get
            {
                LogCallToDriver("ConformanceCheck", "About to get Connected");
                return telescopeDevice.Connected;
            }
            set
            {
                LogCallToDriver("ConformanceCheck", "About to set Connected");
                SetTest("Connected");
                SetAction("Waiting for Connected to become 'true'");
                telescopeDevice.Connected = value;
                ResetTestActionStatus();

                // Make sure that the value set is reflected in Connected GET
                bool connectedState = Connected;
                if (connectedState != value)
                {
                    throw new ASCOM.InvalidOperationException($"Connected was set to {value} but Connected Get returned {connectedState}.");
                }
            }
        }

        public override void PreRunCheck()
        {
            // Get into a consistent state
            if (interfaceVersion > 1)
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
                                WaitWhile("Waiting for scope to unpark", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

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
                LogInfo("Mount Safety", "Skipping AtPark test as this method is not supported in interface V" + interfaceVersion);
                try
                {
                    if (canUnpark)
                    {
                        LogCallToDriver("Mount Safety", "About to call Unpark method");
                        telescopeDevice.Unpark();
                        LogCallToDriver("Mount Safety", "About to get AtPark property repeatedly");
                        WaitWhile("Waiting for scope to unpark", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                        LogOK("Mount Safety", "Scope has been unparked for testing");
                    }
                    else
                    {
                        LogOK("Mount Safety", "Scope reports that it cannot unpark, unparking skipped");
                    }
                }
                catch (Exception ex)
                {
                    LogIssue("Mount Safety", "Driver threw an exception while unparking: " + ex.Message);
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
                    LogInfo("TimeCheck", $"PC UTCDate:    " + DateTime.UtcNow.ToString("dd-MMM-yyyy HH:mm:ss.fff"));
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
            try
            {
                if ((telescopeDevice.AlignmentMode != AlignmentMode.GermanPolar) & canSetPierside)
                    LogIssue("CanSetPierSide", "AlignmentMode is not GermanPolar but CanSetPierSide is true - contrary to ASCOM specification");
            }
            catch (Exception)
            {
            }

            TelescopeCanTest(CanType.CanSetRightAscensionRate, "CanSetRightAscensionRate");
            TelescopeCanTest(CanType.CanSetTracking, "CanSetTracking");
            TelescopeCanTest(CanType.CanSlew, "CanSlew");
            TelescopeCanTest(CanType.CanSlewAltAz, "CanSlewltAz");
            TelescopeCanTest(CanType.CanSlewAltAzAsync, "CanSlewAltAzAsync");
            TelescopeCanTest(CanType.CanSlewAsync, "CanSlewAsync");
            TelescopeCanTest(CanType.CanSync, "CanSync");
            TelescopeCanTest(CanType.CanSyncAltAz, "CanSyncAltAz");
            TelescopeCanTest(CanType.CanUnPark, "CanUnPark");
            if (canUnpark & !canPark)
                LogIssue("CanUnPark", "CanUnPark is true but CanPark is false - this does not comply with ASCOM specification");
        }

        public override void CheckProperties()
        {
            bool l_OriginalTrackingState;
            DriveRate l_DriveRate;
            double l_TimeDifference;
            ITrackingRates l_TrackingRates = null;
            dynamic l_TrackingRate;

            // Test TargetDeclination and TargetRightAscension first because these tests will fail if the telescope has been slewed previously.
            // Slews can happen in the extended guide rate tests for example.
            // The test is made here but reported later so that the properties are tested in mostly alphabetical order.

            // TargetDeclination Read - Optional
            Exception targetDeclinationReadException = null;
            try // First read should fail!
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetDeclination Read", "About to get TargetDeclination property");
                targetDeclination = telescopeDevice.TargetDeclination;
            }
            catch (Exception ex)
            {
                targetDeclinationReadException = ex;
            }

            // TargetRightAscension Read - Optional
            Exception targetRightAscensionReadException = null;
            try // First read should fail!
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetRightAscension Read", "About to get TargetRightAscension property");
                targetRightAscension = telescopeDevice.TargetRightAscension;
            }
            catch (Exception ex)
            {
                targetRightAscensionReadException = ex;
            }

            // AlignmentMode - Optional
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("AlignmentMode", "About to get AlignmentMode property");
                alignmentMode = (AlignmentMode)telescopeDevice.AlignmentMode;
                LogOK("AlignmentMode", alignmentMode.ToString());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("Altitude", "About to get Altitude property");
                altitude = telescopeDevice.Altitude;
                canReadAltitide = true; // Read successfully
                switch (altitude)
                {
                    case var @case when @case < 0.0d:
                        {
                            LogIssue("Altitude", "Altitude is <0.0 degrees: " + FormatAltitude(altitude).Trim());
                            break;
                        }

                    case var case1 when case1 > 90.0000001d:
                        {
                            LogIssue("Altitude", "Altitude is >90.0 degrees: " + FormatAltitude(altitude).Trim());
                            break;
                        }

                    default:
                        {
                            LogOK("Altitude", FormatAltitude(altitude).Trim());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ApertureArea", "About to get ApertureArea property");
                apertureArea = telescopeDevice.ApertureArea;
                switch (apertureArea)
                {
                    case var case2 when case2 < 0d:
                        {
                            LogIssue("ApertureArea", "ApertureArea is < 0.0 : " + apertureArea.ToString());
                            break;
                        }

                    case 0.0d:
                        {
                            LogInfo("ApertureArea", "ApertureArea is 0.0");
                            break;
                        }

                    default:
                        {
                            LogOK("ApertureArea", apertureArea.ToString());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ApertureDiameter", "About to get ApertureDiameter property");
                apertureDiameter = telescopeDevice.ApertureDiameter;
                switch (apertureDiameter)
                {
                    case var case3 when case3 < 0.0d:
                        {
                            LogIssue("ApertureDiameter", "ApertureDiameter is < 0.0 : " + apertureDiameter.ToString());
                            break;
                        }

                    case 0.0d:
                        {
                            LogInfo("ApertureDiameter", "ApertureDiameter is 0.0");
                            break;
                        }

                    default:
                        {
                            LogOK("ApertureDiameter", apertureDiameter.ToString());
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
            if (interfaceVersion > 1)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("AtHome", "About to get AtHome property");
                    atHome = telescopeDevice.AtHome;
                    LogOK("AtHome", atHome.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("AtHome", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("AtHome", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // AtPark - Required
            if (interfaceVersion > 1)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("AtPark", "About to get AtPark property");
                    atPark = telescopeDevice.AtPark;
                    LogOK("AtPark", atPark.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("AtPark", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("AtPark", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Azimuth - Optional
            try
            {
                canReadAzimuth = false;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("Azimuth", "About to get Azimuth property");
                azimuth = telescopeDevice.Azimuth;
                canReadAzimuth = true; // Read successfully
                switch (azimuth)
                {
                    case var case4 when case4 < 0.0d:
                        {
                            LogIssue("Azimuth", "Azimuth is <0.0 degrees: " + FormatAzimuth(azimuth).Trim());
                            break;
                        }

                    case var case5 when case5 > 360.0000000001d:
                        {
                            LogIssue("Azimuth", "Azimuth is >360.0 degrees: " + FormatAzimuth(azimuth).Trim());
                            break;
                        }

                    default:
                        {
                            LogOK("Azimuth", FormatAzimuth(azimuth).Trim());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("Declination", "About to get Declination property");
                declination = telescopeDevice.Declination;
                switch (declination)
                {
                    case var case6 when case6 < -90.0d:
                    case var case7 when case7 > 90.0d:
                        {
                            LogIssue("Declination", "Declination is <-90 or >90 degrees: " + FormatDec(declination).Trim());
                            break;
                        }

                    default:
                        {
                            LogOK("Declination", FormatDec(declination).Trim());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("DeclinationRate Read", "About to get DeclinationRate property");
                declinationRate = telescopeDevice.DeclinationRate;
                // Read has been successful
                if (canSetDeclinationRate) // Any value is acceptable
                {
                    switch (declinationRate)
                    {
                        case var case8 when case8 >= 0.0d:
                            {
                                LogOK("DeclinationRate Read", declinationRate.ToString("0.00"));
                                break;
                            }

                        default:
                            {
                                LogIssue("DeclinationRate Read", "Negative DeclinatioRate: " + declinationRate.ToString("0.00"));
                                break;
                            }
                    }
                }
                else // Only zero is acceptable
                {
                    switch (declinationRate)
                    {
                        case 0.0d:
                            {
                                LogOK("DeclinationRate Read", declinationRate.ToString("0.00"));
                                break;
                            }

                        default:
                            {
                                LogIssue("DeclinationRate Read", "DeclinationRate is non zero when CanSetDeclinationRate is False " + declinationRate.ToString("0.00"));
                                break;
                            }
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
            if (interfaceVersion > 1)
            {
                if (canSetDeclinationRate) // Any value is acceptable
                {
                    SetTest("DeclinationRate Write");
                    if (TestRADecRate("DeclinationRate Write", "Set rate to 0.0", Axis.Dec, 0.0d, true))
                    {
                        TestRADecRate("DeclinationRate Write", "Set rate to 1.5", Axis.Dec, 1.5d, true);
                        TestRADecRate("DeclinationRate Write", "Set rate to -1.5", Axis.Dec, -1.5d, true);
                        TestRADecRate("DeclinationRate Write", "Set rate to 7.5", Axis.Dec, 7.5d, true);
                        TestRADecRate("DeclinationRate Write", "Set rate to -7.5", Axis.Dec, -7.5d, true);
                        TestRADecRate("DeclinationRate Write", "Reset rate to 0.0", Axis.Dec, 0.0d, false); // Reset the rate to zero, skipping the slewing test
                    }
                }
                else // Should generate an error
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("DeclinationRate Write", "About to set DeclinationRate property to 0.0");
                        telescopeDevice.DeclinationRate = 0.0d; // Set to a harmless value
                        LogIssue("DeclinationRate", "CanSetDeclinationRate is False but setting DeclinationRate did not generate an error");
                    }
                    catch (Exception ex)
                    {
                        HandleException("DeclinationRate Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetDeclinationRate is False");
                    }
                }
            }
            else
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("DeclinationRate Write", "About to set DeclinationRate property to 0.0");
                    telescopeDevice.DeclinationRate = 0.0d; // Set to a harmless value
                    LogOK("DeclinationRate Write", declinationRate.ToString("0.00"));
                }
                catch (Exception ex)
                {
                    HandleException("DeclinationRate Write", MemberType.Property, Required.Optional, ex, "");
                }
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // DoesRefraction Read - Optional
            if (interfaceVersion > 1)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("DoesRefraction Read", "About to DoesRefraction get property");
                    doesRefraction = telescopeDevice.DoesRefraction;
                    LogOK("DoesRefraction Read", doesRefraction.ToString());
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
                LogInfo("DoesRefraction Read", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // DoesRefraction Write - Optional
            if (interfaceVersion > 1)
            {
                if (doesRefraction) // Try opposite value
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("DoesRefraction Write", "About to set DoesRefraction property false");
                        telescopeDevice.DoesRefraction = false;
                        LogOK("DoesRefraction Write", "Can set DoesRefraction to False");
                    }
                    catch (Exception ex)
                    {
                        HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "");
                    }
                }
                else // Try other opposite value
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("DoesRefraction Write", "About to set DoesRefraction property true");
                        telescopeDevice.DoesRefraction = true;
                        LogOK("DoesRefraction Write", "Can set DoesRefraction to True");
                    }
                    catch (Exception ex)
                    {
                        HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "");
                    }
                }
            }
            else
            {
                LogInfo("DoesRefraction Write", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // EquatorialSystem - Required
            if (interfaceVersion > 1)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("EquatorialSystem", "About to get EquatorialSystem property");
                    equatorialSystem = (EquatorialCoordinateType)telescopeDevice.EquatorialSystem;
                    LogOK("EquatorialSystem", equatorialSystem.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("EquatorialSystem", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("EquatorialSystem", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // FocalLength - Optional
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("FocalLength", "About to get FocalLength property");
                focalLength = telescopeDevice.FocalLength;
                switch (focalLength)
                {
                    case var case9 when case9 < 0.0d:
                        {
                            LogIssue("FocalLength", "FocalLength is <0.0 : " + focalLength.ToString());
                            break;
                        }

                    case 0.0d:
                        {
                            LogInfo("FocalLength", "FocalLength is 0.0");
                            break;
                        }

                    default:
                        {
                            LogOK("FocalLength", focalLength.ToString());
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
            if (interfaceVersion > 1)
            {
                if (canSetGuideRates) // Can set guide rates so read and write are mandatory
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateDeclination Read", "About to get GuideRateDeclination property");
                        guideRateDeclination = telescopeDevice.GuideRateDeclination; // Read guiderateDEC

                        if (guideRateDeclination >= 0.0)
                        {
                            LogOK("GuideRateDeclination Read", $"{guideRateDeclination:0.000000} ({guideRateDeclination.ToDMS()})");

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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateDeclination Read", "About to set GuideRateDeclination property to " + guideRateDeclination);
                        telescopeDevice.GuideRateDeclination = guideRateDeclination;
                        LogOK("GuideRateDeclination Write", "Can write Declination Guide Rate OK");
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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateDeclination Read", "About to get GuideRateDeclination property");
                        guideRateDeclination = telescopeDevice.GuideRateDeclination;
                        switch (guideRateDeclination)
                        {
                            case var case11 when case11 < 0.0d:
                                {
                                    LogIssue("GuideRateDeclination Read", "GuideRateDeclination is < 0.0 " + guideRateDeclination.ToString("0.00"));
                                    break;
                                }

                            default:
                                {
                                    LogOK("GuideRateDeclination Read", guideRateDeclination.ToString("0.00"));
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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateDeclination Write", "About to set GuideRateDeclination property to " + guideRateDeclination);
                        telescopeDevice.GuideRateDeclination = guideRateDeclination;
                        LogIssue("GuideRateDeclination Write", "CanSetGuideRates is false but no exception generated; value returned: " + guideRateDeclination.ToString("0.00"));
                    }
                    catch (Exception ex) // Some other error so OK
                    {
                        HandleException("GuideRateDeclination Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False");
                    }
                }
            }
            else
            {
                LogInfo("GuideRateDeclination", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // GuideRateRightAscension - Optional
            if (interfaceVersion > 1)
            {
                if (canSetGuideRates)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateRightAscension Read", "About to get GuideRateRightAscension property");
                        guideRateRightAscension = telescopeDevice.GuideRateRightAscension; // Read guide rate RA

                        if (guideRateRightAscension >= 0.0)
                        {
                            LogOK("GuideRateRightAscension Read", $"{guideRateRightAscension:0.000000} ({guideRateRightAscension.ToDMS()})");
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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateRightAscension Read", "About to set GuideRateRightAscension property to " + guideRateRightAscension);
                        telescopeDevice.GuideRateRightAscension = guideRateRightAscension;
                        LogOK("GuideRateRightAscension Write", "Can set RightAscension Guide OK");
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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateRightAscension Read", "About to get GuideRateRightAscension property");
                        guideRateRightAscension = telescopeDevice.GuideRateRightAscension; // Read guiderateRA
                        switch (guideRateDeclination)
                        {
                            case var case13 when case13 < 0.0d:
                                {
                                    LogIssue("GuideRateRightAscension Read", "GuideRateRightAscension is < 0.0 " + guideRateRightAscension.ToString("0.00"));
                                    break;
                                }

                            default:
                                {
                                    LogOK("GuideRateRightAscension Read", guideRateRightAscension.ToString("0.00"));
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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("GuideRateRightAscension Write", "About to set GuideRateRightAscension property to " + guideRateRightAscension);
                        telescopeDevice.GuideRateRightAscension = guideRateRightAscension;
                        LogIssue("GuideRateRightAscension Write", "CanSetGuideRates is false but no exception generated; value returned: " + guideRateRightAscension.ToString("0.00"));
                    }
                    catch (Exception ex) // Some other error so OK
                    {
                        HandleException("GuideRateRightAscension Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False");
                    }
                }
            }
            else
            {
                LogInfo("GuideRateRightAscension", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // IsPulseGuiding - Optional
            if (interfaceVersion > 1)
            {
                if (canPulseGuide) // Can pulse guide so test if we can successfully read IsPulseGuiding
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("IsPulseGuiding", "About to get IsPulseGuiding property");
                        isPulseGuiding = telescopeDevice.IsPulseGuiding;
                        LogOK("IsPulseGuiding", isPulseGuiding.ToString());
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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("IsPulseGuiding", "About to get IsPulseGuiding property");
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
                LogInfo("IsPulseGuiding", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // OperationComplete
            if (interfaceVersion >= 4)
            {
                try
                {
                    LogCallToDriver("OperationComplete", "About to get OperationComplete property");
                    bool operationComplete = telescopeDevice.OperationComplete;

                    if (operationComplete)
                    {
                        LogOK("OperationComplete", $"OperationComplete is True as expected with no operation running.");
                    }
                    else
                    {
                        LogIssue("OperationComplete", $"OperationComplete is False even though no operation has been started.");
                    }
                }
                catch (Exception ex) // Read failed
                {
                    HandleException("OperationComplete", MemberType.Property, Required.MustBeImplemented, ex, "Interface is ITelescopeV4 or later");
                }
            }

            // RightAscension - Required
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("RightAscension", "About to get RightAscension property");
                rightAscension = telescopeDevice.RightAscension;
                switch (rightAscension)
                {
                    case var case14 when case14 < 0.0d:
                    case var case15 when case15 >= 24.0d:
                        {
                            LogIssue("RightAscension", "RightAscension is <0 or >=24 hours: " + rightAscension + " " + FormatRA(rightAscension).Trim());
                            break;
                        }

                    default:
                        {
                            LogOK("RightAscension", FormatRA(rightAscension).Trim());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("RightAscensionRate Read", "About to get RightAscensionRate property");
                rightAscensionRate = telescopeDevice.RightAscensionRate;
                // Read has been successful
                if (canSetRightAscensionRate) // Any value is acceptable
                {
                    switch (declinationRate)
                    {
                        case var case16 when case16 >= 0.0d:
                            {
                                LogOK("RightAscensionRate Read", rightAscensionRate.ToString("0.00"));
                                break;
                            }

                        default:
                            {
                                LogIssue("RightAscensionRate Read", "Negative RightAscensionRate: " + rightAscensionRate.ToString("0.00"));
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
                                LogOK("RightAscensionRate Read", rightAscensionRate.ToString("0.00"));
                                break;
                            }

                        default:
                            {
                                LogIssue("RightAscensionRate Read", "RightAscensionRate is non zero when CanSetRightAscensionRate is False " + declinationRate.ToString("0.00"));
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
            if (interfaceVersion > 1)
            {
                if (canSetRightAscensionRate) // Perform several tests starting with proving we can set a rate of 0.0
                {
                    if (TestRADecRate("RightAscensionRate Write", "Set rate to 0.0", Axis.RA, 0.0d, true))
                    {
                        TestRADecRate("RightAscensionRate Write", "Set rate to 0.1", Axis.RA, 0.1d, true);
                        TestRADecRate("RightAscensionRate Write", "Set rate to -0.1", Axis.RA, -0.1d, true);
                        TestRADecRate("RightAscensionRate Write", "Set rate to 0.5", Axis.RA, 0.5d, true);
                        TestRADecRate("RightAscensionRate Write", "Set rate to -0.5", Axis.RA, -0.5d, true);
                        TestRADecRate("RightAscensionRate Write", "Reset rate to 0.0", Axis.RA, 0.0d, false); // Reset the rate to zero, skipping the slewing test
                    }
                }
                else // Should generate an error
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("RightAscensionRate Write", "About to set RightAscensionRate property to 0.00");
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("RightAscensionRate Write", "About to set RightAscensionRate property to 0.00");
                    telescopeDevice.RightAscensionRate = 0.0d; // Set to a harmless value
                    LogOK("RightAscensionRate Write", rightAscensionRate.ToString("0.00"));
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteElevation Read", "About to get SiteElevation property");
                siteElevation = telescopeDevice.SiteElevation;
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
                            LogOK("SiteElevation Read", siteElevation.ToString());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteElevation Write", "About to set SiteElevation property to -301.0");
                telescopeDevice.SiteElevation = -301.0d;
                LogIssue("SiteElevation Write", "No error generated on set site elevation < -300m");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation < -300m");
            }

            // SiteElevation Write - Invalid high value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteElevation Write", "About to set SiteElevation property to 100001.0");
                telescopeDevice.SiteElevation = 10001.0d;
                LogIssue("SiteElevation Write", "No error generated on set site elevation > 10,000m");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation > 10,000m");
            }

            //SiteElevation Write - Current device value 
            try
            {
                if (siteElevation < -300.0d | siteElevation > 10000.0d)
                    siteElevation = 1000d;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteElevation Write", "About to set SiteElevation property to " + siteElevation);
                telescopeDevice.SiteElevation = siteElevation; // Restore original value
                LogOK("SiteElevation Write", $"Current value {siteElevation}m written successfully");
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteElevation Write", $"About to set SiteElevation property to arbitrary value:{SITE_ELEVATION_TEST_VALUE}");
                    telescopeDevice.SiteElevation = SITE_ELEVATION_TEST_VALUE;

                    // Read the value back
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteElevation Write", "About to get SiteElevation property");
                    double newElevation = telescopeDevice.SiteElevation;

                    // Compare with the expected value
                    if (newElevation == SITE_ELEVATION_TEST_VALUE)
                    {
                        LogOK("SiteElevation Write", $"Test value {SITE_ELEVATION_TEST_VALUE} set and read correctly");
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteElevation Write", $"About to restore original SiteElevation property :{siteElevation}");
                    telescopeDevice.SiteElevation = siteElevation;
                    LogOK("SiteElevation Write", $"Successfully restored original site elevation: {siteElevation}.");
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLatitude Read", "About to get SiteLatitude property");
                siteLatitude = telescopeDevice.SiteLatitude;
                switch (siteLatitude)
                {
                    case var case19 when case19 < -90.0d:
                        {
                            LogIssue("SiteLatitude Read", "SiteLatitude is < -90 degrees");
                            canReadSiteLatitude = false;
                            break;
                        }

                    case var case20 when case20 > 90.0d:
                        {
                            LogIssue("SiteLatitude Read", "SiteLatitude is > 90 degrees");
                            canReadSiteLatitude = false;
                            break;
                        }

                    default:
                        {
                            LogOK("SiteLatitude Read", FormatDec(siteLatitude));
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                canReadSiteLatitude = false;
                HandleException("SiteLatitude Read", MemberType.Property, Required.Optional, ex, "");
            }

            // SiteLatitude Write - Invalid low value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLatitude Write", "About to set SiteLatitude property to -91.0");
                telescopeDevice.SiteLatitude = -91.0d;
                LogIssue("SiteLatitude Write", "No error generated on set site latitude < -90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude < -90 degrees");
            }

            // SiteLatitude Write - Invalid high value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLatitude Write", "About to set SiteLatitude property to 91.0");
                telescopeDevice.SiteLatitude = 91.0d;
                LogIssue("SiteLatitude Write", "No error generated on set site latitude > 90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude > 90 degrees");
            }

            // SiteLatitude Write - Valid value
            bool canWriteSiteLatitude = true;
            try
            {
                if (siteLatitude < -90.0d | siteLatitude > 90.0d)
                    siteLatitude = 45.0d;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLatitude Write", "About to set SiteLatitude property to " + siteLatitude);
                telescopeDevice.SiteLatitude = siteLatitude; // Restore original value
                LogOK("SiteLatitude Write", $"Current value: {siteLatitude.ToDMS()} degrees written successfully");
            }
            catch (Exception ex)
            {
                canWriteSiteLatitude = false;
                HandleException("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "");
            }

            // Change the site latitude value
            if (canReadSiteLatitude & canWriteSiteLatitude & settings.TelescopeExtendedSiteTests)
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteLatitude Write", $"About to set SiteLatitude property to arbitrary value:{testLatitude.ToDMS()}");
                    telescopeDevice.SiteLatitude = testLatitude;

                    // Read the value back
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteLatitude Write", "About to get SiteLatitude property");
                    double newLatitude = telescopeDevice.SiteLatitude;

                    // Compare with the expected value
                    if (newLatitude == testLatitude)
                    {
                        LogOK("SiteLatitude Write", $"Test value {testLatitude.ToDMS()} set and read correctly");
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteLatitude Write", $"About to restore original SiteLatitude property :{siteLatitude.ToDMS()}");
                    telescopeDevice.SiteLatitude = siteLatitude;
                    LogOK("SiteLatitude Write", $"Successfully restored original site latitude: {siteLatitude.ToDMS()}.");
                }
                catch (Exception ex)
                {
                    HandleException("SiteLatitude Write", MemberType.Property, Required.MustBeImplemented, ex, "The original value could not be restored");
                }
            }
            if (cancellationToken.IsCancellationRequested) return;

            #endregion

            #region Site Longitude Tests

            // SiteLongitude Read - Optional
            bool canReadSiteLongitude = true;
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLongitude Read", "About to get SiteLongitude property");
                siteLongitude = telescopeDevice.SiteLongitude;
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
                            LogOK("SiteLongitude Read", FormatDec(siteLongitude));
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLongitude Write", "About to set SiteLongitude property to -181.0");
                telescopeDevice.SiteLongitude = -181.0d;
                LogIssue("SiteLongitude Write", "No error generated on set site longitude < -180 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude < -180 degrees");
            }

            // SiteLongitude Write - Invalid high value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLongitude Write", "About to set SiteLongitude property to 181.0");
                telescopeDevice.SiteLongitude = 181.0d;
                LogIssue("SiteLongitude Write", "No error generated on set site longitude > 180 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude > 180 degrees");
            }

            // SiteLongitude Write - Valid value
            bool canWriteSiteLongitude = true;
            try // Valid value
            {
                if (siteLongitude < -180.0d | siteLongitude > 180.0d)
                    siteLongitude = 60.0d;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiteLongitude Write", "About to set SiteLongitude property to " + siteLongitude);
                telescopeDevice.SiteLongitude = siteLongitude; // Restore original value
                LogOK("SiteLongitude Write", $"Current value {siteLongitude.ToDMS()} degrees written successfully");
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteLongitude Write", $"About to set SiteLongitude property to arbitrary value:{testLongitude.ToDMS()}");
                    telescopeDevice.SiteLongitude = testLongitude;

                    // Read the value back
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteLongitude Write", "About to get SiteLongitude property");
                    double newLongitude = telescopeDevice.SiteLongitude;

                    // Compare with the expected value
                    if (newLongitude == testLongitude)
                    {
                        LogOK("SiteLongitude Write", $"Test value {testLongitude.ToDMS()} set and read correctly");
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SiteLongitude Write", $"About to restore original SiteLongitude property :{siteLongitude.ToDMS()}");
                    telescopeDevice.SiteLongitude = siteLongitude;
                    LogOK("SiteLongitude Write", $"Successfully restored original site longitude: {siteLongitude.ToDMS()}.");
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("Slewing", "About to get Slewing property");
                slewing = telescopeDevice.Slewing;
                switch (slewing)
                {
                    case false:
                        {
                            LogOK("Slewing", slewing.ToString());
                            break;
                        }

                    case true:
                        {
                            LogIssue("Slewing", "Slewing should be false and it reads as " + slewing.ToString());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SlewSettleTime Read", "About to get SlewSettleTime property");
                slewSettleTime = telescopeDevice.SlewSettleTime;
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
                            LogOK("SlewSettleTime Read", slewSettleTime.ToString());
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
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SlewSettleTime Write", "About to set SlewSettleTime property to -1");
                telescopeDevice.SlewSettleTime = -1;
                LogIssue("SlewSettleTime Write", "No error generated on set SlewSettleTime < 0 seconds");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set slew settle time < 0");
            }

            try
            {
                if (slewSettleTime < 0)
                    slewSettleTime = 0;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SlewSettleTime Write", "About to set SlewSettleTime property to " + slewSettleTime);
                telescopeDevice.SlewSettleTime = slewSettleTime; // Restore original value
                LogOK("SlewSettleTime Write", "Legal value " + slewSettleTime.ToString() + " seconds written successfully");
            }
            catch (Exception ex)
            {
                HandleException("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // SideOfPier Read - Optional
            if (interfaceVersion > 1)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SideOfPier Read", "About to get SideOfPier property");
                    sideOfPier = (PointingState)telescopeDevice.SideOfPier;
                    LogOK("SideOfPier Read", sideOfPier.ToString());
                    canReadSideOfPier = true; // Flag that it is OK to read SideOfPier
                }
                catch (Exception ex)
                {
                    HandleException("SideOfPier Read", MemberType.Property, Required.Optional, ex, "");
                }
            }
            else
            {
                LogInfo("SideOfPier Read", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // SideOfPier Write - Optional
            // Moved to methods section as this really is a method rather than a property

            // SiderealTime - Required
            try
            {
                canReadSiderealTime = false;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SiderealTime", "About to get SiderealTime property");
                siderealTimeScope = telescopeDevice.SiderealTime;
                canReadSiderealTime = true;
                siderealTimeASCOM = (18.697374558d + 24.065709824419081d * (DateTime.UtcNow.ToOADate() + 2415018.5 - 2451545.0d) + siteLongitude / 15.0d) % 24.0d;
                switch (siderealTimeScope)
                {
                    case var case25 when case25 < 0.0d:
                    case var case26 when case26 >= 24.0d:
                        {
                            LogIssue("SiderealTime", "SiderealTime is <0 or >=24 hours: " + FormatRA(siderealTimeScope)); // Valid time returned
                            break;
                        }

                    default:
                        {
                            // Now do a sense check on the received value
                            LogOK("SiderealTime", FormatRA(siderealTimeScope));
                            l_TimeDifference = Math.Abs(siderealTimeScope - siderealTimeASCOM); // Get time difference between scope and PC
                                                                                                // Process edge cases where the two clocks are on either side of 0:0:0/24:0:0
                            if (siderealTimeASCOM > 23.0d & siderealTimeASCOM < 23.999d & siderealTimeScope > 0.0d & siderealTimeScope < 1.0d)
                            {
                                l_TimeDifference = Math.Abs(siderealTimeScope - siderealTimeASCOM + 24.0d);
                            }

                            if (siderealTimeScope > 23.0d & siderealTimeScope < 23.999d & siderealTimeASCOM > 0.0d & siderealTimeASCOM < 1.0d)
                            {
                                l_TimeDifference = Math.Abs(siderealTimeScope - siderealTimeASCOM - 24.0d);
                            }

                            switch (l_TimeDifference)
                            {
                                case var case27 when case27 <= 1.0d / 3600.0d: // 1 seconds
                                    {
                                        LogOK("SiderealTime", "Scope and ASCOM sidereal times agree to better than 1 second, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
                                        break;
                                    }

                                case var case28 when case28 <= 2.0d / 3600.0d: // 2 seconds
                                    {
                                        LogOK("SiderealTime", "Scope and ASCOM sidereal times agree to better than 2 seconds, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
                                        break;
                                    }

                                case var case29 when case29 <= 5.0d / 3600.0d: // 5 seconds
                                    {
                                        LogOK("SiderealTime", "Scope and ASCOM sidereal times agree to better than 5 seconds, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
                                        break;
                                    }

                                case var case30 when case30 <= 1.0d / 60.0d: // 1 minute
                                    {
                                        LogOK("SiderealTime", "Scope and ASCOM sidereal times agree to better than 1 minute, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
                                        break;
                                    }

                                case var case31 when case31 <= 5.0d / 60.0d: // 5 minutes
                                    {
                                        LogOK("SiderealTime", "Scope and ASCOM sidereal times agree to better than 5 minutes, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
                                        break;
                                    }

                                case var case32 when case32 <= 0.5d: // 0.5 an hour
                                    {
                                        LogInfo("SiderealTime", "Scope and ASCOM sidereal times are up to 0.5 hour different, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
                                        break;
                                    }

                                case var case33 when case33 <= 1.0d: // 1.0 an hour
                                    {
                                        LogInfo("SiderealTime", "Scope and ASCOM sidereal times are up to 1.0 hour different, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
                                        break;
                                    }

                                default:
                                    {
                                        LogIssue("SiderealTime", "Scope and ASCOM sidereal times are more than 1 hour apart, Scope: " + FormatRA(siderealTimeScope) + ", ASCOM: " + FormatRA(siderealTimeASCOM));
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
                            {
                                LogIssue("TargetDeclination Read", "TargetDeclination is <-90 or >90 degrees: " + FormatDec(targetDeclination));
                                break;
                            }

                        default:
                            {
                                LogOK("TargetDeclination Read", FormatDec(targetDeclination));
                                break;
                            }
                    }
                }
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
            {
                LogOK("TargetDeclination Read", "Not Set exception generated on read before write");
            }
            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
            {
                LogOK("TargetDeclination Read", "Not Set exception generated on read before write");
            }
            catch (Exception ex)
            {
                HandleInvalidOperationExceptionAsOK("TargetDeclination Read", MemberType.Property, Required.Optional, ex, "Incorrect exception received", "InvalidOperationException generated as expected on target read before read");
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
                            {
                                LogIssue("TargetRightAscension Read", "TargetRightAscension is <0 or >=24 hours: " + targetRightAscension + " " + FormatRA(targetRightAscension));
                                break;
                            }

                        default:
                            {
                                LogOK("TargetRightAscension Read", FormatRA(targetRightAscension));
                                break;
                            }
                    }
                }
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
            {
                LogOK("TargetRightAscension Read", "Not Set exception generated on read before write");
            }
            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
            {
                LogOK("TargetRightAscension Read", "Not Set exception generated on read before write");
            }
            catch (Exception ex)
            {
                HandleInvalidOperationExceptionAsOK("TargetRightAscension Read", MemberType.Property, Required.Optional, ex, "Incorrect exception received", "InvalidOperationException generated as expected on target read before read");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetRightAscension Write - Optional
            LogInfo("TargetRightAscension Write", "Tests moved after the SlewToCoordinates tests so that Conform can confirm that target coordinates are set as expected.");

            // Tracking Read - Required
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("Tracking Read", "About to get Tracking property");
                tracking = telescopeDevice.Tracking; // Read of tracking state is mandatory
                LogOK("Tracking Read", tracking.ToString());
            }
            catch (Exception ex)
            {
                HandleException("Tracking Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Tracking Write - Optional
            l_OriginalTrackingState = tracking;
            if (canSetTracking) // Set should work OK
            {
                SetTest("Tracking Write");
                try
                {
                    if (tracking) // OK try turning tracking off
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("Tracking Write", "About to set Tracking property false");
                        telescopeDevice.Tracking = false;
                    }
                    else // OK try turning tracking on
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("Tracking Write", "About to set Tracking property true");
                        telescopeDevice.Tracking = true;
                    }

                    SetAction("Waiting for mount to stabilise");
                    WaitFor(TRACKING_COMMAND_DELAY); // Wait for a short time to allow mounts to implement the tracking state change
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("Tracking Write", "About to get Tracking property");
                    tracking = telescopeDevice.Tracking;
                    if (tracking != l_OriginalTrackingState)
                    {
                        LogOK("Tracking Write", tracking.ToString());
                    }
                    else
                    {
                        LogIssue("Tracking Write", "Tracking didn't change state on write: " + tracking.ToString());
                    }

                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("Tracking Write", "About to set Tracking property " + l_OriginalTrackingState);
                    telescopeDevice.Tracking = l_OriginalTrackingState; // Restore original state
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
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("Tracking Write", "About to set Tracking property false");
                        telescopeDevice.Tracking = false;
                    }
                    else // OK try turning tracking on
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("Tracking Write", "About to set Tracking property true");
                        telescopeDevice.Tracking = true;
                    }

                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("Tracking Write", "About to get Tracking property");
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
            if (interfaceVersion > 1)
            {
                int l_Count = 0;
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("TrackingRates", "About to get TrackingRates property");
                    l_TrackingRates = telescopeDevice.TrackingRates;
                    if (l_TrackingRates is null)
                    {
                        LogDebug("TrackingRates", "ERROR: The driver did NOT return an TrackingRates object!");
                    }
                    else
                    {
                        LogDebug("TrackingRates", "OK - the driver returned an TrackingRates object");
                    }

                    l_Count = l_TrackingRates.Count; // Save count for use later if no members are returned in the for each loop test
                    LogDebug("TrackingRates Count", l_Count.ToString());

                    var loopTo = l_TrackingRates.Count;
                    for (int ii = 1; ii <= loopTo; ii++)
                        LogDebug("TrackingRates Count", "Found drive rate: " + Enum.GetName(typeof(DriveRate), (l_TrackingRates[ii])));
                }
                catch (Exception ex)
                {
                    HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (l_TrackingRates is not null)
                {
                    try
                    {
                        IEnumerator l_Enum;
                        object l_Obj;
                        DriveRate l_Drv;
                        l_Enum = (IEnumerator)l_TrackingRates.GetEnumerator();
                        if (l_Enum is null)
                        {
                            LogDebug("TrackingRates Enum", "ERROR: The driver did NOT return an Enumerator object!");
                        }
                        else
                        {
                            LogDebug("TrackingRates Enum", "OK - the driver returned an Enumerator object");
                        }

                        l_Enum.Reset();
                        LogDebug("TrackingRates Enum", "Reset Enumerator");
                        while (l_Enum.MoveNext())
                        {
                            LogDebug("TrackingRates Enum", "Reading Current");
                            l_Obj = l_Enum.Current;
                            LogDebug("TrackingRates Enum", "Read Current OK, Type: " + l_Obj.GetType().Name);
                            l_Drv = (DriveRate)l_Obj;
                            LogDebug("TrackingRates Enum", "Found drive rate: " + Enum.GetName(typeof(DriveRate), l_Drv));
                        }

                        l_Enum.Reset();
                        l_Enum = null;

                        // Clean up TrackingRates object
                        if (l_TrackingRates is object)
                        {
                            try
                            {
                                l_TrackingRates.Dispose();
                            }
                            catch
                            {
                            }

                            if (OperatingSystem.IsWindows())
                            {
                                try
                                {
                                    Marshal.ReleaseComObject(l_TrackingRates);
                                }
                                catch
                                {
                                }
                            }

                            l_TrackingRates = null;
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("TrackingRates", "About to get TrackingRates property");
                    l_TrackingRates = telescopeDevice.TrackingRates;
                    LogDebug("TrackingRates", $"Read TrackingRates OK, Count: {l_TrackingRates.Count}");
                    int l_RateCount = 0;
                    foreach (DriveRate currentL_DriveRate in (IEnumerable)l_TrackingRates)
                    {
                        l_DriveRate = currentL_DriveRate;
                        LogTestAndMessage("TrackingRates", "Found drive rate: " + l_DriveRate.ToString());
                        l_RateCount += 1;
                    }

                    if (l_RateCount > 0)
                    {
                        LogOK("TrackingRates", "Drive rates read OK");
                    }
                    else if (l_Count > 0) // We did get some members on the first call, but now they have disappeared!
                    {
                        // This can be due to the driver returning the same TrackingRates object on every TrackingRates call but not resetting the iterator pointer
                        LogIssue("TrackingRates", "Multiple calls to TrackingRates returned different answers!");
                        LogInfo("TrackingRates", "");
                        LogInfo("TrackingRates", "The first call to TrackingRates returned " + l_Count + " drive rates; the next call appeared to return no rates.");
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
                if (l_TrackingRates is object)
                {
                    try
                    {
                        l_TrackingRates.Dispose();
                    }
                    catch
                    {
                    }

                    if (OperatingSystem.IsWindows())
                    {
                        try
                        {
                            Marshal.ReleaseComObject(l_TrackingRates);
                        }
                        catch { }
                    }
                }

                // Test the TrackingRates.Dispose() method
                LogDebug("TrackingRates", "Getting tracking rates");
                l_TrackingRates = telescopeDevice.TrackingRates;
                try
                {
                    LogDebug("TrackingRates", "Disposing tracking rates");
                    l_TrackingRates.Dispose();
                    LogOK("TrackingRates", "Disposed tracking rates OK");
                }
                catch (MissingMemberException)
                {
                    LogOK("TrackingRates", "Dispose member not present");
                }
                catch (Exception ex)
                {
                    LogIssue("TrackingRates", "TrackingRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: " + ex.Message);
                    LogDebug("TrackingRates.Dispose", "Exception: " + ex.ToString());
                }

                if (OperatingSystem.IsWindows())
                {
                    try
                    {
                        Marshal.ReleaseComObject(l_TrackingRates);
                    }
                    catch { }
                }

                if (cancellationToken.IsCancellationRequested)
                    return;

                // TrackingRate - Test after TrackingRates so we know what the valid values are
                // TrackingRate Read - Required
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("TrackingRates", "About to get TrackingRates property");
                    l_TrackingRates = telescopeDevice.TrackingRates;
                    if (l_TrackingRates is object) // Make sure that we have received a TrackingRates object after the Dispose() method was called
                    {
                        LogOK("TrackingRates", "Successfully obtained a TrackingRates object after the previous TrackingRates object was disposed");
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("TrackingRate Read", "About to get TrackingRate property");
                        l_TrackingRate = (DriveRate)telescopeDevice.TrackingRate;
                        LogOK("TrackingRate Read", l_TrackingRate.ToString());

                        // TrackingRate Write - Optional
                        // We can read TrackingRate so now test trying to set each tracking rate in turn
                        try
                        {
                            LogDebug("TrackingRate Write", "About to enumerate tracking rates object");
                            foreach (DriveRate currentL_DriveRate1 in (IEnumerable)l_TrackingRates)
                            {
                                l_DriveRate = currentL_DriveRate1;
                                //Application.DoEvents();
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("TrackingRate Write", "About to set TrackingRate property to " + l_DriveRate.ToString());
                                    telescopeDevice.TrackingRate = l_DriveRate;
                                    if (settings.DisplayMethodCalls)
                                        if (telescopeDevice.TrackingRate == l_DriveRate)
                                        {
                                            LogOK("TrackingRate Write", "Successfully set drive rate: " + l_DriveRate.ToString());
                                        }
                                        else
                                        {
                                            LogIssue("TrackingRate Write", "Unable to set drive rate: " + l_DriveRate.ToString());
                                        }
                                }
                                catch (Exception ex)
                                {
                                    HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "");
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
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("TrackingRate Write", "About to set TrackingRate property to invalid value (5)");
                            telescopeDevice.TrackingRate = (DriveRate)5;
                            LogIssue("TrackingRate Write", "No error generated when TrackingRate is set to an invalid value (5)");
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (5)");
                        }

                        // Attempt to write an invalid low tracking rate
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("TrackingRate Write", "About to set TrackingRate property to invalid value (-1)");
                            telescopeDevice.TrackingRate = (DriveRate)(0 - 1); // Done this way to fool the compiler into allowing me to attempt to set a negative, invalid value
                            LogIssue("TrackingRate Write", "No error generated when TrackingRate is set to an invalid value (-1)");
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (-1)");
                        }

                        // Finally restore original TrackingRate
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("TrackingRate Write", "About to set TrackingRate property to " + l_TrackingRate.ToString());
                            telescopeDevice.TrackingRate = l_TrackingRate;
                        }
                        catch (Exception ex)
                        {
                            HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "Unable to restore original tracking rate");
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
                LogInfo("TrackingRate", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // UTCDate Read - Required
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("UTCDate Read", "About to get UTCDate property");
                utcDate = telescopeDevice.UTCDate; // Save starting value
                LogOK("UTCDate Read", utcDate.ToString("dd-MMM-yyyy HH:mm:ss.fff"));

                try // UTCDate Write is optional since if you are using the PC time as UTCTime then you should not write to the PC clock!
                {
                    // Try to write a new UTCDate  1 hour in the future
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("UTCDate Write", $"About to set UTCDate property to {utcDate.AddHours(1.0d)}");
                    telescopeDevice.UTCDate = utcDate.AddHours(1.0d);
                    LogOK("UTCDate Write", $"New UTCDate written successfully: {utcDate.AddHours(1.0d)}");

                    // Restore original value
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("UTCDate Write", $"About to set UTCDate property to {utcDate}");
                    telescopeDevice.UTCDate = utcDate;
                    LogOK("UTCDate Write", $"Original UTCDate restored successfully: {utcDate}");
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
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_CAN_MOVE_AXIS] | telescopeTests[TELTEST_MOVE_AXIS] | telescopeTests[TELTEST_PARK_UNPARK])
                {
                    TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisPrimary, "CanMoveAxis:Primary");
                    if (cancellationToken.IsCancellationRequested) return;
                    TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisSecondary, "CanMoveAxis:Secondary");
                    if (cancellationToken.IsCancellationRequested) return;
                    TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisTertiary, "CanMoveAxis:Tertiary");
                    if (cancellationToken.IsCancellationRequested) return;
                }
                else
                {
                    LogInfo(TELTEST_CAN_MOVE_AXIS, "Tests skipped");
                }
            }
            else
            {
                LogInfo("CanMoveAxis", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // Test Park, Unpark - Optional
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_PARK_UNPARK])
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
                                    SetAction("Parking scope...");
                                    LogTestAndMessage("Park", "Parking scope...");
                                    LogCallToDriver("Park", "About to call Park method");
                                    telescopeDevice.Park();
                                    LogCallToDriver("Park", "About to get AtPark property repeatedly...");

                                    // Wait for the park to complete
                                    WaitWhile("Waiting for scope to park", () => { return !telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                    if (cancellationToken.IsCancellationRequested) return;

                                    SetStatus("Scope parked");
                                    LogOK("Park", "Success");

                                    // Scope Parked OK
                                    try // Confirm second park is harmless
                                    {
                                        LogCallToDriver("Park", "About to Park call method");
                                        telescopeDevice.Park();
                                        LogOK("Park", "Success if already parked");

                                        LogCallToDriver("Park", "About to get AtPark property repeatedly...");

                                        // Wait for the park to complete
                                        WaitWhile("Waiting for scope to park", () => { return !telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                        if (cancellationToken.IsCancellationRequested) return;

                                        SetStatus("Scope parked");
                                        LogOK("Park", "Success if already parked");
                                    }
                                    catch (COMException ex)
                                    {
                                        LogIssue("Park", "Exception when calling Park two times in succession: " + ex.Message + " " + ((int)ex.ErrorCode).ToString("X8"));
                                    }
                                    catch (Exception ex)
                                    {
                                        LogIssue("Park", "Exception when calling Park two times in succession: " + ex.Message);
                                    }

                                    // Check if the operation properties are implemented
                                    if (interfaceVersion >= 4) // Operations are supported
                                    {
                                        try
                                        {
                                            // Prepare the mount to be parked again
                                            LogCallToDriver("Park", "Unpark");
                                            telescopeDevice.Unpark();
                                            LogCallToDriver("Park", "About to get AtPark property repeatedly");
                                            WaitWhile("Waiting for scope to unpark when parked", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                            LogCallToDriver("Park", "Tracking");
                                            telescopeDevice.Tracking = true;

                                            // Slew to an arbitrary test HA
                                            SlewToHa(1.0);

                                            // Park the scope
                                            SetAction("Parking scope for OperationComplete test...");
                                            LogTestAndMessage("Park", "Parking scope for OperationComplete test...");
                                            LogCallToDriver("Park", "About to call Park method");

                                            // Validate OperationComplete state
                                            ValidateOperationComplete("Park", true);

                                            TimeMethod("Park", () => telescopeDevice.Park());

                                            // Wait for the park to complete
                                            LogCallToDriver("Park", "About to get OperationComplete property repeatedly...");
                                            WaitWhile("Waiting for scope to park for OperationComplete test", () => { return !telescopeDevice.OperationComplete; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                            // Validate OperationComplete state
                                            ValidateOperationComplete("Park", true);

                                            if (telescopeDevice.AtPark)
                                            {
                                                SetStatus("Scope parked for OperationComplete test");
                                                LogOK("Park", "Success for OperationComplete test");
                                            }
                                            else
                                            {
                                                LogIssue("Park", $"Failed to park within {settings.TelescopeMaximumSlewTime} seconds while using OperationComplete");
                                            }

                                            if (cancellationToken.IsCancellationRequested) return;
                                        }
                                        catch (Exception ex)
                                        {
                                            LogIssue("Park", $"Exception trying to test Park operation: {ex.Message}");
                                            LogDebug("Park", ex.ToString());
                                        }
                                    }

                                    // Confirm that methods do raise exceptions when scope is parked
                                    if (canSlew | canSlewAsync | canSlewAltAz | canSlewAltAzAsync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepAbortSlew, "AbortSlew");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canFindHome)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepFindHome, "FindHome");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canMoveAxisPrimary)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisPrimary, "MoveAxis Primary");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canMoveAxisSecondary)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisSecondary, "MoveAxis Secondary");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canMoveAxisTertiary)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisTertiary, "MoveAxis Tertiary");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canPulseGuide)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepPulseGuide, "PulseGuide");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canSlew)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinates, "SlewToCoordinates");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canSlewAsync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinatesAsync, "SlewToCoordinatesAsync");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canSlew)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTarget, "SlewToTarget");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canSlewAsync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTargetAsync, "SlewToTargetAsync");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canSync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToCoordinates, "SyncToCoordinates");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    if (canSync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToTarget, "SyncToTarget");
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }

                                    // Test unpark after park
                                    if (canUnpark)
                                    {
                                        try
                                        {
                                            if (settings.DisplayMethodCalls)
                                                LogTestAndMessage("Unpark", "About to call Unpark method");
                                            telescopeDevice.Unpark();
                                            LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                            WaitWhile("Waiting for scope to unpark when parked", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                            if (cancellationToken.IsCancellationRequested)
                                                return;

                                            try // Make sure tracking doesn't generate an error if it is not implemented
                                            {
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("Unpark", "About to set Tracking property true");
                                                telescopeDevice.Tracking = true;
                                            }
                                            catch (Exception)
                                            {
                                            }

                                            SetStatus("Scope Unparked OK");
                                            LogOK("Unpark", "Success");

                                            // Scope unparked
                                            try // Confirm Unpark is harmless if already unparked
                                            {
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("Unpark", "About to call Unpark method");
                                                telescopeDevice.Unpark();
                                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                                WaitWhile("Waiting for scope to unpark when parked", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                                LogOK("Unpark", "Success when already unparked");
                                            }
                                            catch (COMException ex)
                                            {
                                                LogIssue("Unpark", "Exception when calling Unpark two times in succession: " + ex.Message + " " + ((int)ex.ErrorCode).ToString("X8"));
                                            }
                                            catch (Exception ex)
                                            {
                                                LogIssue("Unpark", "Exception when calling Unpark two times in succession: " + ex.Message);
                                            }

                                            // Test OperationComplete if appropriate.
                                            if (interfaceVersion >= 4)
                                            {
                                                // Park the scope 
                                                SetAction("Parking scope...");
                                                LogTestAndMessage("Unpark", "Parking scope for Unpark OperationComplete test...");
                                                LogCallToDriver("Unpark", "About to call Park method");
                                                telescopeDevice.Park();
                                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly...");
                                                WaitWhile("Waiting for scope to park", () => { return !telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                                SetStatus("Scope parked");
                                                if (cancellationToken.IsCancellationRequested) return;

                                                // Validate OperationComplete state
                                                ValidateOperationComplete("Unpark", true);

                                                // Now unpark as an operation

                                                LogCallToDriver("Unpark", "About to call Unpark method");
                                                TimeMethod("Unpark", () => telescopeDevice.Unpark());

                                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                                WaitWhile("Waiting for scope to unpark when parked", () => { return !telescopeDevice.OperationComplete; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                                // Validate OperationComplete state
                                                ValidateOperationComplete("Unpark", true);

                                                if (telescopeDevice.OperationComplete)
                                                {
                                                    LogOK("Unpark", "Success when using OperationComplete");
                                                }
                                                else
                                                {
                                                    LogIssue("Unpark", $"Failed to unpark after {settings.TelescopeMaximumSlewTime} seconds when using waiting for OperationComplete");
                                                }
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
                                            if (settings.DisplayMethodCalls)
                                                LogTestAndMessage("Unpark", "About to call Unpark method");
                                            telescopeDevice.Unpark();
                                            LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                            WaitWhile("Waiting for scope to unpark", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

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
                            LogOK("Park", "Success if already parked");

                            LogCallToDriver("Park", "About to get AtPark property repeatedly...");

                            // Wait for the park to complete
                            WaitWhile("Waiting for scope to park", () => { return !telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                            if (cancellationToken.IsCancellationRequested) return;

                            SetStatus("Scope parked");
                            LogIssue("Park", "CanPark is false but no exception was generated on use");
                        }
                        catch (Exception ex)
                        {
                            HandleException("Park", MemberType.Method, Required.MustNotBeImplemented, ex, "CanPark is False");
                        }
                        // v1.0.12.0 Added test for unpark if CanPark is false
                        // Now test unpark
                        if (canUnpark) // We should already be unparked so confirm that unpark works fine
                        {
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("UnPark", "About to call UnPark method");
                                telescopeDevice.Unpark();
                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                WaitWhile("Waiting for scope to unpark", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

                                LogOK("UnPark", "CanPark is false and CanUnPark is true; no exception generated as expected");
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
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("UnPark", "About to call UnPark method");
                                telescopeDevice.Unpark();
                                LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                WaitWhile("Waiting for scope to unpark", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);

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
                    LogInfo(TELTEST_PARK_UNPARK, "Tests skipped");
                }
            }
            else
            {
                LogInfo("Park", "Skipping tests since behaviour of this method is not well defined in interface V" + interfaceVersion);
            }

            // AbortSlew - Optional
            if (telescopeTests[TELTEST_ABORT_SLEW])
            {
                TelescopeOptionalMethodsTest(OptionalMethodType.AbortSlew, "AbortSlew", true);
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo(TELTEST_ABORT_SLEW, "Tests skipped");
            }

            // AxisRates - Required
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_AXIS_RATE] | telescopeTests[TELTEST_MOVE_AXIS])
                {
                    TelescopeAxisRateTest("AxisRate:Primary", TelescopeAxis.Primary);
                    TelescopeAxisRateTest("AxisRate:Secondary", TelescopeAxis.Secondary);
                    TelescopeAxisRateTest("AxisRate:Tertiary", TelescopeAxis.Tertiary);
                }
                else
                {
                    LogInfo(TELTEST_AXIS_RATE, "Tests skipped");
                }
            }
            else
            {
                LogInfo("AxisRate", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // FindHome - Optional
            if (telescopeTests[TELTEST_FIND_HOME])
            {
                TelescopeOptionalMethodsTest(OptionalMethodType.FindHome, "FindHome", canFindHome);
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo(TELTEST_FIND_HOME, "Tests skipped");
            }

            // MoveAxis - Optional
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_MOVE_AXIS])
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
                    LogInfo(TELTEST_MOVE_AXIS, "Tests skipped");
                }
            }
            else
            {
                LogInfo("MoveAxis", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // PulseGuide - Optional
            if (telescopeTests[TELTEST_PULSE_GUIDE])
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
                            if (applicationCancellationToken.IsCancellationRequested)
                                return;

                            TestPulseGuide(+9.0);
                            if (applicationCancellationToken.IsCancellationRequested)
                                return;

                            TestPulseGuide(-3.0);
                            if (applicationCancellationToken.IsCancellationRequested)
                                return;

                            TestPulseGuide(+3.0);
                        }
                        else
                            LogIssue(TELTEST_PULSE_GUIDE, $"Extended pulse guide tests skipped because at least one of the GuideRateRightAscension or GuideRateDeclination properties was not implemented or returned a zero or negative value.");
                    }
                    else
                        LogInfo(TELTEST_PULSE_GUIDE, $"Extended pulse guide tests skipped because the CanPulseGuide property returned False.");
                }
            }
            else
            {
                LogInfo(TELTEST_PULSE_GUIDE, "Tests skipped");
            }

            // Test Equatorial slewing to coordinates - Optional
            if (telescopeTests[TELTEST_SLEW_TO_COORDINATES])
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
                LogInfo(TELTEST_SLEW_TO_COORDINATES, "Tests skipped");
            }

            // Test Equatorial slewing to coordinates asynchronous - Optional
            if (telescopeTests[TELTEST_SLEW_TO_COORDINATES_ASYNC])
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
                LogInfo(TELTEST_SLEW_TO_COORDINATES_ASYNC, "Tests skipped");
            }

            // Equatorial Sync to Coordinates - Optional - Moved here so that it can be tested before any target coordinates are set - Peter 4th August 2018
            if (telescopeTests[TELTEST_SYNC_TO_COORDINATES])
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
                LogInfo(TELTEST_SYNC_TO_COORDINATES, "Tests skipped");
            }

            // TargetRightAscension Write - Optional - Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetRightAscension Write", "About to set TargetRightAscension property to -1.0");
                telescopeDevice.TargetRightAscension = -1.0d;
                LogIssue("TargetRightAscension Write", "No error generated on set TargetRightAscension < 0 hours");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension < 0 hours");
            }

            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetRightAscension Write", "About to set TargetRightAscension property to 25.0");
                telescopeDevice.TargetRightAscension = 25.0d;
                LogIssue("TargetRightAscension Write", "No error generated on set TargetRightAscension > 24 hours");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension > 24 hours");
            }

            try
            {
                targetRightAscension = TelescopeRAFromSiderealTime("TargetRightAscension Write", -4.0d);
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetRightAscension Write", "About to set TargetRightAscension property to " + targetRightAscension);
                telescopeDevice.TargetRightAscension = targetRightAscension; // Set a valid value
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("TargetRightAscension Write", "About to get TargetRightAscension property");
                    switch (Math.Abs(telescopeDevice.TargetRightAscension - targetRightAscension))
                    {
                        case 0.0d:
                            {
                                LogOK("TargetRightAscension Write", "Legal value " + FormatRA(targetRightAscension) + " HH:MM:SS written successfully");
                                break;
                            }

                        case var @case when @case <= 1.0d / 3600.0d: // 1 seconds
                            {
                                LogOK("TargetRightAscension Write", "Target RightAscension is within 1 second of the value set: " + FormatRA(targetRightAscension));
                                break;
                            }

                        case var case1 when case1 <= 2.0d / 3600.0d: // 2 seconds
                            {
                                LogOK("TargetRightAscension Write", "Target RightAscension is within 2 seconds of the value set: " + FormatRA(targetRightAscension));
                                break;
                            }

                        case var case2 when case2 <= 5.0d / 3600.0d: // 5 seconds
                            {
                                LogOK("TargetRightAscension Write", "Target RightAscension is within 5 seconds of the value set: " + FormatRA(targetRightAscension));
                                break;
                            }

                        default:
                            {
                                LogInfo("TargetRightAscension Write", "Target RightAscension: " + FormatRA(telescopeDevice.TargetRightAscension));
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

            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetDeclination Write - Optional - Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetDeclination Write", "About to set TargetDeclination property to -91.0");
                telescopeDevice.TargetDeclination = -91.0d;
                LogIssue("TargetDeclination Write", "No error generated on set TargetDeclination < -90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
            }

            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetDeclination Write", "About to set TargetDeclination property to 91.0");
                telescopeDevice.TargetDeclination = 91.0d;
                LogIssue("TargetDeclination Write", "No error generated on set TargetDeclination > 90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
            }

            try
            {
                targetDeclination = 1.0d;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("TargetDeclination Write", "About to set TargetDeclination property to " + targetDeclination);
                telescopeDevice.TargetDeclination = targetDeclination; // Set a valid value
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("TargetDeclination Write", "About to get TargetDeclination property");
                    switch (Math.Abs(telescopeDevice.TargetDeclination - targetDeclination))
                    {
                        case 0.0d:
                            {
                                LogOK("TargetDeclination Write", "Legal value " + FormatDec(targetDeclination) + " DD:MM:SS written successfully");
                                break;
                            }

                        case var case3 when case3 <= 1.0d / 3600.0d: // 1 seconds
                            {
                                LogOK("TargetDeclination Write", "Target Declination is within 1 second of the value set: " + FormatDec(targetDeclination));
                                break;
                            }

                        case var case4 when case4 <= 2.0d / 3600.0d: // 2 seconds
                            {
                                LogOK("TargetDeclination Write", "Target Declination is within 2 seconds of the value set: " + FormatDec(targetDeclination));
                                break;
                            }

                        case var case5 when case5 <= 5.0d / 3600.0d: // 5 seconds
                            {
                                LogOK("TargetDeclination Write", "Target Declination is within 5 seconds of the value set: " + FormatDec(targetDeclination));
                                break;
                            }

                        default:
                            {
                                LogInfo("TargetDeclination Write", "Target Declination: " + FormatDec(targetDeclination));
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

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test Equatorial target slewing - Optional
            if (telescopeTests[TELTEST_SLEW_TO_TARGET])
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
                LogInfo(TELTEST_SLEW_TO_TARGET, "Tests skipped");
            }

            // Test Equatorial target slewing asynchronous - Optional
            if (telescopeTests[TELTEST_SLEW_TO_TARGET_ASYNC])
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
                LogInfo(TELTEST_SLEW_TO_TARGET_ASYNC, "Tests skipped");
            }

            // DestinationSideOfPier - Optional
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_DESTINATION_SIDE_OF_PIER])
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
                    LogInfo(TELTEST_DESTINATION_SIDE_OF_PIER, "Tests skipped");
                }
            }
            else
            {
                LogInfo("DestinationSideOfPier", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // Test AltAz Slewing - Optional
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_SLEW_TO_ALTAZ])
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
                    LogInfo(TELTEST_SLEW_TO_ALTAZ, "Tests skipped");
                }
            }
            else
            {
                LogInfo("SlewToAltAz", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // Test AltAz Slewing asynchronous - Optional
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_SLEW_TO_ALTAZ_ASYNC])
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
                    LogInfo(TELTEST_SLEW_TO_ALTAZ_ASYNC, "Tests skipped");
                }
            }
            else
            {
                LogInfo("SlewToAltAzAsync", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            // Equatorial Sync to Target - Optional
            if (telescopeTests[TELTEST_SYNC_TO_TARGET])
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
                LogInfo(TELTEST_SYNC_TO_TARGET, "Tests skipped");
            }

            // AltAz Sync - Optional
            if (interfaceVersion > 1)
            {
                if (telescopeTests[TELTEST_SYNC_TO_ALTAZ])
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
                    LogInfo(TELTEST_SYNC_TO_ALTAZ, "Tests skipped");
                }
            }
            else
            {
                LogInfo("SyncToAltAz", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (settings.TestSideOfPierRead)
            {
                LogNewLine();
                LogTestOnly("SideOfPier Model Tests"); LogDebug("SideOfPier Model Tests", "Starting tests");
                if (interfaceVersion > 1)
                {
                    // 3.0.0.14 - Skip these tests if unable to read SideOfPier
                    if (CanReadSideOfPier("SideOfPier Model Tests"))
                    {

                        // Further side of pier tests
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("SideOfPier Model Tests", "About to get AlignmentMode property");
                        if (telescopeDevice.AlignmentMode == AlignmentMode.GermanPolar)
                        {
                            LogDebug("SideOfPier Model Tests", "Calling SideOfPierTests()");
                            switch (siteLatitude)
                            {
                                case var case6 when -SIDE_OF_PIER_INVALID_LATITUDE <= case6 && case6 <= SIDE_OF_PIER_INVALID_LATITUDE: // Refuse to handle this value because the Conform targeting logic or the mount's SideofPier flip logic may fail when the poles are this close to the horizon
                                    {
                                        LogInfo("SideOfPier Model Tests", "Tests skipped because the site latitude is reported as " + Utilities.DegreesToDMS(siteLatitude, ":", ":", "", 3));
                                        LogInfo("SideOfPier Model Tests", "This places the celestial poles close to the horizon and the mount's flip logic may override Conform's expected behaviour.");
                                        LogInfo("SideOfPier Model Tests", "Please set the site latitude to a value within the ranges " + SIDE_OF_PIER_INVALID_LATITUDE.ToString("+0.0;-0.0") + " to +90.0 or " + (-SIDE_OF_PIER_INVALID_LATITUDE).ToString("+0.0;-0.0") + " to -90.0 to obtain a reliable result.");
                                        break;
                                    }

                                case var case7 when -90.0d <= case7 && case7 <= 90.0d: // Normal case, just run the tests barbecue latitude is outside the invalid range but within -90.0 to +90.0
                                    {
                                        // SideOfPier write property test - Optional
                                        if (settings.TestSideOfPierWrite)
                                        {
                                            LogDebug("SideOfPier Model Tests", "Testing SideOfPier write...");
                                            TelescopeOptionalMethodsTest(OptionalMethodType.SideOfPierWrite, "SideOfPier Write", canSetPierside);
                                            if (cancellationToken.IsCancellationRequested)
                                                return;
                                        }

                                        SideOfPierTests(); // Only run these for German mounts
                                        break; // Values outside the range -90.0 to +90.0 are invalid
                                    }

                                default:
                                    {
                                        LogInfo("SideOfPier Model Tests", "Test skipped because the site latitude Is outside the range -90.0 to +90.0");
                                        break;
                                    }
                            }
                        }
                        else
                        {
                            LogInfo("SideOfPier Model Tests", "Test skipped because this Is Not a German equatorial mount");
                        }
                    }
                    else
                    {
                        LogInfo("SideOfPier Model Tests", "Tests skipped because this driver does Not support SideOfPier Read");
                    }
                }
                else
                {
                    LogInfo("SideOfPier Model Tests", "Skipping test as this method Is Not supported in interface V" + interfaceVersion);
                }
            }

        }

        public override void CheckPerformance()
        {
            SetTest("Performance"); // Clear status messages
            TelescopePerformanceTest(PerformanceType.tstPerfAltitude, "Altitude");
            if (cancellationToken.IsCancellationRequested)
                return;
            if (interfaceVersion > 1)
            {
                TelescopePerformanceTest(PerformanceType.tstPerfAtHome, "AtHome");
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("Performance: AtHome", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            if (interfaceVersion > 1)
            {
                TelescopePerformanceTest(PerformanceType.tstPerfAtPark, "AtPark");
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("Performance: AtPark", "Skipping test as this method is not supported in interface V" + interfaceVersion);
            }

            TelescopePerformanceTest(PerformanceType.tstPerfAzimuth, "Azimuth");
            if (cancellationToken.IsCancellationRequested)
                return;
            TelescopePerformanceTest(PerformanceType.tstPerfDeclination, "Declination");
            if (cancellationToken.IsCancellationRequested)
                return;
            if (interfaceVersion > 1)
            {
                if (canPulseGuide)
                {
                    TelescopePerformanceTest(PerformanceType.tstPerfIsPulseGuiding, "IsPulseGuiding");
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
                LogInfo("Performance: IsPulseGuiding", "Skipping test as this method is not supported in interface v1" + interfaceVersion);
            }

            TelescopePerformanceTest(PerformanceType.tstPerfRightAscension, "RightAscension");
            if (cancellationToken.IsCancellationRequested)
                return;
            if (interfaceVersion > 1)
            {
                if (alignmentMode == AlignmentMode.GermanPolar)
                {
                    if (CanReadSideOfPier("Performance - SideOfPier"))
                    {
                        TelescopePerformanceTest(PerformanceType.tstPerfSideOfPier, "SideOfPier");
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
                LogInfo("Performance: SideOfPier", "Skipping test as this method is not supported in interface v1" + interfaceVersion);
            }

            if (canReadSiderealTime)
            {
                TelescopePerformanceTest(PerformanceType.tstPerfSiderealTime, "SiderealTime");
                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("Performance: SiderealTime", "Skipping test because the SiderealTime property throws an exception.");
            }

            TelescopePerformanceTest(PerformanceType.tstPerfSlewing, "Slewing");
            if (cancellationToken.IsCancellationRequested)
                return;
            TelescopePerformanceTest(PerformanceType.tstPerfUTCDate, "UTCDate");
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
                    LogOK("Mount Safety", "Tracking stopped to protect your mount.");
                }
                else
                {
                    LogInfo("Mount Safety", "Tracking can't be turned off for this mount, please switch off manually.");
                }
            }
            catch (Exception ex)
            {
                LogIssue("Mount Safety", "Exception when disabling tracking to protect mount: " + ex.ToString());
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
                    if (!telescopeTests[TELTEST_CAN_MOVE_AXIS])
                        LogConfigurationAlert("Can move axis tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_PARK_UNPARK])
                        LogConfigurationAlert("Park and Unpark tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_ABORT_SLEW])
                        LogConfigurationAlert("Abort slew tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_AXIS_RATE])
                        LogConfigurationAlert("Axis rate tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_FIND_HOME])
                        LogConfigurationAlert("Find home tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_MOVE_AXIS])
                        LogConfigurationAlert("Move axis tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_PULSE_GUIDE])
                        LogConfigurationAlert("Pulse guide tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SLEW_TO_COORDINATES])
                        LogConfigurationAlert("Synchronous Slew to coordinates tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SLEW_TO_COORDINATES_ASYNC])
                        LogConfigurationAlert("Asynchronous slew to coordinates tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SLEW_TO_TARGET])
                        LogConfigurationAlert("Synchronous slew to target tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SLEW_TO_TARGET_ASYNC])
                        LogConfigurationAlert("Asynchronous slew to target tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_DESTINATION_SIDE_OF_PIER])
                        LogConfigurationAlert("Destination side of pier tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SLEW_TO_ALTAZ])
                        LogConfigurationAlert("Synchronous slew to altitude / azimuth tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SLEW_TO_ALTAZ_ASYNC])
                        LogConfigurationAlert("Asynchronous slew to altitude / azimuth tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SYNC_TO_COORDINATES])
                        LogConfigurationAlert("Sync to coordinates tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SYNC_TO_TARGET])
                        LogConfigurationAlert("Sync to target tests were omitted due to Conform configuration.");

                    if (!telescopeTests[TELTEST_SYNC_TO_ALTAZ])
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

        private void TelescopeSyncTest(SlewSyncType testType, string testName, bool driverSupportsMethod, string canDoItName)
        {
            bool showOutcome = false;
            double difference, syncRA, syncDEC, syncAlt = default, syncAz = default, newAlt, newAz, currentAz = default, currentAlt = default, startRA, startDec, currentRA, currentDec;

            SetTest(testName);
            SetAction("Running test...");

            // Basic test to make sure the method is either implemented OK or fails as expected if it is not supported in this driver.
            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to get RightAscension property");
            syncRA = telescopeDevice.RightAscension;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to get Declination property");
            syncDEC = telescopeDevice.Declination;
            if (!driverSupportsMethod) // Call should fail
            {
                try
                {
                    switch (testType)
                    {
                        case SlewSyncType.SyncToCoordinates: // SyncToCoordinates
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                if (canSetTracking & !telescopeDevice.Tracking)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set Tracking property to true");
                                    telescopeDevice.Tracking = true;
                                }

                                LogDebug(testName, "SyncToCoordinates: " + FormatRA(syncRA) + " " + FormatDec(syncDEC));
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call SyncToCoordinates method, RA: " + FormatRA(syncRA) + ", Declination: " + FormatDec(syncDEC));
                                telescopeDevice.SyncToCoordinates(syncRA, syncDEC);
                                LogIssue(testName, "CanSyncToCoordinates is False but call to SyncToCoordinates did not throw an exception.");
                                break;
                            }

                        case SlewSyncType.SyncToTarget: // SyncToTarget
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                if (canSetTracking & !telescopeDevice.Tracking)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set Tracking property to true");
                                    telescopeDevice.Tracking = true;
                                }

                                try
                                {
                                    LogDebug(testName, "Setting TargetRightAscension: " + FormatRA(syncRA));
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set TargetRightAscension property to " + FormatRA(syncRA));
                                    telescopeDevice.TargetRightAscension = syncRA;
                                    LogDebug(testName, "Completed Set TargetRightAscension");
                                }
                                catch (Exception)
                                {
                                    // Ignore errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }

                                try
                                {
                                    LogDebug(testName, "Setting TargetDeclination: " + FormatDec(syncDEC));
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set TargetDeclination property to " + FormatDec(syncDEC));
                                    telescopeDevice.TargetDeclination = syncDEC;
                                    LogDebug(testName, "Completed Set TargetDeclination");
                                }
                                catch (Exception)
                                {
                                    // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }

                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget(); // Sync to target coordinates
                                LogIssue(testName, "CanSyncToTarget is False but call to SyncToTarget did not throw an exception.");
                                break;
                            }

                        case SlewSyncType.SyncToAltAz:
                            {
                                if (canReadAltitide)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get Altitude property");
                                    syncAlt = telescopeDevice.Altitude;
                                }

                                if (canReadAzimuth)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get Azimuth property");
                                    syncAz = telescopeDevice.Azimuth;
                                }

                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                if (canSetTracking & telescopeDevice.Tracking)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set Tracking property to false");
                                    telescopeDevice.Tracking = false;
                                }

                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call SyncToAltAz method, Altitude: " + FormatDec(syncAlt) + ", Azimuth: " + FormatDec(syncAz));
                                telescopeDevice.SyncToAltAz(syncAz, syncAlt); // Sync to new Alt Az
                                LogIssue(testName, "CanSyncToAltAz is False but call to SyncToAltAz did not throw an exception.");
                                break;
                            }

                        default:
                            {
                                LogIssue(testName, "Conform:SyncTest: Unknown test type " + testType.ToString());
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, canDoItName + " is False");
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
                            {

                                // Calculate the Sync test RA position
                                startRA = TelescopeRAFromHourAngle(testName, +3.0d);
                                LogDebug(testName, string.Format("RA for sync tests: {0}", FormatRA(startRA)));

                                // Calculate the Sync test DEC position
                                if (siteLatitude > 0.0d) // We are in the northern hemisphere
                                {
                                    startDec = 90.0d - (180.0d - siteLatitude) * 0.5d; // Calculate for northern hemisphere
                                }
                                else // We are in the southern hemisphere
                                {
                                    startDec = -90.0d + (180.0d + siteLatitude) * 0.5d;
                                } // Calculate for southern hemisphere

                                LogDebug(testName, string.Format("Declination for sync tests: {0}", FormatDec(startDec)));
                                SlewScope(startRA, startDec, "start position");
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                SetAction("Checking that scope slewed OK");
                                // Now test that we have actually arrived
                                CheckScopePosition(testName, "Slewed to start position", startRA, startDec);

                                // Calculate the sync test RA coordinate as a variation from the current RA coordinate
                                syncRA = startRA - SYNC_SIMULATED_ERROR / (15.0d * 60.0d); // Convert sync error in arc minutes to RA hours
                                if (syncRA < 0.0d)
                                    syncRA += 24.0d; // Ensure legal RA

                                // Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                                syncDEC = startDec - SYNC_SIMULATED_ERROR / 60.0d; // Convert sync error in arc minutes to degrees

                                SetAction("Syncing the scope");
                                // Sync the scope to the offset RA and DEC coordinates
                                SyncScope(testName, canDoItName, testType, syncRA, syncDEC);

                                // Check that the scope's synchronised position is as expected
                                CheckScopePosition(testName, "Synced to sync position", syncRA, syncDEC);

                                // Check that the TargetRA and TargetDec were 
                                SetAction("Checking that the scope synced OK");
                                if (testType == SlewSyncType.SyncToCoordinates)
                                {
                                    // Check that target coordinates are present and set correctly per the ASCOM Telescope specification
                                    try
                                    {
                                        currentRA = telescopeDevice.TargetRightAscension;
                                        LogDebug(testName, string.Format("Current TargetRightAscension: {0}, Set TargetRightAscension: {1}", currentRA, syncRA));
                                        double raDifference;
                                        raDifference = RaDifferenceInArcSeconds(syncRA, currentRA);
                                        switch (raDifference)
                                        {
                                            case var @case when @case <= SLEW_SYNC_OK_TOLERANCE:  // Within specified tolerance
                                                {
                                                    LogOK(testName, string.Format("The TargetRightAscension property {0} matches the expected RA OK. ", FormatRA(syncRA))); // Outside specified tolerance
                                                    break;
                                                }

                                            default:
                                                {
                                                    LogIssue(testName, string.Format("The TargetRightAscension property {0} does not match the expected RA {1}", FormatRA(currentRA), FormatRA(syncRA)));
                                                    break;
                                                }
                                        }
                                    }
                                    catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                    {
                                        LogIssue(testName, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                                    }
                                    catch (ASCOM.InvalidOperationException)
                                    {
                                        LogIssue(testName, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                    }
                                    catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
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
                                        LogDebug(testName, string.Format("Current TargetDeclination: {0}, Set TargetDeclination: {1}", currentDec, syncDEC));
                                        double decDifference;
                                        decDifference = Math.Round(Math.Abs(currentDec - syncDEC) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // Dec difference is in arc seconds from degrees of Declination
                                        switch (decDifference)
                                        {
                                            case var case1 when case1 <= SLEW_SYNC_OK_TOLERANCE: // Within specified tolerance
                                                {
                                                    LogOK(testName, string.Format("The TargetDeclination property {0} matches the expected Declination OK. ", FormatDec(syncDEC))); // Outside specified tolerance
                                                    break;
                                                }

                                            default:
                                                {
                                                    LogIssue(testName, string.Format("The TargetDeclination property {0} does not match the expected Declination {1}", FormatDec(currentDec), FormatDec(syncDEC)));
                                                    break;
                                                }
                                        }
                                    }
                                    catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                    {
                                        LogIssue(testName, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                                    }
                                    catch (ASCOM.InvalidOperationException)
                                    {
                                        LogIssue(testName, "The driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                    }
                                    catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
                                    {
                                        LogIssue(testName, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
                                    }
                                }

                                // Now slew to the scope's original position
                                SlewScope(startRA, startDec, "original position in post-sync coordinates");

                                // Check that the scope's position is the original position
                                CheckScopePosition(testName, "Slewed back to start position", startRA, startDec);

                                // Now "undo" the sync by reversing syncing in the opposition sense than originally made

                                // Calculate the sync test RA coordinate as a variation from the current RA coordinate
                                syncRA = startRA + SYNC_SIMULATED_ERROR / (15.0d * 60.0d); // Convert sync error in arc minutes to RA hours
                                if (syncRA >= 24.0d)
                                    syncRA -= 24.0d; // Ensure legal RA

                                // Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                                syncDEC = startDec + SYNC_SIMULATED_ERROR / 60.0d; // Convert sync error in arc minutes to degrees

                                // Sync back to the original coordinates
                                SetAction("Restoring original sync values");
                                SyncScope(testName, canDoItName, testType, syncRA, syncDEC);

                                // Check that the scope's synchronised position is as expected
                                CheckScopePosition(testName, "Synced to reversed sync position", syncRA, syncDEC);

                                // Now slew to the scope's original position
                                SlewScope(startRA, startDec, "original position in pre-sync coordinates");

                                // Check that the scope's position is the original position
                                CheckScopePosition(testName, "Slewed back to start position", startRA, startDec);
                                break;
                            }

                        case SlewSyncType.SyncToAltAz:
                            {
                                if (canReadAltitide)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get Altitude property");
                                    currentAlt = telescopeDevice.Altitude;
                                }

                                if (canReadAzimuth)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get Azimuth property");
                                    currentAz = telescopeDevice.Azimuth;
                                }

                                syncAlt = currentAlt - 1.0d;
                                syncAz = currentAz + 1.0d;
                                if (syncAlt < 0.0d)
                                    syncAlt = 1.0d; // Ensure legal Alt
                                if (syncAz > 359.0d)
                                    syncAz = 358.0d; // Ensure legal Az
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                if (canSetTracking & telescopeDevice.Tracking)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set Tracking property to false");
                                    telescopeDevice.Tracking = false;
                                }

                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call SyncToAltAz method, Altitude: " + FormatDec(syncAlt) + ", Azimuth: " + FormatDec(syncAz));
                                telescopeDevice.SyncToAltAz(syncAz, syncAlt); // Sync to new Alt Az
                                if (canReadAltitide & canReadAzimuth) // Can check effects of a sync
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get Altitude property");
                                    newAlt = telescopeDevice.Altitude;
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get Azimuth property");
                                    newAz = telescopeDevice.Azimuth;

                                    // Compare old and new values
                                    difference = Math.Abs(syncAlt - newAlt);
                                    switch (difference)
                                    {
                                        case var case2 when case2 <= 1.0d / (60 * 60): // Within 1 seconds
                                            {
                                                LogOK(testName, "Synced Altitude OK");
                                                break;
                                            }

                                        case var case3 when case3 <= 2.0d / (60 * 60): // Within 2 seconds
                                            {
                                                LogOK(testName, "Synced within 2 seconds of Altitude");
                                                showOutcome = true;
                                                break;
                                            }

                                        default:
                                            {
                                                LogInfo(testName, $"Synced to within {TelescopeTester.FormatAltitude(difference)} DD:MM:SS of expected Altitude: {TelescopeTester.FormatAltitude(syncAlt)}");
                                                showOutcome = true;
                                                break;
                                            }
                                    }

                                    difference = Math.Abs(syncAz - newAz);
                                    switch (difference)
                                    {
                                        case var case4 when case4 <= 1.0d / (60 * 60): // Within 1 seconds
                                            {
                                                LogOK(testName, "Synced Azimuth OK");
                                                break;
                                            }

                                        case var case5 when case5 <= 2.0d / (60 * 60): // Within 2 seconds
                                            {
                                                LogOK(testName, "Synced within 2 seconds of Azimuth");
                                                showOutcome = true;
                                                break;
                                            }

                                        default:
                                            {
                                                LogInfo(testName, "Synced to within " + FormatAzimuth(difference) + " DD:MM:SS of expected Azimuth: " + FormatAzimuth(syncAz));
                                                showOutcome = true;
                                                break;
                                            }
                                    }

                                    if (showOutcome)
                                    {
                                        LogTestAndMessage(testName, "           Altitude    Azimuth");
                                        LogTestAndMessage(testName, $"Original:  {TelescopeTester.FormatAltitude(currentAlt)}   {FormatAzimuth(currentAz)}");
                                        LogTestAndMessage(testName, $"Sync to:   {TelescopeTester.FormatAltitude(syncAlt)}   {FormatAzimuth(syncAz)}");
                                        LogTestAndMessage(testName, $"New:       {TelescopeTester.FormatAltitude(newAlt)}   {FormatAzimuth(newAz)}");
                                    }
                                }
                                else // Can't test effects of a sync
                                {
                                    LogInfo(testName, "Can't test SyncToAltAz because Altitude or Azimuth values are not implemented");
                                } // Do nothing

                                break;
                            }

                        default:
                            {
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, canDoItName + " is True");
                }
            }
        }

        private void TelescopeSlewTest(SlewSyncType p_Test, string p_Name, bool p_CanDoIt, string p_CanDoItName)
        {
            SetTest(p_Name);
            if (canSetTracking)
            {
                LogCallToDriver(p_Name, "About to set Tracking property to true");
                telescopeDevice.Tracking = true; // Enable tracking for these tests
            }

            try
            {
                switch (p_Test)
                {
                    case SlewSyncType.SlewToCoordinates:
                        {
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            if (canSetTracking & !telescopeDevice.Tracking)
                            {
                                LogCallToDriver(p_Name, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -1.0d);
                            targetDeclination = 1.0d;
                            SetAction("Slewing synchronously...");

                            LogCallToDriver(p_Name, "About to call SlewToCoordinates method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                            telescopeDevice.SlewToCoordinates(targetRightAscension, targetDeclination);
                            LogDebug(p_Name, "Returned from SlewToCoordinates method");
                            break;
                        }

                    case SlewSyncType.SlewToCoordinatesAsync:
                        {
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            if (canSetTracking & !telescopeDevice.Tracking)
                            {
                                LogCallToDriver(p_Name, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0d);
                            targetDeclination = 2.0d;

                            LogCallToDriver(p_Name, "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                            telescopeDevice.SlewToCoordinatesAsync(targetRightAscension, targetDeclination);
                            if (settings.DisplayMethodCalls) LogDebug(p_Name, $"Asynchronous slew initiated");

                            WaitForSlew(p_Name, $"Slewing to coordinates asynchronously");
                            if (settings.DisplayMethodCalls) LogDebug(p_Name, $"Slew completed");
                            break;
                        }

                    case SlewSyncType.SlewToTarget:
                        {
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            if (canSetTracking & !telescopeDevice.Tracking)
                            {
                                LogCallToDriver(p_Name, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0d);
                            targetDeclination = 3.0d;
                            try
                            {
                                LogCallToDriver(p_Name, "About to set TargetRightAscension property to " + FormatRA(targetRightAscension));
                                telescopeDevice.TargetRightAscension = targetRightAscension;
                            }
                            catch (Exception ex)
                            {
                                if (p_CanDoIt)
                                {
                                    HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName + " is True but can't set TargetRightAscension");
                                }
                                else
                                {
                                    // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }
                            }

                            try
                            {
                                LogCallToDriver(p_Name, "About to set TargetDeclination property to " + FormatDec(targetDeclination));
                                telescopeDevice.TargetDeclination = targetDeclination;
                            }
                            catch (Exception ex)
                            {
                                if (p_CanDoIt)
                                {
                                    HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName + " is True but can't set TargetDeclination");
                                }
                                else
                                {
                                    // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }
                            }

                            SetAction("Slewing synchronously...");
                            LogCallToDriver(p_Name, "About to call SlewToTarget method");
                            telescopeDevice.SlewToTarget();
                            break;
                        }

                    case SlewSyncType.SlewToTargetAsync: // SlewToTargetAsync
                        {
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            if (canSetTracking & !telescopeDevice.Tracking)
                            {
                                LogCallToDriver(p_Name, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -4.0d);
                            targetDeclination = 4.0d;
                            try
                            {
                                LogCallToDriver(p_Name, "About to set TargetRightAscension property to " + FormatRA(targetRightAscension));
                                telescopeDevice.TargetRightAscension = targetRightAscension;
                            }
                            catch (Exception ex)
                            {
                                if (p_CanDoIt)
                                {
                                    HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName + " is True but can't set TargetRightAscension");
                                }
                                else
                                {
                                    // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }
                            }

                            try
                            {
                                LogCallToDriver(p_Name, "About to set TargetDeclination property to " + FormatDec(targetDeclination));
                                telescopeDevice.TargetDeclination = targetDeclination;
                            }
                            catch (Exception ex)
                            {
                                if (p_CanDoIt)
                                {
                                    HandleException(p_Name, MemberType.Property, Required.MustBeImplemented, ex, p_CanDoItName + " is True but can't set TargetDeclination");
                                }
                                else
                                {
                                    // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }
                            }

                            LogCallToDriver(p_Name, "About to call SlewToTargetAsync method");
                            telescopeDevice.SlewToTargetAsync();

                            WaitForSlew(p_Name, $"Slewing to target asynchronously");
                            break;
                        }

                    case SlewSyncType.SlewToAltAz:
                        {
                            LogDebug(p_Name, $"Tracking 1: {telescopeDevice.Tracking}");
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            if (canSetTracking & telescopeDevice.Tracking)
                            {
                                LogCallToDriver(p_Name, "About to set property Tracking to false");
                                telescopeDevice.Tracking = false;
                                LogDebug(p_Name, "Tracking turned off");
                            }

                            LogCallToDriver(p_Name, "About to get Tracking property");
                            LogDebug(p_Name, $"Tracking 2: {telescopeDevice.Tracking}");
                            targetAltitude = 50.0d;
                            targetAzimuth = 150.0d;
                            SetAction("Slewing to Alt/Az synchronously...");

                            LogCallToDriver(p_Name, "About to call SlewToAltAz method, Altitude: " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                            telescopeDevice.SlewToAltAz(targetAzimuth, targetAltitude);

                            LogCallToDriver(p_Name, "About to get Tracking property");
                            LogDebug(p_Name, $"Tracking 3: {telescopeDevice.Tracking}");
                            break;
                        }

                    case SlewSyncType.SlewToAltAzAsync:
                        {
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            LogDebug(p_Name, $"Tracking 1: {telescopeDevice.Tracking}");
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            if (canSetTracking & telescopeDevice.Tracking)
                            {
                                LogCallToDriver(p_Name, "About to set Tracking property false");
                                telescopeDevice.Tracking = false;
                                LogDebug(p_Name, "Tracking turned off");
                            }

                            LogCallToDriver(p_Name, "About to get Tracking property");
                            LogDebug(p_Name, $"Tracking 2: {telescopeDevice.Tracking}");
                            targetAltitude = 55.0d;
                            targetAzimuth = 155.0d;

                            LogCallToDriver(p_Name, "About to call SlewToAltAzAsync method, Altitude: " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                            telescopeDevice.SlewToAltAzAsync(targetAzimuth, targetAltitude);

                            LogCallToDriver(p_Name, "About to get Tracking property");
                            LogDebug(p_Name, $"Tracking 3: {telescopeDevice.Tracking}");

                            WaitForSlew(p_Name, $"Slewing to Alt/Az asynchronously");
                            LogCallToDriver(p_Name, "About to get Tracking property");
                            LogDebug(p_Name, $"Tracking 4: {telescopeDevice.Tracking}");
                            break;
                        }

                    default:
                        {
                            LogError(p_Name, "Conform:SlewTest: Unknown test type " + p_Test.ToString());
                            break;
                        }
                }

                if (cancellationToken.IsCancellationRequested) return;

                if (p_CanDoIt) // Should be able to do this so report what happened
                {
                    SetAction("Slew completed");
                    switch (p_Test)
                    {
                        case SlewSyncType.SlewToCoordinates:
                        case SlewSyncType.SlewToCoordinatesAsync:
                        case SlewSyncType.SlewToTarget:
                        case SlewSyncType.SlewToTargetAsync:
                            {
                                // Test how close the slew was to the required coordinates
                                CheckScopePosition(p_Name, "Slewed", targetRightAscension, targetDeclination);

                                // Check that the slews and syncs set the target RA coordinate correctly per the ASCOM Telescope specification
                                try
                                {
                                    double actualTargetRA = telescopeDevice.TargetRightAscension;
                                    LogDebug(p_Name, $"Current TargetRightAscension: {actualTargetRA}, Set TargetRightAscension: {targetRightAscension}");

                                    if (RaDifferenceInArcSeconds(actualTargetRA, targetRightAscension) <= settings.TelescopeSlewTolerance) // Within specified tolerance
                                    {
                                        LogOK(p_Name, $"The TargetRightAscension property: {FormatRA(actualTargetRA)} matches the expected RA {FormatRA(targetRightAscension)} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds."); // Outside specified tolerance
                                    }
                                    else
                                    {
                                        LogIssue(p_Name, $"The TargetRightAscension property: {FormatRA(actualTargetRA)} does not match the expected RA {FormatRA(targetRightAscension)} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                                    }
                                }
                                catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                {
                                    LogIssue(p_Name, "The Driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                                }
                                catch (ASCOM.InvalidOperationException)
                                {
                                    LogIssue(p_Name, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                }
                                catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
                                {
                                    LogIssue(p_Name, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
                                }

                                // Check that the slews and syncs set the target declination coordinate correctly per the ASCOM Telescope specification
                                try
                                {
                                    double actualTargetDec = telescopeDevice.TargetDeclination;
                                    LogDebug(p_Name, $"Current TargetDeclination: {actualTargetDec}, Set TargetDeclination: {targetDeclination}");

                                    if (Math.Round(Math.Abs(actualTargetDec - targetDeclination) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero) <= settings.TelescopeSlewTolerance) // Within specified tolerance
                                    {
                                        LogOK(p_Name, $"The TargetDeclination property {FormatDec(actualTargetDec)} matches the expected Declination {FormatDec(targetDeclination)} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds."); // Outside specified tolerance
                                    }
                                    else
                                    {
                                        LogIssue(p_Name, $"The TargetDeclination property {FormatDec(actualTargetDec)} does not match the expected Declination {FormatDec(targetDeclination)} within tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                                    }

                                }
                                catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                {
                                    LogIssue(p_Name, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet exception was thrown instead.");
                                }
                                catch (ASCOM.InvalidOperationException)
                                {
                                    LogIssue(p_Name, "The Driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                }
                                catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
                                {
                                    LogIssue(p_Name, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
                                }

                                break;
                            }

                        case SlewSyncType.SlewToAltAz:
                        case SlewSyncType.SlewToAltAzAsync:
                            {
                                // Test how close the slew was to the required coordinates
                                LogCallToDriver(p_Name, "About to get Azimuth property");
                                double actualAzimuth = telescopeDevice.Azimuth;

                                LogCallToDriver(p_Name, "About to get Altitude property");
                                double actualAltitude = telescopeDevice.Altitude;

                                double azimuthDifference = Math.Abs(actualAzimuth - targetAzimuth);
                                if (azimuthDifference > 350.0d) azimuthDifference = 360.0d - azimuthDifference; // Deal with the case where the two elements are on different sides of 360 degrees

                                if (azimuthDifference <= settings.TelescopeSlewTolerance)
                                {
                                    LogOK(p_Name, $"Slewed to target Azimuth OK within tolerance: {settings.TelescopeSlewTolerance} arc seconds. Actual Azimuth: {FormatAzimuth(actualAzimuth)}, Target Azimuth: {FormatAzimuth(targetAzimuth)}");
                                }
                                else
                                {
                                    LogIssue(p_Name, $"Slewed {azimuthDifference:0.0} arc seconds away from Azimuth target: {FormatAzimuth(targetAzimuth)} Actual Azimuth: {FormatAzimuth(actualAzimuth)}. Tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                                }

                                double altitudeDifference = Math.Abs(actualAltitude - targetAltitude);
                                if (altitudeDifference <= settings.DomeSlewTolerance)
                                {
                                    LogOK(p_Name, $"Slewed to target Altitude OK within tolerance: {settings.TelescopeSlewTolerance} arc seconds. Actual Altitude: {FormatAltitude(actualAltitude)}, Target Altitude: {FormatAltitude(targetAltitude)}");
                                }
                                else
                                {
                                    LogIssue(p_Name, $"Slewed {altitudeDifference:0.0} degree(s) away from Altitude target: {FormatAltitude(targetAltitude)} Actual Altitude: {FormatAltitude(actualAltitude)}. Tolerance ±{settings.TelescopeSlewTolerance} arc seconds.");
                                }

                                break;
                            }

                        default: // Do nothing
                            {
                                break;
                            }
                    }
                }
                else // Not supposed to be able to do this but no error generated so report an error
                {
                    LogIssue(p_Name, p_CanDoItName + " is false but no exception was generated on use");
                }
            }
            catch (Exception ex)
            {
                if (p_CanDoIt)
                {
                    HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, p_CanDoItName + " is True");
                }
                else
                {
                    HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, p_CanDoItName + " is False");
                }
            }

        }

        /// <summary>
        /// Confirm that InValidValueExceptions are thrown for invalid values
        /// </summary>
        /// <param name="p_Name"></param>
        /// <param name="p_Test">The method to test</param>
        /// <param name="BadCoordinate1">RA or Altitude</param>
        /// <param name="BadCoordinate2">Dec or Azimuth</param>
        /// <remarks></remarks>
        private void TelescopeBadCoordinateTest(string p_Name, SlewSyncType p_Test, double BadCoordinate1, double BadCoordinate2)
        {
            switch (p_Test)
            {
                case SlewSyncType.SlewToCoordinates:
                case SlewSyncType.SlewToCoordinatesAsync:
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(p_Name, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = BadCoordinate1;
                            targetDeclination = 0.0d;
                            if (p_Test == SlewSyncType.SlewToCoordinates)
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SlewToCoordinates method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                                telescopeDevice.SlewToCoordinates(targetRightAscension, targetDeclination);
                            }
                            else
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                                telescopeDevice.SlewToCoordinatesAsync(targetRightAscension, targetDeclination);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogIssue(p_Name, "Failed to reject bad RA coordinate: " + FormatRA(targetRightAscension));
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(targetRightAscension));
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0d);
                            targetDeclination = BadCoordinate2;
                            if (p_Test == SlewSyncType.SlewToCoordinates)
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SlewToCoordinates method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                                telescopeDevice.SlewToCoordinates(targetRightAscension, targetDeclination);
                            }
                            else
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                                telescopeDevice.SlewToCoordinatesAsync(targetRightAscension, targetDeclination);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogIssue(p_Name, "Failed to reject bad Dec coordinate: " + FormatDec(targetDeclination));
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(targetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SyncToCoordinates:
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(p_Name, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = BadCoordinate1;
                            targetDeclination = 0.0d;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call SyncToCoordinates method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                            telescopeDevice.SyncToCoordinates(targetRightAscension, targetDeclination);
                            LogIssue(p_Name, "Failed to reject bad RA coordinate: " + FormatRA(targetRightAscension));
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(targetRightAscension));
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0d);
                            targetDeclination = BadCoordinate2;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call SyncToCoordinates method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(targetDeclination));
                            telescopeDevice.SyncToCoordinates(targetRightAscension, targetDeclination);
                            LogIssue(p_Name, "Failed to reject bad Dec coordinate: " + FormatDec(targetDeclination));
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(targetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SlewToTarget:
                case SlewSyncType.SlewToTargetAsync:
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(p_Name, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = BadCoordinate1;
                            targetDeclination = 0.0d;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set TargetRightAscension property to " + FormatRA(targetRightAscension));
                            telescopeDevice.TargetRightAscension = targetRightAscension;
                            // Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to set TargetDeclination property to " + FormatDec(targetDeclination));
                                telescopeDevice.TargetDeclination = targetDeclination;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (p_Test == SlewSyncType.SlewToTarget)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(p_Name, "About to call SlewToTarget method");
                                    telescopeDevice.SlewToTarget();
                                }
                                else
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(p_Name, "About to call SlewToTargetAsync method");
                                    telescopeDevice.SlewToTargetAsync();
                                }

                                SetAction("Attempting to abort slew");
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(p_Name, "About to call AbortSlew method");
                                    telescopeDevice.AbortSlew();
                                }
                                catch
                                {
                                } // Attempt to stop any motion that has actually started

                                LogIssue(p_Name, "Failed to reject bad RA coordinate: " + FormatRA(targetRightAscension));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Slew rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(targetRightAscension));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate", "Telescope.TargetRA correctly rejected bad RA coordinate: " + FormatRA(targetRightAscension));
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0d);
                            targetDeclination = BadCoordinate2;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set TargetDeclination property to " + FormatDec(targetDeclination));
                            telescopeDevice.TargetDeclination = targetDeclination;
                            // Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to set TargetRightAscension property to " + FormatRA(targetRightAscension));
                                telescopeDevice.TargetRightAscension = targetRightAscension;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (p_Test == SlewSyncType.SlewToTarget)
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(p_Name, "About to call SlewToTarget method");
                                    telescopeDevice.SlewToTarget();
                                }
                                else
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(p_Name, "About to call SlewToTargetAsync method");
                                    telescopeDevice.SlewToTargetAsync();
                                }

                                SetAction("Attempting to abort slew");
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(p_Name, "About to call AbortSlew method");
                                    telescopeDevice.AbortSlew();
                                }
                                catch
                                {
                                } // Attempt to stop any motion that has actually started

                                LogIssue(p_Name, "Failed to reject bad Dec coordinate: " + FormatDec(targetDeclination));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Slew rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(targetDeclination));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate", "Telescope.TargetDeclination correctly rejected bad Dec coordinate: " + FormatDec(targetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SyncToTarget:
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(p_Name, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = BadCoordinate1;
                            targetDeclination = 0.0d;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set TargetRightAscension property to " + FormatRA(targetRightAscension));
                            telescopeDevice.TargetRightAscension = targetRightAscension;
                            // Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to set TargetDeclination property to " + FormatDec(targetDeclination));
                                telescopeDevice.TargetDeclination = targetDeclination;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                LogIssue(p_Name, "Failed to reject bad RA coordinate: " + FormatRA(targetRightAscension));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Sync rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(targetRightAscension));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate", "Telescope.TargetRA correctly rejected bad RA coordinate: " + FormatRA(targetRightAscension));
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0d);
                            targetDeclination = BadCoordinate2;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set TargetDeclination property to " + FormatDec(targetDeclination));
                            telescopeDevice.TargetDeclination = targetDeclination;
                            // Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to set TargetRightAscension property to " + FormatRA(targetRightAscension));
                                telescopeDevice.TargetRightAscension = targetRightAscension;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                LogIssue(p_Name, "Failed to reject bad Dec coordinate: " + FormatDec(targetDeclination));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                SetAction("Sync rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(targetDeclination));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate", "Telescope.TargetDeclination correctly rejected bad Dec coordinate: " + FormatDec(targetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SlewToAltAz:
                case SlewSyncType.SlewToAltAzAsync:
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(p_Name, "About to get Tracking property");
                        if (canSetTracking & telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set Tracking property to false");
                            telescopeDevice.Tracking = false;
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetAltitude = BadCoordinate1;
                            targetAzimuth = 45.0d;
                            if (p_Test == SlewSyncType.SlewToAltAz)
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SlewToAltAz method, Altitude: " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                                telescopeDevice.SlewToAltAz(targetAzimuth, targetAltitude);
                            }
                            else
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About To call SlewToAltAzAsync method, Altitude:  " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                                telescopeDevice.SlewToAltAzAsync(targetAzimuth, targetAltitude);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogIssue(p_Name, $"Failed to reject bad Altitude coordinate: {TelescopeTester.FormatAltitude(targetAltitude)}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Altitude coordinate", $"Correctly rejected bad Altitude coordinate: {TelescopeTester.FormatAltitude(targetAltitude)}");
                        }

                        try
                        {
                            SetAction("Slew underway");
                            targetAltitude = 45.0d;
                            targetAzimuth = BadCoordinate2;
                            if (p_Test == SlewSyncType.SlewToAltAz)
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SlewToAltAz method, Altitude: " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                                telescopeDevice.SlewToAltAz(targetAzimuth, targetAltitude);
                            }
                            else
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call SlewToAltAzAsync method, Altitude: " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                                telescopeDevice.SlewToAltAzAsync(targetAzimuth, targetAltitude);
                            }

                            SetAction("Attempting to abort slew");
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(p_Name, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogIssue(p_Name, "Failed to reject bad Azimuth coordinate: " + FormatAzimuth(targetAzimuth));
                        }
                        catch (Exception ex)
                        {
                            SetAction("Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Azimuth coordinate", "Correctly rejected bad Azimuth coordinate: " + FormatAzimuth(targetAzimuth));
                        }

                        break;
                    }

                case SlewSyncType.SyncToAltAz:
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(p_Name, "About to get Tracking property");
                        if (canSetTracking & telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to set Tracking property to false");
                            telescopeDevice.Tracking = false;
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetAltitude = BadCoordinate1;
                            targetAzimuth = 45.0d;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call SyncToAltAz method, Altitude: " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                            telescopeDevice.SyncToAltAz(targetAzimuth, targetAltitude);
                            LogIssue(p_Name, $"Failed to reject bad Altitude coordinate: {TelescopeTester.FormatAltitude(targetAltitude)}");
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Altitude coordinate", $"Correctly rejected bad Altitude coordinate: {TelescopeTester.FormatAltitude(targetAltitude)}");
                        }

                        try
                        {
                            SetAction("Sync underway");
                            targetAltitude = 45.0d;
                            targetAzimuth = BadCoordinate2;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call SyncToAltAz method, Altitude: " + FormatDec(targetAltitude) + ", Azimuth: " + FormatDec(targetAzimuth));
                            telescopeDevice.SyncToAltAz(targetAzimuth, targetAltitude);
                            LogIssue(p_Name, "Failed to reject bad Azimuth coordinate: " + FormatAzimuth(targetAzimuth));
                        }
                        catch (Exception ex)
                        {
                            SetAction("Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Azimuth coordinate", "Correctly rejected bad Azimuth coordinate: " + FormatAzimuth(targetAzimuth));
                        }

                        break;
                    }

                default:
                    {
                        LogIssue(p_Name, "Conform:SlewTest: Unknown test type " + p_Test.ToString());
                        break;
                    }
            }

            if (cancellationToken.IsCancellationRequested)
                return;
        }

        private void TelescopePerformanceTest(PerformanceType p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate;
            SetAction(p_Name);
            try
            {
                l_StartTime = DateTime.Now;
                l_Count = 0.0d;
                l_LastElapsedTime = 0.0d;
                do
                {
                    l_Count += 1.0d;
                    switch (p_Type)
                    {
                        case PerformanceType.tstPerfAltitude:
                            {
                                altitude = telescopeDevice.Altitude;
                                break;
                            }

                        case var @case when @case == PerformanceType.tstPerfAtHome:
                            {
                                atHome = telescopeDevice.AtHome;
                                break;
                            }

                        case PerformanceType.tstPerfAtPark:
                            {
                                atPark = telescopeDevice.AtPark;
                                break;
                            }

                        case PerformanceType.tstPerfAzimuth:
                            {
                                azimuth = telescopeDevice.Azimuth;
                                break;
                            }

                        case PerformanceType.tstPerfDeclination:
                            {
                                declination = telescopeDevice.Declination;
                                break;
                            }

                        case PerformanceType.tstPerfIsPulseGuiding:
                            {
                                isPulseGuiding = telescopeDevice.IsPulseGuiding;
                                break;
                            }

                        case PerformanceType.tstPerfRightAscension:
                            {
                                rightAscension = telescopeDevice.RightAscension;
                                break;
                            }

                        case PerformanceType.tstPerfSideOfPier:
                            {
                                sideOfPier = (PointingState)telescopeDevice.SideOfPier;
                                break;
                            }

                        case PerformanceType.tstPerfSiderealTime:
                            {
                                siderealTimeScope = telescopeDevice.SiderealTime;
                                break;
                            }

                        case PerformanceType.tstPerfSlewing:
                            {
                                slewing = telescopeDevice.Slewing;
                                break;
                            }

                        case PerformanceType.tstPerfUTCDate:
                            {
                                utcDate = telescopeDevice.UTCDate;
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, "Conform:PerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0d)
                    {
                        SetStatus(l_Count + " transactions in " + l_ElapsedTime.ToString("0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        //Application.DoEvents();
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (l_ElapsedTime <= PERF_LOOP_TIME);
                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case var case1 when case1 > 10.0d:
                        {
                            LogInfo("Performance: " + p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case var case2 when 2.0d <= case2 && case2 <= 10.0d:
                        {
                            LogInfo("Performance: " + p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case var case3 when 1.0d <= case3 && case3 <= 2.0d:
                        {
                            LogInfo("Performance: " + p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogInfo("Performance: " + p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogIssue("Performance: " + p_Name, $"Exception {ex.Message}");
            }
        }

        private void TelescopeParkedExceptionTest(ParkedExceptionType p_Type, string p_Name)
        {
            double l_TargetRA;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("Parked:" + p_Name, "About to get AtPark property");
            if (telescopeDevice.AtPark) // We are still parked so test AbortSlew
            {
                try
                {
                    switch (p_Type)
                    {
                        case ParkedExceptionType.tstPExcepAbortSlew:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                                break;
                            }

                        case ParkedExceptionType.tstPExcepFindHome:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call FindHome method");
                                telescopeDevice.FindHome();
                                // Wait for mount to find home
                                WaitWhile("Waiting for mount to home...", () => { return !telescopeDevice.AtHome & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000); }, 200, settings.TelescopeMaximumSlewTime);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepMoveAxisPrimary:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call MoveAxis(Primary, 0.0) method");
                                telescopeDevice.MoveAxis(TelescopeAxis.Primary, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepMoveAxisSecondary:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call MoveAxis(Secondary, 0.0) method");
                                telescopeDevice.MoveAxis(TelescopeAxis.Secondary, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepMoveAxisTertiary:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call MoveAxis(Tertiary, 0.0) method");
                                telescopeDevice.MoveAxis(TelescopeAxis.Tertiary, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepPulseGuide:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call PulseGuide(East, 0.0) method");
                                telescopeDevice.PulseGuide(GuideDirection.East, 0);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToCoordinates:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call SlewToCoordinates method");
                                telescopeDevice.SlewToCoordinates(TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d), 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToCoordinatesAsync:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call SlewToCoordinatesAsync method");
                                telescopeDevice.SlewToCoordinatesAsync(TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d), 0.0d);
                                WaitForSlew("Parked:" + p_Name, "Slewing to coordinates asynchronously");
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToTarget:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to set property TargetRightAscension to " + FormatRA(l_TargetRA));
                                telescopeDevice.TargetRightAscension = l_TargetRA;
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to set property TargetDeclination to 0.0");
                                telescopeDevice.TargetDeclination = 0.0d;
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call SlewToTarget method");
                                telescopeDevice.SlewToTarget();
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToTargetAsync:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to set property to " + FormatRA(l_TargetRA));
                                telescopeDevice.TargetRightAscension = l_TargetRA;
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to set property to 0.0");
                                telescopeDevice.TargetDeclination = 0.0d;
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call method");
                                telescopeDevice.SlewToTargetAsync();
                                WaitForSlew("Parked:" + p_Name, "Slewing to target asynchronously");
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSyncToCoordinates:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call method, RA: " + FormatRA(l_TargetRA) + ", Declination: 0.0");
                                telescopeDevice.SyncToCoordinates(l_TargetRA, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSyncToTarget:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to set property to " + FormatRA(l_TargetRA));
                                telescopeDevice.TargetRightAscension = l_TargetRA;
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to set property to 0.0");
                                telescopeDevice.TargetDeclination = 0.0d;
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("Parked:" + p_Name, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                break;
                            }

                        default:
                            {
                                LogIssue("Parked:" + p_Name, "Conform:ParkedExceptionTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    LogIssue("Parked:" + p_Name, p_Name + " didn't raise an error when Parked as required");
                }
                catch (Exception)
                {
                    LogOK("Parked:" + p_Name, p_Name + " did raise an exception when Parked as required");
                }
                // Check that Telescope is still parked after issuing the command!
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("Parked:" + p_Name, "About to get AtPark property");
                if (!telescopeDevice.AtPark)
                    LogIssue("Parked:" + p_Name, "Telescope was unparked by the " + p_Name + " command. This should not happen!");
            }
            else
            {
                LogIssue("Parked:" + p_Name, "Not parked after Telescope.Park command, " + p_Name + " when parked test skipped");
            }

        }

        private void TelescopeAxisRateTest(string p_Name, TelescopeAxis p_Axis)
        {
            int l_NAxisRates, l_i, l_j;
            bool l_AxisRateOverlap = default, l_AxisRateDuplicate, l_CanGetAxisRates = default, l_HasRates = default;
            int l_Count;

            IAxisRates l_AxisRatesIRates;
            IAxisRates l_AxisRates = null;
            IRate l_Rate = null;

            try
            {
                l_NAxisRates = 0;
                l_AxisRates = null;
                switch (p_Axis)
                {
                    case TelescopeAxis.Primary:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call AxisRates method, Axis: " + ((int)TelescopeAxis.Primary).ToString());
                            l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Primary); // Get primary axis rates
                            break;
                        }

                    case TelescopeAxis.Secondary:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call AxisRates method, Axis: " + ((int)TelescopeAxis.Secondary).ToString());
                            l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Secondary); // Get secondary axis rates
                            break;
                        }

                    case TelescopeAxis.Tertiary:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call AxisRates method, Axis: " + ((int)TelescopeAxis.Tertiary).ToString());
                            l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Tertiary); // Get tertiary axis rates
                            break;
                        }

                    default:
                        {
                            LogIssue("TelescopeAxisRateTest", "Unknown telescope axis: " + p_Axis.ToString());
                            break;
                        }
                }

                try
                {
                    if (l_AxisRates is null)
                    {
                        LogDebug(p_Name, "ERROR: The driver did NOT return an AxisRates object!");
                    }
                    else
                    {
                        LogDebug(p_Name, "OK - the driver returned an AxisRates object");
                    }

                    l_Count = l_AxisRates.Count; // Save count for use later if no members are returned in the for each loop test
                    LogDebug(p_Name + " Count", "The driver returned " + l_Count + " rates");
                    int i;
                    var loopTo = l_Count;
                    for (i = 1; i <= loopTo; i++)
                    {
                        IRate AxisRateItem;

                        AxisRateItem = l_AxisRates[i];
                        LogDebug(p_Name + " Count", "Rate " + i + " - Minimum: " + AxisRateItem.Minimum.ToString() + ", Maximum: " + AxisRateItem.Maximum.ToString());
                    }
                }
                catch (Exception ex)
                {
                    LogIssue(p_Name + " Count", $"Unexpected exception: {ex.Message}");
                }

                try
                {
                    IEnumerator l_Enum;
                    dynamic l_Obj;
                    IRate AxisRateItem = null;

                    l_Enum = (IEnumerator)l_AxisRates.GetEnumerator();
                    if (l_Enum is null)
                    {
                        LogDebug(p_Name + " Enum", "ERROR: The driver did NOT return an Enumerator object!");
                    }
                    else
                    {
                        LogDebug(p_Name + " Enum", "OK - the driver returned an Enumerator object");
                    }

                    l_Enum.Reset();
                    LogDebug(p_Name + " Enum", "Reset Enumerator");
                    while (l_Enum.MoveNext())
                    {
                        LogDebug(p_Name + " Enum", "Reading Current");
                        l_Obj = l_Enum.Current;
                        LogDebug(p_Name + " Enum", "Read Current OK, Type: " + l_Obj.GetType().Name);
                        AxisRateItem = l_Obj;

                        LogDebug(p_Name + " Enum", "Found axis rate - Minimum: " + AxisRateItem.Minimum.ToString() + ", Maximum: " + AxisRateItem.Maximum.ToString());
                    }

                    l_Enum.Reset();
                    l_Enum = null;
                    AxisRateItem = null;
                }
                catch (Exception ex)
                {
                    LogIssue(p_Name + " Enum", $"Exception: {ex}");
                }

                if (l_AxisRates.Count > 0)
                {
                    try
                    {
                        l_AxisRatesIRates = l_AxisRates;
                        foreach (IRate currentL_Rate in l_AxisRatesIRates)
                        {
                            l_Rate = currentL_Rate;
                            if ((l_Rate.Minimum < 0.0d) | (l_Rate.Maximum < 0.0d)) // Error because negative values are not allowed
                            {
                                LogIssue(p_Name, "Minimum or maximum rate is negative: " + l_Rate.Minimum.ToString() + ", " + l_Rate.Maximum.ToString());
                            }
                            else  // All positive values so continue tests
                            {
                                if (l_Rate.Minimum <= l_Rate.Maximum) // Minimum <= Maximum so OK
                                {
                                    LogOK(p_Name, "Axis rate minimum: " + l_Rate.Minimum.ToString() + " Axis rate maximum: " + l_Rate.Maximum.ToString());
                                }
                                else // Minimum > Maximum so error!
                                {
                                    LogIssue(p_Name, "Maximum rate is less than minimum rate - minimum: " + l_Rate.Minimum.ToString() + " maximum: " + l_Rate.Maximum.ToString());
                                }
                            }
                            // Save rates for overlap testing
                            l_NAxisRates += 1;
                            axisRatesArray[l_NAxisRates, AXIS_RATE_MINIMUM] = l_Rate.Minimum;
                            axisRatesArray[l_NAxisRates, AXIS_RATE_MAXIMUM] = l_Rate.Maximum;
                            l_HasRates = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogIssue(p_Name, "Unable to read AxisRates object - Exception: " + ex.Message);
                        LogDebug(p_Name, "Unable to read AxisRates object - Exception: " + ex.ToString());
                    }

                    // Overlap testing
                    if (l_NAxisRates > 1) // Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    {
                        int loopTo1 = l_NAxisRates;
                        for (l_i = 1; l_i <= loopTo1; l_i++)
                        {
                            int loopTo2 = l_NAxisRates;
                            for (l_j = 1; l_j <= loopTo2; l_j++)
                            {
                                if (l_i != l_j) // Only test different lines, shouldn't compare same lines!
                                {
                                    if (axisRatesArray[l_i, AXIS_RATE_MINIMUM] >= axisRatesArray[l_j, AXIS_RATE_MINIMUM] & axisRatesArray[l_i, AXIS_RATE_MINIMUM] <= axisRatesArray[l_j, AXIS_RATE_MAXIMUM])
                                        l_AxisRateOverlap = true;
                                }
                            }
                        }
                    }

                    if (l_AxisRateOverlap)
                    {
                        LogIssue(p_Name, "Overlapping axis rates found, suggest these be rationalised to remove overlaps");
                    }
                    else
                    {
                        LogOK(p_Name, "No overlapping axis rates found");
                    }

                    // Duplicate testing
                    l_AxisRateDuplicate = false;
                    if (l_NAxisRates > 1) // Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    {
                        int loopTo3 = l_NAxisRates;
                        for (l_i = 1; l_i <= loopTo3; l_i++)
                        {
                            int loopTo4 = l_NAxisRates;
                            for (l_j = 1; l_j <= loopTo4; l_j++)
                            {
                                if (l_i != l_j) // Only test different lines, shouldn't compare same lines!
                                {
                                    if (axisRatesArray[l_i, AXIS_RATE_MINIMUM] == axisRatesArray[l_j, AXIS_RATE_MINIMUM] & axisRatesArray[l_i, AXIS_RATE_MAXIMUM] == axisRatesArray[l_j, AXIS_RATE_MAXIMUM])
                                        l_AxisRateDuplicate = true;
                                }
                            }
                        }
                    }

                    if (l_AxisRateDuplicate)
                    {
                        LogIssue(p_Name, "Duplicate axis rates found, suggest these be removed");
                    }
                    else
                    {
                        LogOK(p_Name, "No duplicate axis rates found");
                    }
                }
                else
                {
                    LogOK(p_Name, "Empty axis rate returned");
                }

                l_CanGetAxisRates = true; // Record that this driver can deliver a viable AxisRates object that can be tested for AxisRates.Dispose() later
            }
            catch (Exception ex)
            {
                LogIssue(p_Name, "Unable to get or unable to use an AxisRates object - Exception: " + ex.ToString());
            }

            // Clean up AxisRate object if used
            if (l_AxisRates is object)
            {
                try
                {
                    l_AxisRates.Dispose();
                }
                catch
                {
                }

                if (OperatingSystem.IsWindows())
                {
                    try
                    {
                        Marshal.ReleaseComObject(l_AxisRates);
                    }
                    catch
                    {
                    }
                }

                l_AxisRates = null;
            }

            // Clean up and release rate object if used
            if (l_Rate is object)
            {
                try
                {
                    l_Rate.Dispose();
                }
                catch
                {
                }

                if (OperatingSystem.IsWindows())
                {
                    try
                    {
                        Marshal.ReleaseComObject(l_Rate);
                    }
                    catch { }
                }
            }

            if (l_CanGetAxisRates) // The driver does return a viable AxisRates object that can be tested for correct AxisRates.Dispose() and Rate.Dispose() operation
            {
                try
                {
                    // Test Rate.Dispose()
                    switch (p_Axis) // Get the relevant axis rates object for this axis
                    {
                        case TelescopeAxis.Primary:
                            {
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Primary);
                                break;
                            }

                        case TelescopeAxis.Secondary:
                            {
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Secondary);
                                break;
                            }

                        case TelescopeAxis.Tertiary:
                            {
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Tertiary);
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, "AxisRate.Dispose() - Unknown axis: " + p_Axis.ToString());
                                break;
                            }
                    }

                    if (l_HasRates) // This axis does have one or more rates that can be accessed through ForEach so test these for correct Rate.Dispose() action
                    {
                        foreach (IRate currentL_Rate2 in (IEnumerable)l_AxisRates)
                        {
                            l_Rate = currentL_Rate2;
                            try
                            {
                                l_Rate.Dispose();
                                LogOK(p_Name, string.Format("Successfully disposed of rate {0} - {1}", l_Rate.Minimum, l_Rate.Maximum));
                            }
                            catch (MissingMemberException)
                            {
                                LogOK(p_Name, string.Format("Rate.Dispose() member not present for rate {0} - {1}", l_Rate.Minimum, l_Rate.Maximum));
                            }
                            catch (Exception ex1)
                            {
                                LogIssue(p_Name, string.Format("Rate.Dispose() for rate {0} - {1} threw an exception but it is poor practice to throw exceptions in Dispose methods: {2}", l_Rate.Minimum, l_Rate.Maximum, ex1.Message));
                                LogDebug("TrackingRates.Dispose", "Exception: " + ex1.ToString());
                            }
                        }
                    }

                    // Test AxisRates.Dispose()
                    try
                    {
                        LogDebug(p_Name, "Disposing axis rates");
                        l_AxisRates.Dispose();
                        LogOK(p_Name, "Disposed axis rates OK");
                    }
                    catch (MissingMemberException)
                    {
                        LogOK(p_Name, "AxisRates.Dispose() member not present for axis " + p_Axis.ToString());
                    }
                    catch (Exception ex1)
                    {
                        LogIssue(p_Name, "AxisRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: " + ex1.Message);
                        LogDebug("AxisRates.Dispose", "Exception: " + ex1.ToString());
                    }
                }
                catch (Exception ex)
                {
                    LogIssue(p_Name, "AxisRate.Dispose() - Unable to get or unable to use an AxisRates object - Exception: " + ex.ToString());
                }
            }
            else
            {
                LogInfo(p_Name, "AxisRates.Dispose() testing skipped because of earlier issues in obtaining a viable AxisRates object.");
            }

        }

        private void TelescopeRequiredMethodsTest(RequiredMethodType p_Type, string p_Name)
        {
            try
            {
                switch (p_Type)
                {
                    case RequiredMethodType.tstAxisrates:
                        {
                            break;
                        }
                    // This is now done by TelescopeAxisRateTest subroutine 
                    case RequiredMethodType.tstCanMoveAxisPrimary:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call CanMoveAxis method " + ((int)TelescopeAxis.Primary).ToString());
                            canMoveAxisPrimary = telescopeDevice.CanMoveAxis(TelescopeAxis.Primary);
                            LogOK(p_Name, p_Name + " " + canMoveAxisPrimary.ToString());
                            break;
                        }

                    case RequiredMethodType.tstCanMoveAxisSecondary:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call CanMoveAxis method " + ((int)TelescopeAxis.Secondary).ToString());
                            canMoveAxisSecondary = telescopeDevice.CanMoveAxis(TelescopeAxis.Secondary);
                            LogOK(p_Name, p_Name + " " + canMoveAxisSecondary.ToString());
                            break;
                        }

                    case RequiredMethodType.tstCanMoveAxisTertiary:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(p_Name, "About to call CanMoveAxis method " + ((int)TelescopeAxis.Tertiary).ToString());
                            canMoveAxisTertiary = telescopeDevice.CanMoveAxis(TelescopeAxis.Tertiary);
                            LogOK(p_Name, p_Name + " " + canMoveAxisTertiary.ToString());
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "Conform:RequiredMethodsTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Method, Required.Mandatory, ex, "");
            }

            // Clean up and release each object after use
            // If Not (m_Rate Is Nothing) Then Try : Marshal.ReleaseComObject(m_Rate) : Catch : End Try
            // m_Rate = Nothing
        }

        private void TelescopeOptionalMethodsTest(OptionalMethodType testType, string testName, bool canTest)
        {
            double l_TestDec, l_TestRAOffset;
            IAxisRates l_AxisRates = null;

            SetTest(testName);
            LogDebug("TelescopeOptionalMethodsTest", testType.ToString() + " " + testName + " " + canTest.ToString());
            if (canTest) // Confirm that an error is raised if the optional command is not implemented
            {
                try
                {
                    // Set the test declination value depending on whether the scope is in the northern or southern hemisphere
                    if (siteLatitude > 0.0d)
                    {
                        l_TestDec = 45.0d; // Positive for the northern hemisphere
                    }
                    else
                    {
                        l_TestDec = -45.0d;
                    } // Negative for the southern hemisphere

                    l_TestRAOffset = 3.0d; // Set the test RA offset as 3 hours from local sider5eal time
                    LogDebug(testName, string.Format("Test RA offset: {0}, Test declination: {1}", l_TestRAOffset, l_TestDec));
                    switch (testType)
                    {
                        case OptionalMethodType.AbortSlew:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                                LogOK("AbortSlew", "AbortSlew OK when not slewing");
                                break;
                            }

                        case OptionalMethodType.DestinationSideOfPier:
                            {
                                // Get the DestinationSideOfPier for a target in the West i.e. for a German mount when the tube is on the East side of the pier
                                targetRightAscension = TelescopeRAFromSiderealTime(testName, -l_TestRAOffset);
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call DestinationSideOfPier method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(l_TestDec));
                                destinationSideOfPierEast = (PointingState)telescopeDevice.DestinationSideOfPier(targetRightAscension, l_TestDec);
                                LogDebug(testName, "German mount - scope on the pier's East side, target in the West : " + FormatRA(targetRightAscension) + " " + FormatDec(l_TestDec) + " " + destinationSideOfPierEast.ToString());

                                // Get the DestinationSideOfPier for a target in the East i.e. for a German mount when the tube is on the West side of the pier
                                targetRightAscension = TelescopeRAFromSiderealTime(testName, l_TestRAOffset);
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call DestinationSideOfPier method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(l_TestDec));
                                destinationSideOfPierWest = (PointingState)telescopeDevice.DestinationSideOfPier(targetRightAscension, l_TestDec);
                                LogDebug(testName, "German mount - scope on the pier's West side, target in the East: " + FormatRA(targetRightAscension) + " " + FormatDec(l_TestDec) + " " + destinationSideOfPierWest.ToString());

                                // Make sure that we received two valid values i.e. that neither side returned PierSide.Unknown and that the two valid returned values are not the same i.e. we got one PointingState.Normal and one PointingState.ThroughThePole
                                if (destinationSideOfPierEast == PointingState.Unknown | destinationSideOfPierWest == PointingState.Unknown)
                                {
                                    LogIssue(testName, "Invalid SideOfPier value received, Target in West: " + destinationSideOfPierEast.ToString() + ", Target in East: " + destinationSideOfPierWest.ToString());
                                }
                                else if (destinationSideOfPierEast == destinationSideOfPierWest)
                                {
                                    LogIssue(testName, "Same value for DestinationSideOfPier received on both sides of the meridian: " + ((int)destinationSideOfPierEast).ToString());
                                }
                                else
                                {
                                    LogOK(testName, "DestinationSideOfPier is different on either side of the meridian");
                                }

                                break;
                            }

                        case OptionalMethodType.FindHome:
                            {
                                if (interfaceVersion > 1)
                                {
                                    startTime = DateTime.Now;

                                    SetAction("Homing mount...");
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to call FindHome method");
                                    telescopeDevice.FindHome();

                                    // Wait for mount to find home
                                    WaitWhile("Waiting for mount to home...", () => { return !telescopeDevice.AtHome & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000); }, 200, settings.TelescopeMaximumSlewTime);

                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get AtHome property");
                                    if (telescopeDevice.AtHome)
                                    {
                                        LogOK(testName, "Found home OK.");
                                    }
                                    else
                                    {
                                        LogInfo(testName, "Failed to Find home within 1 minute");
                                    }

                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get AtPark property");
                                    if (telescopeDevice.AtPark)
                                    {
                                        LogIssue(testName, "FindHome has parked the scope as well as finding home");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to call UnPark method");
                                        telescopeDevice.Unpark(); // Unpark it ready for further tests
                                        LogCallToDriver("Unpark", "About to get AtPark property repeatedly");

                                        WaitWhile("Waiting for scope to unpark", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                    }

                                    // Check if the operation properties are implemented
                                    if (interfaceVersion >= 4) // Operations are supported
                                    {
                                        SlewToHa(1.0);

                                        startTime = DateTime.Now;

                                        SetAction("Homing mount using operations...");

                                        // Validate OperationComplete state
                                        ValidateOperationComplete("Park", true);

                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to call FindHome method");

                                        TimeMethod("FindHome", () => telescopeDevice.FindHome());

                                        // Wait for mount to find home
                                        WaitWhile("Waiting for mount to home using operations...", () => { return !telescopeDevice.OperationComplete & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000); }, 200, settings.TelescopeMaximumSlewTime);

                                        // Validate OperationComplete state
                                        ValidateOperationComplete("Park", true);

                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to get AtHome property");
                                        if (telescopeDevice.AtHome)
                                        {
                                            LogOK(testName, "Found home OK using operations.");
                                        }
                                        else
                                        {
                                            LogInfo(testName, "Failed to Find home within 1 minute using operations");
                                        }

                                    }
                                }
                                else
                                {
                                    SetAction("Waiting for mount to home");
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to call FindHome method");
                                    telescopeDevice.FindHome();
                                    // Wait for mount to find home
                                    WaitWhile("Waiting for mount to home...", () => { return !telescopeDevice.AtHome & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000); }, 200, settings.TelescopeMaximumSlewTime);

                                    LogOK(testName, "Found home OK.");
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to call Unpark method");
                                    telescopeDevice.Unpark();
                                    LogCallToDriver("Unpark", "About to get AtPark property repeatedly");
                                    WaitWhile("Waiting for scope to unpark", () => { return telescopeDevice.AtPark; }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
                                } // Make sure we are still  unparked!

                                break;
                            }

                        case OptionalMethodType.MoveAxisPrimary:
                            {
                                // Get axis rates for primary axis
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call AxisRates method for axis " + ((int)TelescopeAxis.Primary).ToString());
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Primary);
                                TelescopeMoveAxisTest(testName, TelescopeAxis.Primary, l_AxisRates);
                                break;
                            }

                        case OptionalMethodType.MoveAxisSecondary:
                            {
                                // Get axis rates for secondary axis
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Secondary);
                                TelescopeMoveAxisTest(testName, TelescopeAxis.Secondary, l_AxisRates);
                                break;
                            }

                        case OptionalMethodType.MoveAxisTertiary:
                            {
                                // Get axis rates for tertiary axis
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxis.Tertiary);
                                TelescopeMoveAxisTest(testName, TelescopeAxis.Tertiary, l_AxisRates);
                                break;
                            }

                        case OptionalMethodType.PulseGuide:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get IsPulseGuiding property");
                                if (telescopeDevice.IsPulseGuiding) // IsPulseGuiding is true before we've started so this is an error and voids a real test
                                {
                                    LogIssue(testName, "IsPulseGuiding is True when not pulse guiding - PulseGuide test omitted");
                                }
                                else // OK to test pulse guiding
                                {
                                    SetAction("Calling PulseGuide east");
                                    startTime = DateTime.Now;
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to call PulseGuide method, Direction: " + ((int)GuideDirection.East).ToString() + ", Duration: " + PULSEGUIDE_MOVEMENT_TIME * 1000 + "ms");
                                    telescopeDevice.PulseGuide(GuideDirection.East, PULSEGUIDE_MOVEMENT_TIME * 1000); // Start a 2 second pulse
                                    endTime = DateTime.Now;
                                    LogDebug(testName, "PulseGuide command time: " + PULSEGUIDE_MOVEMENT_TIME * 1000 + " milliseconds, PulseGuide call duration: " + endTime.Subtract(startTime).TotalMilliseconds + " milliseconds");
                                    if (endTime.Subtract(startTime).TotalMilliseconds < PULSEGUIDE_MOVEMENT_TIME * 0.75d * 1000d) // If less than three quarters of the expected duration then assume we have returned early
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to get IsPulseGuiding property");
                                        if (telescopeDevice.IsPulseGuiding)
                                        {
                                            if (settings.DisplayMethodCalls)
                                                LogTestAndMessage(testName, "About to get IsPulseGuiding property multiple times");
                                            WaitWhile("Pulse guiding Eastwards", () => { return telescopeDevice.IsPulseGuiding; }, SLEEP_TIME, PULSEGUIDE_TIMEOUT_TIME);

                                            if (settings.DisplayMethodCalls)
                                                LogTestAndMessage(testName, "About to get IsPulseGuiding property");
                                            if (!telescopeDevice.IsPulseGuiding)
                                            {
                                                LogOK(testName, "Asynchronous pulse guide found OK");
                                                LogDebug(testName, "IsPulseGuiding = True duration: " + DateTime.Now.Subtract(startTime).TotalMilliseconds + " milliseconds");
                                            }
                                            else
                                            {
                                                LogIssue(testName, "Asynchronous pulse guide expected but IsPulseGuiding is still TRUE " + PULSEGUIDE_TIMEOUT_TIME + " seconds beyond expected time");
                                            }
                                        }
                                        else
                                        {
                                            LogIssue(testName, "Asynchronous pulse guide expected but IsPulseGuiding has returned FALSE");
                                        }
                                    }
                                    else // Assume synchronous pulse guide and that IsPulseGuiding is false
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to get IsPulseGuiding property");
                                        if (!telescopeDevice.IsPulseGuiding)
                                        {
                                            LogOK(testName, "Synchronous pulse guide found OK");
                                        }
                                        else
                                        {
                                            LogIssue(testName, "Synchronous pulse guide expected but IsPulseGuiding has returned TRUE");
                                        }
                                    }
                                }

                                break;
                            }

                        case OptionalMethodType.SideOfPierWrite:
                            {
                                // SideOfPier Write
                                if (canSetPierside) // Can set pier side so test if we can
                                {
                                    SlewScope(TelescopeRAFromHourAngle(testName, -3.0d), 0.0d, "far start point");
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    SlewScope(TelescopeRAFromHourAngle(testName, -0.03d), 0.0d, "near start point"); // 2 minutes from zenith
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

                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to get SideOfPier property");
                                    switch (telescopeDevice.SideOfPier)
                                    {
                                        case PointingState.Normal: // We are on pierEast so try pierWest
                                            {
                                                try
                                                {
                                                    LogDebug(testName, "Scope is pierEast so flipping West");
                                                    SetAction("Flipping mount to pointing state pierWest");
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage(testName, "About to set SideOfPier property to " + ((int)PointingState.ThroughThePole).ToString());
                                                    telescopeDevice.SideOfPier = PointingState.ThroughThePole;
                                                    WaitForSlew(testName, $"Moving to the pierEast pointing state asynchronously");

                                                    if (cancellationToken.IsCancellationRequested)
                                                        return;

                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage(testName, "About to get SideOfPier property");
                                                    sideOfPier = (PointingState)telescopeDevice.SideOfPier;
                                                    if (sideOfPier == PointingState.ThroughThePole)
                                                    {
                                                        LogOK(testName, "Successfully flipped pierEast to pierWest");
                                                    }
                                                    else
                                                    {
                                                        LogIssue(testName, "Failed to set SideOfPier to pierWest, got: " + sideOfPier.ToString());
                                                    }
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
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage(testName, "About to set SideOfPier property to " + ((int)PointingState.Normal).ToString());
                                                    telescopeDevice.SideOfPier = PointingState.Normal;
                                                    WaitForSlew(testName, $"Moving to the pierWest pointing state asynchronously");
                                                    if (cancellationToken.IsCancellationRequested)
                                                        return;

                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage(testName, "About to get SideOfPier property");
                                                    sideOfPier = (PointingState)telescopeDevice.SideOfPier;
                                                    if (sideOfPier == PointingState.Normal)
                                                    {
                                                        LogOK(testName, "Successfully flipped pierWest to pierEast");
                                                    }
                                                    else
                                                    {
                                                        LogIssue(testName, "Failed to set SideOfPier to pierEast, got: " + sideOfPier.ToString());
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleException("SideOfPier Write pierEast", MemberType.Method, Required.MustBeImplemented, ex, "CanSetPierSide is True");
                                                }

                                                break;
                                            }

                                        default:
                                            {
                                                LogIssue(testName, "Unknown PierSide: " + sideOfPier.ToString());
                                                break;
                                            }
                                    }
                                }
                                else // Can't set pier side so it should generate an error
                                {
                                    try
                                    {
                                        LogDebug(testName, "Attempting to set SideOfPier");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to set SideOfPier property to " + ((int)PointingState.Normal).ToString());
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

                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to set Tracking property to false");
                                telescopeDevice.Tracking = false;
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                break;
                            }

                        default:
                            {
                                LogIssue(testName, "Conform:OptionalMethodsTest: Unknown test type " + testType.ToString());
                                break;
                            }
                    }

                    // Clean up AxisRate object, if used
                    if (l_AxisRates is object)
                    {
                        try
                        {
                            LogCallToDriver(testName, "About to dispose of AxisRates object");
                            l_AxisRates.Dispose();

                            LogOK(testName, "AxisRates object successfully disposed");
                        }
                        catch (Exception ex)
                        {
                            LogIssue(testName, "AxisRates.Dispose threw an exception but must not: " + ex.Message);
                            LogDebug(testName, "Exception: " + ex.ToString());
                        }

                        try
                        {
#if WINDOWS
                            Marshal.ReleaseComObject(l_AxisRates);
#endif
                        }
                        catch
                        {
                        }

                        l_AxisRates = null;
                    }
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
                catch (Exception ex)
                {
                    LogIssue(testName, $"MoveAxis Exception\r\n{ex}");
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
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                                break;
                            }

                        case OptionalMethodType.DestinationSideOfPier:
                            {
                                targetRightAscension = TelescopeRAFromSiderealTime(testName, -1.0d);
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call DestinationSideOfPier method, RA: " + FormatRA(targetRightAscension) + ", Declination: " + FormatDec(0.0d));
                                destinationSideOfPier = (PointingState)telescopeDevice.DestinationSideOfPier(targetRightAscension, 0.0d);
                                break;
                            }

                        case OptionalMethodType.FindHome:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call FindHome method");
                                telescopeDevice.FindHome();
                                // Wait for mount to find home
                                WaitWhile("Waiting for mount to home...", () => { return !telescopeDevice.AtHome & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000); }, 200, settings.TelescopeMaximumSlewTime);
                                break;
                            }

                        case OptionalMethodType.MoveAxisPrimary:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)TelescopeAxis.Primary).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxis.Primary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.MoveAxisSecondary:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)TelescopeAxis.Secondary).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxis.Secondary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.MoveAxisTertiary:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)TelescopeAxis.Tertiary).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxis.Tertiary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.PulseGuide:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call PulseGuide method, Direction: " + ((int)GuideDirection.East).ToString() + ", Duration: 0ms");
                                telescopeDevice.PulseGuide(GuideDirection.East, 0);
                                break;
                            }

                        case OptionalMethodType.SideOfPierWrite:
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to set SideOfPier property to " + ((int)PointingState.Normal).ToString());
                                telescopeDevice.SideOfPier = PointingState.Normal;
                                break;
                            }

                        default:
                            {
                                LogIssue(testName, "Conform:OptionalMethodsTest: Unknown test type " + testType.ToString());
                                break;
                            }
                    }

                    LogIssue(testName, "Can" + testName + " is false but no exception was generated on use");
                }
                catch (Exception ex)
                {
                    if (IsInvalidValueException(testName, ex))
                    {
                        LogOK(testName, "Received an invalid value exception");
                    }
                    else if (testType == OptionalMethodType.SideOfPierWrite) // PierSide is actually a property even though I have it in the methods section!!
                    {
                        HandleException(testName, MemberType.Property, Required.MustNotBeImplemented, ex, "Can" + testName + " is False");
                    }
                    else
                    {
                        HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, "Can" + testName + " is False");
                    }
                }
            }
            ClearStatus();

        }

        private void TelescopeCanTest(CanType p_Type, string p_Name)
        {
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage(p_Name, string.Format("About to get {0} property", p_Type.ToString()));
                switch (p_Type)
                {
                    case CanType.CanFindHome:
                        {
                            canFindHome = telescopeDevice.CanFindHome;
                            LogOK(p_Name, canFindHome.ToString());
                            break;
                        }

                    case CanType.CanPark:
                        {
                            canPark = telescopeDevice.CanPark;
                            LogOK(p_Name, canPark.ToString());
                            break;
                        }

                    case CanType.CanPulseGuide:
                        {
                            canPulseGuide = telescopeDevice.CanPulseGuide;
                            LogOK(p_Name, canPulseGuide.ToString());
                            break;
                        }

                    case CanType.CanSetDeclinationRate:
                        {
                            if (interfaceVersion > 1)
                            {
                                canSetDeclinationRate = telescopeDevice.CanSetDeclinationRate;
                                LogOK(p_Name, canSetDeclinationRate.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetDeclinationRate", "Skipping test as this method is not supported in interface V" + interfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetGuideRates:
                        {
                            if (interfaceVersion > 1)
                            {
                                canSetGuideRates = telescopeDevice.CanSetGuideRates;
                                LogOK(p_Name, canSetGuideRates.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetGuideRates", "Skipping test as this method is not supported in interface V" + interfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetPark:
                        {
                            canSetPark = telescopeDevice.CanSetPark;
                            LogOK(p_Name, canSetPark.ToString());
                            break;
                        }

                    case CanType.CanSetPierSide:
                        {
                            if (interfaceVersion > 1)
                            {
                                canSetPierside = telescopeDevice.CanSetPierSide;
                                LogOK(p_Name, canSetPierside.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetPierSide", "Skipping test as this method is not supported in interface V" + interfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetRightAscensionRate:
                        {
                            if (interfaceVersion > 1)
                            {
                                canSetRightAscensionRate = telescopeDevice.CanSetRightAscensionRate;
                                LogOK(p_Name, canSetRightAscensionRate.ToString());
                            }
                            else
                            {
                                LogInfo("CanSetRightAscensionRate", "Skipping test as this method is not supported in interface V" + interfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetTracking:
                        {
                            canSetTracking = telescopeDevice.CanSetTracking;
                            LogOK(p_Name, canSetTracking.ToString());
                            break;
                        }

                    case CanType.CanSlew:
                        {
                            canSlew = telescopeDevice.CanSlew;
                            LogOK(p_Name, canSlew.ToString());
                            break;
                        }

                    case CanType.CanSlewAltAz:
                        {
                            if (interfaceVersion > 1)
                            {
                                canSlewAltAz = telescopeDevice.CanSlewAltAz;
                                LogOK(p_Name, canSlewAltAz.ToString());
                            }
                            else
                            {
                                LogInfo("CanSlewAltAz", "Skipping test as this method is not supported in interface V" + interfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSlewAltAzAsync:
                        {
                            if (interfaceVersion > 1)
                            {
                                canSlewAltAzAsync = telescopeDevice.CanSlewAltAzAsync;
                                LogOK(p_Name, canSlewAltAzAsync.ToString());
                            }
                            else
                            {
                                LogInfo("CanSlewAltAzAsync", "Skipping test as this method is not supported in interface V" + interfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSlewAsync:
                        {
                            canSlewAsync = telescopeDevice.CanSlewAsync;
                            LogOK(p_Name, canSlewAsync.ToString());
                            break;
                        }

                    case CanType.CanSync:
                        {
                            canSync = telescopeDevice.CanSync;
                            LogOK(p_Name, canSync.ToString());
                            break;
                        }

                    case CanType.CanSyncAltAz:
                        {
                            if (interfaceVersion > 1)
                            {
                                canSyncAltAz = telescopeDevice.CanSyncAltAz;
                                LogOK(p_Name, canSyncAltAz.ToString());
                            }
                            else
                            {
                                LogInfo("CanSyncAltAz", "Skipping test as this method is not supported in interface V" + interfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanUnPark:
                        {
                            canUnpark = telescopeDevice.CanUnpark;
                            LogOK(p_Name, canUnpark.ToString());
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "Conform:CanTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
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
                IAxisRates l_AxisRatesIRates = axisRates;
                rateCount = 0;
                foreach (IRate currentL_Rate in l_AxisRatesIRates)
                {
                    rate = currentL_Rate;
                    if (rate.Minimum < rateMinimum) rateMinimum = rate.Minimum;
                    if (rate.Maximum > rateMaximum) rateMaximum = rate.Maximum;
                    LogDebug(testName, $"Checking rates: {rate.Minimum} {rate.Maximum} Current rates: {rateMinimum} {rateMaximum}");
                    rateCount += 1;
                }

                if (rateMinimum != double.PositiveInfinity & rateMaximum != double.NegativeInfinity) // Found valid rates
                {
                    LogDebug(testName, "Found minimum rate: " + rateMinimum + " found maximum rate: " + rateMaximum);

                    // Confirm setting a zero rate works
                    SetAction("Set zero rate");
                    canSetZero = false;
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                        telescopeDevice.MoveAxis(testAxis, 0.0d); // Set a value of zero
                        LogOK(testName, "Can successfully set a movement rate of zero");
                        canSetZero = true;
                    }
                    catch (COMException ex)
                    {
                        LogIssue(testName, "Unable to set a movement rate of zero - " + ex.Message + " " + ((int)ex.ErrorCode).ToString("X8"));
                    }
                    catch (DriverException ex)
                    {
                        LogIssue(testName, "Unable to set a movement rate of zero - " + ex.Message + " " + ex.Number.ToString("X8"));
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, "Unable to set a movement rate of zero - " + ex.Message);
                    }

                    SetAction("Set lower rate");

                    // Test that error is generated on attempt to set rate lower than minimum
                    try
                    {
                        if (rateMinimum > 0d) // choose a value between the minimum and zero
                        {
                            moveRate = rateMinimum / 2.0d;
                        }
                        else // Choose a large negative value
                        {
                            moveRate = -rateMaximum - 1.0d;
                        }

                        LogDebug(testName, "Using minimum rate: " + moveRate);
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + moveRate);
                        telescopeDevice.MoveAxis(testAxis, moveRate); // Set a value lower than the minimum
                        LogIssue(testName, "No exception raised when move axis value < minimum rate: " + moveRate);
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
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsOK(testName, MemberType.Method, Required.MustBeImplemented, ex, "when move axis is set below lowest rate (" + moveRate + ")", "Exception correctly generated when move axis is set below lowest rate (" + moveRate + ")");
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

                    // test that error is generated when rate is above maximum set
                    SetAction("Set upper rate");
                    try
                    {
                        moveRate = rateMaximum + 1.0d;
                        LogDebug(testName, "Using maximum rate: " + moveRate);
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + moveRate);
                        telescopeDevice.MoveAxis(testAxis, moveRate); // Set a value higher than the maximum
                        LogIssue(testName, "No exception raised when move axis value > maximum rate: " + moveRate);
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
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsOK(testName, MemberType.Method, Required.MustBeImplemented, ex, "when move axis is set above highest rate (" + moveRate + ")", "Exception correctly generated when move axis is set above highest rate (" + moveRate + ")");
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
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + rateMinimum);
                                telescopeDevice.MoveAxis(testAxis, rateMinimum); // Set the minimum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                SetAction("Stopping movement");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis

                                SetAction("Moving back at minimum rate");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + -rateMinimum);
                                telescopeDevice.MoveAxis(testAxis, -rateMinimum); // Set the minimum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;

                                SetAction("Stopping movement");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                LogOK(testName, "Successfully moved axis at minimum rate: " + rateMinimum);
                            }
                            catch (Exception ex)
                            {
                                HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, "when setting rate: " + rateMinimum);
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
                                // Confirm not slewing first
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Slewing property");
                                if (telescopeDevice.Slewing)
                                {
                                    LogIssue(testName, "Slewing was true before start of MoveAxis but should have been false, remaining tests skipped");
                                    return;
                                }

                                // If ITelescopeV4 or later check OperationComplete state
                                if (interfaceVersion >= 4)
                                    ValidateOperationComplete(testName, true);

                                SetStatus("Moving forward at highest rate");

                                if (interfaceVersion >= 4)// If ITelescopeV4 or later check OperationComplete state
                                {
                                    ValidateOperationComplete(testName, true);
                                    LogCallToDriver(testName, "About to call MoveAxis method (OperationComplete) for axis " + ((int)testAxis).ToString() + " at speed " + rateMaximum);
                                    telescopeDevice.MoveAxis(testAxis, rateMaximum); // Set the maximum rate

                                    // Wait for mount to get to speed
                                    WaitWhile("Waiting for asynchronous move operation to start...", () => { return telescopeDevice.OperationComplete & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000); }, 200, settings.TelescopeMaximumSlewTime);

                                    ValidateOperationComplete(testName + " (Set Maximum)", false);
                                }
                                else // ITelescopeV3 or earlier
                                {
                                    LogCallToDriver(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + rateMaximum);
                                    telescopeDevice.MoveAxis(testAxis, rateMaximum); // Set the maximum rate
                                }

                                // Confirm that slewing is active when the move is underway
                                LogCallToDriver(testName, "About to get Slewing property");
                                if (!telescopeDevice.Slewing)
                                    LogIssue(testName, "Slewing is not true immediately after axis starts moving in positive direction");
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                LogCallToDriver(testName, "About to get Slewing property");
                                if (!telescopeDevice.Slewing)
                                    LogIssue(testName, "Slewing is not true after " + MOVE_AXIS_TIME / 1000d + " seconds moving in positive direction");

                                SetStatus("Stopping movement");

                                if (interfaceVersion >= 4)// If ITelescopeV4 or later check OperationComplete state
                                {
                                    ValidateOperationComplete(testName, false);
                                    LogCallToDriver(testName, "About to call MoveAxis method (OperationComplete) for axis " + ((int)testAxis).ToString() + " at speed 0");
                                    telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis

                                    // Wait for mount to get to speed
                                    WaitWhile("Waiting for asynchronous move operation to start...", () => { return !telescopeDevice.OperationComplete & (DateTime.Now.Subtract(startTime).TotalMilliseconds < 60000); }, 200, settings.TelescopeMaximumSlewTime);

                                    ValidateOperationComplete(testName + " (Set 0.0)", true);
                                }
                                else // ITelescopeV3 or earlier
                                {
                                    LogCallToDriver(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                    telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                }

                                LogCallToDriver(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                                                          // Confirm that slewing is false when movement is stopped
                                LogCallToDriver(testName, "About to get property");
                                if (telescopeDevice.Slewing)
                                {
                                    LogIssue(testName, "Slewing incorrectly remains true after stopping positive axis movement, remaining test skipped");
                                    return;
                                }

                                SetStatus("Moving backward at highest rate");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + -rateMaximum);
                                telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the minimum rate
                                                                                  // Confirm that slewing is active when the move is underway
                                if (!telescopeDevice.Slewing)
                                    LogIssue(testName, "Slewing is not true immediately after axis starts moving in negative direction");
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                if (!telescopeDevice.Slewing)
                                    LogIssue(testName, "Slewing is not true after " + MOVE_AXIS_TIME / 1000d + " seconds moving in negative direction");
                                // Confirm that slewing is false when movement is stopped
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                                                          // Confirm that slewing is false when movement is stopped
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Slewing property");
                                if (telescopeDevice.Slewing)
                                {
                                    LogIssue(testName, "Slewing incorrectly remains true after stopping negative axis movement, remaining test skipped");
                                    return;
                                }

                                LogOK(testName, "Successfully moved axis at maximum rate: " + rateMaximum);
                            }
                            catch (Exception ex)
                            {
                                HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, "when setting rate: " + rateMaximum);
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
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                trackingStart = telescopeDevice.Tracking; // Save the start tracking state
                                SetStatus("Moving forward");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + rateMaximum);
                                telescopeDevice.MoveAxis(testAxis, rateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                SetStatus("Stop movement");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                trackingEnd = telescopeDevice.Tracking; // Save the final tracking state
                                if (trackingStart == trackingEnd) // Successfully retained tracking state
                                {
                                    if (trackingStart) // Tracking is true so switch to false for return movement
                                    {
                                        SetStatus("Set tracking off");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to set Tracking property false");
                                        telescopeDevice.Tracking = false;
                                        SetStatus("Move back");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + -rateMaximum);
                                        telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                        WaitFor(MOVE_AXIS_TIME);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                        telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                        SetStatus("");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to get Tracking property");
                                        if (telescopeDevice.Tracking == false) // tracking correctly retained in both states
                                        {
                                            LogOK(testName, "Tracking state correctly retained for both tracking states");
                                        }
                                        else
                                        {
                                            LogIssue(testName, "Tracking state correctly retained when tracking is " + trackingStart.ToString() + ", but not when tracking is false");
                                        }
                                    }
                                    else // Tracking false so switch to true for return movement
                                    {
                                        SetStatus("Set tracking on");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to set Tracking property true");
                                        telescopeDevice.Tracking = true;
                                        SetStatus("Move back");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + -rateMaximum);
                                        telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                        WaitFor(MOVE_AXIS_TIME);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                        telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                        SetStatus("");
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage(testName, "About to get Tracking property");
                                        if (telescopeDevice.Tracking == true) // tracking correctly retained in both states
                                        {
                                            LogOK(testName, "Tracking state correctly retained for both tracking states");
                                        }
                                        else
                                        {
                                            LogIssue(testName, "Tracking state correctly retained when tracking is " + trackingStart.ToString() + ", but not when tracking is true");
                                        }
                                    }

                                    SetStatus(""); // Clear status flag
                                }
                                else // Tracking state not correctly restored
                                {
                                    SetStatus("Move back");
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + -rateMaximum);
                                    telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                    WaitFor(MOVE_AXIS_TIME);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    SetStatus("");
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                    telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set Tracking property " + trackingStart);
                                    telescopeDevice.Tracking = trackingStart; // Restore original value
                                    LogIssue(testName, "Tracking state not correctly restored after MoveAxis when CanSetTracking is true");
                                }
                            }
                            else // Can't set tracking so just test the current state
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                trackingStart = telescopeDevice.Tracking;
                                SetStatus("Moving forward");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed " + rateMaximum);
                                telescopeDevice.MoveAxis(testAxis, rateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                SetStatus("Stop movement");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to get Tracking property");
                                trackingEnd = telescopeDevice.Tracking; // Save tracking state
                                SetStatus("Move back");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call method MoveAxis for axis " + ((int)testAxis).ToString() + " at speed " + -rateMaximum);
                                telescopeDevice.MoveAxis(testAxis, -rateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                // v1.0.12 next line added because movement wasn't stopped
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage(testName, "About to call MoveAxis method for axis " + ((int)testAxis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(testAxis, 0.0d); // Stop the movement on this axis
                                if (trackingStart == trackingEnd)
                                {
                                    LogOK(testName, "Tracking state correctly restored after MoveAxis when CanSetTracking is false");
                                }
                                else
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage(testName, "About to set Tracking property to " + trackingStart);
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
                    LogInfo(testName, "Found minimum rate: " + rateMinimum + " found maximum rate: " + rateMaximum);
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
            SideOfPierResults l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9;
            double l_Declination3, l_Declination9, l_StartRA;

            // Slew to starting position
            LogDebug("SideofPier", "Starting Side of Pier tests");
            SetTest("Side of pier tests");
            l_StartRA = TelescopeRAFromHourAngle("SideofPier", -3.0d);
            if (siteLatitude > 0.0d) // We are in the northern hemisphere
            {
                l_Declination3 = 90.0d - (180.0d - siteLatitude) * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR; // Calculate for northern hemisphere
                l_Declination9 = 90.0d - siteLatitude * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR;
            }
            else // We are in the southern hemisphere
            {
                l_Declination3 = -90.0d + (180.0d + siteLatitude) * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR; // Calculate for southern hemisphere
                l_Declination9 = -90.0d - siteLatitude * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR;
            }

            LogDebug("SideofPier", "Declination for hour angle = +-3.0 tests: " + FormatDec(l_Declination3) + ", Declination for hour angle = +-9.0 tests: " + FormatDec(l_Declination9));
            SlewScope(l_StartRA, 0.0d, "starting position");
            if (cancellationToken.IsCancellationRequested)
                return;

            // Run tests
            SetAction("Test hour angle -3.0 at declination: " + FormatDec(l_Declination3));
            l_PierSideMinus3 = SOPPierTest(l_StartRA, l_Declination3, "hour angle -3.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            SetAction("Test hour angle +3.0 at declination: " + FormatDec(l_Declination3));
            l_PierSidePlus3 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +3.0d), l_Declination3, "hour angle +3.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            SetAction("Test hour angle -9.0 at declination: " + FormatDec(l_Declination9));
            l_PierSideMinus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", -9.0d), l_Declination9, "hour angle -9.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            SetAction("Test hour angle +9.0 at declination: " + FormatDec(l_Declination9));
            l_PierSidePlus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +9.0d), l_Declination9, "hour angle +9.0");
            if (cancellationToken.IsCancellationRequested)
                return;

            if ((l_PierSideMinus3.SideOfPier == l_PierSidePlus9.SideOfPier) & (l_PierSidePlus3.SideOfPier == l_PierSideMinus9.SideOfPier))// Reporting physical pier side
            {
                LogIssue("SideofPier", "SideofPier reports physical pier side rather than pointing state");
            }
            else if ((l_PierSideMinus3.SideOfPier == l_PierSideMinus9.SideOfPier) & (l_PierSidePlus3.SideOfPier == l_PierSidePlus9.SideOfPier)) // Make other tests
            {
                LogOK("SideofPier", "Reports the pointing state of the mount as expected");
            }
            else // Don't know what this means!
            {
                LogInfo("SideofPier", "Unknown SideofPier reporting model: HA-3: " + l_PierSideMinus3.SideOfPier.ToString() + " HA-9: " + l_PierSideMinus9.SideOfPier.ToString() + " HA+3: " + l_PierSidePlus3.SideOfPier.ToString() + " HA+9: " + l_PierSidePlus9.SideOfPier.ToString());
            }

            LogInfo("SideofPier", "Reported SideofPier at HA -9, +9: " + TranslatePierSide((PointingState)l_PierSideMinus9.SideOfPier, false) + TranslatePierSide((PointingState)l_PierSidePlus9.SideOfPier, false));
            LogInfo("SideofPier", "Reported SideofPier at HA -3, +3: " + TranslatePierSide((PointingState)l_PierSideMinus3.SideOfPier, false) + TranslatePierSide((PointingState)l_PierSidePlus3.SideOfPier, false));

            // Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east at hour angle -3
            if (l_PierSideMinus3.SideOfPier == PointingState.ThroughThePole)
            {
                LogOK("SideofPier", "pierWest is returned when the mount is observing at an hour angle between -6.0 and 0.0");
            }
            else
            {
                LogIssue("SideofPier", "pierEast is returned when the mount is observing at an hour angle between -6.0 and 0.0");
                LogInfo("SideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.");
            }

            if (l_PierSidePlus3.SideOfPier == (int)PointingState.Normal)
            {
                LogOK("SideofPier", "pierEast is returned when the mount is observing at an hour angle between 0.0 and +6.0");
            }
            else
            {
                LogIssue("SideofPier", "pierWest is returned when the mount is observing at an hour angle between 0.0 and +6.0");
                LogInfo("SideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.");
            }

            // Test whether DestinationSideOfPier is implemented
            if (l_PierSideMinus3.DestinationSideOfPier == PointingState.Unknown & l_PierSideMinus9.DestinationSideOfPier == PointingState.Unknown & l_PierSidePlus3.DestinationSideOfPier == PointingState.Unknown & l_PierSidePlus9.DestinationSideOfPier == PointingState.Unknown)
            {
                LogInfo("DestinationSideofPier", "Analysis skipped as this method is not implemented"); // Not implemented
            }
            else // It is implemented so assess the results
            {
                if (l_PierSideMinus3.DestinationSideOfPier == l_PierSidePlus9.DestinationSideOfPier & l_PierSidePlus3.DestinationSideOfPier == l_PierSideMinus9.DestinationSideOfPier) // Reporting physical pier side
                {
                    LogIssue("DestinationSideofPier", "DestinationSideofPier reports physical pier side rather than pointing state");
                }
                else if (l_PierSideMinus3.DestinationSideOfPier == l_PierSideMinus9.DestinationSideOfPier & l_PierSidePlus3.DestinationSideOfPier == l_PierSidePlus9.DestinationSideOfPier) // Make other tests
                {
                    LogOK("DestinationSideofPier", "Reports the pointing state of the mount as expected");
                }
                else // Don't know what this means!
                {
                    LogInfo("DestinationSideofPier", "Unknown DestinationSideofPier reporting model: HA-3: " + l_PierSideMinus3.SideOfPier.ToString() + " HA-9: " + l_PierSideMinus9.SideOfPier.ToString() + " HA+3: " + l_PierSidePlus3.SideOfPier.ToString() + " HA+9: " + l_PierSidePlus9.SideOfPier.ToString());
                }

                // Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east at hour angle -3
                if ((int)l_PierSideMinus3.DestinationSideOfPier == (int)PointingState.ThroughThePole)
                {
                    LogOK("DestinationSideofPier", "pierWest is returned when the mount will observe at an hour angle between -6.0 and 0.0");
                }
                else
                {
                    LogIssue("DestinationSideofPier", "pierEast is returned when the mount will observe at an hour angle between -6.0 and 0.0");
                    LogInfo("DestinationSideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.");
                }

                if (l_PierSidePlus3.DestinationSideOfPier == (int)PointingState.Normal)
                {
                    LogOK("DestinationSideofPier", "pierEast is returned when the mount will observe at an hour angle between 0.0 and +6.0");
                }
                else
                {
                    LogIssue("DestinationSideofPier", "pierWest is returned when the mount will observe at an hour angle between 0.0 and +6.0");
                    LogInfo("DestinationSideofPier", "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.");
                }
            }

            LogInfo("DestinationSideofPier", "Reported DesintationSideofPier at HA -9, +9: " + TranslatePierSide((PointingState)l_PierSideMinus9.DestinationSideOfPier, false) + TranslatePierSide((PointingState)l_PierSidePlus9.DestinationSideOfPier, false));
            LogInfo("DestinationSideofPier", "Reported DesintationSideofPier at HA -3, +3: " + TranslatePierSide((PointingState)l_PierSideMinus3.DestinationSideOfPier, false) + TranslatePierSide((PointingState)l_PierSidePlus3.DestinationSideOfPier, false));

            // Clean up
            // 3.0.0.12 added conditional test to next line
            if (canSetTracking)
                telescopeDevice.Tracking = false;
            ClearStatus();
        }

        public SideOfPierResults SOPPierTest(double p_RA, double p_DEC, string p_Msg)
        {
            // Determine side of pier and destination side of pier results for a particular RA and DEC
            var l_Results = new SideOfPierResults(); // Create result set object
            double l_StartRA, l_StartDEC;
            try
            {
                // Prepare for tests
                l_StartRA = telescopeDevice.RightAscension;
                l_StartDEC = telescopeDevice.Declination;

                // Do destination side of pier test to see what side of pier we should end up on
                LogDebug("", "");
                LogDebug("SOPPierTest", "Testing RA DEC: " + FormatRA(p_RA) + " " + FormatDec(p_DEC) + " Current pierSide: " + TranslatePierSide((PointingState)telescopeDevice.SideOfPier, true));
                try
                {
                    l_Results.DestinationSideOfPier = telescopeDevice.DestinationSideOfPier(p_RA, p_DEC);
                    LogDebug("SOPPierTest", "Target DestinationSideOfPier: " + l_Results.DestinationSideOfPier.ToString());
                }
                catch (COMException ex)
                {
                    switch (ex.ErrorCode)
                    {
                        case var @case when @case == ErrorCodes.NotImplemented:
                            {
                                l_Results.DestinationSideOfPier = PointingState.Unknown;
                                LogDebug("SOPPierTest", "DestinationSideOfPier is not implemented setting result to: " + l_Results.DestinationSideOfPier.ToString());
                                break;
                            }

                        default:
                            {
                                LogIssue("SOPPierTest", "DestinationSideOfPier Exception: " + ex.ToString());
                                break;
                            }
                    }
                }
                catch (MethodNotImplementedException) // DestinationSideOfPier not available so mark as unknown
                {
                    l_Results.DestinationSideOfPier = PointingState.Unknown;
                    LogDebug("SOPPierTest", "DestinationSideOfPier is not implemented setting result to: " + l_Results.DestinationSideOfPier.ToString());
                }
                catch (Exception ex)
                {
                    LogIssue("SOPPierTest", "DestinationSideOfPier Exception: " + ex.ToString());
                }

                // Now do an actual slew and record side of pier we actually get
                SlewScope(p_RA, p_DEC, $"test position {p_Msg}");
                l_Results.SideOfPier = (PointingState)telescopeDevice.SideOfPier;
                LogDebug("SOPPierTest", "Actual SideOfPier: " + l_Results.SideOfPier.ToString());

                // Return to original RA
                SlewScope(l_StartRA, l_StartDEC, "initial start point");
                LogDebug("SOPPierTest", "Returned to: " + FormatRA(l_StartRA) + " " + FormatDec(l_StartDEC));
            }
            catch (Exception ex)
            {
                LogIssue("SOPPierTest", "SideofPierException: " + ex.ToString());
            }

            return l_Results;
        }

        private void DestinationSideOfPierTests()
        {
            PointingState l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9;

            // Slew to one position, then call destination side of pier 4 times and report the pattern
            SlewScope(TelescopeRAFromHourAngle("DestinationSideofPier", -3.0d), 0.0d, "start position");
            l_PierSideMinus3 = telescopeDevice.DestinationSideOfPier(-3.0d, 0.0d);
            l_PierSidePlus3 = telescopeDevice.DestinationSideOfPier(3.0d, 0.0d);
            l_PierSideMinus9 = telescopeDevice.DestinationSideOfPier(-9.0d, 90.0d - siteLatitude);
            l_PierSidePlus9 = telescopeDevice.DestinationSideOfPier(9.0d, 90.0d - siteLatitude);
            if (l_PierSideMinus3 == l_PierSidePlus9 & l_PierSidePlus3 == l_PierSideMinus9) // Reporting physical pier side
            {
                LogIssue("DestinationSideofPier", "The driver appears to be reporting physical pier side rather than pointing state");
            }
            else if (l_PierSideMinus3 == l_PierSideMinus9 & l_PierSidePlus3 == l_PierSidePlus9) // Make other tests
            {
                LogOK("DestinationSideofPier", "The driver reports the pointing state of the mount");
            }
            else // Don't know what this means!
            {
                LogInfo("DestinationSideofPier", "Unknown pier side reporting model: HA-3: " + l_PierSideMinus3.ToString() + " HA-9: " + l_PierSideMinus9.ToString() + " HA+3: " + l_PierSidePlus3.ToString() + " HA+9: " + l_PierSidePlus9.ToString());
            }

            telescopeDevice.Tracking = false;
            LogInfo("DestinationSideofPier", TranslatePierSide(l_PierSideMinus9, false) + TranslatePierSide(l_PierSidePlus9, false));
            LogInfo("DestinationSideofPier", TranslatePierSide(l_PierSideMinus3, false) + TranslatePierSide(l_PierSidePlus3, false));
        }

        #endregion

        #region Support Code

        private void CheckScopePosition(string testName, string functionName, double expectedRA, double expectedDec)
        {
            double actualRA, actualDec, difference;

            LogCallToDriver(testName, "About to get RightAscension property");
            actualRA = telescopeDevice.RightAscension;
            LogDebug(testName, "Read RightAscension: " + FormatRA(actualRA));

            LogCallToDriver(testName, "About to get Declination property");
            actualDec = telescopeDevice.Declination;
            LogDebug(testName, "Read Declination: " + FormatDec(actualDec));

            // Check that we have actually arrived where we are expected to be
            difference = RaDifferenceInArcSeconds(actualRA, expectedRA); // Convert RA difference to arc seconds

            if (difference <= settings.TelescopeSlewTolerance)
            {
                LogOK(testName, $"{functionName} OK within tolerance: ±{settings.TelescopeSlewTolerance} arc seconds. Actual RA: {FormatRA(actualRA)}, Target RA: {FormatRA(expectedRA)}");

            }
            else
            {
                LogIssue(testName, $"{functionName} {difference:0.0} arc seconds away from RA target: {FormatRA(expectedRA)} Actual RA: {FormatRA(actualRA)}. Tolerance: ±{settings.TelescopeSlewTolerance} arc seconds");
            }

            difference = Math.Round(Math.Abs(actualDec - expectedDec) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero);
            if (difference <= settings.TelescopeSlewTolerance) // Dec difference is in arc seconds from degrees of Declination
            {
                LogOK(testName, $"{functionName} OK within tolerance: ±{settings.TelescopeSlewTolerance} arc seconds. Actual DEC: {FormatDec(actualDec)}, Target DEC: {FormatDec(expectedDec)}");
            }
            else
            {
                LogIssue(testName, $"{functionName} {difference:0.0} arc seconds from the expected DEC: {FormatDec(expectedDec)}. Actual DEC: {FormatDec(actualDec)}. Tolerance: ±{settings.TelescopeSlewTolerance} degrees.");
            }
        }

        /// <summary>
        /// Return the difference between two RAs (in hours) as seconds
        /// </summary>
        /// <param name="FirstRA">First RA (hours)</param>
        /// <param name="SecondRA">Second RA (hours)</param>
        /// <returns>Difference (seconds) between the supplied RAs</returns>
        private static double RaDifferenceInArcSeconds(double FirstRA, double SecondRA)
        {
            double RaDifferenceInSecondsRet = Math.Abs(FirstRA - SecondRA); // Calculate the difference allowing for negative outcomes
            if (RaDifferenceInSecondsRet > 12.0d) RaDifferenceInSecondsRet = 24.0d - RaDifferenceInSecondsRet; // Deal with the cases where the two elements are more than 12 hours apart going in the initial direction

            RaDifferenceInSecondsRet = Math.Round(RaDifferenceInSecondsRet * 15.0d * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // RA difference is in arc seconds from hours of RA
            return RaDifferenceInSecondsRet;
        }

        private void SyncScope(string testName, string canDoItName, SlewSyncType testType, double syncRA, double syncDec)
        {
            switch (testType)
            {
                case SlewSyncType.SyncToCoordinates: // SyncToCoordinates
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(testName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(testName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(testName, "About to call SyncToCoordinates method, RA: " + FormatRA(syncRA) + ", Declination: " + FormatDec(syncDec));
                        telescopeDevice.SyncToCoordinates(syncRA, syncDec); // Sync to slightly different coordinates
                        LogDebug(testName, "Completed SyncToCoordinates");
                        break;
                    }

                case SlewSyncType.SyncToTarget: // SyncToTarget
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(testName, "About to get Tracking property");
                        if (canSetTracking & !telescopeDevice.Tracking)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(testName, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(testName, "About to set TargetRightAscension property to " + FormatRA(syncRA));
                            telescopeDevice.TargetRightAscension = syncRA;
                            LogDebug(testName, "Completed Set TargetRightAscension");
                        }
                        catch (Exception ex)
                        {
                            HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, canDoItName + " is True but can't set TargetRightAscension");
                        }

                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage(testName, "About to set TargetDeclination property to " + FormatDec(syncDec));
                            telescopeDevice.TargetDeclination = syncDec;
                            LogDebug(testName, "Completed Set TargetDeclination");
                        }
                        catch (Exception ex)
                        {
                            HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, canDoItName + " is True but can't set TargetDeclination");
                        }

                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage(testName, "About to call SyncToTarget method");
                        telescopeDevice.SyncToTarget(); // Sync to slightly different coordinates
                        LogDebug(testName, "Completed SyncToTarget");
                        break;
                    }

                default:
                    {
                        LogError(testName, "Conform:SyncTest: Unknown test type " + testType.ToString());
                        break;
                    }
            }
        }

        public void SlewScope(double p_RA, double p_DEC, string p_Msg)
        {
            if (canSetTracking)
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SlewScope", "About to set Tracking property to true");
                telescopeDevice.Tracking = true;
            }

            if (canSlew)
            {
                if (canSlewAsync)
                {
                    LogDebug("SlewScope", $"Slewing asynchronously to {p_Msg} {FormatRA(p_RA)} {FormatDec(p_DEC)}");
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SlewScope", "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(p_RA) + ", Declination: " + FormatDec(p_DEC));
                    telescopeDevice.SlewToCoordinatesAsync(p_RA, p_DEC);
                    WaitForSlew(p_Msg, $"Slewing asynchronously to {p_Msg}");
                }
                else
                {
                    LogDebug("SlewScope", "Slewing synchronously to " + p_Msg + " " + FormatRA(p_RA) + " " + FormatDec(p_DEC));
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SlewScope", "About to call SlewToCoordinates method, RA: " + FormatRA(p_RA) + ", Declination: " + FormatDec(p_DEC));
                    SetStatus($"Slewing synchronously to {p_Msg}: {FormatRA(p_RA)} {FormatDec(p_DEC)}");
                    telescopeDevice.SlewToCoordinates(p_RA, p_DEC);
                }

                if (CanReadSideOfPier("SlewScope"))
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("SlewScope", "About to get SideOfPier property");
                    LogDebug("SlewScope", "SideOfPier: " + telescopeDevice.SideOfPier.ToString());
                }
            }
            else
            {
                LogInfo("SlewScope", "Unable to slew this scope as CanSlew is false, slew omitted");
            }

            SetAction("");
        }

        private void WaitForSlew(string testName, string actionMessage)
        {
            Stopwatch sw = Stopwatch.StartNew();

            LogCallToDriver(testName, "About to get Slewing property multiple times");
            WaitWhile(actionMessage, () => { return telescopeDevice.Slewing | (sw.Elapsed.TotalSeconds <= WAIT_FOR_SLEW_MINIMUM_DURATION); }, SLEEP_TIME, settings.TelescopeMaximumSlewTime);
        }

        private double TelescopeRAFromHourAngle(string testName, double p_Offset)
        {
            double TelescopeRAFromHourAngleRet;

            // Handle the possibility that the mandatory SideealTime property has not been implemented
            if (canReadSiderealTime)
            {
                // Create a legal RA based on an offset from Sidereal time
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage(testName, "About to get SiderealTime property");
                TelescopeRAFromHourAngleRet = telescopeDevice.SiderealTime - p_Offset;
                switch (TelescopeRAFromHourAngleRet)
                {
                    case var @case when @case < 0.0d: // Illegal if < 0 hours
                        {
                            TelescopeRAFromHourAngleRet += 24.0d;
                            break;
                        }

                    case var case1 when case1 >= 24.0d: // Illegal if > 24 hours
                        {
                            TelescopeRAFromHourAngleRet -= 24.0d;
                            break;
                        }
                }
            }
            else
            {
                TelescopeRAFromHourAngleRet = 0.0d - p_Offset;
            }

            return TelescopeRAFromHourAngleRet;
        }

        private double TelescopeRAFromSiderealTime(string testName, double p_Offset)
        {
            double TelescopeRAFromSiderealTimeRet;
            double CurrentSiderealTime;

            // Handle the possibility that the mandatory SideealTime property has not been implemented
            if (canReadSiderealTime)
            {
                // Create a legal RA based on an offset from Sidereal time
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage(testName, "About to get SiderealTime property");
                CurrentSiderealTime = telescopeDevice.SiderealTime;
                switch (CurrentSiderealTime) // Deal with possibility that sidereal time from the driver is bad
                {
                    case var @case when @case < 0.0d: // Illegal if < 0 hours
                        {
                            CurrentSiderealTime = 0d;
                            break;
                        }

                    case var case1 when case1 >= 24.0d: // Illegal if > 24 hours
                        {
                            CurrentSiderealTime = 0d;
                            break;
                        }
                }

                TelescopeRAFromSiderealTimeRet = CurrentSiderealTime + p_Offset;
                switch (TelescopeRAFromSiderealTimeRet)
                {
                    case var case2 when case2 < 0.0d: // Illegal if < 0 hours
                        {
                            TelescopeRAFromSiderealTimeRet += 24.0d;
                            break;
                        }

                    case var case3 when case3 >= 24.0d: // Illegal if > 24 hours
                        {
                            TelescopeRAFromSiderealTimeRet -= 24.0d;
                            break;
                        }
                }
            }
            else
            {
                TelescopeRAFromSiderealTimeRet = 0.0d + p_Offset;
            }

            return TelescopeRAFromSiderealTimeRet;
        }
#if WINDOWS
        private void TestEarlyBinding(InterfaceType TestType)
        {
            dynamic l_ITelescope;
            dynamic l_DeviceObject = null;
            string l_ErrMsg;
            int l_TryCount = 0;
            try
            {
                // Try early binding
                l_ITelescope = null;
                do
                {
                    l_TryCount += 1;
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("AccessChecks", "About to create driver object with CreateObject");
                        LogDebug("AccessChecks", "Creating late bound object for interface test");
                        Type driverType = Type.GetTypeFromProgID(settings.ComDevice.ProgId);
                        l_DeviceObject = Activator.CreateInstance(driverType);
                        LogDebug("AccessChecks", "Created late bound object OK");
                        switch (TestType)
                        {
                            case InterfaceType.ITelescopeV2:
                                {
                                    //l_ITelescope = (ASCOM.Interface.ITelescope)l_DeviceObject;
                                    break;
                                }

                            case InterfaceType.ITelescopeV3:
                                {
                                    l_ITelescope = (ITelescopeV3)l_DeviceObject;
                                    break;
                                }

                            default:
                                {
                                    LogIssue("TestEarlyBinding", "Unknown interface type: " + TestType.ToString());
                                    break;
                                }
                        }

                        LogDebug("AccessChecks", "Successfully created driver with interface " + TestType.ToString());
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("AccessChecks", "About to set Connected property true");
                            l_ITelescope.Connected = true;
                            LogInfo("AccessChecks", "Device exposes interface " + TestType.ToString());
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("AccessChecks", "About to set Connected property false");
                            l_ITelescope.Connected = false;
                        }
                        catch (Exception)
                        {
                            LogInfo("AccessChecks", "Device does not expose interface " + TestType.ToString());
                            LogNewLine();
                        }
                    }
                    catch (Exception ex)
                    {
                        l_ErrMsg = ex.ToString();
                        LogDebug("AccessChecks", "Exception: " + ex.Message);
                    }

                    if (l_DeviceObject is null)
                        WaitFor(200);
                }
                while (l_TryCount < 3 & l_ITelescope is not object); // Exit if created OK
                if (l_ITelescope is null)
                {
                    LogInfo("AccessChecks", "Device does not expose interface " + TestType.ToString());
                }
                else
                {
                    LogDebug("AccessChecks", "Created telescope on attempt: " + l_TryCount.ToString());
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

                l_DeviceObject = null;
                l_ITelescope = null;
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
        private static string FormatRA(double ra)
        {
            return Utilities.HoursToHMS(ra, ":", ":", "", DISPLAY_DECIMAL_DIGITS);
        }

        private static string FormatDec(double Dec)
        {
            return Utilities.DegreesToDMS(Dec, ":", ":", "", 1).PadLeft(9 + ((1 > 0) ? 1 + 1 : 0));
        }

        private static dynamic FormatAltitude(double Alt)
        {
            return Utilities.DegreesToDMS(Alt, ":", ":", "", DISPLAY_DECIMAL_DIGITS);
        }

        private static string FormatAzimuth(double Az)
        {
            return Utilities.DegreesToDMS(Az, ":", ":", "", DISPLAY_DECIMAL_DIGITS).PadLeft(9 + ((DISPLAY_DECIMAL_DIGITS > 0) ? DISPLAY_DECIMAL_DIGITS + 1 : 0));
        }

        public static string TranslatePierSide(PointingState p_PierSide, bool p_Long)
        {
            string l_PierSide;
            switch (p_PierSide)
            {
                case PointingState.Normal:
                    {
                        if (p_Long)
                        {
                            l_PierSide = "pierEast";
                        }
                        else
                        {
                            l_PierSide = "E";
                        }

                        break;
                    }

                case PointingState.ThroughThePole:
                    {
                        if (p_Long)
                        {
                            l_PierSide = "pierWest";
                        }
                        else
                        {
                            l_PierSide = "W";
                        }

                        break;
                    }

                default:
                    {
                        if (p_Long)
                        {
                            l_PierSide = "pierUnknown";
                        }
                        else
                        {
                            l_PierSide = "U";
                        }

                        break;
                    }
            }

            return l_PierSide;
        }

        private enum Axis
        {
            RA,
            Dec
        }

        private bool TestRADecRate(string testName, string description, Axis axis, double rate, bool includeSlewiingTest)
        {
            // Initialise the outcome to false
            bool success = false;
            double offsetRate;

            try
            {
                // Tracking must be enabled for this test so make sure that it is enabled
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage(testName, $"{description} - About to get Slewing property");
                tracking = telescopeDevice.Tracking;

                // Test whether we are tracking and if not enable this if possible, abort the test if tracking cannot be set True
                if (!tracking)
                {
                    // We are not tracking so test whether Tracking can be set

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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage(testName, $"{description} - About to get Slewing property");
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
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage(testName, $"{description} - About to set {RateOffsetName(axis)} property to {rate}");

                    // Set the appropriate offset rate
                    if (axis == Axis.RA)
                        telescopeDevice.RightAscensionRate = rate;
                    else
                        telescopeDevice.DeclinationRate = rate;

                    SetAction("Waiting for mount to settle");
                    WaitFor(1000); // Give a short wait to allow the mount to settle at the new rate

                    // If we get here the value was set OK, now check that the new rate is returned by RightAscensionRate Get and that Slewing is false
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage(testName, $"{description} - About to get {RateOffsetName(axis)} property");

                    // Get the appropriate offset rate
                    if (axis == Axis.RA)
                        offsetRate = telescopeDevice.RightAscensionRate;
                    else
                        offsetRate = telescopeDevice.DeclinationRate;

                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage(testName, $"{description} - About to get the Slewing property");
                    slewing = telescopeDevice.Slewing;

                    if (offsetRate == rate & !slewing)
                    {
                        LogOK(testName, $"{description} - successfully set rate to {offsetRate}");
                        success = true;

                        // Run the extended rate offset tests if configured to do so.
                        if (settings.TelescopeExtendedRateOffsetTests)
                        {
                            // Only test movement when the requested rate is not 0.0
                            if (rate != 0.0)
                            {
                                // Now test the actual amount of movement
                                if (axis == Axis.RA)
                                {
                                    TestOffsetRate(testName, -9.0, rate, 0.0);
                                    if (cancellationToken.IsCancellationRequested)
                                        return success;

                                    TestOffsetRate(testName, +3.0, rate, 0.0);
                                    if (cancellationToken.IsCancellationRequested)
                                        return success;

                                    TestOffsetRate(testName, +9.0, rate, 0.0);
                                    if (cancellationToken.IsCancellationRequested)
                                        return success;

                                    TestOffsetRate(testName, -3.0, rate, 0.0);
                                    if (cancellationToken.IsCancellationRequested)
                                        return success;
                                }
                                else
                                {
                                    TestOffsetRate(testName, -9.0, 0.0, rate);
                                    if (cancellationToken.IsCancellationRequested)
                                        return success;

                                    TestOffsetRate(testName, +3.0, 0.0, rate);
                                    if (cancellationToken.IsCancellationRequested)
                                        return success;

                                    TestOffsetRate(testName, +9.0, 0.0, rate);
                                    if (cancellationToken.IsCancellationRequested)
                                        return success;

                                    TestOffsetRate(testName, -3.0, 0.0, rate);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (slewing & offsetRate == rate)
                            LogIssue(testName, $"{RateOffsetName(axis)} was successfully set to {rate} but Slewing is returning True, it should return False.");
                        if (slewing & offsetRate != rate)
                            LogIssue(testName, $"{RateOffsetName(axis)} Read does not return {rate} as set, instead it returns {offsetRate}. Slewing is also returning True, it should return False.");
                        if (!slewing & offsetRate != rate)
                            LogIssue(testName, $"{RateOffsetName(axis)} Read does not return {rate} as set, instead it returns {offsetRate}.");
                    }
                }
                catch (Exception ex)
                {
                    if (IsInvalidOperationException(testName, ex)) // We can't know what the valid range for this telescope is in advance so its possible that our test value will be rejected, if so just report this.
                    {
                        LogInfo(testName, $"Unable to set test rate {rate}, it was rejected as an invalid value.");
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
            if (axis == Axis.RA)
                return "RightAscensionRate";
            return "DeclinationRate";
        }

        /// <summary>
        /// Return the declination which provides the highest elevation that is less than 65 degrees at the given RA.
        /// </summary>
        /// <param name="testRa">RA to use when determining the optimum declination.</param>
        /// <returns>Declination in the range -80.0 to +80.0</returns>
        /// <remarks>The returned declination will correspond to an elevation in the range 0 to 65 degrees.</remarks>
        internal double GetTestDeclination(string testName, double testRa)
        {
            double testDeclination = 0.0;
            double testElevation = double.MinValue;

            // Find a test declination that yields an elevation that is as high as possible but under 65 degrees
            // Initialise transform with site parameters
            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, $"About to get SiteLatitude property");
            transform.SiteLatitude = telescopeDevice.SiteLatitude;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, $"About to get SiteLongitude property");
            transform.SiteLongitude = telescopeDevice.SiteLongitude;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, $"About to get SiteElevation property");
            transform.SiteElevation = telescopeDevice.SiteElevation;

            // Set remaining transform parameters
            transform.SitePressure = 1010.0;
            transform.Refraction = false;

            LogDebug("GetTestDeclination", $"TRANSFORM: Site: Latitude: {transform.SiteLatitude.ToDMS()}, Longitude: {transform.SiteLongitude.ToDMS()}, Elevation: {transform.SiteElevation}, Pressure: {transform.SitePressure}, Temperature: {transform.SiteTemperature}, Refraction: {transform.Refraction}");

            // Iterate from declination -85 to +85 in steps of 10 degrees
            Stopwatch sw = Stopwatch.StartNew();
            for (double declination = -85.0; declination < 90.0; declination += 10.0)
            {
                // Set transform's topocentric coordinates
                transform.SetTopocentric(testRa, declination);

                // Retrieve the corresponding elevation
                double elevation = transform.ElevationTopocentric;

                LogDebug("GetTestDeclination", $"TRANSFORM: RA: {testRa.ToHMS()}, Declination: {declination.ToDMS()}, Azimuth: {transform.AzimuthTopocentric.ToDMS()}, Elevation: {elevation.ToDMS()}");

                // Update the test declination if the new elevation is less that 65 degrees and also greater than the current highest elevation 
                if ((elevation < 65.0) & (elevation > testElevation))
                {
                    testDeclination = declination;
                    testElevation = elevation;
                }
            }
            sw.Stop();
            LogDebug("GetTestDeclination", $"Test RightAscension: {testRa.ToHMS()}, Test Declination: {testDeclination.ToDMS()} at Elevation: {testElevation.ToDMS()} found in {sw.Elapsed.TotalMilliseconds:0.0}ms.");

            // Throw an exception if the test elevation is below 0 degrees
            if (testElevation < 0.0)
            {
                throw new ASCOM.InvalidOperationException($"The highest elevation available: {testElevation.ToDMS()} is below the horizon");
            }
            // Return the test declination
            return testDeclination;
        }

        internal void TestOffsetRate(string testName, double testHa, double expectedRaRate, double expectedDeclinationRate)
        {
            double testDuration = settings.TelescopeRateOffsetTestDuration; // Seconds
            double testDeclination;

            // Update the test name with the test hour angle
            testName = $"{testName} {testHa:+0.0;-0.0;+0.0}";

            // Create the test RA and declination
            double testRa = Utilities.ConditionRA(telescopeDevice.SiderealTime - testHa);
            try
            {
                testDeclination = GetTestDeclination(testName, testRa); // Get the test declination for the test RA
            }
            catch (ASCOM.InvalidOperationException ex)
            {
                LogInfo(testName, $"Test omitted because {ex.Message.ToLowerInvariant()}. This is an expected condition at latitudes close to the equator.");
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

            // Start of test
            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to get RightAscension property");
            double priStart = telescopeDevice.RightAscension;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to get Declination property");
            double secStart = telescopeDevice.Declination;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to set RightAscensionRate property");
            telescopeDevice.RightAscensionRate = expectedRaRate * SIDEREAL_SECONDS_TO_SI_SECONDS;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to set DeclinationRate property");
            telescopeDevice.DeclinationRate = expectedDeclinationRate;

            WaitFor(Convert.ToInt32(testDuration * 1000));

            double priEnd = telescopeDevice.RightAscension;
            double secEnd = telescopeDevice.Declination;

            // Restore previous state
            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to set RightAscensionRate property");
            telescopeDevice.RightAscensionRate = 0.0;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(testName, "About to set DeclinationRate property");
            telescopeDevice.DeclinationRate = 0.0;

            LogDebug(testName, $"Start      - : {priStart.ToHMS()}, {secStart.ToDMS()}");
            LogDebug(testName, $"Finish     - : {priEnd.ToHMS()}, {secEnd.ToDMS()}");
            LogDebug(testName, $"Difference - : {(priEnd - priStart).ToHMS()}, {(secEnd - secStart).ToDMS()}, {priEnd - priStart:N10}, {secEnd - secStart:N10}");

            // Condition results
            double actualPriRate = (priEnd - priStart) / testDuration; // Calculate offset rate in RA hours per SI second
            actualPriRate = actualPriRate * 60.0 * 60.0; // Convert rate in RA hours per SI second to RA seconds per SI second

            double actualSecRate = (secEnd - secStart) / testDuration * 60.0 * 60.0;

            LogDebug(testName, $"Actual primary rate: {actualPriRate}, Expected rate: {expectedRaRate}, Ratio: {actualPriRate / expectedRaRate}, Actual secondary rate: {actualSecRate}, Expected rate: {expectedDeclinationRate}, Ratio: {actualSecRate / expectedDeclinationRate}");
            TestDouble(testName, "RightAscensionRate", actualPriRate, expectedRaRate);
            TestDouble(testName, "DeclinationRate   ", actualSecRate, expectedDeclinationRate);

            LogDebug("", "");
        }

        private void TestDouble(string testName, string name, double actualValue, double expectedValue, double tolerance = 0.0)
        {
            // Tolerance 0 = 2%
            const double TOLERANCE = 0.05; // 2%

            if (tolerance == 0.0)
            {
                tolerance = TOLERANCE;
            }

            if (expectedValue == 0.0)
            {
                if (Math.Abs(actualValue - expectedValue) <= tolerance)
                {
                    LogOK(testName, $"{name} is within expected tolerance. Expected: {expectedValue:+0.000;-0.000;+0.000}, Actual: {actualValue:+0.000;-0.000;+0.000}, Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / tolerance):N2}%.");
                }
                else
                {
                    LogIssue(testName, $"{name} is outside the expected tolerance. Expected: {expectedValue:+0.000;-0.000;+0.000}, Actual: {actualValue:+0.000;-0.000;+0.000}, Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / tolerance):N2}%, Tolerance:{tolerance * 100:N2}.");
                }
            }
            else
            {
                if (Math.Abs(Math.Abs(actualValue - expectedValue) / expectedValue) <= tolerance)
                {
                    LogOK(testName, $"{name} is within expected tolerance. Expected: {expectedValue:+0.000;-0.000;+0.000}, Actual: {actualValue:+0.000;-0.000;+0.000}, Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / expectedValue):N2}%.");
                }
                else
                {
                    LogIssue(testName, $"{name} is outside the expected tolerance. Expected: {expectedValue:+0.000;-0.000;+0.000}, Actual: {actualValue:+0.000;-0.000;+0.000}, Deviation from expected: {Math.Abs((actualValue - expectedValue) * 100.0 / expectedValue):N2}%, Tolerance:{tolerance * 100:N2}.");
                }
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
            if (applicationCancellationToken.IsCancellationRequested)
                return;

            // Test each of the four guide directions
            TestPulseGuideDirection(ha, GuideDirection.North);
            if (applicationCancellationToken.IsCancellationRequested)
                return;

            TestPulseGuideDirection(ha, GuideDirection.South);
            if (applicationCancellationToken.IsCancellationRequested)
                return;

            TestPulseGuideDirection(ha, GuideDirection.East);
            if (applicationCancellationToken.IsCancellationRequested)
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

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(logName, "About to get RightAscension property");
            double initialRACoordinate = telescopeDevice.RightAscension;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(logName, "About to get Declination property");
            double initialDecCoordinate = telescopeDevice.Declination;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(logName, $"About to call PulseGuide. Direction: {direction}, Duration: {PULSE_GUIDE_DURATION * 1000}ms.");
            telescopeDevice.PulseGuide(direction, PULSE_GUIDE_DURATION * 1000);

            WaitWhile($"Pulse guiding {direction} at HA {ha:+0.0;-0.0;+0.0}", () => { return telescopeDevice.IsPulseGuiding; }, SLEEP_TIME, PULSE_GUIDE_DURATION);

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(logName, "About to get RightAscension property");
            double finalRACoordinate = telescopeDevice.RightAscension;

            if (settings.DisplayMethodCalls)
                LogTestAndMessage(logName, "About to get Declination property");
            double finalDecCoordinate = telescopeDevice.Declination;

            double raChange = finalRACoordinate - initialRACoordinate;
            double decChange = finalDecCoordinate - initialDecCoordinate;

            double raDifference = Math.Abs(raChange - expectedRaChange);
            double decDifference = Math.Abs(decChange - expectedDecChange);

            switch (direction)
            {
                case GuideDirection.North:
                    if ((decChange > 0.0) & (decDifference <= changeToleranceDec) & (Math.Abs(raChange) <= changeToleranceRa)) // Moved north by expected amount with no east-west movement
                        LogOK(logName, $"Moved north with no east-west movement  - Declination change (DMS): {decChange.ToDMS()},  Expected: {expectedDecChange.ToDMS()},  Difference: {decDifference.ToDMS()},  Test tolerance: {changeToleranceDec.ToDMS()}.");
                    else
                    {
                        switch (decChange)
                        {
                            case < 0.0: //Moved south
                                LogIssue(logName, $"Moved south instead of north - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The declination axis did not move.");
                                break;

                            default: // Moved north
                                if (decDifference <= changeToleranceDec) // Moved north OK.
                                    LogOK(logName, $"Moved north - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                                else // Moved north but by wrong amount.
                                    LogIssue(logName, $"Moved north but outside test tolerance - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                                break;
                        }
                        if (Math.Abs(raChange) <= changeToleranceRa)
                            LogOK(logName, $"No significant east-west movement as expected. RA change (HMS): {raChange.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                        else
                            LogIssue(logName, $"East-West movement was outside test tolerance: RA change (HMS): {raChange.ToHMS()}, Expected: {0.0.ToHMS()}, Tolerance: {changeToleranceRa.ToHMS()}.");
                    }
                    break;

                case GuideDirection.South:
                    if ((decChange < 0.0) & (decDifference <= changeToleranceDec) & (Math.Abs(raChange) <= changeToleranceRa)) // Moved south by expected amount with no east-west movement
                        LogOK(logName, $"Moved south with no east-west movement  - Declination change (DMS): {decChange.ToDMS()},  Expected: {expectedDecChange.ToDMS()},  Difference: {decDifference.ToDMS()},  Test tolerance: {changeToleranceDec.ToDMS()}.");
                    else
                    {
                        switch (decChange)
                        {
                            case > 0.0: //Moved north
                                LogIssue(logName, $"Moved north instead of south - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The declination axis did not move.");
                                break;

                            default: // Moved south
                                if (decDifference <= changeToleranceDec)// Moved south OK
                                    LogOK(logName, $"Moved south - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                                else// Moved south but by wrong amount.
                                    LogIssue(logName, $"Moved south but outside test tolerance - Declination change (DMS): {decChange.ToDMS()}, Expected: {expectedDecChange.ToDMS()}, Difference: {decDifference.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                                break;
                        }
                        if (Math.Abs(raChange) <= changeToleranceRa)
                            LogOK(logName, $"No significant east-west movement as expected. RA change (HMS): {raChange.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                        else
                            LogIssue(logName, $"East-West movement was outside test tolerance: RA change (HMS): {raChange.ToHMS()}, Expected: {0.0.ToHMS()}, Tolerance: {changeToleranceRa.ToHMS()}.");
                    }
                    break;

                case GuideDirection.East:
                    if ((raChange > 0.0) & (raDifference <= changeToleranceRa) & (Math.Abs(decChange) <= changeToleranceDec)) // Moved east by expected amount with no north-south movement
                        LogOK(logName, $"Moved east with no north-south movement - RA change (HMS):          {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                    else
                    {
                        switch (raChange)
                        {
                            case < 0.0: //Moved west
                                LogIssue(logName, $"Moved west instead of east - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The RA axis did not move.");
                                break;

                            default: // Moved east
                                if (raDifference <= changeToleranceRa)  // Moved east OK
                                    LogOK(logName, $"Moved east - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                                else // Moved east but by wrong amount.
                                    LogIssue(logName, $"Moved east but outside test tolerance - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                                break;
                        }
                        if (Math.Abs(decChange) <= changeToleranceDec)
                            LogOK(logName, $"No significant north-south movement as expected. Declination change (DMS): {decChange.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                        else
                            LogIssue(logName, $"North-south movement was outside test tolerance: Declination change (DMS): {decChange.ToDMS()}, Expected: {0.0.ToDMS()}, Tolerance: {changeToleranceDec.ToDMS()}.");
                    }
                    break;

                case GuideDirection.West:
                    if ((raChange < 0.0) & (raDifference <= changeToleranceRa) & (Math.Abs(decChange) <= changeToleranceDec)) // Moved west by expected amount with no north-south movement
                        LogOK(logName, $"Moved west with no north-south movement - RA change (HMS):          {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                    else
                    {
                        switch (raChange)
                        {
                            case > 0.0: //Moved east
                                LogIssue(logName, $"Moved east instead of west - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                                break;

                            case 0.0: // No movement
                                LogIssue(logName, $"The RA axis did not move.");
                                break;

                            default: // Moved west
                                if (raDifference <= changeToleranceRa) // Moved west OK
                                    LogOK(logName, $"Moved west - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                                else// Moved west but by wrong amount.
                                    LogIssue(logName, $"Moved west but outside test tolerance - RA change (HMS): {raChange.ToHMS()}, Expected: {expectedRaChange.ToHMS()}, Difference: {raDifference.ToHMS()}, Test tolerance: {changeToleranceRa.ToHMS()}.");
                                break;
                        }
                        if (Math.Abs(decChange) <= changeToleranceDec)
                            LogOK(logName, $"No significant north-south movement as expected. Declination change (DMS): {decChange.ToDMS()}, Test tolerance: {changeToleranceDec.ToDMS()}.");
                        else
                            LogIssue(logName, $"North-south movement was outside test tolerance: Declination change (DMS): {decChange.ToDMS()}, Expected: {0.0.ToDMS()}, Tolerance: {changeToleranceDec.ToDMS()}.");
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
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("SlewToHa", "About to get SiderealTime property");
            double targetRa = Utilities.ConditionRA(telescopeDevice.SiderealTime - targetHa);

            // Calculate the target declination
            double targetDeclination = GetTestDeclination("SlewToHa", targetRa);
            LogDebug("SlewToHa", $"Slewing to HA: {targetHa.ToHMS()} (RA: {targetRa.ToHMS()}), Dec: {targetDeclination.ToDMS()}");

            // Slew to the target coordinates
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("SlewToHa", $"About to call SlewToCoordinatesAsync. RA: {targetRa.ToHMS()}, Declination: {targetDeclination.ToDMS()}");
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

        private void ValidateOperationComplete(string test, bool expectedState)
        {
            try
            {
                LogCallToDriver(test, "About to call OperationComplete method");
                bool operationComplete = telescopeDevice.OperationComplete;

                if (operationComplete == expectedState)
                {
                    // Got expected outcome so no action
                }
                else
                {
                    LogIssue(test, $"OperationComplete did not have the executed state: {expectedState}, it was: {operationComplete}.");
                }
            }
            catch (Exception ex)
            {
                LogIssue(test, $"Unexpected exception from OperationComplete: {ex.Message}");
                LogDebug(test, ex.ToString());
            }

        }

        private void TimeMethod(string methodName, Action method)
        {
            Stopwatch sw = Stopwatch.StartNew();
            method();
            if (sw.ElapsedMilliseconds > operationInitiationTime)
                LogIssue(methodName, $"Operation initiation took {sw.Elapsed.TotalSeconds:0.0} seconds, which is more than the configured maximum: {operationInitiationTime:0.0} seconds.");
        }

        #endregion

    }
}

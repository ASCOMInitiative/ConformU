// Option Strict On
using System;
using System.Collections;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ASCOM;
using ASCOM.DeviceInterface;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using static Conform.GlobalVarsAndCode;

namespace Conform
{
    internal class TelescopeTester : DeviceTesterBaseClass
    {

        #region Variables and Constants
        private const int TRACKING_COMMAND_DELAY = 1000; // Time to wait between changing Tracking state
        private const int PERF_LOOP_TIME = 5; // Performance loop run time in seconds
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
        private bool canFindHome, canPark, canPulseGuide, canSetDeclinationRate, canSetGuideRates, canSetPark, canSetPierside, canSetRightAscensionRate;
        private bool canSetTracking, canSlew, canSlewAltAz, canSlewAltAzAsync, canSlewAsync, canSync, canSyncAltAz, canUnpark;
        private AlignmentModes m_AlignmentMode;
        private double m_Altitude;
        private double m_ApertureArea;
        private double m_ApertureDiameter;
        private bool m_AtHome;
        private bool m_AtPark;
        private double m_Azimuth;
        private double m_Declination;
        private double m_DeclinationRate;
        private bool m_DoesRefraction;
        private EquatorialCoordinateType m_EquatorialSystem;
        private double m_FocalLength;
        private double m_GuideRateDeclination;
        private double m_GuideRateRightAscension;
        private bool m_IsPulseGuiding;
        private double m_RightAscension;
        private double m_RightAscensionRate;
        private PierSide m_SideOfPier;
        private double m_SiderealTimeScope;
        private double m_SiteElevation;
        private double m_SiteLatitude;
        private double m_SiteLongitude;
        private bool m_Slewing;
        private short m_SlewSettleTime;
        private double m_TargetDeclination;
        private double m_TargetRightAscension;
        private bool m_Tracking;
        private DateTime m_UTCDate;
        private bool m_CanMoveAxisPrimary, m_CanMoveAxisSecondary, m_CanMoveAxisTertiary;
        private PierSide m_DestinationSideOfPier, m_DestinationSideOfPierEast, m_DestinationSideOfPierWest;
        private double m_SiderealTimeASCOM;
        private DateTime m_StartTime, m_EndTime;
        private bool m_CanReadSideOfPier;
        private double m_TargetAltitude, m_TargetAzimuth;
        private bool canReadAltitide, canReadAzimuth, canReadSiderealTime;

#if DEBUG
        private ASCOM.DriverAccess.Telescope telescopeDevice;
#else
        private dynamic telescopeDevice;
#endif

        private dynamic DriverAsObject;
        // Axis rate checks
        private double[,] m_AxisRatesPrimaryArray = new double[1001, 2], m_AxisRatesArray = new double[1001, 2];
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
            tstPerfAltitude = 1,
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
        public TelescopeTester() : base(true, true, true, true, false, true, true) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        protected override void Dispose(bool disposing)
        {
            LogMsg("Dispose", MessageLevel.msgDebug, "Disposing of Telescope driver: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (true) // Should be True but make False to stop Conform from cleanly dropping the telescope object (useful for retaining driver in memory to change flags)
                    {
                        try
                        {
                            DisposeAndReleaseObject("Telescope Device", telescopeDevice);
                        }
                        catch
                        {
                        }

                        telescopeDevice = null;
                        g_DeviceObject = null;
                        GC.Collect();
                    }
                }
            }

            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        #region Code
        public override void CheckCommonMethods()
        {
            CheckCommonMethods(telescopeDevice, DeviceType.Telescope);
        }

        public override void CheckInitialise()
        {
            unchecked
            {
                // Set the error type numbers according to the standards adopted by individual authors.
                // Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
                // messages to driver authors!
                switch (g_TelescopeProgID ?? "")
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

            CheckInitialise(g_TelescopeProgID);
        }

        public override void CheckAccessibility()
        {
            try
            {
                ASCOM.DriverAccess.Telescope l_DriverAccessTelescope = null;
                string l_ErrMsg = "";
                int l_TryCount = 0;
                try
                {
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Before MyBase.CheckAccessibility");
                    CheckAccessibility(g_TelescopeProgID, DeviceType.Telescope);
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "After MyBase.CheckAccessibility");
                    try
                    {
                        TestEarlyBinding(InterfaceType.ITelescopeV2);
                        TestEarlyBinding(InterfaceType.ITelescopeV3);

                        // Try client access toolkit
                        l_DriverAccessTelescope = null;
                        l_TryCount = 0;
                        do
                        {
                            l_TryCount += 1;
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("AccessChecks", MessageLevel.msgComment, "About to create DriverAccess instance");
                                l_DriverAccessTelescope = new ASCOM.DriverAccess.Telescope(g_TelescopeProgID);
                                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
                                LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using driver access toolkit");
                                try
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property true");
                                    l_DriverAccessTelescope.Connected = true;
                                    LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using driver access toolkit");
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property false");
                                    l_DriverAccessTelescope.Connected = false;
                                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully disconnected using driver access toolkit");
                                }
                                catch (Exception ex)
                                {
                                    LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using driver access toolkit: " + ex.ToString());
                                    LogMsg("", MessageLevel.msgAlways, "");
                                }
                            }
                            catch (Exception ex)
                            {
                                l_ErrMsg = ex.ToString();
                            }

                            if (l_DriverAccessTelescope is null)
                                WaitFor(200);
                        }
                        while (!(l_TryCount == 3 | l_DriverAccessTelescope is object)); // Exit if created OK
                        if (l_DriverAccessTelescope is null)
                        {
                            LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver using driver access toolkit: " + l_ErrMsg);
                            LogMsg("", MessageLevel.msgAlways, "");
                        }
                        else
                        {
                            LogMsg("Telescope:CreateDevice", MessageLevel.msgDebug, "Created telescope on attempt: " + l_TryCount.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMsg("Telescope:CheckAcc.EX3", MessageLevel.msgError, ex.ToString());
                    }
                }
                catch (Exception ex)
                {
                    LogMsg("Telescope:CheckAcc.EX2", MessageLevel.msgError, ex.ToString());
                }

                // Clean up
                try
                {
                    l_DriverAccessTelescope.Dispose();
                }
                catch
                {
                }
                // Try : Marshal.ReleaseComObject(l_DriverAccessTelescope) : Catch : End Try
                l_DriverAccessTelescope = null;
                GC.Collect();
                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for Accessibility Telescope Object to Dispose");
            }
            catch (Exception ex)
            {
                LogMsg("Telescope:CheckAcc.EX1", MessageLevel.msgError, ex.ToString());
            }
        }

        public override void CreateDevice()
        {
            int l_TryCount = 0;
            do
            {
                l_TryCount += 1;
                try
                {
                    LogMsg("Telescope:CreateDevice", MessageLevel.msgDebug, "Creating ProgID: " + g_TelescopeProgID);
#if DEBUG
                    LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Telescope to get a Telescope object");
                    if (g_Settings.DisplayMethodCalls) LogMsg("CreateDevice", MessageLevel.msgComment, "About to create driver using DriverAccess");
                    telescopeDevice = new ASCOM.DriverAccess.Telescope(g_TelescopeProgID);
                    LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");
#else
                    if (g_Settings.UseDriverAccess)
                    {
                        LogMsg("Conform", MessageLevel.msgAlways, "is using ASCOM.DriverAccess.Telescope to get a Telescope object");
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to create driver using DriverAccess");
                        telescopeDevice = new ASCOM.DriverAccess.Telescope(g_TelescopeProgID);
                        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");
                    }
                    else
                    {
                        LogMsg("Conform", MessageLevel.msgAlways, "is using CreateObject to get a Telescope object");
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to create driver using CreateObject");
                        telescopeDevice = Interaction.CreateObject(g_TelescopeProgID);
                        LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");
                    }
#endif

                    WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
                    g_Stop = false;
                }
                catch (Exception ex)
                {
                    LogMsg("", MessageLevel.msgDebug, "Attempt " + l_TryCount + " - exception thrown: " + ex.Message);
                    if (l_TryCount == 3)
                        throw;
                } // Re throw exception if on our third attempt

                if (g_Stop)
                    WaitFor(200);
            }
            while (g_Stop); // Exit if created OK
            LogMsg("CreateDevice", MessageLevel.msgDebug, "Created telescope on attempt: " + l_TryCount.ToString());

            // Create a pointer to the raw COM object that represents the Telescope (Only used for rate object Dispose() tests)
            try
            {
                if (g_Settings.UseDriverAccess) // Use an internal DriverAccess field to get a pointer to the underlying COM driver
                {
                    LogMsg("CreateDevice", MessageLevel.msgDebug, "Using DriverAccess to get underlying driver as an object");
                    DriverAsObject = ((ASCOM.DriverAccess.Telescope)telescopeDevice).memberFactory.GetLateBoundObject; // Have to convert the device from object to DriverAccess.Telescope in order to be able to access the internal GetLateBoundObject field
                }
                else
                {
                    LogMsg("CreateDevice", MessageLevel.msgDebug, "Driver is already an object so using it \"as is\" for driver as an object");
                    DriverAsObject = telescopeDevice;
                }

                LogMsg("CreateDevice", MessageLevel.msgDebug, "Got driver object OK");
            }
            catch (Exception ex)
            {
                LogMsg("CreateDevice", MessageLevel.msgError, "Exception: " + ex.ToString());
            }

            LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver as an object");
        }

        public override bool Connected
        {
            get
            {
                bool ConnectedRet = default;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Connected property");
                ConnectedRet = telescopeDevice.Connected;
                return ConnectedRet;
            }

            set
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Connected property " + value.ToString());
                telescopeDevice.Connected = value;
                g_Stop = false;
            }
        }

        public override void PreRunCheck()
        {
            // Get into a consistent state
            if (g_InterfaceVersion > 1)
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Mount Safety", MessageLevel.msgComment, "About to get AtPark property");
                if (Conversions.ToBoolean(telescopeDevice.AtPark))
                {
                    if (canUnpark)
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("Mount Safety", MessageLevel.msgComment, "About to call Unpark method");
                        telescopeDevice.Unpark();
                        LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope is parked, so it has been unparked for testing");
                    }
                    else
                    {
                        LogMsg("Mount Safety", MessageLevel.msgError, "Scope reports that it is parked but CanUnPark is false - please manually unpark the scope");
                        g_Stop = true;
                    }
                }
                else
                {
                    LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope is not parked, continuing testing");
                }
            }
            else
            {
                LogMsg("Mount Safety", MessageLevel.msgInfo, "Skipping AtPark test as this method is not supported in interface V" + g_InterfaceVersion);
                try
                {
                    if (canUnpark)
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("Mount Safety", MessageLevel.msgComment, "About to call Unpark method");
                        telescopeDevice.Unpark();
                        LogMsg("Mount Safety", MessageLevel.msgOK, "Scope has been unparked for testing");
                    }
                    else
                    {
                        LogMsg("Mount Safety", MessageLevel.msgOK, "Scope reports that it cannot unpark, unparking skipped");
                    }
                }
                catch (Exception ex)
                {
                    LogMsg("Mount Safety", MessageLevel.msgError, "Driver threw an exception while unparking: " + ex.Message);
                }
            }

            if (!TestStop() & canSetTracking)
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Mount Safety", MessageLevel.msgComment, "About to set Tracking property true");
                telescopeDevice.Tracking = true;
                LogMsg("Mount Safety", MessageLevel.msgInfo, "Scope tracking has been enabled");
            }

            if (!TestStop())
            {
                LogMsg("TimeCheck", MessageLevel.msgInfo, "PC Time Zone:  " + g_Util.TimeZoneName + ", offset " + g_Util.TimeZoneOffset.ToString() + " hours.");
                LogMsg("TimeCheck", MessageLevel.msgInfo, "PC UTCDate:    " + Strings.Format(g_Util.UTCDate, "dd-MMM-yyyy HH:mm:ss.fff"));
                // v1.0.12.0 Added catch logic for any UTCDate issues
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("TimeCheck", MessageLevel.msgComment, "About to get UTCDate property");
                    LogMsg("TimeCheck", MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Mount UTCDate Unformatted: ", telescopeDevice.UTCDate)));
                    LogMsg("TimeCheck", MessageLevel.msgInfo, "Mount UTCDate: " + Strings.Format(telescopeDevice.UTCDate, "dd-MMM-yyyy HH:mm:ss.fff"));
                }
                catch (COMException ex)
                {
                    if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented)
                    {
                        LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: COM exception - UTCDate not implemented in this driver");
                    }
                    else
                    {
                        LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: COM Exception - " + ex.Message);
                    }
                }
                catch (PropertyNotImplementedException)
                {
                    LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: .NET exception - UTCDate not implemented in this driver");
                }
                catch (Exception ex)
                {
                    LogMsg("TimeCheck", MessageLevel.msgError, "Mount UTCDate: .NET Exception - " + ex.Message);
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
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectNotEqual(telescopeDevice.AlignmentMode, AlignmentModes.algGermanPolar, false), canSetPierside)))
                    LogMsg("CanSetPierSide", MessageLevel.msgIssue, "AlignmentMode is not GermanPolar but CanSetPierSide is true - contrary to ASCOM specification");
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
                LogMsg("CanUnPark", MessageLevel.msgIssue, "CanUnPark is true but CanPark is false - this does not comply with ASCOM specification");
        }

        public override void CheckProperties()
        {
            bool l_OriginalTrackingState;
            DriveRates l_DriveRate;
            double l_TimeDifference;
#if DEBUG
            ITrackingRates l_TrackingRates = null;
            DriveRates l_TrackingRate;
#else
            dynamic l_TrackingRates = null;
            dynamic l_TrackingRate;
#endif

            // AlignmentMode - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("AlignmentMode", MessageLevel.msgComment, "About to get AlignmentMode property");
                m_AlignmentMode = (AlignmentModes)telescopeDevice.AlignmentMode;
                LogMsg("AlignmentMode", MessageLevel.msgOK, m_AlignmentMode.ToString());
            }
            catch (Exception ex)
            {
                HandleException("AlignmentMode", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // Altitude - Optional
            try
            {
                canReadAltitide = false;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Altitude", MessageLevel.msgComment, "About to get Altitude property");
                m_Altitude = Conversions.ToDouble(telescopeDevice.Altitude);
                canReadAltitide = true; // Read successfully
                switch (m_Altitude)
                {
                    case var @case when @case < 0.0d:
                        {
                            LogMsg("Altitude", MessageLevel.msgWarning, "Altitude is <0.0 degrees: " + Strings.Format(m_Altitude, "0.00000000"));
                            break;
                        }

                    case var case1 when case1 > 90.0000001d:
                        {
                            LogMsg("Altitude", MessageLevel.msgWarning, "Altitude is >90.0 degrees: " + Strings.Format(m_Altitude, "0.00000000"));
                            break;
                        }

                    default:
                        {
                            LogMsg("Altitude", MessageLevel.msgOK, Strings.Format(m_Altitude, "0.00"));
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Altitude", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // ApertureArea - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("ApertureArea", MessageLevel.msgComment, "About to get ApertureArea property");
                m_ApertureArea = Conversions.ToDouble(telescopeDevice.ApertureArea);
                switch (m_ApertureArea)
                {
                    case var case2 when case2 < 0d:
                        {
                            LogMsg("ApertureArea", MessageLevel.msgWarning, "ApertureArea is < 0.0 : " + m_ApertureArea.ToString());
                            break;
                        }

                    case 0.0d:
                        {
                            LogMsg("ApertureArea", MessageLevel.msgInfo, "ApertureArea is 0.0");
                            break;
                        }

                    default:
                        {
                            LogMsg("ApertureArea", MessageLevel.msgOK, m_ApertureArea.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("ApertureArea", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // ApertureDiameter - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("ApertureDiameter", MessageLevel.msgComment, "About to get ApertureDiameter property");
                m_ApertureDiameter = Conversions.ToDouble(telescopeDevice.ApertureDiameter);
                switch (m_ApertureDiameter)
                {
                    case var case3 when case3 < 0.0d:
                        {
                            LogMsg("ApertureDiameter", MessageLevel.msgWarning, "ApertureDiameter is < 0.0 : " + m_ApertureDiameter.ToString());
                            break;
                        }

                    case 0.0d:
                        {
                            LogMsg("ApertureDiameter", MessageLevel.msgInfo, "ApertureDiameter is 0.0");
                            break;
                        }

                    default:
                        {
                            LogMsg("ApertureDiameter", MessageLevel.msgOK, m_ApertureDiameter.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("ApertureDiameter", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // AtHome - Required
            if (g_InterfaceVersion > 1)
            {
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("AtHome", MessageLevel.msgComment, "About to get AtHome property");
                    m_AtHome = telescopeDevice.AtHome;
                    LogMsg("AtHome", MessageLevel.msgOK, m_AtHome.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("AtHome", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogMsg("AtHome", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // AtPark - Required
            if (g_InterfaceVersion > 1)
            {
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("AtPark", MessageLevel.msgComment, "About to get AtPark property");
                    m_AtPark = telescopeDevice.AtPark;
                    LogMsg("AtPark", MessageLevel.msgOK, m_AtPark.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("AtPark", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogMsg("AtPark", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // Azimuth - Optional
            try
            {
                canReadAzimuth = false;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Azimuth", MessageLevel.msgComment, "About to get Azimuth property");
                m_Azimuth = Conversions.ToDouble(telescopeDevice.Azimuth);
                canReadAzimuth = true; // Read successfully
                switch (m_Azimuth)
                {
                    case var case4 when case4 < 0.0d:
                        {
                            LogMsg("Azimuth", MessageLevel.msgWarning, "Azimuth is <0.0 degrees: " + Strings.Format(m_Azimuth, "0.00"));
                            break;
                        }

                    case var case5 when case5 > 360.0000000001d:
                        {
                            LogMsg("Azimuth", MessageLevel.msgWarning, "Azimuth is >360.0 degrees: " + Strings.Format(m_Azimuth, "0.00"));
                            break;
                        }

                    default:
                        {
                            LogMsg("Azimuth", MessageLevel.msgOK, Strings.Format(m_Azimuth, "0.00"));
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Azimuth", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // Declination - Required
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Declination", MessageLevel.msgComment, "About to get Declination property");
                m_Declination = Conversions.ToDouble(telescopeDevice.Declination);
                switch (m_Declination)
                {
                    case var case6 when case6 < -90.0d:
                    case var case7 when case7 > 90.0d:
                        {
                            LogMsg("Declination", MessageLevel.msgWarning, "Declination is <-90 or >90 degrees: " + FormatDec(m_Declination));
                            break;
                        }

                    default:
                        {
                            LogMsg("Declination", MessageLevel.msgOK, FormatDec(m_Declination));
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Declination", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (TestStop())
                return;

            // DeclinationRate Read - Mandatory - must return a number even when CanSetDeclinationRate is False
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("DeclinationRate Read", MessageLevel.msgComment, "About to get DeclinationRate property");
                m_DeclinationRate = Conversions.ToDouble(telescopeDevice.DeclinationRate);
                // Read has been successful
                if (canSetDeclinationRate) // Any value is acceptable
                {
                    switch (m_DeclinationRate)
                    {
                        case var case8 when case8 >= 0.0d:
                            {
                                LogMsg("DeclinationRate Read", MessageLevel.msgOK, Strings.Format(m_DeclinationRate, "0.00"));
                                break;
                            }

                        default:
                            {
                                LogMsg("DeclinationRate Read", MessageLevel.msgWarning, "Negative DeclinatioRate: " + Strings.Format(m_DeclinationRate, "0.00"));
                                break;
                            }
                    }
                }
                else // Only zero is acceptable
                {
                    switch (m_DeclinationRate)
                    {
                        case 0.0d:
                            {
                                LogMsg("DeclinationRate Read", MessageLevel.msgOK, Strings.Format(m_DeclinationRate, "0.00"));
                                break;
                            }

                        default:
                            {
                                LogMsg("DeclinationRate Read", MessageLevel.msgIssue, "DeclinationRate is non zero when CanSetDeclinationRate is False " + Strings.Format(m_DeclinationRate, "0.00"));
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!canSetDeclinationRate)
                    LogMsg("DeclinationRate Read", MessageLevel.msgIssue, "DeclinationRate must return 0 even when CanSetDeclinationRate is false.");
                HandleException("DeclinationRate Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (TestStop())
                return;

            // DeclinationRate Write - Optional
            if (g_InterfaceVersion > 1)
            {
                if (canSetDeclinationRate) // Any value is acceptable
                {
                    if (TestRADecRate("DeclinationRate Write", "Set rate to 0.0", Axis.Dec, 0.0d, false))
                    {
                        TestRADecRate("DeclinationRate Write", "Set rate to 0.1", Axis.Dec, 0.1d, false);
                        TestRADecRate("DeclinationRate Write", "Set rate to 1.0", Axis.Dec, 1.0d, false);
                        TestRADecRate("DeclinationRate Write", "Reset rate to 0.0", Axis.Dec, 0.0d, true); // Reset the rate to zero, skipping the slewing test
                    }
                }
                else // Should generate an error
                {
                    try
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("DeclinationRate Write", MessageLevel.msgComment, "About to set DeclinationRate property to 0.0");
                        telescopeDevice.DeclinationRate = 0.0d; // Set to a harmless value
                        LogMsg("DeclinationRate", MessageLevel.msgIssue, "CanSetDeclinationRate is False but setting DeclinationRate did not generate an error");
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
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("DeclinationRate Write", MessageLevel.msgComment, "About to set DeclinationRate property to 0.0");
                    telescopeDevice.DeclinationRate = 0.0d; // Set to a harmless value
                    LogMsg("DeclinationRate Write", MessageLevel.msgOK, Strings.Format(m_DeclinationRate, "0.00"));
                }
                catch (Exception ex)
                {
                    HandleException("DeclinationRate Write", MemberType.Property, Required.Optional, ex, "");
                }
            }

            if (TestStop())
                return;

            // DoesRefraction Read - Optional
            if (g_InterfaceVersion > 1)
            {
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("DoesRefraction Read", MessageLevel.msgComment, "About to DoesRefraction get property");
                    m_DoesRefraction = telescopeDevice.DoesRefraction;
                    LogMsg("DoesRefraction Read", MessageLevel.msgOK, m_DoesRefraction.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("DoesRefraction Read", MemberType.Property, Required.Optional, ex, "");
                }

                if (TestStop())
                    return;
            }
            else
            {
                LogMsg("DoesRefraction Read", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // DoesRefraction Write - Optional
            if (g_InterfaceVersion > 1)
            {
                if (m_DoesRefraction) // Try opposite value
                {
                    try
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("DoesRefraction Write", MessageLevel.msgComment, "About to set DoesRefraction property false");
                        telescopeDevice.DoesRefraction = false;
                        LogMsg("DoesRefraction Write", MessageLevel.msgOK, "Can set DoesRefraction to False");
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("DoesRefraction Write", MessageLevel.msgComment, "About to set DoesRefraction property true");
                        telescopeDevice.DoesRefraction = true;
                        LogMsg("DoesRefraction Write", MessageLevel.msgOK, "Can set DoesRefraction to True");
                    }
                    catch (Exception ex)
                    {
                        HandleException("DoesRefraction Write", MemberType.Property, Required.Optional, ex, "");
                    }
                }
            }
            else
            {
                LogMsg("DoesRefraction Write", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // EquatorialSystem - Required
            if (g_InterfaceVersion > 1)
            {
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("EquatorialSystem", MessageLevel.msgComment, "About to get EquatorialSystem property");
                    m_EquatorialSystem = (EquatorialCoordinateType)telescopeDevice.EquatorialSystem;
                    LogMsg("EquatorialSystem", MessageLevel.msgOK, m_EquatorialSystem.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("EquatorialSystem", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogMsg("EquatorialSystem", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // FocalLength - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("FocalLength", MessageLevel.msgComment, "About to get FocalLength property");
                m_FocalLength = Conversions.ToDouble(telescopeDevice.FocalLength);
                switch (m_FocalLength)
                {
                    case var case9 when case9 < 0.0d:
                        {
                            LogMsg("FocalLength", MessageLevel.msgWarning, "FocalLength is <0.0 : " + m_FocalLength.ToString());
                            break;
                        }

                    case 0.0d:
                        {
                            LogMsg("FocalLength", MessageLevel.msgInfo, "FocalLength is 0.0");
                            break;
                        }

                    default:
                        {
                            LogMsg("FocalLength", MessageLevel.msgOK, m_FocalLength.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("FocalLength", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // GuideRateDeclination - Optional
            if (g_InterfaceVersion > 1)
            {
                if (canSetGuideRates) // Can set guide rates so read and write are mandatory
                {
                    try
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateDeclination Read", MessageLevel.msgComment, "About to get GuideRateDeclination property");
                        m_GuideRateDeclination = Conversions.ToDouble(telescopeDevice.GuideRateDeclination); // Read guiderateDEC
                        switch (m_GuideRateDeclination)
                        {
                            case var case10 when case10 < 0.0d:
                                {
                                    LogMsg("GuideRateDeclination Read", MessageLevel.msgWarning, "GuideRateDeclination is < 0.0 " + Strings.Format(m_GuideRateDeclination, "0.00"));
                                    break;
                                }

                            default:
                                {
                                    LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, Strings.Format(m_GuideRateDeclination, "0.00"));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex) // Read failed
                    {
                        HandleException("GuideRateDeclination Read", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True");
                    }

                    try // Read OK so now try to write
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateDeclination Read", MessageLevel.msgComment, "About to set GuideRateDeclination property to " + m_GuideRateDeclination);
                        telescopeDevice.GuideRateDeclination = m_GuideRateDeclination;
                        LogMsg("GuideRateDeclination Write", MessageLevel.msgOK, "Can write Declination Guide Rate OK");
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateDeclination Read", MessageLevel.msgComment, "About to get GuideRateDeclination property");
                        m_GuideRateDeclination = Conversions.ToDouble(telescopeDevice.GuideRateDeclination);
                        switch (m_GuideRateDeclination)
                        {
                            case var case11 when case11 < 0.0d:
                                {
                                    LogMsg("GuideRateDeclination Read", MessageLevel.msgWarning, "GuideRateDeclination is < 0.0 " + Strings.Format(m_GuideRateDeclination, "0.00"));
                                    break;
                                }

                            default:
                                {
                                    LogMsg("GuideRateDeclination Read", MessageLevel.msgOK, Strings.Format(m_GuideRateDeclination, "0.00"));
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateDeclination Write", MessageLevel.msgComment, "About to set GuideRateDeclination property to " + m_GuideRateDeclination);
                        telescopeDevice.GuideRateDeclination = m_GuideRateDeclination;
                        LogMsg("GuideRateDeclination Write", MessageLevel.msgIssue, "CanSetGuideRates is false but no exception generated; value returned: " + Strings.Format(m_GuideRateDeclination, "0.00"));
                    }
                    catch (Exception ex) // Some other error so OK
                    {
                        HandleException("GuideRateDeclination Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False");
                    }
                }
            }
            else
            {
                LogMsg("GuideRateDeclination", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // GuideRateRightAscension - Optional
            if (g_InterfaceVersion > 1)
            {
                if (canSetGuideRates)
                {
                    try
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateRightAscension Read", MessageLevel.msgComment, "About to get GuideRateRightAscension property");
                        m_GuideRateRightAscension = Conversions.ToDouble(telescopeDevice.GuideRateRightAscension); // Read guiderateRA
                        switch (m_GuideRateDeclination)
                        {
                            case var case12 when case12 < 0.0d:
                                {
                                    LogMsg("GuideRateRightAscension Read", MessageLevel.msgWarning, "GuideRateRightAscension is < 0.0 " + Strings.Format(m_GuideRateRightAscension, "0.00"));
                                    break;
                                }

                            default:
                                {
                                    LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, Strings.Format(m_GuideRateRightAscension, "0.00"));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex) // Read failed
                    {
                        HandleException("GuideRateRightAscension Read", MemberType.Property, Required.MustBeImplemented, ex, "CanSetGuideRates is True");
                    }

                    try // Read OK so now try to write
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateRightAscension Read", MessageLevel.msgComment, "About to set GuideRateRightAscension property to " + m_GuideRateRightAscension);
                        telescopeDevice.GuideRateRightAscension = m_GuideRateRightAscension;
                        LogMsg("GuideRateRightAscension Write", MessageLevel.msgOK, "Can set RightAscension Guide OK");
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateRightAscension Read", MessageLevel.msgComment, "About to get GuideRateRightAscension property");
                        m_GuideRateRightAscension = Conversions.ToDouble(telescopeDevice.GuideRateRightAscension); // Read guiderateRA
                        switch (m_GuideRateDeclination)
                        {
                            case var case13 when case13 < 0.0d:
                                {
                                    LogMsg("GuideRateRightAscension Read", MessageLevel.msgWarning, "GuideRateRightAscension is < 0.0 " + Strings.Format(m_GuideRateRightAscension, "0.00"));
                                    break;
                                }

                            default:
                                {
                                    LogMsg("GuideRateRightAscension Read", MessageLevel.msgOK, Strings.Format(m_GuideRateRightAscension, "0.00"));
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("GuideRateRightAscension Write", MessageLevel.msgComment, "About to set GuideRateRightAscension property to " + m_GuideRateRightAscension);
                        telescopeDevice.GuideRateRightAscension = m_GuideRateRightAscension;
                        LogMsg("GuideRateRightAscension Write", MessageLevel.msgIssue, "CanSetGuideRates is false but no exception generated; value returned: " + Strings.Format(m_GuideRateRightAscension, "0.00"));
                    }
                    catch (Exception ex) // Some other error so OK
                    {
                        HandleException("GuideRateRightAscension Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetGuideRates is False");
                    }
                }
            }
            else
            {
                LogMsg("GuideRateRightAscension", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // IsPulseGuiding - Optional
            if (g_InterfaceVersion > 1)
            {
                if (canPulseGuide) // Can pulse guide so test if we can successfully read IsPulseGuiding
                {
                    try
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("IsPulseGuiding", MessageLevel.msgComment, "About to get IsPulseGuiding property");
                        m_IsPulseGuiding = telescopeDevice.IsPulseGuiding;
                        LogMsg("IsPulseGuiding", MessageLevel.msgOK, m_IsPulseGuiding.ToString());
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("IsPulseGuiding", MessageLevel.msgComment, "About to get IsPulseGuiding property");
                        m_IsPulseGuiding = telescopeDevice.IsPulseGuiding;
                        LogMsg("IsPulseGuiding", MessageLevel.msgIssue, "CanPulseGuide is False but no error was raised on calling IsPulseGuiding");
                    }
                    catch (Exception ex)
                    {
                        HandleException("IsPulseGuiding", MemberType.Property, Required.MustNotBeImplemented, ex, "CanPulseGuide is False");
                    }
                }
            }
            else
            {
                LogMsg("IsPulseGuiding", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // RightAscension - Required
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("RightAscension", MessageLevel.msgComment, "About to get RightAscension property");
                m_RightAscension = Conversions.ToDouble(telescopeDevice.RightAscension);
                switch (m_RightAscension)
                {
                    case var case14 when case14 < 0.0d:
                    case var case15 when case15 >= 24.0d:
                        {
                            LogMsg("RightAscension", MessageLevel.msgWarning, "RightAscension is <0 or >=24 hours: " + m_RightAscension + " " + FormatRA(m_RightAscension));
                            break;
                        }

                    default:
                        {
                            LogMsg("RightAscension", MessageLevel.msgOK, FormatRA(m_RightAscension));
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("RightAscension", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (TestStop())
                return;

            // RightAscensionRate Read - Mandatory because read must always return 0.0
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("RightAscensionRate Read", MessageLevel.msgComment, "About to get RightAscensionRate property");
                m_RightAscensionRate = Conversions.ToDouble(telescopeDevice.RightAscensionRate);
                // Read has been successful
                if (canSetRightAscensionRate) // Any value is acceptable
                {
                    switch (m_DeclinationRate)
                    {
                        case var case16 when case16 >= 0.0d:
                            {
                                LogMsg("RightAscensionRate Read", MessageLevel.msgOK, Strings.Format(m_RightAscensionRate, "0.00"));
                                break;
                            }

                        default:
                            {
                                LogMsg("RightAscensionRate Read", MessageLevel.msgWarning, "Negative RightAscensionRate: " + Strings.Format(m_RightAscensionRate, "0.00"));
                                break;
                            }
                    }
                }
                else // Only zero is acceptable
                {
                    switch (m_RightAscensionRate)
                    {
                        case 0.0d:
                            {
                                LogMsg("RightAscensionRate Read", MessageLevel.msgOK, Strings.Format(m_RightAscensionRate, "0.00"));
                                break;
                            }

                        default:
                            {
                                LogMsg("RightAscensionRate Read", MessageLevel.msgIssue, "RightAscensionRate is non zero when CanSetRightAscensionRate is False " + Strings.Format(m_DeclinationRate, "0.00"));
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!canSetRightAscensionRate)
                    LogMsg("RightAscensionRate Read", MessageLevel.msgInfo, "RightAscensionRate must return 0 if CanSetRightAscensionRate is false.");
                HandleException("RightAscensionRate Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (TestStop())
                return;

            // RightAscensionRate Write - Optional
            if (g_InterfaceVersion > 1)
            {
                if (canSetRightAscensionRate) // Perform several tests starting with proving we can set a rate of 0.0
                {
                    if (TestRADecRate("RightAscensionRate Write", "Set rate to 0.0", Axis.RA, 0.0d, false))
                    {
                        TestRADecRate("RightAscensionRate Write", "Set rate to 0.1", Axis.RA, 0.1d, false);
                        TestRADecRate("RightAscensionRate Write", "Set rate to 1.0", Axis.RA, 1.0d, false);
                        TestRADecRate("RightAscensionRate Write", "Reset rate to 0.0", Axis.RA, 0.0d, true); // Reset the rate to zero, skipping the slewing test
                    }
                }
                else // Should generate an error
                {
                    try
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("RightAscensionRate Write", MessageLevel.msgComment, "About to set RightAscensionRate property to 0.00");
                        telescopeDevice.RightAscensionRate = 0.0d; // Set to a harmless value
                        LogMsg("RightAscensionRate Write", MessageLevel.msgIssue, "CanSetRightAscensionRate is False but setting RightAscensionRate did not generate an error");
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
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("RightAscensionRate Write", MessageLevel.msgComment, "About to set RightAscensionRate property to 0.00");
                    telescopeDevice.RightAscensionRate = 0.0d; // Set to a harmless value
                    LogMsg("RightAscensionRate Write", MessageLevel.msgOK, Strings.Format(m_RightAscensionRate, "0.00"));
                }
                catch (Exception ex)
                {
                    HandleException("RightAscensionRate Write", MemberType.Property, Required.Optional, ex, "");
                }
            }

            if (TestStop())
                return;

            // SiteElevation Read - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteElevation Read", MessageLevel.msgComment, "About to get SiteElevation property");
                m_SiteElevation = Conversions.ToDouble(telescopeDevice.SiteElevation);
                switch (m_SiteElevation)
                {
                    case var case17 when case17 < -300.0d:
                        {
                            LogMsg("SiteElevation Read", MessageLevel.msgIssue, "SiteElevation is <-300m");
                            break;
                        }

                    case var case18 when case18 > 10000.0d:
                        {
                            LogMsg("SiteElevation Read", MessageLevel.msgIssue, "SiteElevation is >10,000m");
                            break;
                        }

                    default:
                        {
                            LogMsg("SiteElevation Read", MessageLevel.msgOK, m_SiteElevation.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("SiteElevation Read", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SiteElevation Write - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteElevation Write", MessageLevel.msgComment, "About to set SiteElevation property to -301.0");
                telescopeDevice.SiteElevation = -301.0d;
                LogMsg("SiteElevation Write", MessageLevel.msgIssue, "No error generated on set site elevation < -300m");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation < -300m");
            }

            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteElevation Write", MessageLevel.msgComment, "About to set SiteElevation property to 100001.0");
                telescopeDevice.SiteElevation = 10001.0d;
                LogMsg("SiteElevation Write", MessageLevel.msgIssue, "No error generated on set site elevation > 10,000m");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteElevation Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site elevation > 10,000m");
            }

            try
            {
                if (m_SiteElevation < -300.0d | m_SiteElevation > 10000.0d)
                    m_SiteElevation = 1000d;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteElevation Write", MessageLevel.msgComment, "About to set SiteElevation property to " + m_SiteElevation);
                telescopeDevice.SiteElevation = m_SiteElevation; // Restore original value
                LogMsg("SiteElevation Write", MessageLevel.msgOK, "Legal value " + m_SiteElevation.ToString() + "m written successfully");
            }
            catch (Exception ex)
            {
                HandleException("SiteElevation Write", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SiteLatitude Read - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLatitude Read", MessageLevel.msgComment, "About to get SiteLatitude property");
                m_SiteLatitude = Conversions.ToDouble(telescopeDevice.SiteLatitude);
                switch (m_SiteLatitude)
                {
                    case var case19 when case19 < -90.0d:
                        {
                            LogMsg("SiteLatitude Read", MessageLevel.msgWarning, "SiteLatitude is < -90 degrees");
                            break;
                        }

                    case var case20 when case20 > 90.0d:
                        {
                            LogMsg("SiteLatitude Read", MessageLevel.msgWarning, "SiteLatitude is > 90 degrees");
                            break;
                        }

                    default:
                        {
                            LogMsg("SiteLatitude Read", MessageLevel.msgOK, FormatDec(m_SiteLatitude));
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("SiteLatitude Read", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SiteLatitude Write - Optional
            try // Invalid low value
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLatitude Write", MessageLevel.msgComment, "About to set SiteLatitude property to -91.0");
                telescopeDevice.SiteLatitude = -91.0d;
                LogMsg("SiteLatitude Write", MessageLevel.msgIssue, "No error generated on set site latitude < -90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude < -90 degrees");
            }

            try // Invalid high value
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLatitude Write", MessageLevel.msgComment, "About to set SiteLatitude property to 91.0");
                telescopeDevice.SiteLatitude = 91.0d;
                LogMsg("SiteLatitude Write", MessageLevel.msgIssue, "No error generated on set site latitude > 90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site latitude > 90 degrees");
            }

            try // Valid value
            {
                if (m_SiteLatitude < -90.0d | m_SiteLatitude > 90.0d)
                    m_SiteLatitude = 45.0d;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLatitude Write", MessageLevel.msgComment, "About to set SiteLatitude property to " + m_SiteLatitude);
                telescopeDevice.SiteLatitude = m_SiteLatitude; // Restore original value
                LogMsg("SiteLatitude Write", MessageLevel.msgOK, "Legal value " + FormatDec(m_SiteLatitude) + " degrees written successfully");
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented)
                {
                    LogMsg("SiteLatitude Write", MessageLevel.msgOK, NOT_IMP_COM);
                }
                else
                {
                    ExTest("SiteLatitude Write", ex.Message, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                }
            }
            catch (Exception ex)
            {
                HandleException("SiteLatitude Write", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SiteLongitude Read - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLongitude Read", MessageLevel.msgComment, "About to get SiteLongitude property");
                m_SiteLongitude = Conversions.ToDouble(telescopeDevice.SiteLongitude);
                switch (m_SiteLongitude)
                {
                    case var case21 when case21 < -180.0d:
                        {
                            LogMsg("SiteLongitude Read", MessageLevel.msgWarning, "SiteLongitude is < -180 degrees");
                            break;
                        }

                    case var case22 when case22 > 180.0d:
                        {
                            LogMsg("SiteLongitude Read", MessageLevel.msgWarning, "SiteLongitude is > 180 degrees");
                            break;
                        }

                    default:
                        {
                            LogMsg("SiteLongitude Read", MessageLevel.msgOK, FormatDec(m_SiteLongitude));
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("SiteLongitude Read", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SiteLongitude Write - Optional
            try // Invalid low value
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLongitude Write", MessageLevel.msgComment, "About to set SiteLongitude property to -181.0");
                telescopeDevice.SiteLongitude = -181.0d;
                LogMsg("SiteLongitude Write", MessageLevel.msgIssue, "No error generated on set site longitude < -180 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude < -180 degrees");
            }

            try // Invalid high value
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLongitude Write", MessageLevel.msgComment, "About to set SiteLongitude property to 181.0");
                telescopeDevice.SiteLongitude = 181.0d;
                LogMsg("SiteLongitude Write", MessageLevel.msgIssue, "No error generated on set site longitude > 180 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set site longitude > 180 degrees");
            }

            try // Valid value
            {
                if (m_SiteLongitude < -180.0d | m_SiteLongitude > 180.0d)
                    m_SiteLongitude = 60.0d;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiteLongitude Write", MessageLevel.msgComment, "About to set SiteLongitude property to " + m_SiteLongitude);
                telescopeDevice.SiteLongitude = m_SiteLongitude; // Restore original value
                LogMsg("SiteLongitude Write", MessageLevel.msgOK, "Legal value " + FormatDec(m_SiteLongitude) + " degrees written successfully");
            }
            catch (Exception ex)
            {
                HandleException("SiteLongitude Write", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // Slewing - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Slewing", MessageLevel.msgComment, "About to get Slewing property");
                m_Slewing = telescopeDevice.Slewing;
                switch (m_Slewing)
                {
                    case false:
                        {
                            LogMsg("Slewing", MessageLevel.msgOK, m_Slewing.ToString());
                            break;
                        }

                    case true:
                        {
                            LogMsg("Slewing", MessageLevel.msgIssue, "Slewing should be false and it reads as " + m_Slewing.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Slewing", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SlewSettleTime Read - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SlewSettleTime Read", MessageLevel.msgComment, "About to get SlewSettleTime property");
                m_SlewSettleTime = Conversions.ToShort(telescopeDevice.SlewSettleTime);
                switch (m_SlewSettleTime)
                {
                    case var case23 when case23 < 0:
                        {
                            LogMsg("SlewSettleTime Read", MessageLevel.msgWarning, "SlewSettleTime is < 0 seconds");
                            break;
                        }

                    case var case24 when case24 > (short)Math.Round(30.0d):
                        {
                            LogMsg("SlewSettleTime Read", MessageLevel.msgInfo, "SlewSettleTime is > 30 seconds");
                            break;
                        }

                    default:
                        {
                            LogMsg("SlewSettleTime Read", MessageLevel.msgOK, m_SlewSettleTime.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("SlewSettleTime Read", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SlewSettleTime Write - Optional
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SlewSettleTime Write", MessageLevel.msgComment, "About to set SlewSettleTime property to -1");
                telescopeDevice.SlewSettleTime = -1;
                LogMsg("SlewSettleTime Write", MessageLevel.msgIssue, "No error generated on set SlewSettleTime < 0 seconds");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set slew settle time < 0");
            }

            try
            {
                if (m_SlewSettleTime < 0)
                    m_SlewSettleTime = 0;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SlewSettleTime Write", MessageLevel.msgComment, "About to set SlewSettleTime property to " + m_SlewSettleTime);
                telescopeDevice.SlewSettleTime = m_SlewSettleTime; // Restore original value
                LogMsg("SlewSettleTime Write", MessageLevel.msgOK, "Legal value " + m_SlewSettleTime.ToString() + " seconds written successfully");
            }
            catch (Exception ex)
            {
                HandleException("SlewSettleTime Write", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // SideOfPier Read - Optional
            m_CanReadSideOfPier = false; // Start out assuming that we actually can't read side of pier so the performance test can be omitted
            if (g_InterfaceVersion > 1)
            {
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("SideOfPier Read", MessageLevel.msgComment, "About to get SideOfPier property");
                    m_SideOfPier = (PierSide)telescopeDevice.SideOfPier;
                    LogMsg("SideOfPier Read", MessageLevel.msgOK, m_SideOfPier.ToString());
                    m_CanReadSideOfPier = true; // Flag that it is OK to read SideOfPier
                }
                catch (Exception ex)
                {
                    HandleException("SideOfPier Read", MemberType.Property, Required.Optional, ex, "");
                }
            }
            else
            {
                LogMsg("SideOfPier Read", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // SideOfPier Write - Optional
            // Moved to methods section as this really is a method rather than a property

            // SiderealTime - Required
            try
            {
                canReadSiderealTime = false;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SiderealTime", MessageLevel.msgComment, "About to get SiderealTime property");
                m_SiderealTimeScope = Conversions.ToDouble(telescopeDevice.SiderealTime);
                canReadSiderealTime = true;
                m_SiderealTimeASCOM = (18.697374558d + 24.065709824419081d * (g_Util.DateLocalToJulian(DateAndTime.Now) - 2451545.0d) + m_SiteLongitude / 15.0d) % 24.0d;
                switch (m_SiderealTimeScope)
                {
                    case var case25 when case25 < 0.0d:
                    case var case26 when case26 >= 24.0d:
                        {
                            LogMsg("SiderealTime", MessageLevel.msgWarning, "SiderealTime is <0 or >=24 hours: " + FormatRA(m_SiderealTimeScope)); // Valid time returned
                            break;
                        }

                    default:
                        {
                            // Now do a sense check on the received value
                            LogMsg("SiderealTime", MessageLevel.msgOK, FormatRA(m_SiderealTimeScope));
                            l_TimeDifference = Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM); // Get time difference between scope and PC
                                                                                                    // Process edge cases where the two clocks are on either side of 0:0:0/24:0:0
                            if (m_SiderealTimeASCOM > 23.0d & m_SiderealTimeASCOM < 23.999d & m_SiderealTimeScope > 0.0d & m_SiderealTimeScope < 1.0d)
                            {
                                l_TimeDifference = Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM + 24.0d);
                            }

                            if (m_SiderealTimeScope > 23.0d & m_SiderealTimeScope < 23.999d & m_SiderealTimeASCOM > 0.0d & m_SiderealTimeASCOM < 1.0d)
                            {
                                l_TimeDifference = Math.Abs(m_SiderealTimeScope - m_SiderealTimeASCOM - 24.0d);
                            }

                            switch (l_TimeDifference)
                            {
                                case var case27 when case27 <= 1.0d / 3600.0d: // 1 seconds
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 1 second, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        break;
                                    }

                                case var case28 when case28 <= 2.0d / 3600.0d: // 2 seconds
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 2 seconds, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        break;
                                    }

                                case var case29 when case29 <= 5.0d / 3600.0d: // 5 seconds
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 5 seconds, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        break;
                                    }

                                case var case30 when case30 <= 1.0d / 60.0d: // 1 minute
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 1 minute, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        break;
                                    }

                                case var case31 when case31 <= 5.0d / 60.0d: // 5 minutes
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgOK, "Scope and ASCOM sidereal times agree to better than 5 minutes, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        break;
                                    }

                                case var case32 when case32 <= 0.5d: // 0.5 an hour
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgInfo, "Scope and ASCOM sidereal times are up to 0.5 hour different, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        break;
                                    }

                                case var case33 when case33 <= 1.0d: // 1.0 an hour
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgInfo, "Scope and ASCOM sidereal times are up to 1.0 hour different, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        break;
                                    }

                                default:
                                    {
                                        LogMsg("SiderealTime", MessageLevel.msgError, "Scope and ASCOM sidereal times are more than 1 hour apart, Scope: " + FormatRA(m_SiderealTimeScope) + ", ASCOM: " + FormatRA(m_SiderealTimeASCOM));
                                        MessageBox.Show("Following tests rely on correct sidereal time to calculate target RAs. The sidereal time returned by this driver is more than 1 hour from the expected value based on your computer clock and site longitude, so this program will end now to protect your mount from potential harm caused by slewing to an inappropriate location." + Constants.vbCrLf + Constants.vbCrLf + "Please check the longitude set by the driver and your PC clock (time, time zone and summer time) before checking the sidereal time code in your driver or your mount. Thanks, Peter", "CONFORM - MOUNT PROTECTION", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        g_Stop = true;
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

            if (TestStop())
                return;

            // TargetDeclination Read - Optional
            try // First read should fail!
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetDeclination Read", MessageLevel.msgComment, "About to get TargetDeclination property");
                m_TargetDeclination = Conversions.ToDouble(telescopeDevice.TargetDeclination);
                LogMsg("TargetDeclination Read", MessageLevel.msgIssue, "Read before write should generate an error and didn't");
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
            {
                LogMsg("TargetDeclination Read", MessageLevel.msgOK, "COM Not Set exception generated on read before write");
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.InvalidOperationException)
            {
                LogMsg("TargetDeclination Read", MessageLevel.msgOK, "COM InvalidOperationException generated on read before write");
            }
            catch (ASCOM.InvalidOperationException)
            {
                LogMsg("TargetDeclination Read", MessageLevel.msgOK, ".NET InvalidOperationException generated on read before write");
            }
            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
            {
                LogMsg("TargetDeclination Read", MessageLevel.msgOK, ".NET Not Set exception generated on read before write");
            }
            catch (System.InvalidOperationException)
            {
                LogMsg("TargetDeclination Read", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException");
            }
            catch (Exception ex)
            {
                HandleException("TargetDeclination Read", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // TargetDeclination Write - Optional
            LogMsg("TargetDeclination Write", MessageLevel.msgInfo, "Tests moved after the SlewToCoordinates tests so that Conform can check they properly set target coordinates.");

            // TargetRightAscension Read - Optional
            try // First read should fail!
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetRightAscension Read", MessageLevel.msgComment, "About to get TargetRightAscension property");
                m_TargetRightAscension = Conversions.ToDouble(telescopeDevice.TargetRightAscension);
                LogMsg("TargetRightAscension Read", MessageLevel.msgIssue, "Read before write should generate an error and didn't");
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
            {
                LogMsg("TargetRightAscension Read", MessageLevel.msgOK, "COM Not Set exception generated on read before write");
            }
            catch (COMException ex) when (ex.ErrorCode == ErrorCodes.InvalidOperationException)
            {
                LogMsg("TargetDeclination Read", MessageLevel.msgOK, "COM InvalidOperationException generated on read before write");
            }
            catch (ASCOM.InvalidOperationException)
            {
                LogMsg("TargetRightAscension Read", MessageLevel.msgOK, ".NET InvalidOperationException generated on read before write");
            }
            catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
            {
                LogMsg("TargetRightAscension Read", MessageLevel.msgOK, ".NET Not Set exception generated on read before write");
            }
            catch (System.InvalidOperationException)
            {
                LogMsg("TargetRightAscension Read", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException");
            }
            catch (Exception ex)
            {
                HandleException("TargetRightAscension Read", MemberType.Property, Required.Optional, ex, "");
            }

            if (TestStop())
                return;

            // TargetRightAscension Write - Optional
            LogMsg("TargetRightAscension Write", MessageLevel.msgInfo, "Tests moved after the SlewToCoordinates tests so that Conform can check they properly set target coordinates.");

            // Tracking Read - Required
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Tracking Read", MessageLevel.msgComment, "About to get Tracking property");
                m_Tracking = telescopeDevice.Tracking; // Read of tracking state is mandatory
                LogMsg("Tracking Read", MessageLevel.msgOK, m_Tracking.ToString());
            }
            catch (Exception ex)
            {
                HandleException("Tracking Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (TestStop())
                return;

            // Tracking Write - Optional
            l_OriginalTrackingState = m_Tracking;
            if (canSetTracking) // Set should work OK
            {
                try
                {
                    if (m_Tracking) // OK try turning tracking off
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property false");
                        telescopeDevice.Tracking = false;
                    }
                    else // OK try turning tracking on
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property true");
                        telescopeDevice.Tracking = true;
                    }

                    WaitFor(TRACKING_COMMAND_DELAY); // Wait for a short time to allow mounts to implement the tracking state change
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("Tracking Write", MessageLevel.msgComment, "About to get Tracking property");
                    m_Tracking = telescopeDevice.Tracking;
                    if (m_Tracking != l_OriginalTrackingState)
                    {
                        LogMsg("Tracking Write", MessageLevel.msgOK, m_Tracking.ToString());
                    }
                    else
                    {
                        LogMsg("Tracking Write", MessageLevel.msgIssue, "Tracking didn't change state on write: " + m_Tracking.ToString());
                    }

                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property " + l_OriginalTrackingState);
                    telescopeDevice.Tracking = l_OriginalTrackingState; // Restore original state
                    WaitFor(TRACKING_COMMAND_DELAY); // Wait for a short time to allow mounts to implement the tracking state change
                }
                catch (Exception ex)
                {
                    HandleException("Tracking Write", MemberType.Property, Required.MustBeImplemented, ex, "CanSetTracking is True");
                }
            }
            else // Can read OK but Set tracking should fail
            {
                try
                {
                    if (m_Tracking) // OK try turning tracking off
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property false");
                        telescopeDevice.Tracking = false;
                    }
                    else // OK try turning tracking on
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("Tracking Write", MessageLevel.msgComment, "About to set Tracking property true");
                        telescopeDevice.Tracking = true;
                    }

                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("Tracking Write", MessageLevel.msgComment, "About to get Tracking property");
                    m_Tracking = telescopeDevice.Tracking;
                    LogMsg("Tracking Write", MessageLevel.msgIssue, "CanSetTracking is false but no error generated when value is set");
                }
                catch (Exception ex)
                {
                    HandleException("Tracking Write", MemberType.Property, Required.MustNotBeImplemented, ex, "CanSetTracking is False");
                }
            }

            if (TestStop())
                return;

            // TrackingRates - Required
            if (g_InterfaceVersion > 1)
            {
                int l_Count = 0;
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("TrackingRates", MessageLevel.msgComment, "About to get TrackingRates property");
                    l_TrackingRates = telescopeDevice.TrackingRates;
                    if (l_TrackingRates is null)
                    {
                        LogMsg("TrackingRates", MessageLevel.msgDebug, "ERROR: The driver did NOT return an TrackingRates object!");
                    }
                    else
                    {
                        LogMsg("TrackingRates", MessageLevel.msgDebug, "OK - the driver returned an TrackingRates object");
                    }

                    l_Count = Conversions.ToInteger(l_TrackingRates.Count); // Save count for use later if no members are returned in the for each loop test
                    LogMsg("TrackingRates Count", MessageLevel.msgDebug, l_Count.ToString());

                    var loopTo = Conversions.ToInteger(l_TrackingRates.Count);
                    for (int ii = 1; ii <= loopTo; ii++)
                        LogMsg("TrackingRates Count", MessageLevel.msgDebug, "Found drive rate: " + Enum.GetName(typeof(DriveRates), (l_TrackingRates[ii])));
                }
                catch (Exception ex)
                {
                    HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "");
                }

                try
                {
                    IEnumerator l_Enum;
                    object l_Obj;
                    DriveRates l_Drv;
                    l_Enum = (IEnumerator)l_TrackingRates.GetEnumerator();
                    if (l_Enum is null)
                    {
                        LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "ERROR: The driver did NOT return an Enumerator object!");
                    }
                    else
                    {
                        LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "OK - the driver returned an Enumerator object");
                    }

                    l_Enum.Reset();
                    LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Reset Enumerator");
                    while (l_Enum.MoveNext())
                    {
                        LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Reading Current");
                        l_Obj = l_Enum.Current;
                        LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Read Current OK, Type: " + l_Obj.GetType().Name);
                        l_Drv = (DriveRates)Conversions.ToInteger(l_Obj);
                        LogMsg("TrackingRates Enum", MessageLevel.msgDebug, "Found drive rate: " + Enum.GetName(typeof(DriveRates), l_Drv));
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

                        try
                        {
                            Marshal.ReleaseComObject(l_TrackingRates);
                        }
                        catch
                        {
                        }

                        l_TrackingRates = null;
                    }
                }
                catch (Exception ex)
                {
                    HandleException("TrackingRates", MemberType.Property, Required.Mandatory, ex, "");
                }

                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("TrackingRates", MessageLevel.msgComment, "About to get TrackingRates property");
                    l_TrackingRates = telescopeDevice.TrackingRates;
                    LogMsg("TrackingRates", MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Read TrackingRates OK, Count: ", l_TrackingRates.Count)));
                    int l_RateCount = 0;
                    foreach (DriveRates currentL_DriveRate in (IEnumerable)l_TrackingRates)
                    {
                        l_DriveRate = currentL_DriveRate;
                        LogMsg("TrackingRates", MessageLevel.msgComment, "Found drive rate: " + l_DriveRate.ToString());
                        l_RateCount += 1;
                    }

                    if (l_RateCount > 0)
                    {
                        LogMsg("TrackingRates", MessageLevel.msgOK, "Drive rates read OK");
                    }
                    else if (l_Count > 0) // We did get some members on the first call, but now they have disappeared!
                    {
                        // This can be due to the driver returning the same TrackingRates object on every TrackingRates call but not resetting the iterator pointer
                        LogMsg("TrackingRates", MessageLevel.msgError, "Multiple calls to TrackingRates returned different answers!");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "The first call to TrackingRates returned " + l_Count + " drive rates; the next call appeared to return no rates.");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "This can arise when the SAME TrackingRates object is returned on every TrackingRates call.");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "The root cause is usually that the enumeration pointer in the object is set to the end of the");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "collection through the application's use of the first object; subsequent uses see the pointer at the end");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "of the collection, which indicates no more members and is interpreted as meaning the collection is empty.");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "It is recommended to return a new TrackingRates object on each call. Alternatively, you could reset the");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "object's enumeration pointer every time the GetEnumerator method is called.");
                        LogMsg("TrackingRates", MessageLevel.msgInfo, "");
                    }
                    else
                    {
                        LogMsg("TrackingRates", MessageLevel.msgIssue, "No drive rates returned");
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

                    try
                    {
                        Marshal.ReleaseComObject(l_TrackingRates);
                    }
                    catch
                    {
                    }

                    l_TrackingRates = null;
                }

                // Test the TrackingRates.Dispose() method
                LogMsg("TrackingRates", MessageLevel.msgDebug, "Getting tracking rates");
                l_TrackingRates = DriverAsObject.TrackingRates;
                try
                {
                    LogMsg("TrackingRates", MessageLevel.msgDebug, "Disposing tracking rates");
                    l_TrackingRates.Dispose();
                    LogMsg("TrackingRates", MessageLevel.msgOK, "Disposed tracking rates OK");
                }
                catch (MissingMemberException)
                {
                    LogMsg("TrackingRates", MessageLevel.msgOK, "Dispose member not present");
                }
                catch (Exception ex)
                {
                    LogMsgWarning("TrackingRates", "TrackingRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: " + ex.Message);
                    LogMsg("TrackingRates.Dispose", MessageLevel.msgDebug, "Exception: " + ex.ToString());
                }

                try
                {
                    Marshal.ReleaseComObject(l_TrackingRates);
                }
                catch
                {
                }

                l_TrackingRates = null;
                if (TestStop())
                    return;

                // TrackingRate - Test after TrackingRates so we know what the valid values are
                // TrackingRate Read - Required
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("TrackingRates", MessageLevel.msgComment, "About to get TrackingRates property");
                    l_TrackingRates = telescopeDevice.TrackingRates;
                    if (l_TrackingRates is object) // Make sure that we have received a TrackingRates object after the Dispose() method was called
                    {
                        LogMsgOK("TrackingRates", "Successfully obtained a TrackingRates object after the previous TrackingRates object was disposed");
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("TrackingRate Read", MessageLevel.msgComment, "About to get TrackingRate property");
                        l_TrackingRate = telescopeDevice.TrackingRate;
                        LogMsg("TrackingRate Read", MessageLevel.msgOK, l_TrackingRate.ToString());

                        // TrackingRate Write - Optional
                        // We can read TrackingRate so now test trying to set each tracking rate in turn
                        try
                        {
                            LogMsgDebug("TrackingRate Write", "About to enumerate tracking rates object");
                            foreach (DriveRates currentL_DriveRate1 in (IEnumerable)l_TrackingRates)
                            {
                                l_DriveRate = currentL_DriveRate1;
                                Application.DoEvents();
                                if (TestStop())
                                    return;
                                try
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to " + l_DriveRate.ToString());
                                    telescopeDevice.TrackingRate = l_DriveRate;
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to get TrackingRate property");
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(telescopeDevice.TrackingRate, l_DriveRate, false)))
                                    {
                                        LogMsg("TrackingRate Write", MessageLevel.msgOK, "Successfully set drive rate: " + l_DriveRate.ToString());
                                    }
                                    else
                                    {
                                        LogMsg("TrackingRate Write", MessageLevel.msgIssue, "Unable to set drive rate: " + l_DriveRate.ToString());
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
                            LogMsgError("TrackingRate Write 1", "A NullReferenceException was thrown while iterating a new TrackingRates instance after a previous TrackingRates instance was disposed. TrackingRate.Write testing skipped");
                            LogMsgInfo("TrackingRate Write 1", "This may indicate that the TrackingRates.Dispose method cleared a global variable shared by all TrackingRates instances.");
                        }
                        catch (Exception ex)
                        {
                            HandleException("TrackingRate Write 1", MemberType.Property, Required.Mandatory, ex, "");
                        }

                        // Attempt to write an invalid high tracking rate
                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to invalid value (5)");
                            telescopeDevice.TrackingRate = (DriveRates)5;
                            LogMsg("TrackingRate Write", MessageLevel.msgIssue, "No error generated when TrackingRate is set to an invalid value (5)");
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (5)");
                        }

                        // Attempt to write an invalid low tracking rate
                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to invalid value (-1)");
                            telescopeDevice.TrackingRate = (DriveRates)(0 - 1); // Done this way to fool the compiler into allowing me to attempt to set a negative, invalid value
                            LogMsg("TrackingRate Write", MessageLevel.msgIssue, "No error generated when TrackingRate is set to an invalid value (-1)");
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK("TrackingRate Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected when TrackingRate is set to an invalid value (-1)");
                        }

                        // Finally restore original TrackingRate
                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg("TrackingRate Write", MessageLevel.msgComment, "About to set TrackingRate property to " + l_TrackingRate.ToString());
                            telescopeDevice.TrackingRate = l_TrackingRate;
                        }
                        catch (Exception ex)
                        {
                            HandleException("TrackingRate Write", MemberType.Property, Required.Optional, ex, "Unable to restore original tracking rate");
                        }
                    }
                    else // No TrackingRates object received after disposing of a previous instance
                    {
                        LogMsgError("TrackingRate Write", "TrackingRates did not return an object after calling Disposed() on a previous instance, TrackingRate.Write testing skipped");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("TrackingRate Read", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogMsg("TrackingRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (TestStop())
                return;

            // UTCDate Read - Required
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("UTCDate Read", MessageLevel.msgComment, "About to get UTCDate property");
                m_UTCDate = Conversions.ToDate(telescopeDevice.UTCDate); // Save starting value
                LogMsg("UTCDate Read", MessageLevel.msgOK, m_UTCDate.ToString("dd-MMM-yyyy HH:mm:ss.fff"));
                try // UTCDate Write is optional since if you are using the PC time as UTCTime then you should not write to the PC clock!
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("UTCDate Write", MessageLevel.msgComment, "About to set UTCDate property to " + m_UTCDate.AddHours(1.0d).ToString());
                    telescopeDevice.UTCDate = m_UTCDate.AddHours(1.0d); // Try and write a new UTCDate in the future
                    LogMsg("UTCDate Write", MessageLevel.msgOK, "New UTCDate written successfully: " + m_UTCDate.AddHours(1.0d).ToString());
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("UTCDate Write", MessageLevel.msgComment, "About to set UTCDate property to " + m_UTCDate.ToString());
                    telescopeDevice.UTCDate = m_UTCDate; // Restore original value
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

            if (TestStop())
                return;
        }

        public override void CheckMethods()
        {

            // CanMoveAxis - Required - This must be first test as Parked tests use its results
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_CAN_MOVE_AXIS] == CheckState.Checked | g_TelescopeTests[TELTEST_MOVE_AXIS] == CheckState.Checked | g_TelescopeTests[TELTEST_PARK_UNPARK] == CheckState.Checked)
                {
                    TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisPrimary, "CanMoveAxis:Primary");
                    if (TestStop())
                        return;
                    TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisSecondary, "CanMoveAxis:Secondary");
                    if (TestStop())
                        return;
                    TelescopeRequiredMethodsTest(RequiredMethodType.tstCanMoveAxisTertiary, "CanMoveAxis:Tertiary");
                    if (TestStop())
                        return;
                }
                else
                {
                    LogMsg(TELTEST_CAN_MOVE_AXIS, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("CanMoveAxis", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // Test Park, Unpark - Optional
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_PARK_UNPARK] == CheckState.Checked)
                {
                    if (canPark) // Can Park
                    {
                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg("Park", MessageLevel.msgComment, "About to get AtPark property");
                            if (Conversions.ToBoolean(!telescopeDevice.AtPark)) // OK We are unparked so check that no error is generated
                            {
                                Status(StatusType.staTest, "Park");
                                try
                                {
                                    Status(StatusType.staAction, "Park scope");
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("Park", MessageLevel.msgComment, "About to call Park method");
                                    telescopeDevice.Park();
                                    Status(StatusType.staStatus, "Waiting for scope to park");
                                    do
                                    {
                                        WaitFor(SLEEP_TIME);
                                        Application.DoEvents();
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg("Park", MessageLevel.msgComment, "About to get AtPark property");
                                    }
                                    while (!telescopeDevice.AtPark & !TestStop());
                                    if (TestStop())
                                        return;
                                    Status(StatusType.staStatus, "Scope parked");
                                    LogMsg("Park", MessageLevel.msgOK, "Success");

                                    // Scope Parked OK
                                    try // Confirm second park is harmless
                                    {
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg("Park", MessageLevel.msgComment, "About to Park call method");
                                        telescopeDevice.Park();
                                        LogMsg("Park", MessageLevel.msgOK, "Success if already parked");
                                    }
                                    catch (COMException ex)
                                    {
                                        LogMsg("Park", MessageLevel.msgIssue, "Exception when calling Park two times in succession: " + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                                    }
                                    catch (Exception ex)
                                    {
                                        LogMsg("Park", MessageLevel.msgIssue, "Exception when calling Park two times in succession: " + ex.Message);
                                    }

                                    // Confirm that methods do raise exceptions when scope is parked
                                    if (canSlew | canSlewAsync | canSlewAltAz | canSlewAltAzAsync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepAbortSlew, "AbortSlew");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canFindHome)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepFindHome, "FindHome");
                                        if (TestStop())
                                            return;
                                    }

                                    if (m_CanMoveAxisPrimary)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisPrimary, "MoveAxis Primary");
                                        if (TestStop())
                                            return;
                                    }

                                    if (m_CanMoveAxisSecondary)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisSecondary, "MoveAxis Secondary");
                                        if (TestStop())
                                            return;
                                    }

                                    if (m_CanMoveAxisTertiary)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepMoveAxisTertiary, "MoveAxis Tertiary");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canPulseGuide)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepPulseGuide, "PulseGuide");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canSlew)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinates, "SlewToCoordinates");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canSlewAsync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToCoordinatesAsync, "SlewToCoordinatesAsync");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canSlew)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTarget, "SlewToTarget");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canSlewAsync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSlewToTargetAsync, "SlewToTargetAsync");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canSync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToCoordinates, "SyncToCoordinates");
                                        if (TestStop())
                                            return;
                                    }

                                    if (canSync)
                                    {
                                        TelescopeParkedExceptionTest(ParkedExceptionType.tstPExcepSyncToTarget, "SyncToTarget");
                                        if (TestStop())
                                            return;
                                    }

                                    // Test unpark after park
                                    if (canUnpark)
                                    {
                                        try
                                        {
                                            Status(StatusType.staAction, "UnPark scope after park");
                                            if (g_Settings.DisplayMethodCalls)
                                                LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method");
                                            telescopeDevice.Unpark();
                                            do
                                            {
                                                WaitFor(SLEEP_TIME);
                                                Application.DoEvents();
                                                if (g_Settings.DisplayMethodCalls)
                                                    LogMsg("UnPark", MessageLevel.msgComment, "About to get AtPark property");
                                            }
                                            while (telescopeDevice.AtPark & !TestStop());
                                            if (TestStop())
                                                return;
                                            try // Make sure tracking doesn't generate an error if it is not implemented
                                            {
                                                if (g_Settings.DisplayMethodCalls)
                                                    LogMsg("UnPark", MessageLevel.msgComment, "About to set Tracking property true");
                                                telescopeDevice.Tracking = true;
                                            }
                                            catch (Exception)
                                            {
                                            }

                                            Status(StatusType.staStatus, "Scope UnParked");
                                            LogMsg("UnPark", MessageLevel.msgOK, "Success");

                                            // Scope unparked
                                            try // Confirm UnPark is harmless if already unparked
                                            {
                                                if (g_Settings.DisplayMethodCalls)
                                                    LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method");
                                                telescopeDevice.Unpark();
                                                LogMsg("UnPark", MessageLevel.msgOK, "Success if already unparked");
                                            }
                                            catch (COMException ex)
                                            {
                                                LogMsg("UnPark", MessageLevel.msgIssue, "Exception when calling UnPark two times in succession: " + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                                            }
                                            catch (Exception ex)
                                            {
                                                LogMsg("UnPark", MessageLevel.msgIssue, "Exception when calling UnPark two times in succession: " + ex.Message);
                                            }
                                        }
                                        catch (COMException ex)
                                        {
                                            LogMsg("UnPark", MessageLevel.msgError, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                                        }
                                        catch (Exception ex)
                                        {
                                            LogMsg("UnPark", MessageLevel.msgError, EX_NET + ex.Message);
                                        }
                                    }
                                    else // Can't UnPark
                                    {
                                        // Confirm that UnPark generates an error
                                        try
                                        {
                                            if (g_Settings.DisplayMethodCalls)
                                                LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method");
                                            telescopeDevice.Unpark();
                                            LogMsg("UnPark", MessageLevel.msgIssue, "No exception thrown by UnPark when CanUnPark is false");
                                        }
                                        catch (COMException ex)
                                        {
                                            if (ex.ErrorCode == g_ExNotImplemented | ex.ErrorCode == ErrorCodes.NotImplemented)
                                            {
                                                LogMsg("UnPark", MessageLevel.msgOK, NOT_IMP_COM);
                                            }
                                            else
                                            {
                                                ExTest("UnPark", ex.Message, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                                            }
                                        }
                                        catch (MethodNotImplementedException)
                                        {
                                            LogMsg("UnPark", MessageLevel.msgOK, NOT_IMP_NET);
                                        }
                                        catch (Exception ex)
                                        {
                                            ExTest("UnPark", ex.Message, EX_NET + ex.Message);
                                        }
                                        // Create user interface message asking for manual scope UnPark
                                        LogMsg("UnPark", MessageLevel.msgComment, "CanUnPark is false so you need to unpark manually");
                                        MessageBox.Show("This scope cannot be unparked automatically, please unpark it now", "UnPark", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                                catch (COMException ex)
                                {
                                    LogMsg("Park", MessageLevel.msgError, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                                }
                                catch (Exception ex)
                                {
                                    LogMsg("Park", MessageLevel.msgError, EX_NET + ex.Message);
                                }
                            }
                            else // We are still in parked status despite a successful UnPark
                            {
                                LogMsg("Park", MessageLevel.msgError, "AtPark still true despite an earlier successful unpark");
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleException("Park", MemberType.Method, Required.MustBeImplemented, ex, "CanPark is True");
                        }
                    }
                    else // Can't park
                    {
                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg("UnPark", MessageLevel.msgComment, "About to call Park method");
                            telescopeDevice.Park();
                            LogMsg("Park", MessageLevel.msgError, "CanPark is false but no exception was generated on use");
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
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method");
                                telescopeDevice.Unpark();
                                LogMsg("UnPark", MessageLevel.msgOK, "CanPark is false and CanUnPark is true; no exception generated as expected");
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
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("UnPark", MessageLevel.msgComment, "About to call UnPark method");
                                telescopeDevice.Unpark();
                                LogMsg("UnPark", MessageLevel.msgError, "CanPark and CanUnPark are false but no exception was generated on use");
                            }
                            catch (Exception ex)
                            {
                                HandleException("UnPark", MemberType.Method, Required.MustNotBeImplemented, ex, "CanUnPark is False");
                            }
                        }
                    }

                    g_Status.Clear(); // Clear status messages
                    if (TestStop())
                        return;
                }
                else
                {
                    LogMsg(TELTEST_PARK_UNPARK, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("Park", MessageLevel.msgInfo, "Skipping tests since behaviour of this method is not well defined in interface V" + g_InterfaceVersion);
            }

            // AbortSlew - Optional
            if (g_TelescopeTests[TELTEST_ABORT_SLEW] == CheckState.Checked)
            {
                TelescopeOptionalMethodsTest(OptionalMethodType.AbortSlew, "AbortSlew", true);
                if (TestStop())
                    return;
            }
            else
            {
                LogMsg(TELTEST_ABORT_SLEW, MessageLevel.msgInfo, "Tests skipped");
            }

            // AxisRates - Required
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_AXIS_RATE] == CheckState.Checked | g_TelescopeTests[TELTEST_MOVE_AXIS] == CheckState.Checked)
                {
                    TelescopeAxisRateTest("AxisRate:Primary", TelescopeAxes.axisPrimary);
                    TelescopeAxisRateTest("AxisRate:Secondary", TelescopeAxes.axisSecondary);
                    TelescopeAxisRateTest("AxisRate:Tertiary", TelescopeAxes.axisTertiary);
                }
                else
                {
                    LogMsg(TELTEST_AXIS_RATE, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("AxisRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // FindHome - Optional
            if (g_TelescopeTests[TELTEST_FIND_HOME] == CheckState.Checked)
            {
                TelescopeOptionalMethodsTest(OptionalMethodType.FindHome, "FindHome", canFindHome);
                if (TestStop())
                    return;
            }
            else
            {
                LogMsg(TELTEST_FIND_HOME, MessageLevel.msgInfo, "Tests skipped");
            }

            // MoveAxis - Optional
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_MOVE_AXIS] == CheckState.Checked)
                {
                    TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisPrimary, "MoveAxis Primary", m_CanMoveAxisPrimary);
                    if (TestStop())
                        return;
                    TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisSecondary, "MoveAxis Secondary", m_CanMoveAxisSecondary);
                    if (TestStop())
                        return;
                    TelescopeOptionalMethodsTest(OptionalMethodType.MoveAxisTertiary, "MoveAxis Tertiary", m_CanMoveAxisTertiary);
                    if (TestStop())
                        return;
                }
                else
                {
                    LogMsg(TELTEST_MOVE_AXIS, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("MoveAxis", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // PulseGuide - Optional
            if (g_TelescopeTests[TELTEST_PULSE_GUIDE] == CheckState.Checked)
            {
                TelescopeOptionalMethodsTest(OptionalMethodType.PulseGuide, "PulseGuide", canPulseGuide);
                if (TestStop())
                    return;
            }
            else
            {
                LogMsg(TELTEST_PULSE_GUIDE, MessageLevel.msgInfo, "Tests skipped");
            }

            // Test Equatorial slewing to coordinates - Optional
            if (g_TelescopeTests[TELTEST_SLEW_TO_COORDINATES] == CheckState.Checked)
            {
                TelescopeSlewTest(SlewSyncType.SlewToCoordinates, "SlewToCoordinates", canSlew, "CanSlew");
                if (TestStop())
                    return;
                if (canSlew) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToCoordinates (Bad L)", SlewSyncType.SlewToCoordinates, BAD_RA_LOW, BAD_DEC_LOW);
                    if (TestStop())
                        return;
                    TelescopeBadCoordinateTest("SlewToCoordinates (Bad H)", SlewSyncType.SlewToCoordinates, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (TestStop())
                        return;
                }
            }
            else
            {
                LogMsg(TELTEST_SLEW_TO_COORDINATES, MessageLevel.msgInfo, "Tests skipped");
            }

            // Test Equatorial slewing to coordinates asynchronous - Optional
            if (g_TelescopeTests[TELTEST_SLEW_TO_COORDINATES_ASYNC] == CheckState.Checked)
            {
                TelescopeSlewTest(SlewSyncType.SlewToCoordinatesAsync, "SlewToCoordinatesAsync", canSlewAsync, "CanSlewAsync");
                if (TestStop())
                    return;
                if (canSlewAsync) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToCoordinatesAsync (Bad L)", SlewSyncType.SlewToCoordinatesAsync, BAD_RA_LOW, BAD_DEC_LOW);
                    if (TestStop())
                        return;
                    TelescopeBadCoordinateTest("SlewToCoordinatesAsync (Bad H)", SlewSyncType.SlewToCoordinatesAsync, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (TestStop())
                        return;
                }
            }
            else
            {
                LogMsg(TELTEST_SLEW_TO_COORDINATES_ASYNC, MessageLevel.msgInfo, "Tests skipped");
            }

            // Equatorial Sync to Coordinates - Optional - Moved here so that it can be tested before any target coordinates are set - Peter 4th August 2018
            if (g_TelescopeTests[TELTEST_SYNC_TO_COORDINATES] == CheckState.Checked)
            {
                TelescopeSyncTest(SlewSyncType.SyncToCoordinates, "SyncToCoordinates", canSync, "CanSync");
                if (TestStop())
                    return;
                if (canSync) // Test syncing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SyncToCoordinates (Bad L)", SlewSyncType.SyncToCoordinates, BAD_RA_LOW, BAD_DEC_LOW);
                    if (TestStop())
                        return;
                    TelescopeBadCoordinateTest("SyncToCoordinates (Bad H)", SlewSyncType.SyncToCoordinates, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (TestStop())
                        return;
                }
            }
            else
            {
                LogMsg(TELTEST_SYNC_TO_COORDINATES, MessageLevel.msgInfo, "Tests skipped");
            }

            // TargetRightAscension Write - Optional - Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to set TargetRightAscension property to -1.0");
                telescopeDevice.TargetRightAscension = -1.0d;
                LogMsg("TargetRightAscension Write", MessageLevel.msgIssue, "No error generated on set TargetRightAscension < 0 hours");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension < 0 hours");
            }

            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to set TargetRightAscension property to 25.0");
                telescopeDevice.TargetRightAscension = 25.0d;
                LogMsg("TargetRightAscension Write", MessageLevel.msgIssue, "No error generated on set TargetRightAscension > 24 hours");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetRightAscension Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetRightAscension > 24 hours");
            }

            try
            {
                m_TargetRightAscension = TelescopeRAFromSiderealTime("TargetRightAscension Write", -4.0d);
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to set TargetRightAscension property to " + m_TargetRightAscension);
                telescopeDevice.TargetRightAscension = m_TargetRightAscension; // Set a valid value
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("TargetRightAscension Write", MessageLevel.msgComment, "About to get TargetRightAscension property");
                    switch (Math.Abs(telescopeDevice.TargetRightAscension - m_TargetRightAscension))
                    {
                        case 0.0d:
                            {
                                LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Legal value " + FormatRA(m_TargetRightAscension) + " HH:MM:SS written successfully");
                                break;
                            }

                        case var @case when @case <= 1.0d / 3600.0d: // 1 seconds
                            {
                                LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 1 second of the value set: " + FormatRA(m_TargetRightAscension));
                                break;
                            }

                        case var case1 when case1 <= 2.0d / 3600.0d: // 2 seconds
                            {
                                LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 2 seconds of the value set: " + FormatRA(m_TargetRightAscension));
                                break;
                            }

                        case var case2 when case2 <= 5.0d / 3600.0d: // 5 seconds
                            {
                                LogMsg("TargetRightAscension Write", MessageLevel.msgOK, "Target RightAscension is within 5 seconds of the value set: " + FormatRA(m_TargetRightAscension));
                                break;
                            }

                        default:
                            {
                                LogMsg("TargetRightAscension Write", MessageLevel.msgInfo, "Target RightAscension: " + FormatRA(Conversions.ToDouble(telescopeDevice.TargetRightAscension)));
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

            if (TestStop())
                return;

            // TargetDeclination Write - Optional - Test moved here so that Conform can check that the SlewTo... methods properly set target coordinates.")
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to set TargetDeclination property to -91.0");
                telescopeDevice.TargetDeclination = -91.0d;
                LogMsg("TargetDeclination Write", MessageLevel.msgIssue, "No error generated on set TargetDeclination < -90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
            }

            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to set TargetDeclination property to 91.0");
                telescopeDevice.TargetDeclination = 91.0d;
                LogMsg("TargetDeclination Write", MessageLevel.msgIssue, "No error generated on set TargetDeclination > 90 degrees");
            }
            catch (Exception ex)
            {
                HandleInvalidValueExceptionAsOK("TargetDeclination Write", MemberType.Property, Required.Optional, ex, "", "Invalid Value exception generated as expected on set TargetDeclination < -90 degrees");
            }

            try
            {
                m_TargetDeclination = 1.0d;
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to set TargetDeclination property to " + m_TargetDeclination);
                telescopeDevice.TargetDeclination = m_TargetDeclination; // Set a valid value
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("TargetDeclination Write", MessageLevel.msgComment, "About to get TargetDeclination property");
                    switch (Math.Abs(telescopeDevice.TargetDeclination - m_TargetDeclination))
                    {
                        case 0.0d:
                            {
                                LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Legal value " + FormatDec(m_TargetDeclination) + " DD:MM:SS written successfully");
                                break;
                            }

                        case var case3 when case3 <= 1.0d / 3600.0d: // 1 seconds
                            {
                                LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 1 second of the value set: " + FormatDec(m_TargetDeclination));
                                break;
                            }

                        case var case4 when case4 <= 2.0d / 3600.0d: // 2 seconds
                            {
                                LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 2 seconds of the value set: " + FormatDec(m_TargetDeclination));
                                break;
                            }

                        case var case5 when case5 <= 5.0d / 3600.0d: // 5 seconds
                            {
                                LogMsg("TargetDeclination Write", MessageLevel.msgOK, "Target Declination is within 5 seconds of the value set: " + FormatDec(m_TargetDeclination));
                                break;
                            }

                        default:
                            {
                                LogMsg("TargetDeclination Write", MessageLevel.msgInfo, "Target Declination: " + FormatDec(m_TargetDeclination));
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

            if (TestStop())
                return;

            // Test Equatorial target slewing - Optional
            if (g_TelescopeTests[TELTEST_SLEW_TO_TARGET] == CheckState.Checked)
            {
                TelescopeSlewTest(SlewSyncType.SlewToTarget, "SlewToTarget", canSlew, "CanSlew");
                if (TestStop())
                    return;
                if (canSlew) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToTarget (Bad L)", SlewSyncType.SlewToTarget, BAD_RA_LOW, BAD_DEC_LOW);
                    if (TestStop())
                        return;
                    TelescopeBadCoordinateTest("SlewToTarget (Bad H)", SlewSyncType.SlewToTarget, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (TestStop())
                        return;
                }
            }
            else
            {
                LogMsg(TELTEST_SLEW_TO_TARGET, MessageLevel.msgInfo, "Tests skipped");
            }

            // Test Equatorial target slewing asynchronous - Optional
            if (g_TelescopeTests[TELTEST_SLEW_TO_TARGET_ASYNC] == CheckState.Checked)
            {
                TelescopeSlewTest(SlewSyncType.SlewToTargetAsync, "SlewToTargetAsync", canSlewAsync, "CanSlewAsync");
                if (TestStop())
                    return;
                if (canSlewAsync) // Test slewing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SlewToTargetAsync (Bad L)", SlewSyncType.SlewToTargetAsync, BAD_RA_LOW, BAD_DEC_LOW);
                    if (TestStop())
                        return;
                    TelescopeBadCoordinateTest("SlewToTargetAsync (Bad H)", SlewSyncType.SlewToTargetAsync, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (TestStop())
                        return;
                }
            }
            else
            {
                LogMsg(TELTEST_SLEW_TO_TARGET_ASYNC, MessageLevel.msgInfo, "Tests skipped");
            }

            // DestinationSideOfPier - Optional
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_DESTINATION_SIDE_OF_PIER] == CheckState.Checked)
                {
                    if (m_AlignmentMode == AlignmentModes.algGermanPolar)
                    {
                        TelescopeOptionalMethodsTest(OptionalMethodType.DestinationSideOfPier, "DestinationSideOfPier", true);
                        if (TestStop())
                            return;
                    }
                    else
                    {
                        LogMsg("DestinationSideOfPier", MessageLevel.msgComment, "Test skipped as AligmentMode is not German Polar");
                    }
                }
                else
                {
                    LogMsg(TELTEST_DESTINATION_SIDE_OF_PIER, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("DestinationSideOfPier", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // Test AltAz Slewing - Optional
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_SLEW_TO_ALTAZ] == CheckState.Checked)
                {
                    TelescopeSlewTest(SlewSyncType.SlewToAltAz, "SlewToAltAz", canSlewAltAz, "CanSlewAltAz");
                    if (TestStop())
                        return;
                    if (canSlewAltAz) // Test slewing to bad co-ordinates
                    {
                        TelescopeBadCoordinateTest("SlewToAltAz (Bad L)", SlewSyncType.SlewToAltAz, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW);
                        if (TestStop())
                            return; // -100 is used for the Altitude limit to enable -90 to be used for parking the scope
                        TelescopeBadCoordinateTest("SlewToAltAz (Bad H)", SlewSyncType.SlewToAltAz, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH);
                        if (TestStop())
                            return;
                    }
                }
                else
                {
                    LogMsg(TELTEST_SLEW_TO_ALTAZ, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("SlewToAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // Test AltAz Slewing asynchronous - Optional
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_SLEW_TO_ALTAZ_ASYNC] == CheckState.Checked)
                {
                    TelescopeSlewTest(SlewSyncType.SlewToAltAzAsync, "SlewToAltAzAsync", canSlewAltAzAsync, "CanSlewAltAzAsync");
                    if (TestStop())
                        return;
                    if (canSlewAltAzAsync) // Test slewing to bad co-ordinates
                    {
                        TelescopeBadCoordinateTest("SlewToAltAzAsync (Bad L)", SlewSyncType.SlewToAltAzAsync, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW);
                        if (TestStop())
                            return;
                        TelescopeBadCoordinateTest("SlewToAltAzAsync (Bad H)", SlewSyncType.SlewToAltAzAsync, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH);
                        if (TestStop())
                            return;
                    }
                }
                else
                {
                    LogMsg(TELTEST_SLEW_TO_ALTAZ_ASYNC, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("SlewToAltAzAsync", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // Equatorial Sync to Target - Optional
            if (g_TelescopeTests[TELTEST_SYNC_TO_TARGET] == CheckState.Checked)
            {
                TelescopeSyncTest(SlewSyncType.SyncToTarget, "SyncToTarget", canSync, "CanSync");
                if (TestStop())
                    return;
                if (canSync) // Test syncing to bad co-ordinates
                {
                    TelescopeBadCoordinateTest("SyncToTarget (Bad L)", SlewSyncType.SyncToTarget, BAD_RA_LOW, BAD_DEC_LOW);
                    if (TestStop())
                        return;
                    TelescopeBadCoordinateTest("SyncToTarget (Bad H)", SlewSyncType.SyncToTarget, BAD_RA_HIGH, BAD_DEC_HIGH);
                    if (TestStop())
                        return;
                }
            }
            else
            {
                LogMsg(TELTEST_SYNC_TO_TARGET, MessageLevel.msgInfo, "Tests skipped");
            }

            // AltAz Sync - Optional
            if (g_InterfaceVersion > 1)
            {
                if (g_TelescopeTests[TELTEST_SYNC_TO_ALTAZ] == CheckState.Checked)
                {
                    TelescopeSyncTest(SlewSyncType.SyncToAltAz, "SyncToAltAz", canSyncAltAz, "CanSyncAltAz");
                    if (TestStop())
                        return;
                    if (canSyncAltAz) // Test syncing to bad co-ordinates
                    {
                        TelescopeBadCoordinateTest("SyncToAltAz (Bad L)", SlewSyncType.SyncToAltAz, BAD_ALTITUDE_LOW, BAD_AZIMUTH_LOW);
                        if (TestStop())
                            return;
                        TelescopeBadCoordinateTest("SyncToAltAz (Bad H)", SlewSyncType.SyncToAltAz, BAD_ALTITUDE_HIGH, BAD_AZIMUTH_HIGH);
                        if (TestStop())
                            return;
                    }
                }
                else
                {
                    LogMsg(TELTEST_SYNC_TO_ALTAZ, MessageLevel.msgInfo, "Tests skipped");
                }
            }
            else
            {
                LogMsg("SyncToAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (g_Settings.TestSideOfPierRead)
            {
                LogMsg("", MessageLevel.msgAlways, "");
                LogMsg("SideOfPier Model Tests", MessageLevel.msgAlways, "");
                LogMsg("SideOfPier Model Tests", MessageLevel.msgDebug, "Starting tests");
                if (g_InterfaceVersion > 1)
                {
                    // 3.0.0.14 - Skip these tests if unable to read SideOfPier
                    if (m_CanReadSideOfPier)
                    {

                        // Further side of pier tests
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("SideOfPier Model Tests", MessageLevel.msgComment, "About to get AlignmentMode property");
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(telescopeDevice.AlignmentMode, AlignmentModes.algGermanPolar, false)))
                        {
                            LogMsg("SideOfPier Model Tests", MessageLevel.msgDebug, "Calling SideOfPierTests()");
                            switch (m_SiteLatitude)
                            {
                                case var case6 when -SIDE_OF_PIER_INVALID_LATITUDE <= case6 && case6 <= SIDE_OF_PIER_INVALID_LATITUDE: // Refuse to handle this value because the Conform targeting logic or the mount's SideofPier flip logic may fail when the poles are this close to the horizon
                                    {
                                        LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Tests skipped because the site latitude is reported as " + g_Util.DegreesToDMS(m_SiteLatitude, ":", ":", "", 3));
                                        LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "This places the celestial poles close to the horizon and the mount's flip logic may override Conform's expected behaviour.");
                                        LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Please set the site latitude to a value within the ranges " + SIDE_OF_PIER_INVALID_LATITUDE.ToString("+0.0;-0.0") + " to +90.0 or " + (-SIDE_OF_PIER_INVALID_LATITUDE).ToString("+0.0;-0.0") + " to -90.0 to obtain a reliable result.");
                                        break;
                                    }

                                case var case7 when -90.0d <= case7 && case7 <= 90.0d: // Normal case, just run the tests barbecue latitude is outside the invalid range but within -90.0 to +90.0
                                    {
                                        // SideOfPier write property test - Optional
                                        if (g_Settings.TestSideOfPierWrite)
                                        {
                                            LogMsg("SideOfPier Model Tests", MessageLevel.msgDebug, "Testing SideOfPier write...");
                                            TelescopeOptionalMethodsTest(OptionalMethodType.SideOfPierWrite, "SideOfPier Write", canSetPierside);
                                            if (TestStop())
                                                return;
                                        }

                                        SideOfPierTests(); // Only run these for German mounts
                                        break; // Values outside the range -90.0 to +90.0 are invalid
                                    }

                                default:
                                    {
                                        LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Test skipped because the site latitude Is outside the range -90.0 to +90.0");
                                        break;
                                    }
                            }
                        }
                        else
                        {
                            LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Test skipped because this Is Not a German equatorial mount");
                        }
                    }
                    else
                    {
                        LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Tests skipped because this driver does Not support SideOfPier Read");
                    }
                }
                else
                {
                    LogMsg("SideOfPier Model Tests", MessageLevel.msgInfo, "Skipping test as this method Is Not supported in interface V" + g_InterfaceVersion);
                }
            }

            g_Status.Clear();  // Clear status messages
        }

        public override void CheckPerformance()
        {
            Status(StatusType.staTest, "Performance"); // Clear status messages
            TelescopePerformanceTest(PerformanceType.tstPerfAltitude, "Altitude");
            if (TestStop())
                return;
            if (g_InterfaceVersion > 1)
            {
                TelescopePerformanceTest(PerformanceType.tstPerfAtHome, "AtHome");
                if (TestStop())
                    return;
            }
            else
            {
                LogMsg("Performance: AtHome", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            if (g_InterfaceVersion > 1)
            {
                TelescopePerformanceTest(PerformanceType.tstPerfAtPark, "AtPark");
                if (TestStop())
                    return;
            }
            else
            {
                LogMsg("Performance: AtPark", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            TelescopePerformanceTest(PerformanceType.tstPerfAzimuth, "Azimuth");
            if (TestStop())
                return;
            TelescopePerformanceTest(PerformanceType.tstPerfDeclination, "Declination");
            if (TestStop())
                return;
            if (g_InterfaceVersion > 1)
            {
                if (canPulseGuide)
                {
                    TelescopePerformanceTest(PerformanceType.tstPerfIsPulseGuiding, "IsPulseGuiding");
                    if (TestStop())
                        return;
                }
                else
                {
                    LogMsg("Performance: IsPulseGuiding", MessageLevel.msgInfo, "Test omitted since IsPulseGuiding is not implemented");
                }
            }
            else
            {
                LogMsg("Performance: IsPulseGuiding", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface v1" + g_InterfaceVersion);
            }

            TelescopePerformanceTest(PerformanceType.tstPerfRightAscension, "RightAscension");
            if (TestStop())
                return;
            if (g_InterfaceVersion > 1)
            {
                if (m_AlignmentMode == AlignmentModes.algGermanPolar)
                {
                    if (m_CanReadSideOfPier)
                    {
                        TelescopePerformanceTest(PerformanceType.tstPerfSideOfPier, "SideOfPier");
                        if (TestStop())
                            return;
                    }
                    else
                    {
                        LogMsg("Performance: SideOfPier", MessageLevel.msgInfo, "Test omitted since SideOfPier is not implemented");
                    }
                }
                else
                {
                    LogMsg("Performance: SideOfPier", MessageLevel.msgInfo, "Test omitted since alignment mode is not German Polar");
                }
            }
            else
            {
                LogMsg("Performance: SideOfPier", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface v1" + g_InterfaceVersion);
            }

            if (canReadSiderealTime)
            {
                TelescopePerformanceTest(PerformanceType.tstPerfSiderealTime, "SiderealTime");
                if (TestStop())
                    return;
            }
            else
            {
                LogMsgInfo("Performance: SiderealTime", "Skipping test because the SiderealTime property throws an exception.");
            }

            TelescopePerformanceTest(PerformanceType.tstPerfSlewing, "Slewing");
            if (TestStop())
                return;
            TelescopePerformanceTest(PerformanceType.tstPerfUTCDate, "UTCDate");
            if (TestStop())
                return;
            g_Status.Clear();
        }

        public override void PostRunCheck()
        {
            // Make things safe
            // LogMsg("", MessageLevel.msgAlways, "") 'Blank line
            try
            {
                if (Conversions.ToBoolean(telescopeDevice.CanSetTracking))
                {
                    telescopeDevice.Tracking = false;
                    LogMsg("Mount Safety", MessageLevel.msgOK, "Tracking stopped to protect your mount.");
                }
                else
                {
                    LogMsg("Mount Safety", MessageLevel.msgInfo, "Tracking can't be turned off for this mount, please switch off manually.");
                }
            }
            catch (Exception ex)
            {
                LogMsg("Mount Safety", MessageLevel.msgError, "Exception when disabling tracking to protect mount: " + ex.ToString());
            }
        }

        private void TelescopeSyncTest(SlewSyncType testType, string testName, bool driverSupportsMethod, string canDoItName)
        {
            bool showOutcome = false;
            double difference, syncRA, syncDEC, syncAlt = default, syncAz = default, newAlt, newAz, currentAz = default, currentAlt = default, startRA, startDec, currentRA, currentDec;

            // Basic test to make sure the method is either implemented OK or fails as expected if it is not supported in this driver.
            if (g_Settings.DisplayMethodCalls)
                LogMsg(testName, MessageLevel.msgComment, "About to get RightAscension property");
            syncRA = Conversions.ToDouble(telescopeDevice.RightAscension);
            if (g_Settings.DisplayMethodCalls)
                LogMsg(testName, MessageLevel.msgComment, "About to get Declination property");
            syncDEC = Conversions.ToDouble(telescopeDevice.Declination);
            if (!driverSupportsMethod) // Call should fail
            {
                try
                {
                    switch (testType)
                    {
                        case SlewSyncType.SyncToCoordinates: // SyncToCoordinates
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property");
                                if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to true");
                                    telescopeDevice.Tracking = true;
                                }

                                LogMsg(testName, MessageLevel.msgDebug, "SyncToCoordinates: " + FormatRA(syncRA) + " " + FormatDec(syncDEC));
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " + FormatRA(syncRA) + ", Declination: " + FormatDec(syncDEC));
                                telescopeDevice.SyncToCoordinates(syncRA, syncDEC);
                                LogMsg(testName, MessageLevel.msgError, "CanSyncToCoordinates is False but call to SyncToCoordinates did not throw an exception.");
                                break;
                            }

                        case SlewSyncType.SyncToTarget: // SyncToTarget
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property");
                                if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to true");
                                    telescopeDevice.Tracking = true;
                                }

                                try
                                {
                                    LogMsg(testName, MessageLevel.msgDebug, "Setting TargetRightAscension: " + FormatRA(syncRA));
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(syncRA));
                                    telescopeDevice.TargetRightAscension = syncRA;
                                    LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetRightAscension");
                                }
                                catch (Exception)
                                {
                                    // Ignore errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }

                                try
                                {
                                    LogMsg(testName, MessageLevel.msgDebug, "Setting TargetDeclination: " + FormatDec(syncDEC));
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(syncDEC));
                                    telescopeDevice.TargetDeclination = syncDEC;
                                    LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetDeclination");
                                }
                                catch (Exception)
                                {
                                    // Ignore other errors at this point as we aren't trying to test Telescope.TargetRightAscension
                                }

                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget(); // Sync to target coordinates
                                LogMsg(testName, MessageLevel.msgError, "CanSyncToTarget is False but call to SyncToTarget did not throw an exception.");
                                break;
                            }

                        case SlewSyncType.SyncToAltAz:
                            {
                                if (canReadAltitide)
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to get Altitude property");
                                    syncAlt = Conversions.ToDouble(telescopeDevice.Altitude);
                                }

                                if (canReadAzimuth)
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to get Azimuth property");
                                    syncAz = Conversions.ToDouble(telescopeDevice.Azimuth);
                                }

                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property");
                                if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, telescopeDevice.Tracking)))
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to false");
                                    telescopeDevice.Tracking = false;
                                }

                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " + FormatDec(syncAlt) + ", Azimuth: " + FormatDec(syncAz));
                                telescopeDevice.SyncToAltAz(syncAz, syncAlt); // Sync to new Alt Az
                                LogMsg(testName, MessageLevel.msgError, "CanSyncToAltAz is False but call to SyncToAltAz did not throw an exception.");
                                break;
                            }

                        default:
                            {
                                LogMsg(testName, MessageLevel.msgError, "Conform:SyncTest: Unknown test type " + testType.ToString());
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
                                LogMsg(testName, MessageLevel.msgDebug, string.Format("RA for sync tests: {0}", FormatRA(startRA)));

                                // Calculate the Sync test DEC position
                                if (m_SiteLatitude > 0.0d) // We are in the northern hemisphere
                                {
                                    startDec = 90.0d - (180.0d - m_SiteLatitude) * 0.5d; // Calculate for northern hemisphere
                                }
                                else // We are in the southern hemisphere
                                {
                                    startDec = -90.0d + (180.0d + m_SiteLatitude) * 0.5d;
                                } // Calculate for southern hemisphere

                                LogMsg(testName, MessageLevel.msgDebug, string.Format("Declination for sync tests: {0}", FormatDec(startDec)));
                                SlewScope(startRA, startDec, string.Format("Start position - RA: {0}, Dec: {1}", FormatRA(startRA), FormatDec(startDec)));
                                if (TestStop())
                                    return;

                                // Now test that we have actually arrived
                                CheckScopePosition(testName, "Slewed to start position", startRA, startDec);

                                // Calculate the sync test RA coordinate as a variation from the current RA coordinate
                                syncRA = startRA - SYNC_SIMULATED_ERROR / (15.0d * 60.0d); // Convert sync error in arc minutes to RA hours
                                if (syncRA < 0.0d)
                                    syncRA = syncRA + 24.0d; // Ensure legal RA

                                // Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                                syncDEC = startDec - SYNC_SIMULATED_ERROR / 60.0d; // Convert sync error in arc minutes to degrees

                                // Sync the scope to the offset RA and DEC coordinates
                                SyncScope(testName, canDoItName, testType, syncRA, syncDEC);

                                // Check that the scope's synchronised position is as expected
                                CheckScopePosition(testName, "Synced to sync position", syncRA, syncDEC);

                                // Check that the TargetRA and TargetDec were 
                                if (testType == SlewSyncType.SyncToCoordinates)
                                {
                                    // Check that target coordinates are present and set correctly per the ASCOM Telescope specification
                                    try
                                    {
                                        currentRA = Conversions.ToDouble(telescopeDevice.TargetRightAscension);
                                        LogMsg(testName, MessageLevel.msgDebug, string.Format("Current TargetRightAscension: {0}, Set TargetRightAscension: {1}", currentRA, syncRA));
                                        double raDifference;
                                        raDifference = RaDifferenceInSeconds(syncRA, currentRA);
                                        switch (raDifference)
                                        {
                                            case var @case when @case <= SLEW_SYNC_OK_TOLERANCE:  // Within specified tolerance
                                                {
                                                    LogMsg(testName, MessageLevel.msgOK, string.Format("The TargetRightAscension property {0} matches the expected RA OK. ", FormatRA(syncRA))); // Outside specified tolerance
                                                    break;
                                                }

                                            default:
                                                {
                                                    LogMsg(testName, MessageLevel.msgError, string.Format("The TargetRightAscension property {0} does not match the expected RA {1}", FormatRA(currentRA), FormatRA(syncRA)));
                                                    break;
                                                }
                                        }
                                    }
                                    catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                    {
                                        LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.");
                                    }
                                    catch (ASCOM.InvalidOperationException)
                                    {
                                        LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                    }
                                    catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
                                    {
                                        LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
                                    }

                                    try
                                    {
                                        currentDec = Conversions.ToDouble(telescopeDevice.TargetDeclination);
                                        LogMsg(testName, MessageLevel.msgDebug, string.Format("Current TargetDeclination: {0}, Set TargetDeclination: {1}", currentDec, syncDEC));
                                        double decDifference;
                                        decDifference = Math.Round(Math.Abs(currentDec - syncDEC) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // Dec difference is in arc seconds from degrees of Declination
                                        switch (decDifference)
                                        {
                                            case var case1 when case1 <= SLEW_SYNC_OK_TOLERANCE: // Within specified tolerance
                                                {
                                                    LogMsg(testName, MessageLevel.msgOK, string.Format("The TargetDeclination property {0} matches the expected Declination OK. ", FormatDec(syncDEC))); // Outside specified tolerance
                                                    break;
                                                }

                                            default:
                                                {
                                                    LogMsg(testName, MessageLevel.msgError, string.Format("The TargetDeclination property {0} does not match the expected Declination {1}", FormatDec(currentDec), FormatDec(syncDEC)));
                                                    break;
                                                }
                                        }
                                    }
                                    catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                    {
                                        LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.");
                                    }
                                    catch (ASCOM.InvalidOperationException)
                                    {
                                        LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                    }
                                    catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
                                    {
                                        LogMsg(testName, MessageLevel.msgError, "The driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
                                    }
                                }

                                // Now slew to the scope's original position
                                SlewScope(startRA, startDec, string.Format("Slewing back to start position - RA: {0}, Dec: {1}", FormatRA(startRA), FormatDec(startDec)));

                                // Check that the scope's position is the original position
                                CheckScopePosition(testName, "Slewed back to start position", startRA, startDec);

                                // Now "undo" the sync by reversing syncing in the opposition sense than originally made

                                // Calculate the sync test RA coordinate as a variation from the current RA coordinate
                                syncRA = startRA + SYNC_SIMULATED_ERROR / (15.0d * 60.0d); // Convert sync error in arc minutes to RA hours
                                if (syncRA >= 24.0d)
                                    syncRA = syncRA - 24.0d; // Ensure legal RA

                                // Calculate the sync test DEC coordinate as a variation from the current DEC coordinate
                                syncDEC = startDec + SYNC_SIMULATED_ERROR / 60.0d; // Convert sync error in arc minutes to degrees

                                // Sync back to the original coordinates
                                SyncScope(testName, canDoItName, testType, syncRA, syncDEC);

                                // Check that the scope's synchronised position is as expected
                                CheckScopePosition(testName, "Synced to reversed sync position", syncRA, syncDEC);

                                // Now slew to the scope's original position
                                SlewScope(startRA, startDec, string.Format("Slewing back to start position - RA: {0}, Dec: {1}", FormatRA(startRA), FormatDec(startDec)));

                                // Check that the scope's position is the original position
                                CheckScopePosition(testName, "Slewed back to start position", startRA, startDec);
                                break;
                            }

                        case SlewSyncType.SyncToAltAz:
                            {
                                if (canReadAltitide)
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to get Altitude property");
                                    currentAlt = Conversions.ToDouble(telescopeDevice.Altitude);
                                }

                                if (canReadAzimuth)
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to get Azimuth property");
                                    currentAz = Conversions.ToDouble(telescopeDevice.Azimuth);
                                }

                                syncAlt = currentAlt - 1.0d;
                                syncAz = currentAz + 1.0d;
                                if (syncAlt < 0.0d)
                                    syncAlt = 1.0d; // Ensure legal Alt
                                if (syncAz > 359.0d)
                                    syncAz = 358.0d; // Ensure legal Az
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property");
                                if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, telescopeDevice.Tracking)))
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to false");
                                    telescopeDevice.Tracking = false;
                                }

                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(testName, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " + FormatDec(syncAlt) + ", Azimuth: " + FormatDec(syncAz));
                                telescopeDevice.SyncToAltAz(syncAz, syncAlt); // Sync to new Alt Az
                                if (canReadAltitide & canReadAzimuth) // Can check effects of a sync
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to get Altitude property");
                                    newAlt = Conversions.ToDouble(telescopeDevice.Altitude);
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(testName, MessageLevel.msgComment, "About to get Azimuth property");
                                    newAz = Conversions.ToDouble(telescopeDevice.Azimuth);

                                    // Compare old and new values
                                    difference = Math.Abs(syncAlt - newAlt);
                                    switch (difference)
                                    {
                                        case var case2 when case2 <= 1.0d / (60 * 60): // Within 1 seconds
                                            {
                                                LogMsg(testName, MessageLevel.msgOK, "Synced Altitude OK");
                                                break;
                                            }

                                        case var case3 when case3 <= 2.0d / (60 * 60): // Within 2 seconds
                                            {
                                                LogMsg(testName, MessageLevel.msgOK, "Synced within 2 seconds of Altitude");
                                                showOutcome = true;
                                                break;
                                            }

                                        default:
                                            {
                                                LogMsg(testName, MessageLevel.msgInfo, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Synced to within ", FormatAltitude(difference)), " DD:MM:SS of expected Altitude: "), FormatAltitude(syncAlt))));
                                                showOutcome = true;
                                                break;
                                            }
                                    }

                                    difference = Math.Abs(syncAz - newAz);
                                    switch (difference)
                                    {
                                        case var case4 when case4 <= 1.0d / (60 * 60): // Within 1 seconds
                                            {
                                                LogMsg(testName, MessageLevel.msgOK, "Synced Azimuth OK");
                                                break;
                                            }

                                        case var case5 when case5 <= 2.0d / (60 * 60): // Within 2 seconds
                                            {
                                                LogMsg(testName, MessageLevel.msgOK, "Synced within 2 seconds of Azimuth");
                                                showOutcome = true;
                                                break;
                                            }

                                        default:
                                            {
                                                LogMsg(testName, MessageLevel.msgInfo, "Synced to within " + FormatAzimuth(difference) + " DD:MM:SS of expected Azimuth: " + FormatAzimuth(syncAz));
                                                showOutcome = true;
                                                break;
                                            }
                                    }

                                    if (showOutcome)
                                    {
                                        LogMsg(testName, MessageLevel.msgComment, "           Altitude    Azimuth");
                                        LogMsg(testName, MessageLevel.msgComment, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Original:  ", FormatAltitude(currentAlt)), "   "), FormatAzimuth(currentAz))));
                                        LogMsg(testName, MessageLevel.msgComment, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Sync to:   ", FormatAltitude(syncAlt)), "   "), FormatAzimuth(syncAz))));
                                        LogMsg(testName, MessageLevel.msgComment, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("New:       ", FormatAltitude(newAlt)), "   "), FormatAzimuth(newAz))));
                                    }
                                }
                                else // Can't test effects of a sync
                                {
                                    LogMsg(testName, MessageLevel.msgInfo, "Can't test SyncToAltAz because Altitude or Azimuth values are not implemented");
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
            double l_Difference, l_ActualAltitude, l_ActualAzimuth, actualRA, actualDec;
            Status(StatusType.staTest, p_Name);
            if (g_Settings.DisplayMethodCalls)
                LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
            if (canSetTracking)
                telescopeDevice.Tracking = true; // Enable tracking for these tests
            try
            {
                switch (p_Test)
                {
                    case SlewSyncType.SlewToCoordinates:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -1.0d);
                            m_TargetDeclination = 1.0d;
                            Status(StatusType.staAction, "Slewing");
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                            telescopeDevice.SlewToCoordinates(m_TargetRightAscension, m_TargetDeclination);
                            break;
                        }

                    case SlewSyncType.SlewToCoordinatesAsync:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0d);
                            m_TargetDeclination = 2.0d;
                            Status(StatusType.staAction, "Slewing");
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                            telescopeDevice.SlewToCoordinatesAsync(m_TargetRightAscension, m_TargetDeclination);
                            WaitForSlew(p_Name);
                            break;
                        }

                    case SlewSyncType.SlewToTarget:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0d);
                            m_TargetDeclination = 3.0d;
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(m_TargetRightAscension));
                                telescopeDevice.TargetRightAscension = m_TargetRightAscension;
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
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(m_TargetDeclination));
                                telescopeDevice.TargetDeclination = m_TargetDeclination;
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

                            Status(StatusType.staAction, "Slewing");
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTarget method");
                            telescopeDevice.SlewToTarget();
                            break;
                        }

                    case SlewSyncType.SlewToTargetAsync: // SlewToTargetAsync
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                                telescopeDevice.Tracking = true;
                            }

                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -4.0d);
                            m_TargetDeclination = 4.0d;
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(m_TargetRightAscension));
                                telescopeDevice.TargetRightAscension = m_TargetRightAscension;
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
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(m_TargetDeclination));
                                telescopeDevice.TargetDeclination = m_TargetDeclination;
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

                            Status(StatusType.staAction, "Slewing");
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTargetAsync method");
                            telescopeDevice.SlewToTargetAsync();
                            WaitForSlew(p_Name);
                            break;
                        }

                    case SlewSyncType.SlewToAltAz:
                        {
                            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Tracking 1: ", telescopeDevice.Tracking)));
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, telescopeDevice.Tracking)))
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set property Tracking to false");
                                telescopeDevice.Tracking = false;
                                LogMsg(p_Name, MessageLevel.msgDebug, "Tracking turned off");
                            }

                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Tracking 2: ", telescopeDevice.Tracking)));
                            m_TargetAltitude = 50.0d;
                            m_TargetAzimuth = 150.0d;
                            Status(StatusType.staAction, "Slewing to Alt/Az: " + FormatDec(m_TargetAltitude) + " " + FormatDec(m_TargetAzimuth));
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAz method, Altitude: " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                            telescopeDevice.SlewToAltAz(m_TargetAzimuth, m_TargetAltitude);
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Tracking 3: ", telescopeDevice.Tracking)));
                            break;
                        }

                    case SlewSyncType.SlewToAltAzAsync:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Tracking 1: ", telescopeDevice.Tracking)));
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, telescopeDevice.Tracking)))
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property false");
                                telescopeDevice.Tracking = false;
                                LogMsg(p_Name, MessageLevel.msgDebug, "Tracking turned off");
                            }

                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Tracking 2: ", telescopeDevice.Tracking)));
                            m_TargetAltitude = 55.0d;
                            m_TargetAzimuth = 155.0d;
                            Status(StatusType.staAction, "Slewing to Alt/Az: " + FormatDec(m_TargetAltitude) + " " + FormatDec(m_TargetAzimuth));
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAzAsync method, Altitude: " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                            telescopeDevice.SlewToAltAzAsync(m_TargetAzimuth, m_TargetAltitude);
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Tracking 3: ", telescopeDevice.Tracking)));
                            WaitForSlew(p_Name);
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Tracking 4: ", telescopeDevice.Tracking)));
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Conform:SlewTest: Unknown test type " + p_Test.ToString());
                            break;
                        }
                }

                if (TestStop())
                    return;
                if (p_CanDoIt) // Should be able to do this so report what happened
                {
                    switch (p_Test)
                    {
                        case SlewSyncType.SlewToCoordinates:
                        case SlewSyncType.SlewToCoordinatesAsync:
                        case SlewSyncType.SlewToTarget:
                        case SlewSyncType.SlewToTargetAsync:
                            {
                                Status(StatusType.staAction, "Slew completed");
                                // Test how close the slew was to the required coordinates
                                CheckScopePosition(p_Name, "Slewed", m_TargetRightAscension, m_TargetDeclination);

                                // Check that the slews and syncs set the target coordinates correctly per the ASCOM Telescope specification
                                try
                                {
                                    actualRA = Conversions.ToDouble(telescopeDevice.TargetRightAscension);
                                    LogMsg(p_Name, MessageLevel.msgDebug, string.Format("Current TargetRightAscension: {0}, Set TargetRightAscension: {1}", actualRA, m_TargetRightAscension));
                                    double raDifference;
                                    raDifference = RaDifferenceInSeconds(actualRA, m_TargetRightAscension);
                                    switch (raDifference)
                                    {
                                        case var @case when @case <= SLEW_SYNC_OK_TOLERANCE:  // Within specified tolerance
                                            {
                                                LogMsg(p_Name, MessageLevel.msgOK, string.Format("The TargetRightAscension property {0} matches the expected RA OK. ", FormatRA(m_TargetRightAscension))); // Outside specified tolerance
                                                break;
                                            }

                                        default:
                                            {
                                                LogMsg(p_Name, MessageLevel.msgError, string.Format("The TargetRightAscension property {0} does not match the expected RA {1}", FormatRA(actualRA), FormatRA(m_TargetRightAscension)));
                                                break;
                                            }
                                    }
                                }
                                catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.");
                                }
                                catch (ASCOM.InvalidOperationException)
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                }
                                catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "The driver did not set the TargetRightAscension property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
                                }

                                try
                                {
                                    actualDec = Conversions.ToDouble(telescopeDevice.TargetDeclination);
                                    LogMsg(p_Name, MessageLevel.msgDebug, string.Format("Current TargetDeclination: {0}, Set TargetDeclination: {1}", actualDec, m_TargetDeclination));
                                    double decDifference;
                                    decDifference = Math.Round(Math.Abs(actualDec - m_TargetDeclination) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // Dec difference is in arc seconds from degrees of Declination
                                    switch (decDifference)
                                    {
                                        case var case1 when case1 <= SLEW_SYNC_OK_TOLERANCE: // Within specified tolerance
                                            {
                                                LogMsg(p_Name, MessageLevel.msgOK, string.Format("The TargetDeclination property {0} matches the expected Declination OK. ", FormatDec(m_TargetDeclination))); // Outside specified tolerance
                                                break;
                                            }

                                        default:
                                            {
                                                LogMsg(p_Name, MessageLevel.msgError, string.Format("The TargetDeclination property {0} does not match the expected Declination {1}", FormatDec(actualDec), FormatDec(m_TargetDeclination)));
                                                break;
                                            }
                                    }
                                }
                                catch (COMException ex) when (ex.ErrorCode == ErrorCodes.ValueNotSet | ex.ErrorCode == g_ExNotSet1 | ex.ErrorCode == g_ExNotSet2)
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet COM exception was thrown instead.");
                                }
                                catch (ASCOM.InvalidOperationException)
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetDeclination property as required by the Telescope specification, An InvalidOperationException was thrown instead.");
                                }
                                catch (DriverException ex) when (ex.Number == ErrorCodes.ValueNotSet | ex.Number == g_ExNotSet1 | ex.Number == g_ExNotSet2)
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "The Driver did not set the TargetDeclination property as required by the Telescope specification, A ValueNotSet DriverException was thrown instead.");
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
                                Status(StatusType.staAction, "Slew completed");
                                // Test how close the slew was to the required coordinates
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Azimuth property");
                                l_ActualAzimuth = Conversions.ToDouble(telescopeDevice.Azimuth);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Altitude property");
                                l_ActualAltitude = Conversions.ToDouble(telescopeDevice.Altitude);
                                l_Difference = Math.Abs(l_ActualAzimuth - m_TargetAzimuth);
                                if (l_Difference > 350.0d)
                                    l_Difference = 360.0d - l_Difference; // Deal with the case where the two elements are on different sides of 360 degrees
                                switch (l_Difference)
                                {
                                    case var case2 when case2 <= 1.0d / 3600.0d: // seconds
                                        {
                                            LogMsg(p_Name, MessageLevel.msgOK, "Slewed to target Azimuth OK: " + FormatAzimuth(m_TargetAzimuth));
                                            break;
                                        }

                                    case var case3 when case3 <= 2.0d / 3600.0d: // 2 seconds
                                        {
                                            LogMsg(p_Name, MessageLevel.msgOK, "Slewed to within 2 seconds of Azimuth target: " + FormatAzimuth(m_TargetAzimuth) + " Actual Azimuth " + FormatAzimuth(l_ActualAzimuth));
                                            break;
                                        }

                                    default:
                                        {
                                            LogMsg(p_Name, MessageLevel.msgInfo, "Slewed to within " + FormatAzimuth(l_Difference) + " DD:MM:SS of expected Azimuth: " + FormatAzimuth(m_TargetAzimuth));
                                            break;
                                        }
                                }

                                l_Difference = Math.Abs(l_ActualAltitude - m_TargetAltitude);
                                switch (l_Difference)
                                {
                                    case var case4 when case4 <= 1.0d / 3600.0d: // <1 seconds
                                        {
                                            LogMsg(p_Name, MessageLevel.msgOK, Conversions.ToString(Operators.ConcatenateObject("Slewed to target Altitude OK: ", FormatAltitude(m_TargetAltitude))));
                                            break;
                                        }

                                    case var case5 when case5 <= 2.0d / 3600.0d: // 2 seconds
                                        {
                                            LogMsg(p_Name, MessageLevel.msgOK, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Slewed to within 2 seconds of Altitude target: ", FormatAltitude(m_TargetAltitude)), " Actual Altitude "), FormatAltitude(l_ActualAltitude))));
                                            break;
                                        }

                                    default:
                                        {
                                            LogMsg(p_Name, MessageLevel.msgInfo, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Slewed to within ", FormatAltitude(l_Difference)), " DD:MM:SS of expected Altitude: "), FormatAltitude(m_TargetAltitude))));
                                            break;
                                        }
                                } // Do nothing

                                break;
                            }

                        default:
                            {
                                break;
                            }
                    }
                }
                else // Not supposed to be able to do this but no error generated so report an error
                {
                    LogMsg(p_Name, MessageLevel.msgIssue, p_CanDoItName + " is false but no exception was generated on use");
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

            g_Status.Clear();  // Clear status messages
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            Status(StatusType.staAction, "Slew underway");
                            m_TargetRightAscension = BadCoordinate1;
                            m_TargetDeclination = 0.0d;
                            if (p_Test == SlewSyncType.SlewToCoordinates)
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                                telescopeDevice.SlewToCoordinates(m_TargetRightAscension, m_TargetDeclination);
                            }
                            else
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                                telescopeDevice.SlewToCoordinatesAsync(m_TargetRightAscension, m_TargetDeclination);
                            }

                            Status(StatusType.staAction, "Attempting to abort slew");
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                        }

                        try
                        {
                            Status(StatusType.staAction, "Slew underway");
                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0d);
                            m_TargetDeclination = BadCoordinate2;
                            if (p_Test == SlewSyncType.SlewToCoordinates)
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                                telescopeDevice.SlewToCoordinates(m_TargetRightAscension, m_TargetDeclination);
                            }
                            else
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                                telescopeDevice.SlewToCoordinatesAsync(m_TargetRightAscension, m_TargetDeclination);
                            }

                            Status(StatusType.staAction, "Attempting to abort slew");
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SyncToCoordinates:
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            Status(StatusType.staAction, "Sync underway");
                            m_TargetRightAscension = BadCoordinate1;
                            m_TargetDeclination = 0.0d;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                            telescopeDevice.SyncToCoordinates(m_TargetRightAscension, m_TargetDeclination);
                            LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                        }

                        try
                        {
                            Status(StatusType.staAction, "Sync underway");
                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0d);
                            m_TargetDeclination = BadCoordinate2;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(m_TargetDeclination));
                            telescopeDevice.SyncToCoordinates(m_TargetRightAscension, m_TargetDeclination);
                            LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SlewToTarget:
                case SlewSyncType.SlewToTargetAsync:
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            Status(StatusType.staAction, "Slew underway");
                            m_TargetRightAscension = BadCoordinate1;
                            m_TargetDeclination = 0.0d;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(m_TargetRightAscension));
                            telescopeDevice.TargetRightAscension = m_TargetRightAscension;
                            // Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(m_TargetDeclination));
                                telescopeDevice.TargetDeclination = m_TargetDeclination;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (p_Test == SlewSyncType.SlewToTarget)
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTarget method");
                                    telescopeDevice.SlewToTarget();
                                }
                                else
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTargetAsync method");
                                    telescopeDevice.SlewToTargetAsync();
                                }

                                Status(StatusType.staAction, "Attempting to abort slew");
                                try
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                    telescopeDevice.AbortSlew();
                                }
                                catch
                                {
                                } // Attempt to stop any motion that has actually started

                                LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                Status(StatusType.staAction, "Slew rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate", "Telescope.TargetRA correctly rejected bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                        }

                        try
                        {
                            Status(StatusType.staAction, "Slew underway");
                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -2.0d);
                            m_TargetDeclination = BadCoordinate2;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(m_TargetDeclination));
                            telescopeDevice.TargetDeclination = m_TargetDeclination;
                            // Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(m_TargetRightAscension));
                                telescopeDevice.TargetRightAscension = m_TargetRightAscension;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (p_Test == SlewSyncType.SlewToTarget)
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTarget method");
                                    telescopeDevice.SlewToTarget();
                                }
                                else
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToTargetAsync method");
                                    telescopeDevice.SlewToTargetAsync();
                                }

                                Status(StatusType.staAction, "Attempting to abort slew");
                                try
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                    telescopeDevice.AbortSlew();
                                }
                                catch
                                {
                                } // Attempt to stop any motion that has actually started

                                LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                Status(StatusType.staAction, "Slew rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate", "Telescope.TargetDeclination correctly rejected bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SyncToTarget:
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            Status(StatusType.staAction, "Sync underway");
                            m_TargetRightAscension = BadCoordinate1;
                            m_TargetDeclination = 0.0d;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(m_TargetRightAscension));
                            telescopeDevice.TargetRightAscension = m_TargetRightAscension;
                            // Successfully set bad RA coordinate so now set the good Dec coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(m_TargetDeclination));
                                telescopeDevice.TargetDeclination = m_TargetDeclination;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                Status(StatusType.staAction, "Sync rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad RA coordinate", "Correctly rejected bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad RA coordinate", "Telescope.TargetRA correctly rejected bad RA coordinate: " + FormatRA(m_TargetRightAscension));
                        }

                        try
                        {
                            Status(StatusType.staAction, "Sync underway");
                            m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -3.0d);
                            m_TargetDeclination = BadCoordinate2;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(m_TargetDeclination));
                            telescopeDevice.TargetDeclination = m_TargetDeclination;
                            // Successfully set bad Dec coordinate so now set the good RA coordinate and see whether the move fails when the slew is attempted
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(m_TargetRightAscension));
                                telescopeDevice.TargetRightAscension = m_TargetRightAscension;
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                            }
                            catch (Exception ex) // Attempt to set bad coordinate failed, so check whether an invalid value exception was thrown or something else
                            {
                                Status(StatusType.staAction, "Sync rejected");
                                HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Dec coordinate", "Correctly rejected bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Property, Required.MustBeImplemented, ex, "Exception setting bad Dec coordinate", "Telescope.TargetDeclination correctly rejected bad Dec coordinate: " + FormatDec(m_TargetDeclination));
                        }

                        break;
                    }

                case SlewSyncType.SlewToAltAz:
                case SlewSyncType.SlewToAltAzAsync:
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to false");
                            telescopeDevice.Tracking = false;
                        }

                        try
                        {
                            Status(StatusType.staAction, "Slew underway");
                            m_TargetAltitude = BadCoordinate1;
                            m_TargetAzimuth = 45.0d;
                            if (p_Test == SlewSyncType.SlewToAltAz)
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAz method, Altitude: " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                                telescopeDevice.SlewToAltAz(m_TargetAzimuth, m_TargetAltitude);
                            }
                            else
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About To call SlewToAltAzAsync method, Altitude:  " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                                telescopeDevice.SlewToAltAzAsync(m_TargetAzimuth, m_TargetAltitude);
                            }

                            Status(StatusType.staAction, "Attempting to abort slew");
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogMsg(p_Name, MessageLevel.msgError, Conversions.ToString(Operators.ConcatenateObject("Failed to reject bad Altitude coordinate: ", FormatAltitude(m_TargetAltitude))));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Altitude coordinate", Conversions.ToString(Operators.ConcatenateObject("Correctly rejected bad Altitude coordinate: ", FormatAltitude(m_TargetAltitude))));
                        }

                        try
                        {
                            Status(StatusType.staAction, "Slew underway");
                            m_TargetAltitude = 45.0d;
                            m_TargetAzimuth = BadCoordinate2;
                            if (p_Test == SlewSyncType.SlewToAltAz)
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAz method, Altitude: " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                                telescopeDevice.SlewToAltAz(m_TargetAzimuth, m_TargetAltitude);
                            }
                            else
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call SlewToAltAzAsync method, Altitude: " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                                telescopeDevice.SlewToAltAzAsync(m_TargetAzimuth, m_TargetAltitude);
                            }

                            Status(StatusType.staAction, "Attempting to abort slew");
                            try
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                            }
                            catch
                            {
                            } // Attempt to stop any motion that has actually started

                            LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Azimuth coordinate: " + FormatAzimuth(m_TargetAzimuth));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Slew rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "slewing to bad Azimuth coordinate", "Correctly rejected bad Azimuth coordinate: " + FormatAzimuth(m_TargetAzimuth));
                        }

                        break;
                    }

                case SlewSyncType.SyncToAltAz:
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to false");
                            telescopeDevice.Tracking = false;
                        }

                        try
                        {
                            Status(StatusType.staAction, "Sync underway");
                            m_TargetAltitude = BadCoordinate1;
                            m_TargetAzimuth = 45.0d;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                            telescopeDevice.SyncToAltAz(m_TargetAzimuth, m_TargetAltitude);
                            LogMsg(p_Name, MessageLevel.msgError, Conversions.ToString(Operators.ConcatenateObject("Failed to reject bad Altitude coordinate: ", FormatAltitude(m_TargetAltitude))));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Altitude coordinate", Conversions.ToString(Operators.ConcatenateObject("Correctly rejected bad Altitude coordinate: ", FormatAltitude(m_TargetAltitude))));
                        }

                        try
                        {
                            Status(StatusType.staAction, "Sync underway");
                            m_TargetAltitude = 45.0d;
                            m_TargetAzimuth = BadCoordinate2;
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call SyncToAltAz method, Altitude: " + FormatDec(m_TargetAltitude) + ", Azimuth: " + FormatDec(m_TargetAzimuth));
                            telescopeDevice.SyncToAltAz(m_TargetAzimuth, m_TargetAltitude);
                            LogMsg(p_Name, MessageLevel.msgError, "Failed to reject bad Azimuth coordinate: " + FormatAzimuth(m_TargetAzimuth));
                        }
                        catch (Exception ex)
                        {
                            Status(StatusType.staAction, "Sync rejected");
                            HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "syncing to bad Azimuth coordinate", "Correctly rejected bad Azimuth coordinate: " + FormatAzimuth(m_TargetAzimuth));
                        }

                        break;
                    }

                default:
                    {
                        LogMsg(p_Name, MessageLevel.msgError, "Conform:SlewTest: Unknown test type " + p_Test.ToString());
                        break;
                    }
            }

            if (TestStop())
                return;
        }

        private void TelescopePerformanceTest(PerformanceType p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate;
            Status(StatusType.staAction, p_Name);
            try
            {
                l_StartTime = DateAndTime.Now;
                l_Count = 0.0d;
                l_LastElapsedTime = 0.0d;
                do
                {
                    l_Count += 1.0d;
                    switch (p_Type)
                    {
                        case PerformanceType.tstPerfAltitude:
                            {
                                m_Altitude = Conversions.ToDouble(telescopeDevice.Altitude);
                                break;
                            }

                        case var @case when @case == PerformanceType.tstPerfAtHome:
                            {
                                m_AtHome = telescopeDevice.AtHome;
                                break;
                            }

                        case PerformanceType.tstPerfAtPark:
                            {
                                m_AtPark = telescopeDevice.AtPark;
                                break;
                            }

                        case PerformanceType.tstPerfAzimuth:
                            {
                                m_Azimuth = Conversions.ToDouble(telescopeDevice.Azimuth);
                                break;
                            }

                        case PerformanceType.tstPerfDeclination:
                            {
                                m_Declination = Conversions.ToDouble(telescopeDevice.Declination);
                                break;
                            }

                        case PerformanceType.tstPerfIsPulseGuiding:
                            {
                                m_IsPulseGuiding = telescopeDevice.IsPulseGuiding;
                                break;
                            }

                        case PerformanceType.tstPerfRightAscension:
                            {
                                m_RightAscension = Conversions.ToDouble(telescopeDevice.RightAscension);
                                break;
                            }

                        case PerformanceType.tstPerfSideOfPier:
                            {
                                m_SideOfPier = (PierSide)telescopeDevice.SideOfPier;
                                break;
                            }

                        case PerformanceType.tstPerfSiderealTime:
                            {
                                m_SiderealTimeScope = Conversions.ToDouble(telescopeDevice.SiderealTime);
                                break;
                            }

                        case PerformanceType.tstPerfSlewing:
                            {
                                m_Slewing = telescopeDevice.Slewing;
                                break;
                            }

                        case PerformanceType.tstPerfUTCDate:
                            {
                                m_UTCDate = Conversions.ToDate(telescopeDevice.UTCDate);
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "Conform:PerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateAndTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0d)
                    {
                        Status(StatusType.staStatus, l_Count + " transactions in " + Strings.Format(l_ElapsedTime, "0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        Application.DoEvents();
                        if (TestStop())
                            return;
                    }
                }
                while (l_ElapsedTime <= PERF_LOOP_TIME);
                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case var case1 when case1 > 10.0d:
                        {
                            LogMsg("Performance: " + p_Name, MessageLevel.msgInfo, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    case var case2 when 2.0d <= case2 && case2 <= 10.0d:
                        {
                            LogMsg("Performance: " + p_Name, MessageLevel.msgInfo, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    case var case3 when 1.0d <= case3 && case3 <= 2.0d:
                        {
                            LogMsg("Performance: " + p_Name, MessageLevel.msgInfo, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogMsg("Performance: " + p_Name, MessageLevel.msgInfo, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogMsg("Performance: " + p_Name, MessageLevel.msgError, EX_NET + ex.Message);
            }
        }

        private void TelescopeParkedExceptionTest(ParkedExceptionType p_Type, string p_Name)
        {
            double l_TargetRA;
            g_Status.Action = p_Name;
            g_Status.Test = p_Type.ToString();
            if (g_Settings.DisplayMethodCalls)
                LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to get AtPark property");
            if (Conversions.ToBoolean(telescopeDevice.AtPark)) // We are still parked so test AbortSlew
            {
                try
                {
                    switch (p_Type)
                    {
                        case ParkedExceptionType.tstPExcepAbortSlew:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                                break;
                            }

                        case ParkedExceptionType.tstPExcepFindHome:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call FindHome method");
                                telescopeDevice.FindHome();
                                break;
                            }

                        case ParkedExceptionType.tstPExcepMoveAxisPrimary:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call MoveAxis(Primary, 0.0) method");
                                telescopeDevice.MoveAxis(TelescopeAxes.axisPrimary, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepMoveAxisSecondary:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call MoveAxis(Secondary, 0.0) method");
                                telescopeDevice.MoveAxis(TelescopeAxes.axisSecondary, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepMoveAxisTertiary:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call MoveAxis(Tertiary, 0.0) method");
                                telescopeDevice.MoveAxis(TelescopeAxes.axisTertiary, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepPulseGuide:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call PulseGuide(East, 0.0) method");
                                telescopeDevice.PulseGuide(GuideDirections.guideEast, 0);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToCoordinates:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call SlewToCoordinates method");
                                telescopeDevice.SlewToCoordinates(TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d), 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToCoordinatesAsync:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method");
                                telescopeDevice.SlewToCoordinatesAsync(TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d), 0.0d);
                                WaitForSlew("Parked:" + p_Name);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToTarget:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to set property TargetRightAscension to " + FormatRA(l_TargetRA));
                                telescopeDevice.TargetRightAscension = l_TargetRA;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to set property TargetDeclination to 0.0");
                                telescopeDevice.TargetDeclination = 0.0d;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call SlewToTarget method");
                                telescopeDevice.SlewToTarget();
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSlewToTargetAsync:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to set property to " + FormatRA(l_TargetRA));
                                telescopeDevice.TargetRightAscension = l_TargetRA;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to set property to 0.0");
                                telescopeDevice.TargetDeclination = 0.0d;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call method");
                                telescopeDevice.SlewToTargetAsync();
                                WaitForSlew("Parked:" + p_Name);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSyncToCoordinates:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call method, RA: " + FormatRA(l_TargetRA) + ", Declination: 0.0");
                                telescopeDevice.SyncToCoordinates(l_TargetRA, 0.0d);
                                break;
                            }

                        case ParkedExceptionType.tstPExcepSyncToTarget:
                            {
                                l_TargetRA = TelescopeRAFromSiderealTime("Parked:" + p_Name, 1.0d);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to set property to " + FormatRA(l_TargetRA));
                                telescopeDevice.TargetRightAscension = l_TargetRA;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to set property to 0.0");
                                telescopeDevice.TargetDeclination = 0.0d;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to call SyncToTarget method");
                                telescopeDevice.SyncToTarget();
                                break;
                            }

                        default:
                            {
                                LogMsg("Parked:" + p_Name, MessageLevel.msgError, "Conform:ParkedExceptionTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    LogMsg("Parked:" + p_Name, MessageLevel.msgIssue, p_Name + " didn't raise an error when Parked as required");
                }
                catch (Exception)
                {
                    LogMsg("Parked:" + p_Name, MessageLevel.msgOK, p_Name + " did raise an exception when Parked as required");
                }
                // Check that Telescope is still parked after issuing the command!
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("Parked:" + p_Name, MessageLevel.msgComment, "About to get AtPark property");
                if (Conversions.ToBoolean(!telescopeDevice.AtPark))
                    LogMsg("Parked:" + p_Name, MessageLevel.msgIssue, "Telescope was unparked by the " + p_Name + " command. This should not happen!");
            }
            else
            {
                LogMsg("Parked:" + p_Name, MessageLevel.msgIssue, "Not parked after Telescope.Park command, " + p_Name + " when parked test skipped");
            }

            g_Status.Clear();
        }

        private void TelescopeAxisRateTest(string p_Name, TelescopeAxes p_Axis)
        {
            int l_NAxisRates, l_i, l_j;
            bool l_AxisRateOverlap = default, l_AxisRateDuplicate, l_CanGetAxisRates = default, l_HasRates = default;
            int l_Count = 0;

#if DEBUG
            IAxisRates l_AxisRatesIRates;
            IAxisRates l_AxisRates = null;
            IRate l_Rate = null;
#else
            dynamic l_AxisRatesIRates;
            dynamic l_AxisRates = null;
            dynamic l_Rate = null;
#endif

            try
            {
                l_NAxisRates = 0;
                l_AxisRates = null;
                switch (p_Axis)
                {
                    case TelescopeAxes.axisPrimary:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method, Axis: " + ((int)TelescopeAxes.axisPrimary).ToString());
                            l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisPrimary); // Get primary axis rates
                            break;
                        }

                    case TelescopeAxes.axisSecondary:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method, Axis: " + ((int)TelescopeAxes.axisSecondary).ToString());
                            l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisSecondary); // Get secondary axis rates
                            break;
                        }

                    case TelescopeAxes.axisTertiary:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method, Axis: " + ((int)TelescopeAxes.axisTertiary).ToString());
                            l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisTertiary); // Get tertiary axis rates
                            break;
                        }

                    default:
                        {
                            LogMsg("TelescopeAxisRateTest", MessageLevel.msgError, "Unknown telescope axis: " + p_Axis.ToString());
                            break;
                        }
                }

                try
                {
                    if (l_AxisRates is null)
                    {
                        LogMsg(p_Name, MessageLevel.msgDebug, "ERROR: The driver did NOT return an AxisRates object!");
                    }
                    else
                    {
                        LogMsg(p_Name, MessageLevel.msgDebug, "OK - the driver returned an AxisRates object");
                    }

                    l_Count = Conversions.ToInteger(l_AxisRates.Count); // Save count for use later if no members are returned in the for each loop test
                    LogMsg(p_Name + " Count", MessageLevel.msgDebug, "The driver returned " + l_Count + " rates");
                    int i;
                    var loopTo = l_Count;
                    for (i = 1; i <= loopTo; i++)
                    {
#if DEBUG
                        IRate AxisRateItem;
#else
                        dynamic AxisRateItem;
#endif

                        AxisRateItem = l_AxisRates[i];
                        LogMsg(p_Name + " Count", MessageLevel.msgDebug, "Rate " + i + " - Minimum: " + AxisRateItem.Minimum.ToString() + ", Maximum: " + AxisRateItem.Maximum.ToString());
                    }
                }
                catch (COMException ex)
                {
                    LogMsg(p_Name + " Count", MessageLevel.msgError, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                }
                catch (Exception ex)
                {
                    LogMsg(p_Name + " Count", MessageLevel.msgError, EX_NET + ex.ToString());
                }

                try
                {
                    IEnumerator l_Enum;
                    dynamic l_Obj;
#if DEBUG
                    IRate AxisRateItem = null;
#else
                    dynamic AxisRateItem = null;
#endif

                    l_Enum = (IEnumerator)l_AxisRates.GetEnumerator();
                    if (l_Enum is null)
                    {
                        LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "ERROR: The driver did NOT return an Enumerator object!");
                    }
                    else
                    {
                        LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "OK - the driver returned an Enumerator object");
                    }

                    l_Enum.Reset();
                    LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Reset Enumerator");
                    while (l_Enum.MoveNext())
                    {
                        LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Reading Current");
                        l_Obj = l_Enum.Current;
                        LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Read Current OK, Type: " + l_Obj.GetType().Name);
#if DEBUG
                        AxisRateItem = (IRate)l_Obj;
#else
                        AxisRateItem = l_Obj;
#endif

                        LogMsg(p_Name + " Enum", MessageLevel.msgDebug, "Found axis rate - Minimum: " + AxisRateItem.Minimum.ToString() + ", Maximum: " + AxisRateItem.Maximum.ToString());
                    }

                    l_Enum.Reset();
                    l_Enum = null;
                    AxisRateItem = null;
                }
                catch (COMException ex)
                {
                    LogMsg(p_Name + " Enum", MessageLevel.msgError, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                }
                catch (Exception ex)
                {
                    LogMsg(p_Name + " Enum", MessageLevel.msgError, EX_NET + ex.ToString());
                }

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(l_AxisRates.Count, 0, false)))
                {
                    try
                    {
#if DEBUG
                        l_AxisRatesIRates = (IAxisRates)l_AxisRates;
                        foreach (IRate currentL_Rate in l_AxisRatesIRates)
                        {
                            l_Rate = currentL_Rate;
                            if (Conversions.ToBoolean(Operators.OrObject(Operators.ConditionalCompareObjectLess(l_Rate.Minimum, 0, false), Operators.ConditionalCompareObjectLess(l_Rate.Maximum, 0, false)))) // Error because negative values are not allowed
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " + l_Rate.Minimum.ToString() + ", " + l_Rate.Maximum.ToString());
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLessEqual(l_Rate.Minimum, l_Rate.Maximum, false))) // All positive values so continue tests
                                                                                                                                                // Minimum <= Maximum so OK
                            {
                                LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " + l_Rate.Minimum.ToString() + " Axis rate maximum: " + l_Rate.Maximum.ToString());
                            }
                            else // Minimum > Maximum so error!
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " + l_Rate.Minimum.ToString() + " maximum: " + l_Rate.Maximum.ToString());
                            }

                            // Save rates for overlap testing
                            l_NAxisRates += 1;
                            m_AxisRatesArray[l_NAxisRates, AXIS_RATE_MINIMUM] = Conversions.ToDouble(l_Rate.Minimum);
                            m_AxisRatesArray[l_NAxisRates, AXIS_RATE_MAXIMUM] = Conversions.ToDouble(l_Rate.Maximum);
                            l_HasRates = true;
                        }
                    }
#else
                        if (g_Settings.UseDriverAccess)
                        {
                            l_AxisRatesIRates = (IAxisRates)l_AxisRates;
                            foreach (var currentL_Rate in l_AxisRatesIRates)
                            {
                                l_Rate = currentL_Rate;
                                if (Conversions.ToBoolean(Operators.OrObject(Operators.ConditionalCompareObjectLess(l_Rate.Minimum, 0, false), Operators.ConditionalCompareObjectLess(l_Rate.Maximum, 0, false)))) // Error because negative values are not allowed
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " + l_Rate.Minimum.ToString() + ", " + l_Rate.Maximum.ToString());
                                }
                                else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLessEqual(l_Rate.Minimum, l_Rate.Maximum, false))) // All positive values so continue tests
                                                                                                                                                    // Minimum <= Maximum so OK
                                {
                                    LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " + l_Rate.Minimum.ToString() + " Axis rate maximum: " + l_Rate.Maximum.ToString());
                                }
                                else // Minimum > Maximum so error!
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " + l_Rate.Minimum.ToString() + " maximum: " + l_Rate.Maximum.ToString());
                                }

                                // Save rates for overlap testing
                                l_NAxisRates += 1;
                                m_AxisRatesArray[l_NAxisRates, AXIS_RATE_MINIMUM] = Conversions.ToDouble(l_Rate.Minimum);
                                m_AxisRatesArray[l_NAxisRates, AXIS_RATE_MAXIMUM] = Conversions.ToDouble(l_Rate.Maximum);
                                l_HasRates = true;
                            }
                        }
                        else
                        {
                            foreach (var currentL_Rate1 in (IEnumerable)l_AxisRates)
                            {
                                l_Rate = currentL_Rate1;
                                if (Conversions.ToBoolean(Operators.OrObject(Operators.ConditionalCompareObjectLess(l_Rate.Minimum, 0, false), Operators.ConditionalCompareObjectLess(l_Rate.Maximum, 0, false)))) // Error because negative values are not allowed
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "Minimum or maximum rate is negative: " + l_Rate.Minimum.ToString() + ", " + l_Rate.Maximum.ToString());
                                }
                                else // All positive values so continue tests
                                {
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLessEqual(l_Rate.Minimum, l_Rate.Maximum, false))) // Minimum <= Maximum so OK
                                    {
                                        LogMsg(p_Name, MessageLevel.msgOK, "Axis rate minimum: " + l_Rate.Minimum.ToString() + " Axis rate maximum: " + l_Rate.Maximum.ToString());
                                    }
                                    else // Minimum > Maximum so error!
                                    {
                                        LogMsg(p_Name, MessageLevel.msgError, "Maximum rate is less than minimum rate - minimum: " + l_Rate.Minimum.ToString() + " maximum: " + l_Rate.Maximum.ToString());
                                    }

                                    l_HasRates = true;
                                }

                                // Save rates for overlap testing
                                l_NAxisRates += 1;
                                m_AxisRatesArray[l_NAxisRates, AXIS_RATE_MINIMUM] = Conversions.ToDouble(l_Rate.Minimum);
                                m_AxisRatesArray[l_NAxisRates, AXIS_RATE_MAXIMUM] = Conversions.ToDouble(l_Rate.Maximum);
                            }
                        }
                    }
#endif

                    catch (COMException ex)
                    {
                        LogMsg(p_Name, MessageLevel.msgError, "COM Unable to read AxisRates object - Exception: " + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                        LogMsg(p_Name, MessageLevel.msgDebug, "COM Unable to read AxisRates object - Exception: " + ex.ToString());
                    }
                    catch (DriverException ex)
                    {
                        LogMsg(p_Name, MessageLevel.msgError, ".NET Unable to read AxisRates object - Exception: " + ex.Message + " " + Conversion.Hex(ex.Number));
                        LogMsg(p_Name, MessageLevel.msgDebug, ".NET Unable to read AxisRates object - Exception: " + ex.ToString());
                    }
                    catch (Exception ex)
                    {
                        LogMsg(p_Name, MessageLevel.msgError, "Unable to read AxisRates object - Exception: " + ex.Message);
                        LogMsg(p_Name, MessageLevel.msgDebug, "Unable to read AxisRates object - Exception: " + ex.ToString());
                    }

                    // Overlap testing
                    if (l_NAxisRates > 1) // Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    {
                        var loopTo1 = l_NAxisRates;
                        for (l_i = 1; l_i <= loopTo1; l_i++)
                        {
                            var loopTo2 = l_NAxisRates;
                            for (l_j = 1; l_j <= loopTo2; l_j++)
                            {
                                if (l_i != l_j) // Only test different lines, shouldn't compare same lines!
                                {
                                    if (m_AxisRatesArray[l_i, AXIS_RATE_MINIMUM] >= m_AxisRatesArray[l_j, AXIS_RATE_MINIMUM] & m_AxisRatesArray[l_i, AXIS_RATE_MINIMUM] <= m_AxisRatesArray[l_j, AXIS_RATE_MAXIMUM])
                                        l_AxisRateOverlap = true;
                                }
                            }
                        }
                    }

                    if (l_AxisRateOverlap)
                    {
                        LogMsg(p_Name, MessageLevel.msgIssue, "Overlapping axis rates found, suggest these be rationalised to remove overlaps");
                    }
                    else
                    {
                        LogMsg(p_Name, MessageLevel.msgOK, "No overlapping axis rates found");
                    }

                    // Duplicate testing
                    l_AxisRateDuplicate = false;
                    if (l_NAxisRates > 1) // Confirm whether there are overlaps if number of axis rate pairs exceeds 1
                    {
                        var loopTo3 = l_NAxisRates;
                        for (l_i = 1; l_i <= loopTo3; l_i++)
                        {
                            var loopTo4 = l_NAxisRates;
                            for (l_j = 1; l_j <= loopTo4; l_j++)
                            {
                                if (l_i != l_j) // Only test different lines, shouldn't compare same lines!
                                {
                                    if (m_AxisRatesArray[l_i, AXIS_RATE_MINIMUM] == m_AxisRatesArray[l_j, AXIS_RATE_MINIMUM] & m_AxisRatesArray[l_i, AXIS_RATE_MAXIMUM] == m_AxisRatesArray[l_j, AXIS_RATE_MAXIMUM])
                                        l_AxisRateDuplicate = true;
                                }
                            }
                        }
                    }

                    if (l_AxisRateDuplicate)
                    {
                        LogMsg(p_Name, MessageLevel.msgIssue, "Duplicate axis rates found, suggest these be removed");
                    }
                    else
                    {
                        LogMsg(p_Name, MessageLevel.msgOK, "No duplicate axis rates found");
                    }
                }
                else
                {
                    LogMsg(p_Name, MessageLevel.msgOK, "Empty axis rate returned");
                }

                l_CanGetAxisRates = true; // Record that this driver can deliver a viable AxisRates object that can be tested for AxisRates.Dispose() later
            }
            catch (COMException ex)
            {
                LogMsg(p_Name, MessageLevel.msgError, "COM Unable to get an AxisRates object - Exception: " + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
            }
            catch (DriverException ex)
            {
                LogMsg(p_Name, MessageLevel.msgError, ".NET Unable to get an AxisRates object - Exception: " + ex.Message + " " + Conversion.Hex(ex.Number));
            }
            catch (NullReferenceException ex) // Report null objects returned by the driver that are caught by DriverAccess.
            {
                LogMsg(p_Name, MessageLevel.msgError, ex.Message);
                LogMsg(p_Name, MessageLevel.msgDebug, ex.ToString());
            } // If debug then give full information
            catch (Exception ex)
            {
                LogMsg(p_Name, MessageLevel.msgError, "Unable to get or unable to use an AxisRates object - Exception: " + ex.ToString());
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

                try
                {
                    Marshal.ReleaseComObject(l_AxisRates);
                }
                catch
                {
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

                try
                {
                    Marshal.ReleaseComObject(l_Rate);
                }
                catch
                {
                }

                l_Rate = null;
            }

            if (l_CanGetAxisRates) // The driver does return a viable AxisRates object that can be tested for correct AxisRates.Dispose() and Rate.Dispose() operation
            {
                try
                {
                    // Test Rate.Dispose()
                    switch (p_Axis) // Get the relevant axis rates object for this axis
                    {
                        case TelescopeAxes.axisPrimary:
                            {
                                l_AxisRates = DriverAsObject.AxisRates(TelescopeAxes.axisPrimary);
                                break;
                            }

                        case TelescopeAxes.axisSecondary:
                            {
                                l_AxisRates = DriverAsObject.AxisRates(TelescopeAxes.axisSecondary);
                                break;
                            }

                        case TelescopeAxes.axisTertiary:
                            {
                                l_AxisRates = DriverAsObject.AxisRates(TelescopeAxes.axisTertiary);
                                break;
                            }

                        default:
                            {
                                LogMsgError(p_Name, "AxisRate.Dispose() - Unknown axis: " + p_Axis.ToString());
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
                                LogMsgOK(p_Name, string.Format("Successfully disposed of rate {0} - {1}", l_Rate.Minimum, l_Rate.Maximum));
                            }
                            catch (MissingMemberException)
                            {
                                LogMsg(p_Name, MessageLevel.msgOK, string.Format("Rate.Dispose() member not present for rate {0} - {1}", l_Rate.Minimum, l_Rate.Maximum));
                            }
                            catch (Exception ex1)
                            {
                                LogMsgWarning(p_Name, string.Format("Rate.Dispose() for rate {0} - {1} threw an exception but it is poor practice to throw exceptions in Dispose methods: {2}", l_Rate.Minimum, l_Rate.Maximum, ex1.Message));
                                LogMsg("TrackingRates.Dispose", MessageLevel.msgDebug, "Exception: " + ex1.ToString());
                            }
                        }
                    }

                    // Test AxisRates.Dispose()
                    try
                    {
                        LogMsg(p_Name, MessageLevel.msgDebug, "Disposing axis rates");
                        l_AxisRates.Dispose();
                        LogMsg(p_Name, MessageLevel.msgOK, "Disposed axis rates OK");
                    }
                    catch (MissingMemberException)
                    {
                        LogMsg(p_Name, MessageLevel.msgOK, "AxisRates.Dispose() member not present for axis " + p_Axis.ToString());
                    }
                    catch (Exception ex1)
                    {
                        LogMsgWarning(p_Name, "AxisRates.Dispose() threw an exception but it is poor practice to throw exceptions in Dispose() methods: " + ex1.Message);
                        LogMsg("AxisRates.Dispose", MessageLevel.msgDebug, "Exception: " + ex1.ToString());
                    }
                }
                catch (Exception ex)
                {
                    LogMsg(p_Name, MessageLevel.msgError, "AxisRate.Dispose() - Unable to get or unable to use an AxisRates object - Exception: " + ex.ToString());
                }
            }
            else
            {
                LogMsgInfo(p_Name, "AxisRates.Dispose() testing skipped because of earlier issues in obtaining a viable AxisRates object.");
            }

            g_Status.Clear();  // Clear status messages
        }

        private void TelescopeRequiredMethodsTest(RequiredMethodType p_Type, string p_Name)
        {
            try
            {
                g_Status.Test = p_Name;
                switch (p_Type)
                {
                    case RequiredMethodType.tstAxisrates:
                        {
                            break;
                        }
                    // This is now done by TelescopeAxisRateTest subroutine 
                    case RequiredMethodType.tstCanMoveAxisPrimary:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call CanMoveAxis method " + ((int)TelescopeAxes.axisPrimary).ToString());
                            m_CanMoveAxisPrimary = Conversions.ToBoolean(telescopeDevice.CanMoveAxis(TelescopeAxes.axisPrimary));
                            LogMsg(p_Name, MessageLevel.msgOK, p_Name + " " + m_CanMoveAxisPrimary.ToString());
                            break;
                        }

                    case RequiredMethodType.tstCanMoveAxisSecondary:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call CanMoveAxis method " + ((int)TelescopeAxes.axisSecondary).ToString());
                            m_CanMoveAxisSecondary = Conversions.ToBoolean(telescopeDevice.CanMoveAxis(TelescopeAxes.axisSecondary));
                            LogMsg(p_Name, MessageLevel.msgOK, p_Name + " " + m_CanMoveAxisSecondary.ToString());
                            break;
                        }

                    case RequiredMethodType.tstCanMoveAxisTertiary:
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(p_Name, MessageLevel.msgComment, "About to call CanMoveAxis method " + ((int)TelescopeAxes.axisTertiary).ToString());
                            m_CanMoveAxisTertiary = Conversions.ToBoolean(telescopeDevice.CanMoveAxis(TelescopeAxes.axisTertiary));
                            LogMsg(p_Name, MessageLevel.msgOK, p_Name + " " + m_CanMoveAxisTertiary.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Conform:RequiredMethodsTest: Unknown test type " + p_Type.ToString());
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
            g_Status.Clear(); // Clear status messages
        }

        private void TelescopeOptionalMethodsTest(OptionalMethodType p_Type, string p_Name, bool p_CanTest)
        {
            int l_ct;
            double l_TestDec, l_TestRAOffset;
#if DEBUG
            IAxisRates l_AxisRates = null;
#else
            dynamic l_AxisRates = null;
#endif

            Status(StatusType.staTest, p_Name);
            LogMsg("TelescopeOptionalMethodsTest", MessageLevel.msgDebug, p_Type.ToString() + " " + p_Name + " " + p_CanTest.ToString());
            if (p_CanTest) // Confirm that an error is raised if the optional command is not implemented
            {
                try
                {
                    // Set the test declination value depending on whether the scope is in the northern or southern hemisphere
                    if (m_SiteLatitude > 0.0d)
                    {
                        l_TestDec = 45.0d; // Positive for the northern hemisphere
                    }
                    else
                    {
                        l_TestDec = -45.0d;
                    } // Negative for the southern hemisphere

                    l_TestRAOffset = 3.0d; // Set the test RA offset as 3 hours from local sider5eal time
                    LogMsg(p_Name, MessageLevel.msgDebug, string.Format("Test RA offset: {0}, Test declination: {1}", l_TestRAOffset, l_TestDec));
                    switch (p_Type)
                    {
                        case OptionalMethodType.AbortSlew:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                                LogMsg("AbortSlew", MessageLevel.msgOK, "AbortSlew OK when not slewing");
                                break;
                            }

                        case OptionalMethodType.DestinationSideOfPier:
                            {
                                // Get the DestinationSideOfPier for a target in the West i.e. for a German mount when the tube is on the East side of the pier
                                m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -l_TestRAOffset);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call DestinationSideOfPier method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(l_TestDec));
                                m_DestinationSideOfPierEast = (PierSide)telescopeDevice.DestinationSideOfPier(m_TargetRightAscension, l_TestDec);
                                LogMsg(p_Name, MessageLevel.msgDebug, "German mount - scope on the pier's East side, target in the West : " + FormatRA(m_TargetRightAscension) + " " + FormatDec(l_TestDec) + " " + m_DestinationSideOfPierEast.ToString());

                                // Get the DestinationSideOfPier for a target in the East i.e. for a German mount when the tube is on the West side of the pier
                                m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, l_TestRAOffset);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call DestinationSideOfPier method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(l_TestDec));
                                m_DestinationSideOfPierWest = (PierSide)telescopeDevice.DestinationSideOfPier(m_TargetRightAscension, l_TestDec);
                                LogMsg(p_Name, MessageLevel.msgDebug, "German mount - scope on the pier's West side, target in the East: " + FormatRA(m_TargetRightAscension) + " " + FormatDec(l_TestDec) + " " + m_DestinationSideOfPierWest.ToString());

                                // Make sure that we received two valid values i.e. that neither side returned PierSide.Unknown and that the two valid returned values are not the same i.e. we got one PierSide.PierEast and one PierSide.PierWest
                                if (m_DestinationSideOfPierEast == PierSide.pierUnknown | m_DestinationSideOfPierWest == PierSide.pierUnknown)
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "Invalid SideOfPier value received, Target in West: " + m_DestinationSideOfPierEast.ToString() + ", Target in East: " + m_DestinationSideOfPierWest.ToString());
                                }
                                else if (m_DestinationSideOfPierEast == m_DestinationSideOfPierWest)
                                {
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Same value for DestinationSideOfPier received on both sides of the meridian: " + ((int)m_DestinationSideOfPierEast).ToString());
                                }
                                else
                                {
                                    LogMsg(p_Name, MessageLevel.msgOK, "DestinationSideOfPier is different on either side of the meridian");
                                }

                                break;
                            }

                        case OptionalMethodType.FindHome:
                            {
                                if (g_InterfaceVersion > 1)
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call FindHome method");
                                    telescopeDevice.FindHome();
                                    m_StartTime = DateAndTime.Now;
                                    Status(StatusType.staAction, "Waiting for mount to home");
                                    l_ct = 0;
                                    do
                                    {
                                        WaitFor(SLEEP_TIME);
                                        l_ct += 1;
                                        g_Status.Status = l_ct.ToString();
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to get AtHome property");
                                    }
                                    while (!telescopeDevice.AtHome & TestStop() & (DateAndTime.Now.Subtract(m_StartTime).TotalMilliseconds < 60000)); // Wait up to a minute to find home
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to get AtHome property");
                                    if (Conversions.ToBoolean(telescopeDevice.AtHome))
                                    {
                                        LogMsg(p_Name, MessageLevel.msgOK, "Found home OK.");
                                    }
                                    else
                                    {
                                        LogMsg(p_Name, MessageLevel.msgInfo, "Failed to Find home within 1 minute");
                                    }

                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to get AtPark property");
                                    if (Conversions.ToBoolean(telescopeDevice.AtPark))
                                    {
                                        LogMsg(p_Name, MessageLevel.msgIssue, "FindHome has parked the scope as well as finding home");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to call UnPark method");
                                        telescopeDevice.Unpark(); // Unpark it ready for further tests
                                    }
                                }
                                else
                                {
                                    Status(StatusType.staAction, "Waiting for mount to home");
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call FindHome method");
                                    telescopeDevice.FindHome();
                                    g_Status.Clear();
                                    LogMsg(p_Name, MessageLevel.msgOK, "Found home OK.");
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call Unpark method");
                                    telescopeDevice.Unpark();
                                } // Make sure we are still  unparked!

                                break;
                            }

                        case OptionalMethodType.MoveAxisPrimary:
                            {
                                // Get axis rates for primary axis
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call AxisRates method for axis " + ((int)TelescopeAxes.axisPrimary).ToString());
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisPrimary);
                                TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisPrimary, l_AxisRates);
                                break;
                            }

                        case OptionalMethodType.MoveAxisSecondary:
                            {
                                // Get axis rates for secondary axis
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisSecondary);
                                TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisSecondary, l_AxisRates);
                                break;
                            }

                        case OptionalMethodType.MoveAxisTertiary:
                            {
                                // Get axis rates for tertiary axis
                                l_AxisRates = telescopeDevice.AxisRates(TelescopeAxes.axisTertiary);
                                TelescopeMoveAxisTest(p_Name, TelescopeAxes.axisTertiary, l_AxisRates);
                                break;
                            }

                        case OptionalMethodType.PulseGuide:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property");
                                if (Conversions.ToBoolean(telescopeDevice.IsPulseGuiding)) // IsPulseGuiding is true before we've started so this is an error and voids a real test
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "IsPulseGuiding is True when not pulse guiding - PulseGuide test omitted");
                                }
                                else // OK to test pulse guiding
                                {
                                    Status(StatusType.staAction, "Start PulseGuide");
                                    m_StartTime = DateAndTime.Now;
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call PulseGuide method, Direction: " + ((int)GuideDirections.guideEast).ToString() + ", Duration: " + PULSEGUIDE_MOVEMENT_TIME * 1000 + "ms");
                                    telescopeDevice.PulseGuide(GuideDirections.guideEast, PULSEGUIDE_MOVEMENT_TIME * 1000); // Start a 2 second pulse
                                    m_EndTime = DateAndTime.Now;
                                    LogMsg(p_Name, MessageLevel.msgDebug, "PulseGuide command time: " + PULSEGUIDE_MOVEMENT_TIME * 1000 + " milliseconds, PulseGuide call duration: " + m_EndTime.Subtract(m_StartTime).TotalMilliseconds + " milliseconds");
                                    if (m_EndTime.Subtract(m_StartTime).TotalMilliseconds < PULSEGUIDE_MOVEMENT_TIME * 0.75d * 1000d) // If less than three quarters of the expected duration then assume we have returned early
                                    {
                                        l_ct = 0;
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property");
                                        if (Conversions.ToBoolean(telescopeDevice.IsPulseGuiding))
                                        {
                                            do
                                            {
                                                WaitFor(SLEEP_TIME);
                                                l_ct += 1;
                                                g_Status.Status = l_ct.ToString();
                                                if (TestStop())
                                                    return;
                                                if (g_Settings.DisplayMethodCalls)
                                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property");
                                            }
                                            while (telescopeDevice.IsPulseGuiding & DateTime.Now.Subtract(m_StartTime).TotalMilliseconds < PULSEGUIDE_TIMEOUT_TIME * 1000); // Wait for success or timeout
                                            if (g_Settings.DisplayMethodCalls)
                                                LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property");
                                            if (Conversions.ToBoolean(!telescopeDevice.IsPulseGuiding))
                                            {
                                                LogMsg(p_Name, MessageLevel.msgOK, "Asynchronous pulse guide found OK");
                                                LogMsg(p_Name, MessageLevel.msgDebug, "IsPulseGuiding = True duration: " + DateTime.Now.Subtract(m_StartTime).TotalMilliseconds + " milliseconds");
                                            }
                                            else
                                            {
                                                LogMsg(p_Name, MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding is still TRUE " + PULSEGUIDE_TIMEOUT_TIME + " seconds beyond expected time");
                                            }
                                        }
                                        else
                                        {
                                            LogMsg(p_Name, MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding has returned FALSE");
                                        }
                                    }
                                    else // Assume synchronous pulse guide and that IsPulseGuiding is false
                                    {
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to get IsPulseGuiding property");
                                        if (Conversions.ToBoolean(!telescopeDevice.IsPulseGuiding))
                                        {
                                            LogMsg(p_Name, MessageLevel.msgOK, "Synchronous pulse guide found OK");
                                        }
                                        else
                                        {
                                            LogMsg(p_Name, MessageLevel.msgIssue, "Synchronous pulse guide expected but IsPulseGuiding has returned TRUE");
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
                                    SlewScope(TelescopeRAFromHourAngle(p_Name, -3.0d), 0.0d, "Slewing to far start point");
                                    if (TestStop())
                                        return;
                                    SlewScope(TelescopeRAFromHourAngle(p_Name, -0.03d), 0.0d, "Slewing to near start point"); // 2 minutes from zenith
                                    if (TestStop())
                                        return;

                                    // We are now 2 minutes from the meridian looking east so allow the mount to track for 7 minutes 
                                    // so it passes through the meridian and ends up 5 minutes past the meridian
                                    LogMsg(p_Name, MessageLevel.msgInfo, "This test will now wait for 7 minutes while the mount tracks through the Meridian");

                                    // Wait for mount to move
                                    m_StartTime = DateAndTime.Now;
                                    do
                                    {
                                        System.Threading.Thread.Sleep(SLEEP_TIME);
                                        Application.DoEvents();
                                        SetStatus(p_Name, "Waiting for transit through Meridian", Convert.ToInt32(DateAndTime.Now.Subtract(m_StartTime).TotalSeconds) + "/" + SIDEOFPIER_MERIDIAN_TRACKING_PERIOD / 1000d + " seconds");
                                    }
                                    while (!(DateAndTime.Now.Subtract(m_StartTime).TotalMilliseconds > SIDEOFPIER_MERIDIAN_TRACKING_PERIOD | TestStop()));

                                    // SlewScope(TelescopeRAFromHourAngle(+0.0833333), 0.0, "Slewing to flip point") '5 minutes past zenith
                                    if (TestStop())
                                        return;
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to get SideOfPier property");
                                    switch (telescopeDevice.SideOfPier)
                                    {
                                        case PierSide.pierEast: // We are on pierEast so try pierWest
                                            {
                                                try
                                                {
                                                    LogMsg(p_Name, MessageLevel.msgDebug, "Scope is pierEast so flipping West");
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " + ((int)PierSide.pierWest).ToString());
                                                    telescopeDevice.SideOfPier = PierSide.pierWest;
                                                    WaitForSlew(p_Name);
                                                    if (TestStop())
                                                        return;
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg(p_Name, MessageLevel.msgComment, "About to get SideOfPier property");
                                                    m_SideOfPier = (PierSide)telescopeDevice.SideOfPier;
                                                    if (m_SideOfPier == PierSide.pierWest)
                                                    {
                                                        LogMsg(p_Name, MessageLevel.msgOK, "Successfully flipped pierEast to pierWest");
                                                    }
                                                    else
                                                    {
                                                        LogMsg(p_Name, MessageLevel.msgIssue, "Failed to set SideOfPier to pierWest, got: " + m_SideOfPier.ToString());
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleException("SideOfPier Write pierWest", MemberType.Method, Required.MustBeImplemented, ex, "CanSetPierSide is True");
                                                }

                                                break;
                                            }

                                        case PierSide.pierWest: // We are on pierWest so try pierEast
                                            {
                                                try
                                                {
                                                    LogMsg(p_Name, MessageLevel.msgDebug, "Scope is pierWest so flipping East");
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " + ((int)PierSide.pierEast).ToString());
                                                    telescopeDevice.SideOfPier = PierSide.pierEast;
                                                    WaitForSlew(p_Name);
                                                    if (TestStop())
                                                        return;
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg(p_Name, MessageLevel.msgComment, "About to get SideOfPier property");
                                                    m_SideOfPier = (PierSide)telescopeDevice.SideOfPier;
                                                    if (m_SideOfPier == PierSide.pierEast)
                                                    {
                                                        LogMsg(p_Name, MessageLevel.msgOK, "Successfully flipped pierWest to pierEast");
                                                    }
                                                    else
                                                    {
                                                        LogMsg(p_Name, MessageLevel.msgIssue, "Failed to set SideOfPier to pierEast, got: " + m_SideOfPier.ToString());
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleException("SideOfPier Write pierEast", MemberType.Method, Required.MustBeImplemented, ex, "CanSetPierSide is True");
                                                } // Unknown pier side

                                                break;
                                            }

                                        default:
                                            {
                                                LogMsg(p_Name, MessageLevel.msgError, "Unknown PierSide: " + m_SideOfPier.ToString());
                                                break;
                                            }
                                    }
                                }
                                else // Can't set pier side so it should generate an error
                                {
                                    try
                                    {
                                        LogMsg(p_Name, MessageLevel.msgDebug, "Attempting to set SideOfPier");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " + ((int)PierSide.pierEast).ToString());
                                        telescopeDevice.SideOfPier = PierSide.pierEast;
                                        LogMsg(p_Name, MessageLevel.msgDebug, "SideOfPier set OK to pierEast but should have thrown an error");
                                        WaitForSlew(p_Name);
                                        LogMsg(p_Name, MessageLevel.msgIssue, "CanSetPierSide is false but no exception was generated when set was attempted");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "CanSetPierSide is False");
                                    }
                                    finally
                                    {
                                        WaitForSlew(p_Name);
                                    } // Make sure slewing is stopped if an exception was thrown
                                }

                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to false");
                                telescopeDevice.Tracking = false;
                                if (TestStop())
                                    return;
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "Conform:OptionalMethodsTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    // Clean up AxisRate object, if used
                    if (l_AxisRates is object)
                    {
                        try
                        {
#if DEBUG
                            if (g_Settings.DisplayMethodCalls) LogMsg(p_Name, MessageLevel.msgComment, "About to dispose of AxisRates object");
                            l_AxisRates.Dispose();
#else
                            if (g_Settings.UseDriverAccess)
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to dispose of AxisRates object");
                                l_AxisRates.Dispose();
                            }
#endif

                            LogMsg(p_Name, MessageLevel.msgOK, "AxisRates object successfully disposed");
                        }
                        catch (Exception ex)
                        {
                            LogMsgError(p_Name, "AxisRates.Dispose threw an exception but must not: " + ex.Message);
                            LogMsg(p_Name, MessageLevel.msgDebug, "Exception: " + ex.ToString());
                        }

                        try
                        {
                            Marshal.ReleaseComObject(l_AxisRates);
                        }
                        catch
                        {
                        }

                        l_AxisRates = null;
                    }

                    if (TestStop())
                        return;
                }
                catch (Exception ex)
                {
                    HandleException(p_Name, MemberType.Method, Required.Optional, ex, "");
                }
            }
            else // Can property is false so confirm that an error is generated
            {
                try
                {
                    switch (p_Type)
                    {
                        case OptionalMethodType.AbortSlew:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call AbortSlew method");
                                telescopeDevice.AbortSlew();
                                break;
                            }

                        case OptionalMethodType.DestinationSideOfPier:
                            {
                                m_TargetRightAscension = TelescopeRAFromSiderealTime(p_Name, -1.0d);
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call DestinationSideOfPier method, RA: " + FormatRA(m_TargetRightAscension) + ", Declination: " + FormatDec(0.0d));
                                m_DestinationSideOfPier = (PierSide)telescopeDevice.DestinationSideOfPier(m_TargetRightAscension, 0.0d);
                                break;
                            }

                        case OptionalMethodType.FindHome:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call FindHome method");
                                telescopeDevice.FindHome();
                                break;
                            }

                        case OptionalMethodType.MoveAxisPrimary:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)TelescopeAxes.axisPrimary).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxes.axisPrimary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.MoveAxisSecondary:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)TelescopeAxes.axisSecondary).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxes.axisSecondary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.MoveAxisTertiary:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)TelescopeAxes.axisTertiary).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(TelescopeAxes.axisTertiary, 0.0d);
                                break;
                            }

                        case OptionalMethodType.PulseGuide:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call PulseGuide method, Direction: " + ((int)GuideDirections.guideEast).ToString() + ", Duration: 0ms");
                                telescopeDevice.PulseGuide(GuideDirections.guideEast, 0);
                                break;
                            }

                        case OptionalMethodType.SideOfPierWrite:
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to set SideOfPier property to " + ((int)PierSide.pierEast).ToString());
                                telescopeDevice.SideOfPier = PierSide.pierEast;
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "Conform:OptionalMethodsTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    LogMsg(p_Name, MessageLevel.msgIssue, "Can" + p_Name + " is false but no exception was generated on use");
                }
                catch (Exception ex)
                {
                    if (IsInvalidValueException(p_Name, ex))
                    {
                        LogMsg(p_Name, MessageLevel.msgOK, "Received an invalid value exception");
                    }
                    else if (p_Type == OptionalMethodType.SideOfPierWrite) // PierSide is actually a property even though I have it in the methods section!!
                    {
                        HandleException(p_Name, MemberType.Property, Required.MustNotBeImplemented, ex, "Can" + p_Name + " is False");
                    }
                    else
                    {
                        HandleException(p_Name, MemberType.Method, Required.MustNotBeImplemented, ex, "Can" + p_Name + " is False");
                    }
                }
            }

            g_Status.Clear();  // Clear status messages
        }

        private void TelescopeCanTest(CanType p_Type, string p_Name)
        {
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg(p_Name, MessageLevel.msgComment, string.Format("About to get {0} property", p_Type.ToString()));
                switch (p_Type)
                {
                    case CanType.CanFindHome:
                        {
                            canFindHome = telescopeDevice.CanFindHome;
                            LogMsg(p_Name, MessageLevel.msgOK, canFindHome.ToString());
                            break;
                        }

                    case CanType.CanPark:
                        {
                            canPark = telescopeDevice.CanPark;
                            LogMsg(p_Name, MessageLevel.msgOK, canPark.ToString());
                            break;
                        }

                    case CanType.CanPulseGuide:
                        {
                            canPulseGuide = telescopeDevice.CanPulseGuide;
                            LogMsg(p_Name, MessageLevel.msgOK, canPulseGuide.ToString());
                            break;
                        }

                    case CanType.CanSetDeclinationRate:
                        {
                            if (g_InterfaceVersion > 1)
                            {
                                canSetDeclinationRate = telescopeDevice.CanSetDeclinationRate;
                                LogMsg(p_Name, MessageLevel.msgOK, canSetDeclinationRate.ToString());
                            }
                            else
                            {
                                LogMsg("CanSetDeclinationRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetGuideRates:
                        {
                            if (g_InterfaceVersion > 1)
                            {
                                canSetGuideRates = telescopeDevice.CanSetGuideRates;
                                LogMsg(p_Name, MessageLevel.msgOK, canSetGuideRates.ToString());
                            }
                            else
                            {
                                LogMsg("CanSetGuideRates", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetPark:
                        {
                            canSetPark = telescopeDevice.CanSetPark;
                            LogMsg(p_Name, MessageLevel.msgOK, canSetPark.ToString());
                            break;
                        }

                    case CanType.CanSetPierSide:
                        {
                            if (g_InterfaceVersion > 1)
                            {
                                canSetPierside = telescopeDevice.CanSetPierSide;
                                LogMsg(p_Name, MessageLevel.msgOK, canSetPierside.ToString());
                            }
                            else
                            {
                                LogMsg("CanSetPierSide", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetRightAscensionRate:
                        {
                            if (g_InterfaceVersion > 1)
                            {
                                canSetRightAscensionRate = telescopeDevice.CanSetRightAscensionRate;
                                LogMsg(p_Name, MessageLevel.msgOK, canSetRightAscensionRate.ToString());
                            }
                            else
                            {
                                LogMsg("CanSetRightAscensionRate", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSetTracking:
                        {
                            canSetTracking = telescopeDevice.CanSetTracking;
                            LogMsg(p_Name, MessageLevel.msgOK, canSetTracking.ToString());
                            break;
                        }

                    case CanType.CanSlew:
                        {
                            canSlew = telescopeDevice.CanSlew;
                            LogMsg(p_Name, MessageLevel.msgOK, canSlew.ToString());
                            break;
                        }

                    case CanType.CanSlewAltAz:
                        {
                            if (g_InterfaceVersion > 1)
                            {
                                canSlewAltAz = telescopeDevice.CanSlewAltAz;
                                LogMsg(p_Name, MessageLevel.msgOK, canSlewAltAz.ToString());
                            }
                            else
                            {
                                LogMsg("CanSlewAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSlewAltAzAsync:
                        {
                            if (g_InterfaceVersion > 1)
                            {
                                canSlewAltAzAsync = telescopeDevice.CanSlewAltAzAsync;
                                LogMsg(p_Name, MessageLevel.msgOK, canSlewAltAzAsync.ToString());
                            }
                            else
                            {
                                LogMsg("CanSlewAltAzAsync", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanSlewAsync:
                        {
                            canSlewAsync = telescopeDevice.CanSlewAsync;
                            LogMsg(p_Name, MessageLevel.msgOK, canSlewAsync.ToString());
                            break;
                        }

                    case CanType.CanSync:
                        {
                            canSync = telescopeDevice.CanSync;
                            LogMsg(p_Name, MessageLevel.msgOK, canSync.ToString());
                            break;
                        }

                    case CanType.CanSyncAltAz:
                        {
                            if (g_InterfaceVersion > 1)
                            {
                                canSyncAltAz = telescopeDevice.CanSyncAltAz;
                                LogMsg(p_Name, MessageLevel.msgOK, canSyncAltAz.ToString());
                            }
                            else
                            {
                                LogMsg("CanSyncAltAz", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
                            }

                            break;
                        }

                    case CanType.CanUnPark:
                        {
                            canUnpark = telescopeDevice.CanUnpark;
                            LogMsg(p_Name, MessageLevel.msgOK, canUnpark.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Conform:CanTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
            }
        }

#if DEBUG
        private void TelescopeMoveAxisTest(string p_Name, TelescopeAxes p_Axis, IAxisRates p_AxisRates)
        {
            IRate l_Rate = null;
#else
        private void TelescopeMoveAxisTest(string p_Name, TelescopeAxes p_Axis, dynamic p_AxisRates)
        {
            dynamic l_Rate = null;
#endif

            double l_MoveRate = default, l_RateMinimum, l_RateMaximum;
            bool l_TrackingStart, l_TrackingEnd, l_CanSetZero;
            int l_RateCount;

            // Determine lowest and highest tracking rates
            l_RateMinimum = double.PositiveInfinity; // Set to invalid values
            l_RateMaximum = double.NegativeInfinity;
            LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject("Number of rates found: ", p_AxisRates.Count)));
            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(p_AxisRates.Count, 0, false)))
            {
#if DEBUG
                IAxisRates l_AxisRatesIRates = (IAxisRates)p_AxisRates;
                l_RateCount = 0;
                foreach (IRate currentL_Rate in l_AxisRatesIRates)
                {
                    l_Rate = currentL_Rate;
                    if (l_Rate.Minimum < l_RateMinimum) l_RateMinimum = l_Rate.Minimum;
                    if (l_Rate.Maximum > l_RateMaximum) l_RateMaximum = l_Rate.Maximum;
                    LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Checking rates: ", l_Rate.Minimum), " "), l_Rate.Maximum), ", Current rates: "), l_RateMinimum), " "), l_RateMaximum)));
                    l_RateCount += 1;
                }
#else
                IAxisRates l_AxisRatesIRates;
                l_RateCount = 0;
                if (g_Settings.UseDriverAccess)
                {
                    l_AxisRatesIRates = (IAxisRates)p_AxisRates;
                    foreach (IRate currentL_Rate in l_AxisRatesIRates)
                    {
                        l_Rate = currentL_Rate;
                        if (l_Rate.Minimum < l_RateMinimum) l_RateMinimum = l_Rate.Minimum;
                        if (l_Rate.Maximum > l_RateMaximum) l_RateMaximum = l_Rate.Maximum;
                        LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Checking rates: ", l_Rate.Minimum), " "), l_Rate.Maximum), ", Current rates: "), l_RateMinimum), " "), l_RateMaximum)));
                        l_RateCount += 1;
                    }
                }
                else
                {
                    foreach (IRate currentL_Rate1 in (IEnumerable)p_AxisRates)
                    {
                        l_Rate = currentL_Rate1;
                        if (l_Rate.Minimum < l_RateMinimum) l_RateMinimum = l_Rate.Minimum;
                        if (l_Rate.Maximum > l_RateMaximum) l_RateMaximum = l_Rate.Maximum;
                        LogMsg(p_Name, MessageLevel.msgDebug, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Checking rates: ", l_Rate.Minimum), " "), l_Rate.Maximum), ", Current rates: "), l_RateMinimum), " "), l_RateMaximum)));
                        l_RateCount += 1;
                    }
                }
#endif

                if (l_RateMinimum != double.PositiveInfinity & l_RateMaximum != double.NegativeInfinity) // Found valid rates
                {
                    LogMsg(p_Name, MessageLevel.msgDebug, "Found minimum rate: " + l_RateMinimum + " found maximum rate: " + l_RateMaximum);

                    // Confirm setting a zero rate works
                    Status(StatusType.staAction, "Set zero rate");
                    l_CanSetZero = false;
                    try
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                        telescopeDevice.MoveAxis(p_Axis, 0.0d); // Set a value of zero
                        LogMsg(p_Name, MessageLevel.msgOK, "Can successfully set a movement rate of zero");
                        l_CanSetZero = true;
                    }
                    catch (COMException ex)
                    {
                        LogMsg(p_Name, MessageLevel.msgError, "Unable to set a movement rate of zero - " + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                    }
                    catch (DriverException ex)
                    {
                        LogMsg(p_Name, MessageLevel.msgError, "Unable to set a movement rate of zero - " + ex.Message + " " + Conversion.Hex(ex.Number));
                    }
                    catch (Exception ex)
                    {
                        LogMsg(p_Name, MessageLevel.msgError, "Unable to set a movement rate of zero - " + ex.Message);
                    }

                    Status(StatusType.staAction, "Set lower rate");
                    // Test that error is generated on attempt to set rate lower than minimum
                    try
                    {
                        if (l_RateMinimum > 0d) // choose a value between the minimum and zero
                        {
                            l_MoveRate = l_RateMinimum / 2.0d;
                        }
                        else // Choose a large negative value
                        {
                            l_MoveRate = -l_RateMaximum - 1.0d;
                        }

                        LogMsg(p_Name, MessageLevel.msgDebug, "Using minimum rate: " + l_MoveRate);
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + l_MoveRate);
                        telescopeDevice.MoveAxis(p_Axis, l_MoveRate); // Set a value lower than the minimum
                        LogMsg(p_Name, MessageLevel.msgIssue, "No exception raised when move axis value < minimum rate: " + l_MoveRate);
                        // Clean up and release each object after use
                        try
                        {
                            Marshal.ReleaseComObject(l_Rate);
                        }
                        catch
                        {
                        }

                        l_Rate = null;
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when move axis is set below lowest rate (" + l_MoveRate + ")", "Exception correctly generated when move axis is set below lowest rate (" + l_MoveRate + ")");
                    }
                    // Clean up and release each object after use
                    try
                    {
                        Marshal.ReleaseComObject(l_Rate);
                    }
                    catch
                    {
                    }

                    l_Rate = null;
                    if (TestStop())
                        return;

                    // test that error is generated when rate is above maximum set
                    Status(StatusType.staAction, "Set upper rate");
                    try
                    {
                        l_MoveRate = l_RateMaximum + 1.0d;
                        LogMsg(p_Name, MessageLevel.msgDebug, "Using maximum rate: " + l_MoveRate);
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + l_MoveRate);
                        telescopeDevice.MoveAxis(p_Axis, l_MoveRate); // Set a value higher than the maximum
                        LogMsg(p_Name, MessageLevel.msgIssue, "No exception raised when move axis value > maximum rate: " + l_MoveRate);
                        // Clean up and release each object after use
                        try
                        {
                            Marshal.ReleaseComObject(l_Rate);
                        }
                        catch
                        {
                        }

                        l_Rate = null;
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when move axis is set above highest rate (" + l_MoveRate + ")", "Exception correctly generated when move axis is set above highest rate (" + l_MoveRate + ")");
                    }
                    // Clean up and release each object after use
                    try
                    {
                        Marshal.ReleaseComObject(l_Rate);
                    }
                    catch
                    {
                    }

                    l_Rate = null;
                    if (TestStop())
                        return;
                    if (l_CanSetZero) // Can set a rate of zero so undertake these tests
                    {
                        // Confirm that lowest tracking rate can be set
                        Status(StatusType.staAction, "Move at minimum rate");
                        if (l_RateMinimum != double.PositiveInfinity) // Valid value found so try and set it
                        {
                            try
                            {
                                Status(StatusType.staStatus, "Moving forward");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + l_RateMinimum);
                                telescopeDevice.MoveAxis(p_Axis, l_RateMinimum); // Set the minimum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (TestStop())
                                    return;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                Status(StatusType.staStatus, "Moving back");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + -l_RateMinimum);
                                telescopeDevice.MoveAxis(p_Axis, -l_RateMinimum); // Set the minimum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (TestStop())
                                    return;
                                // v1.0.12 Next line added because movement wasn't stopped
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                LogMsg(p_Name, MessageLevel.msgOK, "Successfully moved axis at minimum rate: " + l_RateMinimum);
                            }
                            catch (Exception ex)
                            {
                                HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when setting rate: " + l_RateMinimum);
                            }

                            Status(StatusType.staStatus, ""); // Clear status flag
                        }
                        else // No valid rate was found so print an error
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Minimum rate test - unable to find lowest axis rate");
                        }

                        if (TestStop())
                            return;

                        // Confirm that highest tracking rate can be set
                        Status(StatusType.staAction, "Move at maximum rate");
                        if (l_RateMaximum != double.NegativeInfinity) // Valid value found so try and set it
                        {
                            try
                            {
                                // Confirm not slewing first
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property");
                                if (Conversions.ToBoolean(telescopeDevice.Slewing))
                                {
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Slewing was true before start of MoveAxis but should have been false, remaining tests skipped");
                                    return;
                                }

                                Status(StatusType.staStatus, "Moving forward");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + l_RateMaximum);
                                telescopeDevice.MoveAxis(p_Axis, l_RateMaximum); // Set the minimum rate
                                                                                 // Confirm that slewing is active when the move is underway
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property");
                                if (Conversions.ToBoolean(!telescopeDevice.Slewing))
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true immediately after axis starts moving in positive direction");
                                WaitFor(MOVE_AXIS_TIME);
                                if (TestStop())
                                    return;
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property");
                                if (Conversions.ToBoolean(!telescopeDevice.Slewing))
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true after " + MOVE_AXIS_TIME / 1000d + " seconds moving in positive direction");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                                                        // Confirm that slewing is false when movement is stopped
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get property");
                                if (Conversions.ToBoolean(telescopeDevice.Slewing))
                                {
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Slewing incorrectly remains true after stopping positive axis movement, remaining test skipped");
                                    return;
                                }

                                Status(StatusType.staStatus, "Moving back");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + -l_RateMaximum);
                                telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum); // Set the minimum rate
                                                                                  // Confirm that slewing is active when the move is underway
                                if (Conversions.ToBoolean(!telescopeDevice.Slewing))
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true immediately after axis starts moving in negative direction");
                                WaitFor(MOVE_AXIS_TIME);
                                if (TestStop())
                                    return;
                                if (Conversions.ToBoolean(!telescopeDevice.Slewing))
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Slewing is not true after " + MOVE_AXIS_TIME / 1000d + " seconds moving in negative direction");
                                // Confirm that slewing is false when movement is stopped
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                                                        // Confirm that slewing is false when movement is stopped
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Slewing property");
                                if (Conversions.ToBoolean(telescopeDevice.Slewing))
                                {
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Slewing incorrectly remains true after stopping negative axis movement, remaining test skipped");
                                    return;
                                }

                                LogMsg(p_Name, MessageLevel.msgOK, "Successfully moved axis at maximum rate: " + l_RateMaximum);
                            }
                            catch (Exception ex)
                            {
                                HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "when setting rate: " + l_RateMaximum);
                            }

                            Status(StatusType.staStatus, ""); // Clear status flag
                        }
                        else // No valid rate was found so print an error
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Maximum rate test - unable to find lowest axis rate");
                        }

                        if (TestStop())
                            return;

                        // Confirm that tracking state is correctly restored after a move axis command
                        try
                        {
                            Status(StatusType.staAction, "Tracking state restore");
                            if (canSetTracking)
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                                l_TrackingStart = telescopeDevice.Tracking; // Save the start tracking state
                                Status(StatusType.staStatus, "Moving forward");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + l_RateMaximum);
                                telescopeDevice.MoveAxis(p_Axis, l_RateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (TestStop())
                                    return;
                                Status(StatusType.staStatus, "Stop movement");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                                l_TrackingEnd = telescopeDevice.Tracking; // Save the final tracking state
                                if (l_TrackingStart == l_TrackingEnd) // Successfully retained tracking state
                                {
                                    if (l_TrackingStart) // Tracking is true so switch to false for return movement
                                    {
                                        Status(StatusType.staStatus, "Set tracking off");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property false");
                                        telescopeDevice.Tracking = false;
                                        Status(StatusType.staStatus, "Move back");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + -l_RateMaximum);
                                        telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum); // Set the maximum rate
                                        WaitFor(MOVE_AXIS_TIME);
                                        if (TestStop())
                                            return;
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                        telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                        Status(StatusType.staStatus, "");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property false");
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(telescopeDevice.Tracking, false, false))) // tracking correctly retained in both states
                                        {
                                            LogMsg(p_Name, MessageLevel.msgOK, "Tracking state correctly retained for both tracking states");
                                        }
                                        else
                                        {
                                            LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state correctly retained when tracking is " + l_TrackingStart.ToString() + ", but not when tracking is false");
                                        }
                                    }
                                    else // Tracking false so switch to true for return movement
                                    {
                                        Status(StatusType.staStatus, "Set tracking on");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property true");
                                        telescopeDevice.Tracking = true;
                                        Status(StatusType.staStatus, "Move back");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + -l_RateMaximum);
                                        telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum); // Set the maximum rate
                                        WaitFor(MOVE_AXIS_TIME);
                                        if (TestStop())
                                            return;
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                        telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                        Status(StatusType.staStatus, "");
                                        if (g_Settings.DisplayMethodCalls)
                                            LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(telescopeDevice.Tracking, true, false))) // tracking correctly retained in both states
                                        {
                                            LogMsg(p_Name, MessageLevel.msgOK, "Tracking state correctly retained for both tracking states");
                                        }
                                        else
                                        {
                                            LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state correctly retained when tracking is " + l_TrackingStart.ToString() + ", but not when tracking is true");
                                        }
                                    }

                                    Status(StatusType.staStatus, ""); // Clear status flag
                                }
                                else // Tracking state not correctly restored
                                {
                                    Status(StatusType.staStatus, "Move back");
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + -l_RateMaximum);
                                    telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum); // Set the maximum rate
                                    WaitFor(MOVE_AXIS_TIME);
                                    if (TestStop())
                                        return;
                                    Status(StatusType.staStatus, "");
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                    telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property " + l_TrackingStart);
                                    telescopeDevice.Tracking = l_TrackingStart; // Restore original value
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state not correctly restored after MoveAxis when CanSetTracking is true");
                                }
                            }
                            else // Can't set tracking so just test the current state
                            {
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                                l_TrackingStart = telescopeDevice.Tracking;
                                Status(StatusType.staStatus, "Moving forward");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed " + l_RateMaximum);
                                telescopeDevice.MoveAxis(p_Axis, l_RateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (TestStop())
                                    return;
                                Status(StatusType.staStatus, "Stop movement");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to get Tracking property");
                                l_TrackingEnd = telescopeDevice.Tracking; // Save tracking state
                                Status(StatusType.staStatus, "Move back");
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call method MoveAxis for axis " + ((int)p_Axis).ToString() + " at speed " + -l_RateMaximum);
                                telescopeDevice.MoveAxis(p_Axis, -l_RateMaximum); // Set the maximum rate
                                WaitFor(MOVE_AXIS_TIME);
                                if (TestStop())
                                    return;
                                // v1.0.12 next line added because movement wasn't stopped
                                if (g_Settings.DisplayMethodCalls)
                                    LogMsg(p_Name, MessageLevel.msgComment, "About to call MoveAxis method for axis " + ((int)p_Axis).ToString() + " at speed 0");
                                telescopeDevice.MoveAxis(p_Axis, 0.0d); // Stop the movement on this axis
                                if (l_TrackingStart == l_TrackingEnd)
                                {
                                    LogMsg(p_Name, MessageLevel.msgOK, "Tracking state correctly restored after MoveAxis when CanSetTracking is false");
                                }
                                else
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(p_Name, MessageLevel.msgComment, "About to set Tracking property to " + l_TrackingStart);
                                    telescopeDevice.Tracking = l_TrackingStart; // Restore correct value
                                    LogMsg(p_Name, MessageLevel.msgIssue, "Tracking state not correctly restored after MoveAxis when CanSetTracking is false");
                                }

                                Status(StatusType.staStatus, "");
                            } // Clear status flag
                        }
                        catch (Exception ex)
                        {
                            HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "");
                        }
                    }
                    else // Cant set zero so tests skipped 
                    {
                        LogMsg(p_Name, MessageLevel.msgInfo, "Remaining MoveAxis tests skipped because unable to set a movement rate of zero");
                    }

                    Status(StatusType.staStatus, ""); // Clear status flag
                    Status(StatusType.staAction, ""); // Clear action flag
                }
                else // Some problem in finding rates inside the AxisRates object
                {
                    LogMsg(p_Name, MessageLevel.msgInfo, "Found minimum rate: " + l_RateMinimum + " found maximum rate: " + l_RateMaximum);
                    LogMsg(p_Name, MessageLevel.msgError, Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(Operators.ConcatenateObject("Unable to determine lowest or highest rates, expected ", p_AxisRates.Count), " rates, found "), l_RateCount)));
                }
            }
            else
            {
                LogMsg(p_Name, MessageLevel.msgWarning, "MoveAxis tests skipped because there are no AxisRate values");
            }
        }

        private void SideOfPierTests()
        {
            SideOfPierResults l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9;
            double l_Declination3, l_Declination9, l_StartRA;

            // Slew to starting position
            LogMsg("SideofPier", MessageLevel.msgDebug, "Starting Side of Pier tests");
            Status(StatusType.staTest, "Side of pier tests");
            l_StartRA = TelescopeRAFromHourAngle("SideofPier", -3.0d);
            if (m_SiteLatitude > 0.0d) // We are in the northern hemisphere
            {
                l_Declination3 = 90.0d - (180.0d - m_SiteLatitude) * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR; // Calculate for northern hemisphere
                l_Declination9 = 90.0d - m_SiteLatitude * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR;
            }
            else // We are in the southern hemisphere
            {
                l_Declination3 = -90.0d + (180.0d + m_SiteLatitude) * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR; // Calculate for southern hemisphere
                l_Declination9 = -90.0d - m_SiteLatitude * SIDE_OF_PIER_TARGET_DECLINATION_ESTIMATOR;
            }

            LogMsg("SideofPier", MessageLevel.msgDebug, "Declination for hour angle = +-3.0 tests: " + FormatDec(l_Declination3) + ", Declination for hour angle = +-9.0 tests: " + FormatDec(l_Declination9));
            SlewScope(l_StartRA, 0.0d, "Move to starting position " + FormatRA(l_StartRA) + " " + FormatDec(0.0d));
            if (TestStop())
                return;

            // Run tests
            Status(StatusType.staAction, "Test hour angle -3.0 at declination: " + FormatDec(l_Declination3));
            l_PierSideMinus3 = SOPPierTest(l_StartRA, l_Declination3, "hour angle -3.0");
            if (TestStop())
                return;
            Status(StatusType.staAction, "Test hour angle +3.0 at declination: " + FormatDec(l_Declination3));
            l_PierSidePlus3 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +3.0d), l_Declination3, "hour angle +3.0");
            if (TestStop())
                return;
            Status(StatusType.staAction, "Test hour angle -9.0 at declination: " + FormatDec(l_Declination9));
            l_PierSideMinus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", -9.0d), l_Declination9, "hour angle -9.0");
            if (TestStop())
                return;
            Status(StatusType.staAction, "Test hour angle +9.0 at declination: " + FormatDec(l_Declination9));
            l_PierSidePlus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +9.0d), l_Declination9, "hour angle +9.0");
            if (TestStop())
                return;
            if (l_PierSideMinus3.SideOfPier == l_PierSidePlus9.SideOfPier & l_PierSidePlus3.SideOfPier == l_PierSideMinus9.SideOfPier) // Reporting physical pier side
            {
                LogMsg("SideofPier", MessageLevel.msgIssue, "SideofPier reports physical pier side rather than pointing state");
            }
            else if (l_PierSideMinus3.SideOfPier == l_PierSideMinus9.SideOfPier & l_PierSidePlus3.SideOfPier == l_PierSidePlus9.SideOfPier) // Make other tests
            {
                LogMsg("SideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected");
            }
            else // Don't know what this means!
            {
                LogMsg("SideofPier", MessageLevel.msgInfo, "Unknown SideofPier reporting model: HA-3: " + l_PierSideMinus3.SideOfPier.ToString() + " HA-9: " + l_PierSideMinus9.SideOfPier.ToString() + " HA+3: " + l_PierSidePlus3.SideOfPier.ToString() + " HA+9: " + l_PierSidePlus9.SideOfPier.ToString());
            }

            LogMsg("SideofPier", MessageLevel.msgInfo, "Reported SideofPier at HA -9, +9: " + TranslatePierSide((PierSide)l_PierSideMinus9.SideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus9.SideOfPier, false));
            LogMsg("SideofPier", MessageLevel.msgInfo, "Reported SideofPier at HA -3, +3: " + TranslatePierSide((PierSide)l_PierSideMinus3.SideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus3.SideOfPier, false));

            // Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east at hour angle -3
            if ((int)l_PierSideMinus3.SideOfPier == (int)PierSide.pierWest)
            {
                LogMsg("SideofPier", MessageLevel.msgOK, "pierWest is returned when the mount is observing at an hour angle between -6.0 and 0.0");
            }
            else
            {
                LogMsg("SideofPier", MessageLevel.msgIssue, "pierEast is returned when the mount is observing at an hour angle between -6.0 and 0.0");
                LogMsg("SideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.");
            }

            if (l_PierSidePlus3.SideOfPier == (int)PierSide.pierEast)
            {
                LogMsg("SideofPier", MessageLevel.msgOK, "pierEast is returned when the mount is observing at an hour angle between 0.0 and +6.0");
            }
            else
            {
                LogMsg("SideofPier", MessageLevel.msgIssue, "pierWest is returned when the mount is observing at an hour angle between 0.0 and +6.0");
                LogMsg("SideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when observing at hour angles from -6.0 to -0.0 and that pierEast must be returned at hour angles from 0.0 to +6.0.");
            }

            // Test whether DestinationSideOfPier is implemented
            if ((int)l_PierSideMinus3.DestinationSideOfPier == (int)PierSide.pierUnknown & (int)l_PierSideMinus9.DestinationSideOfPier == (int)PierSide.pierUnknown & (int)l_PierSidePlus3.DestinationSideOfPier == (int)PierSide.pierUnknown & (int)l_PierSidePlus9.DestinationSideOfPier == (int)PierSide.pierUnknown)
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Analysis skipped as this method is not implemented"); // Not implemented
            }
            else // It is implemented so assess the results
            {
                if (l_PierSideMinus3.DestinationSideOfPier == l_PierSidePlus9.DestinationSideOfPier & l_PierSidePlus3.DestinationSideOfPier == l_PierSideMinus9.DestinationSideOfPier) // Reporting physical pier side
                {
                    LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "DestinationSideofPier reports physical pier side rather than pointing state");
                }
                else if (l_PierSideMinus3.DestinationSideOfPier == l_PierSideMinus9.DestinationSideOfPier & l_PierSidePlus3.DestinationSideOfPier == l_PierSidePlus9.DestinationSideOfPier) // Make other tests
                {
                    LogMsg("DestinationSideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected");
                }
                else // Don't know what this means!
                {
                    LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Unknown DestinationSideofPier reporting model: HA-3: " + l_PierSideMinus3.SideOfPier.ToString() + " HA-9: " + l_PierSideMinus9.SideOfPier.ToString() + " HA+3: " + l_PierSidePlus3.SideOfPier.ToString() + " HA+9: " + l_PierSidePlus9.SideOfPier.ToString());
                }

                // Now test the ASCOM convention that pierWest is returned when the mount is on the west side of the pier facing east at hour angle -3
                if ((int)l_PierSideMinus3.DestinationSideOfPier == (int)PierSide.pierWest)
                {
                    LogMsg("DestinationSideofPier", MessageLevel.msgOK, "pierWest is returned when the mount will observe at an hour angle between -6.0 and 0.0");
                }
                else
                {
                    LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "pierEast is returned when the mount will observe at an hour angle between -6.0 and 0.0");
                    LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.");
                }

                if (l_PierSidePlus3.DestinationSideOfPier == (int)PierSide.pierEast)
                {
                    LogMsg("DestinationSideofPier", MessageLevel.msgOK, "pierEast is returned when the mount will observe at an hour angle between 0.0 and +6.0");
                }
                else
                {
                    LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "pierWest is returned when the mount will observe at an hour angle between 0.0 and +6.0");
                    LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "ASCOM has adopted a convention that, for German Equatorial mounts, pierWest must be returned when the mount will observe at hour angles from -6.0 to -0.0 and that pierEast must be returned for hour angles from 0.0 to +6.0.");
                }
            }

            LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Reported DesintationSideofPier at HA -9, +9: " + TranslatePierSide((PierSide)l_PierSideMinus9.DestinationSideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus9.DestinationSideOfPier, false));
            LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Reported DesintationSideofPier at HA -3, +3: " + TranslatePierSide((PierSide)l_PierSideMinus3.DestinationSideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus3.DestinationSideOfPier, false));

            // Clean up
            // 3.0.0.12 added conditional test to next line
            if (canSetTracking)
                telescopeDevice.Tracking = false;
            Status(StatusType.staStatus, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staTest, "");
        }

        public SideOfPierResults SOPPierTest(double p_RA, double p_DEC, string p_Msg)
        {
            // Determine side of pier and destination side of pier results for a particular RA and DEC
            var l_Results = new SideOfPierResults(); // Create result set object
            double l_StartRA, l_StartDEC;
            try
            {
                // Prepare for tests
                l_StartRA = Conversions.ToDouble(telescopeDevice.RightAscension);
                l_StartDEC = Conversions.ToDouble(telescopeDevice.Declination);

                // Do destination side of pier test to see what side of pier we should end up on
                LogMsg("", MessageLevel.msgDebug, "");
                LogMsg("SOPPierTest", MessageLevel.msgDebug, "Testing RA DEC: " + FormatRA(p_RA) + " " + FormatDec(p_DEC) + " Current pierSide: " + TranslatePierSide((PierSide)telescopeDevice.SideOfPier, true));
                try
                {
                    l_Results.DestinationSideOfPier = (ASCOM.Interface.PierSide)telescopeDevice.DestinationSideOfPier(p_RA, p_DEC);
                    LogMsg("SOPPierTest", MessageLevel.msgDebug, "Target DestinationSideOfPier: " + l_Results.DestinationSideOfPier.ToString());
                }
                catch (COMException ex)
                {
                    switch (ex.ErrorCode)
                    {
                        case var @case when @case == ErrorCodes.NotImplemented:
                            {
                                l_Results.DestinationSideOfPier = (ASCOM.Interface.PierSide)PierSide.pierUnknown;
                                LogMsg("SOPPierTest", MessageLevel.msgDebug, "COM DestinationSideOfPier is not implemented setting result to: " + l_Results.DestinationSideOfPier.ToString());
                                break;
                            }

                        default:
                            {
                                LogMsg("SOPPierTest", MessageLevel.msgError, "COM DestinationSideOfPier Exception: " + ex.ToString());
                                break;
                            }
                    }
                }
                catch (MethodNotImplementedException) // DestinationSideOfPier not available so mark as unknown
                {
                    l_Results.DestinationSideOfPier = (ASCOM.Interface.PierSide)PierSide.pierUnknown;
                    LogMsg("SOPPierTest", MessageLevel.msgDebug, ".NET DestinationSideOfPier is not implemented setting result to: " + l_Results.DestinationSideOfPier.ToString());
                }
                catch (Exception ex)
                {
                    LogMsg("SOPPierTest", MessageLevel.msgError, ".NET DestinationSideOfPier Exception: " + ex.ToString());
                }
                // Now do an actual slew and record side of pier we actually get
                SlewScope(p_RA, p_DEC, "Testing " + p_Msg + ", co-ordinates: " + FormatRA(p_RA) + " " + FormatDec(p_DEC));
                l_Results.SideOfPier = (ASCOM.Interface.PierSide)telescopeDevice.SideOfPier;
                LogMsg("SOPPierTest", MessageLevel.msgDebug, "Actual SideOfPier: " + l_Results.SideOfPier.ToString());

                // Return to original RA
                SlewScope(l_StartRA, l_StartDEC, "Returning to start point");
                LogMsg("SOPPierTest", MessageLevel.msgDebug, "Returned to: " + FormatRA(l_StartRA) + " " + FormatDec(l_StartDEC));
            }
            catch (Exception ex)
            {
                LogMsg("SOPPierTest", MessageLevel.msgError, "SideofPierException: " + ex.ToString());
            }

            return l_Results;
        }

        private void DestinationSideOfPierTests()
        {
            PierSide l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9;

            // Slew to one position, then call destination side of pier 4 times and report the pattern
            SlewScope(TelescopeRAFromHourAngle("DestinationSideofPier", -3.0d), 0.0d, "Slew to start position");
            l_PierSideMinus3 = (PierSide)telescopeDevice.DestinationSideOfPier(-3.0d, 0.0d);
            l_PierSidePlus3 = (PierSide)telescopeDevice.DestinationSideOfPier(3.0d, 0.0d);
            l_PierSideMinus9 = (PierSide)telescopeDevice.DestinationSideOfPier(-9.0d, 90.0d - m_SiteLatitude);
            l_PierSidePlus9 = (PierSide)telescopeDevice.DestinationSideOfPier(9.0d, 90.0d - m_SiteLatitude);
            if (l_PierSideMinus3 == l_PierSidePlus9 & l_PierSidePlus3 == l_PierSideMinus9) // Reporting physical pier side
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "The driver appears to be reporting physical pier side rather than pointing state");
            }
            else if (l_PierSideMinus3 == l_PierSideMinus9 & l_PierSidePlus3 == l_PierSidePlus9) // Make other tests
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgOK, "The driver reports the pointing state of the mount");
            }
            else // Don't know what this means!
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Unknown pier side reporting model: HA-3: " + l_PierSideMinus3.ToString() + " HA-9: " + l_PierSideMinus9.ToString() + " HA+3: " + l_PierSidePlus3.ToString() + " HA+9: " + l_PierSidePlus9.ToString());
            }

            telescopeDevice.Tracking = false;
            LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus9, false) + TranslatePierSide(l_PierSidePlus9, false));
            LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide(l_PierSideMinus3, false) + TranslatePierSide(l_PierSidePlus3, false));
        }

        #endregion

        #region Support Code

        private void CheckScopePosition(string testName, string functionName, double expectedRA, double expectedDec)
        {
            double actualRA, actualDec, difference;
            if (g_Settings.DisplayMethodCalls)
                LogMsg(testName, MessageLevel.msgComment, "About to get RightAscension property");
            actualRA = Conversions.ToDouble(telescopeDevice.RightAscension);
            LogMsg(testName, MessageLevel.msgDebug, "Read RightAscension: " + FormatRA(actualRA));
            if (g_Settings.DisplayMethodCalls)
                LogMsg(testName, MessageLevel.msgComment, "About to get Declination property");
            actualDec = Conversions.ToDouble(telescopeDevice.Declination);
            LogMsg(testName, MessageLevel.msgDebug, "Read Declination: " + FormatDec(actualDec));

            // Check that we have actually arrived where we are expected to be
            difference = RaDifferenceInSeconds(actualRA, expectedRA);
            switch (difference)
            {
                case var @case when @case <= SLEW_SYNC_OK_TOLERANCE:  // Convert arc seconds to hours of RA
                    {
                        LogMsg(testName, MessageLevel.msgOK, string.Format("{0} OK. RA:   {1}", functionName, FormatRA(expectedRA)));
                        break;
                    }

                default:
                    {
                        LogMsg(testName, MessageLevel.msgInfo, string.Format("{0} within {1} arc seconds of expected RA: {2}, actual RA: {3}", functionName, difference.ToString("0.0"), FormatRA(expectedRA), FormatRA(actualRA)));
                        break;
                    }
            }

            difference = Math.Round(Math.Abs(actualDec - expectedDec) * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // Dec difference is in arc seconds from degrees of Declination
            switch (difference)
            {
                case var case1 when case1 <= SLEW_SYNC_OK_TOLERANCE:
                    {
                        LogMsg(testName, MessageLevel.msgOK, string.Format("{0} OK. DEC: {1}", functionName, FormatDec(expectedDec)));
                        break;
                    }

                default:
                    {
                        LogMsg(testName, MessageLevel.msgInfo, string.Format("{0} within {1} arc seconds of expected DEC: {2}, actual DEC: {3}", functionName, difference.ToString("0.0"), FormatDec(expectedDec), FormatDec(actualDec)));
                        break;
                    }
            }
        }

        /// <summary>
        /// Return the difference between two RAs (in hours) as seconds
        /// </summary>
        /// <param name="FirstRA">First RA (hours)</param>
        /// <param name="SecondRA">Second RA (hours)</param>
        /// <returns>Difference (seconds) between the supplied RAs</returns>
        private double RaDifferenceInSeconds(double FirstRA, double SecondRA)
        {
            double RaDifferenceInSecondsRet = default;
            RaDifferenceInSecondsRet = Math.Abs(FirstRA - SecondRA); // Calculate the difference allowing for negative outcomes
            if (RaDifferenceInSecondsRet > 12.0d)
                RaDifferenceInSecondsRet = 24.0d - RaDifferenceInSecondsRet; // Deal with the cases where the two elements are more than 12 hours apart going in the initial direction
            RaDifferenceInSecondsRet = Math.Round(RaDifferenceInSecondsRet * 15.0d * 60.0d * 60.0d, 1, MidpointRounding.AwayFromZero); // RA difference is in arc seconds from hours of RA
            return RaDifferenceInSecondsRet;
        }

        private void SyncScope(string testName, string canDoItName, SlewSyncType testType, double syncRA, double syncDec)
        {
            switch (testType)
            {
                case SlewSyncType.SyncToCoordinates: // SyncToCoordinates
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(testName, MessageLevel.msgComment, "About to call SyncToCoordinates method, RA: " + FormatRA(syncRA) + ", Declination: " + FormatDec(syncDec));
                        telescopeDevice.SyncToCoordinates(syncRA, syncDec); // Sync to slightly different coordinates
                        LogMsg(testName, MessageLevel.msgDebug, "Completed SyncToCoordinates");
                        break;
                    }

                case SlewSyncType.SyncToTarget: // SyncToTarget
                    {
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(testName, MessageLevel.msgComment, "About to get Tracking property");
                        if (Conversions.ToBoolean(Operators.AndObject(canSetTracking, !telescopeDevice.Tracking)))
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(testName, MessageLevel.msgComment, "About to set Tracking property to true");
                            telescopeDevice.Tracking = true;
                        }

                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(testName, MessageLevel.msgComment, "About to set TargetRightAscension property to " + FormatRA(syncRA));
                            telescopeDevice.TargetRightAscension = syncRA;
                            LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetRightAscension");
                        }
                        catch (Exception ex)
                        {
                            HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, canDoItName + " is True but can't set TargetRightAscension");
                        }

                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg(testName, MessageLevel.msgComment, "About to set TargetDeclination property to " + FormatDec(syncDec));
                            telescopeDevice.TargetDeclination = syncDec;
                            LogMsg(testName, MessageLevel.msgDebug, "Completed Set TargetDeclination");
                        }
                        catch (Exception ex)
                        {
                            HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, canDoItName + " is True but can't set TargetDeclination");
                        }

                        if (g_Settings.DisplayMethodCalls)
                            LogMsg(testName, MessageLevel.msgComment, "About to call SyncToTarget method");
                        telescopeDevice.SyncToTarget(); // Sync to slightly different coordinates
                        LogMsg(testName, MessageLevel.msgDebug, "Completed SyncToTarget");
                        break;
                    }

                default:
                    {
                        LogMsg(testName, MessageLevel.msgError, "Conform:SyncTest: Unknown test type " + testType.ToString());
                        break;
                    }
            }
        }

        public void SlewScope(double p_RA, double p_DEC, string p_Msg)
        {
            if (canSetTracking)
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SlewScope", MessageLevel.msgComment, "About to set Tracking property to true");
                telescopeDevice.Tracking = true;
            }

            Status(StatusType.staAction, p_Msg);
            if (canSlew)
            {
                if (canSlewAsync)
                {
                    LogMsg("SlewScope", MessageLevel.msgDebug, "Slewing asynchronously to " + p_Msg + " " + FormatRA(p_RA) + " " + FormatDec(p_DEC));
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("SlewScope", MessageLevel.msgComment, "About to call SlewToCoordinatesAsync method, RA: " + FormatRA(p_RA) + ", Declination: " + FormatDec(p_DEC));
                    telescopeDevice.SlewToCoordinatesAsync(p_RA, p_DEC);
                    WaitForSlew("SlewScope");
                }
                else
                {
                    LogMsg("SlewScope", MessageLevel.msgDebug, "Slewing synchronously to " + p_Msg + " " + FormatRA(p_RA) + " " + FormatDec(p_DEC));
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("SlewScope", MessageLevel.msgComment, "About to call SlewToCoordinates method, RA: " + FormatRA(p_RA) + ", Declination: " + FormatDec(p_DEC));
                    telescopeDevice.SlewToCoordinates(p_RA, p_DEC);
                }

                if (m_CanReadSideOfPier)
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("SlewScope", MessageLevel.msgComment, "About to get SideOfPier property");
                    LogMsg("SlewScope", MessageLevel.msgDebug, "SideOfPier: " + telescopeDevice.SideOfPier.ToString());
                }
            }
            else
            {
                LogMsg("SlewScope", MessageLevel.msgInfo, "Unable to slew this scope as CanSlew is false, slew omitted");
            }

            Status(StatusType.staAction, "");
        }

        private void WaitForSlew(string testName)
        {
            DateTime WaitStartTime;
            WaitStartTime = DateAndTime.Now;
            do
            {
                WaitFor(SLEEP_TIME);
                My.MyProject.Forms.FrmConformMain.staStatus.Text = "Slewing";
                Application.DoEvents();
                if (g_Settings.DisplayMethodCalls)
                    LogMsg(testName, MessageLevel.msgComment, "About to get Slewing property");
            }
            while (telescopeDevice.Slewing & DateAndTime.Now.Subtract(WaitStartTime).TotalSeconds < (double)WAIT_FOR_SLEW_MINIMUM_DURATION & TestStop());
            My.MyProject.Forms.FrmConformMain.staStatus.Text = "Slew completed";
        }

        private double TelescopeRAFromHourAngle(string testName, double p_Offset)
        {
            double TelescopeRAFromHourAngleRet = default;

            // Handle the possibility that the mandatory SideealTime property has not been implemented
            if (canReadSiderealTime)
            {
                // Create a legal RA based on an offset from Sidereal time
                if (g_Settings.DisplayMethodCalls)
                    LogMsg(testName, MessageLevel.msgComment, "About to get SiderealTime property");
                TelescopeRAFromHourAngleRet = Conversions.ToDouble(Operators.SubtractObject(telescopeDevice.SiderealTime, p_Offset));
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
            double TelescopeRAFromSiderealTimeRet = default;
            double CurrentSiderealTime;

            // Handle the possibility that the mandatory SideealTime property has not been implemented
            if (canReadSiderealTime)
            {
                // Create a legal RA based on an offset from Sidereal time
                if (g_Settings.DisplayMethodCalls)
                    LogMsg(testName, MessageLevel.msgComment, "About to get SiderealTime property");
                CurrentSiderealTime = Conversions.ToDouble(telescopeDevice.SiderealTime);
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
                            TelescopeRAFromSiderealTimeRet = TelescopeRAFromSiderealTimeRet + 24.0d;
                            break;
                        }

                    case var case3 when case3 >= 24.0d: // Illegal if > 24 hours
                        {
                            TelescopeRAFromSiderealTimeRet = TelescopeRAFromSiderealTimeRet - 24.0d;
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

        private void TestEarlyBinding(InterfaceType TestType)
        {
            dynamic l_ITelescope;
            dynamic l_DeviceObject = null;
            string l_ErrMsg = "";
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
                        if (g_Settings.DisplayMethodCalls)
                            LogMsg("AccessChecks", MessageLevel.msgComment, "About to create driver object with CreateObject");
                        LogMsg("AccessChecks", MessageLevel.msgDebug, "Creating late bound object for interface test");
                        l_DeviceObject = Interaction.CreateObject(g_TelescopeProgID);
                        LogMsg("AccessChecks", MessageLevel.msgDebug, "Created late bound object OK");
                        switch (TestType)
                        {
                            case InterfaceType.ITelescopeV2:
                                {
                                    l_ITelescope = (ASCOM.Interface.ITelescope)l_DeviceObject;
                                    break;
                                }

                            case InterfaceType.ITelescopeV3:
                                {
                                    l_ITelescope = (ITelescopeV3)l_DeviceObject;
                                    break;
                                }

                            default:
                                {
                                    LogMsg("TestEarlyBinding", MessageLevel.msgError, "Unknown interface type: " + TestType.ToString());
                                    break;
                                }
                        }

                        LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver with interface " + TestType.ToString());
                        try
                        {
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property true");
                            l_ITelescope.Connected = true;
                            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device exposes interface " + TestType.ToString());
                            if (g_Settings.DisplayMethodCalls)
                                LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property false");
                            l_ITelescope.Connected = false;
                        }
                        catch (Exception)
                        {
                            LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface " + TestType.ToString());
                            LogMsg("", MessageLevel.msgAlways, "");
                        }
                    }
                    catch (Exception ex)
                    {
                        l_ErrMsg = ex.ToString();
                        LogMsg("AccessChecks", MessageLevel.msgDebug, "Exception: " + ex.Message);
                    }

                    if (l_DeviceObject is null)
                        WaitFor(200);
                }
                while (!(l_TryCount == 3 | l_ITelescope is object)); // Exit if created OK
                if (l_ITelescope is null)
                {
                    LogMsg("AccessChecks", MessageLevel.msgInfo, "Device does not expose interface " + TestType.ToString());
                }
                else
                {
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Created telescope on attempt: " + l_TryCount.ToString());
                }

                // Clean up
                try
                {
                    DisposeAndReleaseObject("Telescope V1", l_ITelescope);
                }
                catch
                {
                }

                try
                {
                    DisposeAndReleaseObject("Telescope V3", l_DeviceObject);
                }
                catch
                {
                }

                l_DeviceObject = null;
                l_ITelescope = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                WaitForAbsolute(DEVICE_DESTROY_WAIT, "TestEarlyBinding waiting for Telescope Object to Dispose");
            }
            catch (Exception ex)
            {
                LogMsg("Telescope:TestEarlyBinding.EX1", MessageLevel.msgError, ex.ToString());
            }
        }

        private string FormatRA(double ra)
        {
            return g_Util.HoursToHMS(ra, ":", ":", "", DISPLAY_DECIMAL_DIGITS);
        }

        private string FormatDec(double Dec)
        {
            return g_Util.DegreesToDMS(Dec, ":", ":", "", DISPLAY_DECIMAL_DIGITS).PadLeft(Conversions.ToInteger(Operators.AddObject(9, Interaction.IIf(DISPLAY_DECIMAL_DIGITS > 0, DISPLAY_DECIMAL_DIGITS + 1, 0))));
        }

        private dynamic FormatAltitude(double Alt)
        {
            return g_Util.DegreesToDMS(Alt, ":", ":", "", DISPLAY_DECIMAL_DIGITS);
        }

        private string FormatAzimuth(double Az)
        {
            return g_Util.DegreesToDMS(Az, ":", ":", "", DISPLAY_DECIMAL_DIGITS).PadLeft(Conversions.ToInteger(Operators.AddObject(9, Interaction.IIf(DISPLAY_DECIMAL_DIGITS > 0, DISPLAY_DECIMAL_DIGITS + 1, 0))));
        }

        public string TranslatePierSide(PierSide p_PierSide, bool p_Long)
        {
            string l_PierSide;
            switch (p_PierSide)
            {
                case PierSide.pierEast:
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

                case PierSide.pierWest:
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

        private bool TestRADecRate(string TestName, string Description, Axis Axis, double Rate, bool SkipSlewiingTest)
        {
            bool success = false;
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg(TestName, MessageLevel.msgComment, string.Format("{0} - About to get Slewing property", Description));
                m_Slewing = telescopeDevice.Slewing;
                if (!m_Slewing | SkipSlewiingTest) // Slewing should be false at this point or we are ignoring the test!
                {
                    // Check that we can set the rate to a non-zero value
                    try
                    {
                        switch (Axis)
                        {
                            case Axis.RA:
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(TestName, MessageLevel.msgComment, string.Format("{0} - About to set RightAscensionRate property to {1}", Description, Rate));
                                    telescopeDevice.RightAscensionRate = Rate;
                                    SetStatus(string.Format("Watling for mount to settle after setting RightAcensionRate to {0}", Rate), "", "");
                                    WaitFor(2000); // Give a short wait to allow the mount to settle

                                    // Value set OK, now check that the new rate is returned by RightAscensionRate Get and that Slewing is false
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(TestName, MessageLevel.msgComment, string.Format("{0} - About to get RightAscensionRate property", Description));
                                    m_RightAscensionRate = Conversions.ToDouble(telescopeDevice.RightAscensionRate);
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(TestName, MessageLevel.msgComment, string.Format("{0} - About to get Slewing property", Description));
                                    m_Slewing = telescopeDevice.Slewing;
                                    if (m_RightAscensionRate == Rate & !m_Slewing)
                                    {
                                        LogMsg(TestName, MessageLevel.msgOK, string.Format("{0} - successfully set rate to {1}", Description, m_RightAscensionRate));
                                        success = true;
                                    }
                                    else
                                    {
                                        if (m_Slewing & m_RightAscensionRate == Rate)
                                            LogMsg(TestName, MessageLevel.msgError, string.Format("RightAscensionRate was successfully set to {0} but Slewing is returning True, it should return False.", Rate, m_RightAscensionRate));
                                        if (m_Slewing & m_RightAscensionRate != Rate)
                                            LogMsg(TestName, MessageLevel.msgError, string.Format("RightAscensionRate Read does not return {0} as set, instead it returns {1}. Slewing is also returning True, it should return False.", Rate, m_RightAscensionRate));
                                        if (!m_Slewing & m_RightAscensionRate != Rate)
                                            LogMsg(TestName, MessageLevel.msgError, string.Format("RightAscensionRate Read does not return {0} as set, instead it returns {1}.", Rate, m_RightAscensionRate));
                                    }

                                    break;
                                }

                            case Axis.Dec:
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(TestName, MessageLevel.msgComment, string.Format("{0} - About to set DeclinationRate property to {1}", Description, Rate));
                                    telescopeDevice.DeclinationRate = Rate;
                                    SetStatus(string.Format("Watling for mount to settle after setting DeclinationRate to {0}", Rate), "", "");
                                    WaitFor(2000); // Give a short wait to allow the mount to settle

                                    // Value set OK, now check that the new rate is returned by DeclinationRate Get and that Slewing is false
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(TestName, MessageLevel.msgComment, string.Format("{0} - About to get DeclinationRate property", Description));
                                    m_DeclinationRate = Conversions.ToDouble(telescopeDevice.DeclinationRate);
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg(TestName, MessageLevel.msgComment, string.Format("{0} - About to get Slewing property", Description));
                                    m_Slewing = telescopeDevice.Slewing;
                                    if (m_DeclinationRate == Rate & !m_Slewing)
                                    {
                                        LogMsg(TestName, MessageLevel.msgOK, string.Format("{0} - successfully set rate to {1}", Description, m_DeclinationRate));
                                        success = true;
                                    }
                                    else
                                    {
                                        if (m_Slewing & m_DeclinationRate == Rate)
                                            LogMsg(TestName, MessageLevel.msgError, string.Format("DeclinationRate was successfully set to {0} but Slewing is returning True, it should return False.", Rate, m_DeclinationRate));
                                        if (m_Slewing & m_DeclinationRate != Rate)
                                            LogMsg(TestName, MessageLevel.msgError, string.Format("DeclinationRate Read does not return {0} as set, instead it returns {1}. Slewing is also returning True, it should return False.", Rate, m_DeclinationRate));
                                        if (!m_Slewing & m_DeclinationRate != Rate)
                                            LogMsg(TestName, MessageLevel.msgError, string.Format("DeclinationRate Read does not return {0} as set, instead it returns {1}.", Rate, m_DeclinationRate));
                                    }

                                    break;
                                }

                            default:
                                {
                                    MessageBox.Show(string.Format("Conform internal error - Unknown Axis value: {0}", Axis.ToString()));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (IsInvalidOperationException(TestName, ex)) // We can't know what the valid range for this telescope is in advance so its possible that our test value will be rejected, if so just report this.
                        {
                            LogMsg(TestName, MessageLevel.msgInfo, string.Format("Unable to set test rate {0}, it was rejected as an invalid value.", Rate));
                        }
                        else
                        {
                            HandleException(TestName, MemberType.Property, Required.MustBeImplemented, ex, "CanSetRightAscensionRate is True");
                        }
                    }
                }
                else
                {
                    LogMsg(TestName, MessageLevel.msgError, string.Format("{0} - Telescope.Slewing should be False at the start of this test but is returning True, test abandoned", Description));
                }
            }
            catch (Exception ex)
            {
                HandleException(TestName, MemberType.Property, Required.MustBeImplemented, ex, "Tried to read Slewing property");
            }

            SetStatus("", "", "");
            return success;
        }

        #endregion

        #region Private Test Code

        protected override void SpecialTelescopeDestinationSideOfPier()
        {
            double l_FlipPointEast, l_FlipPointWest;
            CheckInitialise(); // Set up error codes
            CreateDevice(); // Create the telescope device
            Connected = true; // Connect
            PreRunCheck(); // Make sure we are ready to move the scope
            telescopeDevice.Tracking = true;
            if (g_Settings.DSOPSide == "pierEast only" | g_Settings.DSOPSide == "pierEast and pierWest")
            {
                l_FlipPointEast = FlipCheck(FlipTestType.DestinationSideOfPier, g_Settings.get_FlipTestDECStart(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestDECEnd(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestDECStep(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestHAStart(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestHAEnd(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierEast));
            }

            if (g_Settings.DSOPSide == "pierWest only" | g_Settings.DSOPSide == "pierEast and pierWest")
            {
                l_FlipPointWest = FlipCheck(FlipTestType.DestinationSideOfPier, g_Settings.get_FlipTestDECStart(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestDECEnd(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestDECStep(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestHAStart(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestHAEnd(SpecialTest.TelescopeDestinationSideOfPier, PierSide.pierWest));
            }

            PostRunCheck(); // Make the mount safe
            Connected = false; // Disconnect
        }

        protected override void SpecialTelescopeSideOfPier()
        {
            double l_FlipPointEast, l_FlipPointWest;
            CheckInitialise(); // Set up error codes
            CreateDevice(); // Create the telescope device
            Connected = true; // Connect
            PreRunCheck(); // Make sure we are ready to move the scope
            telescopeDevice.Tracking = true;
            if (g_Settings.DSOPSide == "pierEast only" | g_Settings.DSOPSide == "pierEast and pierWest")
            {
                l_FlipPointEast = FlipCheck(FlipTestType.SideOfPier, g_Settings.get_FlipTestDECStart(SpecialTest.TelescopeSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestDECEnd(SpecialTest.TelescopeSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestDECStep(SpecialTest.TelescopeSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestHAStart(SpecialTest.TelescopeSideOfPier, PierSide.pierEast), g_Settings.get_FlipTestHAEnd(SpecialTest.TelescopeSideOfPier, PierSide.pierEast));
            }

            if (g_Settings.DSOPSide == "pierWest only" | g_Settings.DSOPSide == "pierEast and pierWest")
            {
                l_FlipPointWest = FlipCheck(FlipTestType.SideOfPier, g_Settings.get_FlipTestDECStart(SpecialTest.TelescopeSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestDECEnd(SpecialTest.TelescopeSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestDECStep(SpecialTest.TelescopeSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestHAStart(SpecialTest.TelescopeSideOfPier, PierSide.pierWest), g_Settings.get_FlipTestHAEnd(SpecialTest.TelescopeSideOfPier, PierSide.pierWest));
            }

            PostRunCheck(); // Make the mount safe
            Connected = false; // Disconnect
        }

        protected override void SpecialTelescopeSideOfPierAnalysis()
        {
            CheckInitialise(); // Set up error codes
            CreateDevice(); // Create the telescope device
            Connected = true; // Connect
            PreRunCheck(); // Make sure we are ready to move the scope
            telescopeDevice.Tracking = true;
            canSlew = true;

            // SideOfPierTests()
            SideOfPierResults l_PierSideMinus3, l_PierSideMinus9, l_PierSidePlus3, l_PierSidePlus9;
            double l_Declination, l_StartRA;

            // Slew to starting position
            Status(StatusType.staTest, "Side of pier tests");
            l_StartRA = TelescopeRAFromHourAngle("SideofPier", -3.0d);
            l_Declination = 60.0d;
            SlewScope(l_StartRA, 0.0d, "Move to starting position " + FormatRA(l_StartRA) + " " + FormatDec(0.0d));
            if (TestStop())
                return;

            // Run tests
            Status(StatusType.staAction, "Test hour angle -3.0");
            l_PierSideMinus3 = SOPPierTest(l_StartRA, 10.0d, "hour angle -3.0 +");
            if (TestStop())
                return;
            l_PierSideMinus3 = SOPPierTest(l_StartRA, -10.0d, "hour angle -3.0 -");
            if (TestStop())
                return;
            Status(StatusType.staAction, "Test hour angle +3.0");
            l_PierSidePlus3 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +3.0d), 10.0d, "hour angle +3.0 +");
            if (TestStop())
                return;
            l_PierSidePlus3 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +3.0d), -10.0d, "hour angle +3.0 -");
            if (TestStop())
                return;
            Status(StatusType.staAction, "Test hour angle -9.0");
            l_PierSideMinus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", -9.0d), l_Declination, "hour angle -9.0");
            if (TestStop())
                return;
            Status(StatusType.staAction, "Test hour angle +9.0");
            l_PierSidePlus9 = SOPPierTest(TelescopeRAFromHourAngle("SideofPier", +9.0d), l_Declination, "hour angle +9.0");
            if (TestStop())
                return;
            if (l_PierSideMinus3.SideOfPier == l_PierSidePlus9.SideOfPier & l_PierSidePlus3.SideOfPier == l_PierSideMinus9.SideOfPier) // Reporting physical pier side
            {
                LogMsg("SideofPier", MessageLevel.msgIssue, "SideofPier reports physical pier side rather than pointing state");
                LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus9.SideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus9.SideOfPier, false));
                LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus3.SideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus3.SideOfPier, false));
            }
            else if (l_PierSideMinus3.SideOfPier == l_PierSideMinus9.SideOfPier & l_PierSidePlus3.SideOfPier == l_PierSidePlus9.SideOfPier) // Make other tests
            {
                LogMsg("SideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected");
            }
            else // Don't know what this means!
            {
                LogMsg("SideofPier", MessageLevel.msgInfo, "Unknown SideofPier reporting model: HA-3: " + l_PierSideMinus3.SideOfPier.ToString() + " HA-9: " + l_PierSideMinus9.SideOfPier.ToString() + " HA+3: " + l_PierSidePlus3.SideOfPier.ToString() + " HA+9: " + l_PierSidePlus9.SideOfPier.ToString());
                LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus9.SideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus9.SideOfPier, false));
                LogMsg("SideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus3.SideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus3.SideOfPier, false));
            }

            // Test whether DestinationSideOfPier is implemented
            if ((int)l_PierSideMinus3.DestinationSideOfPier == (int)PierSide.pierUnknown & (int)l_PierSideMinus9.DestinationSideOfPier == (int)PierSide.pierUnknown & (int)l_PierSidePlus3.DestinationSideOfPier == (int)PierSide.pierUnknown & (int)l_PierSidePlus9.DestinationSideOfPier == (int)PierSide.pierUnknown)
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Analysis skipped as this method is not implemented"); // Not implemented
            }
            else if (l_PierSideMinus3.DestinationSideOfPier == l_PierSidePlus9.DestinationSideOfPier & l_PierSidePlus3.DestinationSideOfPier == l_PierSideMinus9.DestinationSideOfPier) // It is implemented so assess the results
                                                                                                                                                                                        // Reporting physical pier side
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgIssue, "DestinationSideofPier reports physical pier side rather than pointing state");
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus9.DestinationSideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus9.DestinationSideOfPier, false));
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus3.DestinationSideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus3.DestinationSideOfPier, false));
            }
            else if (l_PierSideMinus3.DestinationSideOfPier == l_PierSideMinus9.DestinationSideOfPier & l_PierSidePlus3.DestinationSideOfPier == l_PierSidePlus9.DestinationSideOfPier) // Make other tests
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgOK, "Reports the pointing state of the mount as expected");
            }
            else // Don't know what this means!
            {
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, "Unknown DestinationSideofPier reporting model: HA-3: " + l_PierSideMinus3.SideOfPier.ToString() + " HA-9: " + l_PierSideMinus9.SideOfPier.ToString() + " HA+3: " + l_PierSidePlus3.SideOfPier.ToString() + " HA+9: " + l_PierSidePlus9.SideOfPier.ToString());
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus9.DestinationSideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus9.DestinationSideOfPier, false));
                LogMsg("DestinationSideofPier", MessageLevel.msgInfo, TranslatePierSide((PierSide)l_PierSideMinus3.DestinationSideOfPier, false) + TranslatePierSide((PierSide)l_PierSidePlus3.DestinationSideOfPier, false));
            }

            // Clean up
            // 3.0.0.12 added conditional test to next line
            if (canSetTracking)
                telescopeDevice.Tracking = false;
            Status(StatusType.staStatus, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staTest, "");
            PostRunCheck(); // Make the mount safe
            Connected = false; // Disconnected
        }

        protected override void SpecialTelescopeCommands()
        {
            string Response;
            CheckInitialise(); // Set up error codes
            CreateDevice(); // Create the telescope device
            Connected = true; // Connect
            PreRunCheck(); // Make sure we are ready to move the scope
            telescopeDevice.Tracking = true;

            // Command tests
            Response = Conversions.ToString(telescopeDevice.CommandString("Gc", false));
            LogMsg("TelescopeCommandTest", MessageLevel.msgInfo, "Sent Gc False - Received: " + Response);
            Response = Conversions.ToString(telescopeDevice.CommandString(":Gc#", true));
            LogMsg("TelescopeCommandTest", MessageLevel.msgInfo, "Sent :Gc# True - Received: " + Response);
            PostRunCheck(); // Make the mount safe
            Connected = false; // Disconnected
        }

        public double FlipCheck(FlipTestType p_FlipTestType, double p_DECStart, double p_DECEnd, double p_DECStep, double p_HAStart, double p_HAEnd)
        {
            double l_FlipPoint = default, l_TargetDEC;
            LogMsg("FlipCheck", MessageLevel.msgOK, "Test parameters - Start DEC: " + p_DECStart.ToString() + " End DEC: " + p_DECEnd.ToString() + " DEC Step size: " + p_DECStep.ToString() + " Start HA: " + p_HAStart + " End HA: " + p_HAEnd);
            LogMsg("", MessageLevel.msgAlways, "");
            var loopTo = p_DECEnd;
            for (l_TargetDEC = p_DECStart; p_DECStep >= 0 ? l_TargetDEC <= loopTo : l_TargetDEC >= loopTo; l_TargetDEC += p_DECStep)
            {
                if (!TestStop())
                    l_FlipPoint = TelescopeFindFlipPoint(p_FlipTestType, p_HAStart, p_HAEnd, l_TargetDEC);
                if (!TestStop())
                    LogMsg(p_FlipTestType.ToString(), MessageLevel.msgOK, "Hour angle flip point when scope is east of pier (HH:MM:SS):" + FormatRA(l_FlipPoint) + " at DEC: " + FormatDec(l_TargetDEC));
            }

            return l_FlipPoint;
        }

        public double TelescopeFindFlipPoint(FlipTestType p_TestType, double p_HARangeStart, double p_HARangeEnd, double p_TargetDEC)
        {
            // Returns the flip point returned either by SideofPier or DestinationSideOfPier
            PierSide l_SideOfPier = default, l_TargetSideOfPier = default;
            double l_TargetHA;
            int l_Iteration;
            double l_StartHARange, l_EndHARange, l_Temp;
            l_StartHARange = p_HARangeStart;
            l_TargetHA = p_HARangeEnd;
            l_EndHARange = p_HARangeEnd;

            // Use a halving method to determine where the flip point is
            for (l_Iteration = 1; l_Iteration <= 30; l_Iteration++)
            {
                // Move scope to requested starting position
                SlewScope(TelescopeRAFromHourAngle("FlipRange: " + l_SideOfPier.ToString(), p_HARangeStart), p_TargetDEC, "Moving to start position");
                l_SideOfPier = (PierSide)telescopeDevice.SideOfPier; // Save current side of pier as reference to check when a flip happens
                switch (p_TestType) // Carry out relevant test
                {
                    case FlipTestType.DestinationSideOfPier:
                        {
                            l_TargetSideOfPier = (PierSide)telescopeDevice.DestinationSideOfPier(TelescopeRAFromHourAngle("FlipRange: " + l_SideOfPier.ToString(), l_TargetHA), p_TargetDEC);
                            break;
                        }

                    case FlipTestType.SideOfPier:
                        {
                            // Move scope to requested position
                            SlewScope(TelescopeRAFromHourAngle("FlipRange: " + l_SideOfPier.ToString(), l_TargetHA), p_TargetDEC, "Moving to requested position");
                            l_TargetSideOfPier = (PierSide)telescopeDevice.SideOfPier; // Record pier side of target RA, DEC
                            break;
                        }

                    default:
                        {
                            MessageBox.Show("TelescopeFindFlipPoint - unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        }
                }

                switch (l_TargetSideOfPier) // Check outcome in terms of pier side and then adjust parameters to home in on correct value
                {
                    case var @case when @case == l_SideOfPier: // We have come back too far so go further out
                        {
                            l_Temp = l_TargetHA;
                            l_TargetHA = l_TargetHA + (l_EndHARange - l_StartHARange) / 2.0d;
                            l_StartHARange = l_Temp;
                            LogMsg("FlipRange: " + l_SideOfPier.ToString(), MessageLevel.msgDebug, "No Flip - Start, Target, End HA: " + FormatRA(l_StartHARange) + " " + FormatRA(l_TargetHA) + " " + FormatRA(l_EndHARange) + " " + FormatDec(p_TargetDEC));
                            break;
                        }

                    case PierSide.pierUnknown:
                        {
                            LogMsg("TelescopeFlipRange", MessageLevel.msgError, "Unexpected pier side: " + l_TargetSideOfPier.ToString()); // Side of pier has changed so come back a bit
                            break;
                        }

                    default:
                        {
                            l_Temp = l_TargetHA;
                            l_TargetHA = l_TargetHA - (l_EndHARange - l_StartHARange) / 2.0d;
                            l_EndHARange = l_Temp;
                            LogMsg("FlipRange: " + l_SideOfPier.ToString(), MessageLevel.msgDebug, "Flipped - Start, Target, End HA: " + FormatRA(l_StartHARange) + " " + FormatRA(l_TargetHA) + " " + FormatRA(l_EndHARange) + " " + FormatDec(p_TargetDEC));
                            break;
                        }
                }

                if (TestStop())
                    break;
            }

            return l_TargetHA;
        }

        #endregion

    }
}
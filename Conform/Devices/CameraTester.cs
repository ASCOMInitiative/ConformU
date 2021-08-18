using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using Microsoft.VisualBasic;
//using ASCOM.DeviceInterface;
using System.Collections;
using System.Threading;
using ASCOM.Standard.Interfaces;
using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.COM.DriverAccess;
using System.Runtime.InteropServices;
using ASCOM;

namespace ConformU
{

    internal class CameraTester : DeviceTesterBaseClass
    {
        #region Constants and variables

        const int CAMERA_PULSE_DURATION = 2000; // Duration of camera pulse guide test (ms)
        const int CAMERA_PULSE_TOLERANCE = 300; // Tolerance for acceptable performance (ms)

        const int MAX_BIN_X = 16; // Values of MaxBin above which warnings are given. Implemented to warn developers if they are returning "silly" values
        const int MAX_BIN_Y = 16;

        const double MIN_CAMERA_SETPOINT_TEMPERATURE = -280.0; // Value below which CCD temperatures will be flagged as "silly" values
        const double MAX_CAMERA_SETPOINT_TEMPERATURE = 100.0; // Value above which CCD setpoint temperatures will be flagged as "silly" values
        const double MAX_CAMERA_REPORTED_TEMPERATURE = 1000.0; // Value above which the CCD reported temperature will be flagged as a "silly" value. It is higher than the SetPoint temperature because this value is not specified in the Interface Standard.
        const double CAMERA_SETPOINT_TEST_INCREMENT = 0.000000001; // Value to add to MAX_CAMERA_SETPOINT_TEMPERATURE to test whether an exception is thrown at this setpoint

        // Camera variables
        private bool m_CanAbortExposure, m_CanAsymmetricBin, m_CanGetCoolerPower, m_CanSetCCDTemperature, m_CanStopExposure, m_CanFastReadout;
        private bool m_CoolerOn, m_HasShutter, m_ImageReady;
        private int m_CameraXSize, m_CameraYSize, m_MaxADU, m_NumX, m_NumY, m_StartX, m_StartY;
        private short m_MaxBinX, m_MaxBinY, m_BinX, m_BinY;
        private double m_CCDTemperature, m_CoolerPower, m_ElectronsPerADU, m_FullWellCapacity, m_HeatSinkTemperature, m_LastExposureDuration;
        private double m_PixelSizeX, m_PixelSizeY, m_SetCCDTemperature;
        private string m_LastExposureStartTime;
        private CameraState m_CameraState;
        private Array m_ImageArray, m_ImageArrayVariant;
        private bool m_IsPulseGuidingSupported; // Confirm that IsPulseGuiding command will work
        private bool m_CanPulseGuide;
        private bool m_IsPulseGuiding;
        // ICameraV2 properties
        private short m_BayerOffsetX, m_BayerOffsetY, m_Gain, m_GainMax, m_GainMin, m_PercentCompleted, m_ReadoutMode;
        private double m_ExposureMax, m_ExposureMin, m_ExposureResolution;
        private bool m_FastReadout, m_CanReadGain, m_CanReadGainMax, m_CanReadGainMin, m_CanReadGains, m_CanReadReadoutModes;
        private IList<string> m_Gains;
        private IList<string> m_ReadoutModes;
        private string m_SensorName;
        private SensorType m_SensorType;
        private bool m_CanReadSensorType = false;
        private readonly Stopwatch sw = new();

        // ICameraV3 properties
        private int m_Offset, m_OffsetMax, m_OffsetMin;
        private bool m_CanReadOffset, m_CanReadOffsetMax, m_CanReadOffsetMin, m_CanReadOffsets;
        private IList<string> m_Offsets;
        private double m_SubExposureDuration;
        private GainOffsetMode m_OffsetMode = GainOffsetMode.Unknown;
        private GainOffsetMode m_GainMode = GainOffsetMode.Unknown;

        private ICameraV3 m_Camera;

        private enum CanType
        {
            tstCanFindHome = 1,
            tstCanPark = 2,
            tstCanPulseGuide = 3,
            tstCanSetDeclinationRate = 4,
            tstCanSetGuideRates = 5,
            tstCanSetPark = 6,
            tstCanSetPierSide = 7,
            tstCanSetRightAscensionRate = 8,
            tstCanSetTracking = 9,
            tstCanSlew = 10,
            tstCanSlewAsync = 11,
            tstCanSlewAltAz = 12,
            tstCanSlewAltAzAsync = 13,
            tstCanSync = 14,
            tstCanSyncAltAz = 15,
            tstCanUnPark = 16,
            tstCanAbortExposure = 17,
            tstCanAsymmetricBin = 18,
            tstCanGetCoolerPower = 19,
            tstCanSetCCDTemperature = 20,
            tstCanStopExposure = 21,
            // ICameraV2 property
            tstCanFastReadout = 22
        }
        private enum CameraPerformance : int
        {
            CameraState,
            CCDTemperature,
            CoolerPower,
            HeatSinkTemperature,
            ImageReady,
            IsPulseGuiding,
            ImageArray,
            ImageArrayVariant
        }
        private enum CamPropertyType
        {
            // ICameraV1 Properties
            BinX = 1,
            BinY = 2,
            CameraState = 3,
            CameraXSize = 4,
            CameraYSize = 5,
            CCDTemperature = 6,
            CoolerOn = 7,
            CoolerPower = 8,
            Description = 9,
            DriverInfo = 10,
            ElectronsPerADU = 11,
            FullWellCapacity = 12,
            HasShutter = 13,
            HeatSinkTemperature = 14,
            ImageReady = 15,
            IsPulseGuiding = 16,
            MaxADU = 17,
            MaxBinX = 18,
            MaxBinY = 19,
            NumX = 20,
            NumY = 21,
            PixelSizeX = 22,
            PixelSizeY = 23,
            SetCCDTemperature = 24,
            StartX = 25,
            StartY = 26,

            // ICameraV2 Properties
            BayerOffsetX = 27,
            BayerOffsetY = 28,
            ExposureMax = 29,
            ExposureMin = 30,
            ExposureResolution = 31,
            FastReadout = 32,
            Gain = 33,
            GainMax = 34,
            GainMin = 35,
            Gains = 36,
            PercentCompleted = 37,
            ReadoutMode = 38,
            ReadoutModes = 39,
            SensorName = 40,
            SensorType = 41,

            // ICameraV3 Properties
            SubExposureDuration = 42
        }
        private enum GainOffsetMode
        {
            Unknown = 0,
            IndexMode = 1,
            ValueMode = 2
        }

        readonly Settings settings;
        readonly CancellationToken cancellationToken;
        readonly ConformLogger logger;

        #endregion

        public CameraTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, true, true, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls
        protected override void Dispose(bool disposing)
        {
            LogMsg("Dispose", MessageLevel.msgDebug, "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (m_Camera is not null) m_Camera.Dispose();
                    m_Camera = null;
                    m_ImageArray = null;
                    m_ImageArrayVariant = null;
                    GC.Collect();
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_Camera, DeviceType.Camera);
        }

        public override void CheckInitialise()
        {
            // Set the error type numbers according to the standards adopted by individual authors.
            // Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
            // messages to driver authors!
            unchecked
            {
                switch (settings.ComDevice.ProgId ?? "")
                {
                    default:
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040405;
                            g_ExInvalidValue2 = (int)0x80040405;
                            g_ExNotSet1 = (int)0x80040403;
                            break;
                        }
                }
            }
            base.CheckInitialise(settings.ComDevice.ProgId);
        }
        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_Camera = new AlpacaCamera(settings.AlpacaConfiguration.AccessServiceType.ToString(), settings.AlpacaDevice.IpAddress, settings.AlpacaDevice.IpPort, settings.AlpacaDevice.AlpacaDeviceNumber, logger);
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_Camera = new CameraFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                m_Camera = new Camera(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
                baseClassDevice = m_Camera; // Assign the driver to the base class
                g_Stop = false;
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");

            }
            catch (Exception ex)
            {
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Exception thrown: " + ex.Message);
                throw; // Re throw exception 
            }

            if (g_Stop) WaitFor(200);
        }

        public override bool Connected
        {
            get
            {
                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Connected");
                return m_Camera.Connected;
            }
            set
            {
                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Connected");
                m_Camera.Connected = value;
                g_Stop = false;
            }
        }
        public override void ReadCanProperties()
        {
            // ICameraV1 properties
            CameraCanTest(CanType.tstCanAbortExposure, "CanAbortExposure");
            CameraCanTest(CanType.tstCanAsymmetricBin, "CanAsymmetricBin");
            CameraCanTest(CanType.tstCanGetCoolerPower, "CanGetCoolerPower");
            CameraCanTest(CanType.tstCanPulseGuide, "CanPulseGuide");
            CameraCanTest(CanType.tstCanSetCCDTemperature, "CanSetCCDTemperature");
            CameraCanTest(CanType.tstCanStopExposure, "CanStopExposure");

            // ICameraV2 properties
            CameraCanTest(CanType.tstCanFastReadout, "CanFastReadout");
        }
        private void CameraCanTest(CanType p_Type, string p_Name)
        {
            try
            {
                switch (p_Type)
                {
                    case CanType.tstCanAbortExposure:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CanAbortExposure");
                            m_CanAbortExposure = m_Camera.CanAbortExposure;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanAbortExposure.ToString());
                            break;
                        }

                    case CanType.tstCanAsymmetricBin:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CanAsymmetricBin");
                            m_CanAsymmetricBin = m_Camera.CanAsymmetricBin;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanAsymmetricBin.ToString());
                            break;
                        }

                    case CanType.tstCanGetCoolerPower:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CanGetCoolerPower");
                            m_CanGetCoolerPower = m_Camera.CanGetCoolerPower;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanGetCoolerPower.ToString());
                            break;
                        }

                    case CanType.tstCanPulseGuide:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CanPulseGuide");
                            m_CanPulseGuide = m_Camera.CanPulseGuide;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanPulseGuide.ToString());
                            break;
                        }

                    case CanType.tstCanSetCCDTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CanSetCCDTemperature");
                            m_CanSetCCDTemperature = m_Camera.CanSetCCDTemperature;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanSetCCDTemperature.ToString());
                            break;
                        }

                    case CanType.tstCanStopExposure:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CanStopExposure");
                            m_CanStopExposure = m_Camera.CanStopExposure;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanStopExposure.ToString());
                            break;
                        }

                    case CanType.tstCanFastReadout:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CanFastReadout");
                            m_CanFastReadout = m_Camera.CanFastReadout;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanFastReadout.ToString());
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
        public override void PreRunCheck()
        {
            int l_VStringPtr, l_V1, l_V2, l_V3;
            // Add a test for a back level version of the Camera simulator - just abandon this process if any errors occur
            if (settings.ComDevice.ProgId.ToUpper() == "CCDSIMULATOR.CAMERA")
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Description");
                    l_VStringPtr = Strings.InStr(m_Camera.Description.ToUpper(), "VERSION "); // Point at the start of the version string
                    if (l_VStringPtr > 0)
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Description");
                        string l_VString = Strings.Mid(m_Camera.Description.ToUpper(), l_VStringPtr + 8);
                        l_VStringPtr = Strings.InStr(l_VString, ".");
                        if (l_VStringPtr > 1)
                        {
                            l_V1 = System.Convert.ToInt32(Strings.Mid(l_VString, 1, l_VStringPtr - 1)); // Extract the number
                            l_VString = Strings.Mid(l_VString, l_VStringPtr + 1); // Get the second version number part
                            l_VStringPtr = Strings.InStr(l_VString, ".");
                            if (l_VStringPtr > 1)
                            {
                                l_V2 = System.Convert.ToInt32(Strings.Mid(l_VString, 1, l_VStringPtr - 1)); // Extract the number
                                l_VString = Strings.Mid(l_VString, l_VStringPtr + 1); // Get the third version number part
                                l_V3 = System.Convert.ToInt32(l_VString); // Extract the number
                                                                          // Turn the version parts into a whole number
                                l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3;
                                if (l_V1 < 5000008)
                                {
                                    LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the camera simulator has known conformance issues, ***");
                                    LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***");
                                    LogMsg("", MessageLevel.msgAlways, "");
                                }
                            }
                        }
                    }
                    else
                    {
                        LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the camera simulator has known conformance issues, ***");
                        LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***");
                        LogMsg("", MessageLevel.msgAlways, "");
                    }
                }
                catch (Exception ex)
                {
                    LogMsg("ConformanceCheck", MessageLevel.msgError, ex.ToString());
                }
            }

            // Run camera tests
            if (!cancellationToken.IsCancellationRequested)
            {
                LogMsg("", MessageLevel.msgAlways, "");
                // Check LastError throws an exception
                LogMsg("Last Tests", MessageLevel.msgAlways, "");
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get LastExposureDuration");
                    m_LastExposureDuration = m_Camera.LastExposureDuration;
                    LogMsg("LastExposureDuration", MessageLevel.msgError, "LastExposureDuration did not generate an exception when called before an exposure was made");
                }
                catch (COMException)
                {
                    LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration correctly generated a COM exception before an exposure was made");
                }
                catch (Exception)
                {
                    LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration correctly generated a .NET exception before an exposure was made");
                }

                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get LastExposureStartTime");
                    m_LastExposureStartTime = m_Camera.LastExposureStartTime;
                    LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime did not generate an exception when called before an exposure was made");
                }
                catch (COMException)
                {
                    LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime correctly generated a COM exception before an exposure was made");
                }
                catch (Exception)
                {
                    LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime correctly generated a .NET exception before an exposure was made");
                }
            }
        }

        public override void CheckProperties()
        {
            int l_BinX, l_BinY, l_MaxBinX, l_MaxBinY;


            // Basic read tests
            m_MaxBinX = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.MaxBinX, "MaxBinX", 1, MAX_BIN_X)); if (cancellationToken.IsCancellationRequested)
                return;
            m_MaxBinY = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.MaxBinY, "MaxBinY", 1, MAX_BIN_Y)); if (cancellationToken.IsCancellationRequested)
                return;

            if (!m_CanAsymmetricBin)
            {
                if (m_MaxBinX != m_MaxBinY)
                    LogMsg("CanAsymmetricBin", MessageLevel.msgError, "CanAsymmetricBin is false but MaxBinX and MaxBinY are not equal!");
            }

            m_BinX = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.BinX, "BinX Read", 1, 1)); if (cancellationToken.IsCancellationRequested)
                return; // Must default to 1 on start-up
            m_BinY = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.BinY, "BinY Read", 1, 1)); if (cancellationToken.IsCancellationRequested)
                return; // Must default to 1 on start-up
            if (!m_CanAsymmetricBin)
            {
                if (m_BinX != m_BinY)
                    LogMsg("CanAsymmetricBin", MessageLevel.msgError, "CanAsymmetricBin is false but BinX and BinY are not equal!");
            }

            // Test writing low and high Bin values outside maximum range
            try // Invalid low value
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinX");
                m_Camera.BinX = 0;
                LogMsg("BinX Write", MessageLevel.msgError, "Invalid value 0 written but no exception generated");
            }
            catch (Exception)
            {
                LogMsg("BinX Write", MessageLevel.msgOK, "Exception correctly generated on setting BinX to 0");
            }
            try // Invalid high value
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinX");
                m_Camera.BinX = (short)(m_MaxBinX + 1);
                LogMsg("BinX Write", MessageLevel.msgError, "Invalid value " + m_MaxBinX + 1 + " written but no exception generated");
            }
            catch (Exception)
            {
                LogMsg("BinX Write", MessageLevel.msgOK, "Exception correctly generated on setting BinX to " + m_MaxBinX + 1);
            }
            try // Invalid low value
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinY");
                m_Camera.BinY = 0;
                LogMsg("BinY Write", MessageLevel.msgError, "Invalid value 0 written but no exception generated");
            }
            catch (Exception)
            {
                LogMsg("BinY Write", MessageLevel.msgOK, "Exception correctly generated on setting BinY to 0");
            }
            try // Invalid high value
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinY");
                m_Camera.BinY = (short)(m_MaxBinY + 1);
                LogMsg("BinY Write", MessageLevel.msgError, "Invalid value " + m_MaxBinY + 1 + " written but no exception generated");
            }
            catch (Exception)
            {
                LogMsg("BinY Write", MessageLevel.msgOK, "Exception correctly generated on setting BinY to " + m_MaxBinY + 1);
            }

            // Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogMsg("", MessageLevel.msgAlways, ""); // Insert a blank line if required
            if (settings.CameraMaxBinX > 0)
            {
                l_MaxBinX = settings.CameraMaxBinX;
                LogMsg("BinXY Write", MessageLevel.msgComment, string.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", l_MaxBinX, m_MaxBinX));
            }
            else
                l_MaxBinX = m_MaxBinX;
            if (settings.CameraMaxBinY > 0)
            {
                l_MaxBinY = settings.CameraMaxBinY;
                LogMsg("BinXY Write", MessageLevel.msgComment, string.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", l_MaxBinY, m_MaxBinY));
            }
            else
                l_MaxBinY = m_MaxBinY;

            if ((settings.CameraMaxBinX > m_MaxBinX) | (settings.CameraMaxBinY > m_MaxBinY))
                LogMsg("", MessageLevel.msgAlways, ""); // Insert a blank line if required
            if (settings.CameraMaxBinX > m_MaxBinX)
                LogMsg("BinXY Write", MessageLevel.msgComment, string.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", l_MaxBinX, m_MaxBinX));
            if (settings.CameraMaxBinY > m_MaxBinY)
                LogMsg("BinXY Write", MessageLevel.msgComment, string.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", l_MaxBinY, m_MaxBinY));

            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogMsg("", MessageLevel.msgAlways, ""); // Insert a blank line if required

            // Write BinX and BinY
            if (m_CanAsymmetricBin)
            {
                for (l_BinY = 1; l_BinY <= l_MaxBinY; l_BinY++)
                {
                    for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                    {
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinY");
                            m_Camera.BinY = (short)l_BinY;
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinX");
                                m_Camera.BinX = (short)l_BinX;
                                LogMsg("BinXY Write", MessageLevel.msgOK, "Successfully set asymmetric xy binning: " + l_BinX + " x " + l_BinY);
                            }
                            catch (Exception ex1)
                            {
                                HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex1, "when setting BinX to " + l_BinX, "InvalidValueException received when setting BinX to " + l_BinX);
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex, "when setting BinY to " + l_BinY, "InvalidValueException received when setting BinY to " + l_BinY);
                        }
                    }
                }
            }
            else // Can only bin symmetrically
                for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinY");
                        m_Camera.BinY = (short)l_BinX;
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinX");
                            m_Camera.BinX = (short)l_BinX;
                            LogMsg("BinXY Write", MessageLevel.msgOK, "Successfully set symmetric xy binning: " + l_BinX + " x " + l_BinX);
                        }
                        catch (Exception ex1)
                        {
                            HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex1, "when setting BinX to " + l_BinX, "InvalidValueException received when setting BinX to " + l_BinX);
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex, "when setting BinY to " + l_BinX, "InvalidValueException received when setting BinY to " + l_BinX);
                    }
                }

            // Reset binning to 1x1 state
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinX");
            try
            {
                m_Camera.BinX = 1;
            }
            catch (Exception)
            {
            }
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinY");
            try
            {
                m_Camera.BinY = 1;
            }
            catch (Exception)
            {
            }

            m_CameraState = CameraPropertyTestCameraState(CamPropertyType.CameraState, "CameraState"); if (cancellationToken.IsCancellationRequested)
                return;
            m_CameraXSize = CameraPropertyTestInteger(CamPropertyType.CameraXSize, "CameraXSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested)
                return;
            m_CameraYSize = CameraPropertyTestInteger(CamPropertyType.CameraYSize, "CameraYSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested)
                return;

            m_CCDTemperature = CameraPropertyTestDouble(CamPropertyType.CCDTemperature, "CCDTemperature", MIN_CAMERA_SETPOINT_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested)
                return;
            m_CoolerOn = CameraPropertyTestBoolean(CamPropertyType.CoolerOn, "CoolerOn Read", false); if (cancellationToken.IsCancellationRequested)
                return;

            // Write CoolerOn
            bool l_OriginalCoolerState;
            string l_TargetCoolerState;
            try
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set CoolerOn");
                l_OriginalCoolerState = m_Camera.CoolerOn;
                if (l_OriginalCoolerState)
                    l_TargetCoolerState = "off";
                else
                    l_TargetCoolerState = "on";
                try
                {
                    if (l_OriginalCoolerState)
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set CoolerOn");
                        m_Camera.CoolerOn = false;
                    }
                    else
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set CoolerOn");
                        m_Camera.CoolerOn = true;
                    }
                    LogMsg("CoolerOn Write", MessageLevel.msgOK, "Successfully changed CoolerOn state");
                }
                catch (Exception ex)
                {
                    HandleException("CoolerOn Write", MemberType.Property, Required.Optional, ex, "turning Cooler " + l_TargetCoolerState);
                }
                // Restore Cooler state
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set CoolerOn");
                try
                {
                    m_Camera.CoolerOn = l_OriginalCoolerState;
                }
                catch
                {
                }
            }
            catch (Exception ex)
            {
                // Skip tests as we can't even read the cooler state
                HandleException("CoolerOn Read", MemberType.Property, Required.Optional, ex, "");
            }

            m_CoolerPower = CameraPropertyTestDouble(CamPropertyType.CoolerPower, "CoolerPower", 0.0, 100.0, false); if (cancellationToken.IsCancellationRequested)
                return;
            m_ElectronsPerADU = CameraPropertyTestDouble(CamPropertyType.ElectronsPerADU, "ElectronsPerADU", 0.00001, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;
            m_FullWellCapacity = CameraPropertyTestDouble(CamPropertyType.FullWellCapacity, "FullWellCapacity", 0.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;
            m_HasShutter = CameraPropertyTestBoolean(CamPropertyType.HasShutter, "HasShutter", false); if (cancellationToken.IsCancellationRequested)
                return;
            m_HeatSinkTemperature = CameraPropertyTestDouble(CamPropertyType.HeatSinkTemperature, "HeatSinkTemperature", MIN_CAMERA_SETPOINT_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested)
                return;

            m_ImageReady = CameraPropertyTestBoolean(CamPropertyType.ImageReady, "ImageReady", false); if (cancellationToken.IsCancellationRequested)
                return;
            if (m_ImageReady)
                LogMsg("ImageReady", MessageLevel.msgError, "Image is flagged as ready but no exposure has been started!");
            if (m_ImageReady)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageArray");
                    m_ImageArray = (int[,])m_Camera.ImageArray;
                    LogMsg("ImageArray", MessageLevel.msgError, "No image has been taken but ImageArray has not generated an exception");
                }
                catch (Exception)
                {
                    LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated before an image has been taken");
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageArray");
                    m_ImageArray = (int[,])m_Camera.ImageArray;
                    LogMsg("ImageArray", MessageLevel.msgError, "ImageReady is false and no image has been taken but ImageArray has not generated an exception");
                }
                catch (Exception)
                {
                    LogMsg("ImageArray", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false");
                }

            m_ImageArray = null;
            m_ImageArrayVariant = null;
            GC.Collect();

            if (m_ImageReady)
            {
                try
                {
                    object ImageArrayVariantObject;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageArrayVariant");
                    ImageArrayVariantObject = m_Camera.ImageArrayVariant;
                    m_ImageArrayVariant = (int[,])ImageArrayVariantObject;
                    LogMsg("ImageArrayVariant", MessageLevel.msgError, "No image has been taken but ImageArray has not generated an exception");
                }
                catch (Exception)
                {
                    LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated before an image has been taken");
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageArrayVariant");
                    m_ImageArrayVariant = (int[,])m_Camera.ImageArrayVariant;
                    LogMsg("ImageArrayVariant", MessageLevel.msgError, "ImageReady is false and no image has been taken but ImageArray has not generated an exception");
                }
                catch (Exception)
                {
                    LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Exception correctly generated when ImageReady is false");
                }

            try
            {
                if (OperatingSystem.IsWindows())
                {
                    if (m_ImageArray is not null) Marshal.ReleaseComObject(m_ImageArray);
                }
            }
            catch
            {
            }
            try
            {
                if (OperatingSystem.IsWindows())
                {
                    if (m_ImageArray is not null) Marshal.ReleaseComObject(m_ImageArrayVariant);
                }
            }
            catch
            {
            }
            m_ImageArray = null;
            m_ImageArrayVariant = null;
            GC.Collect();

            m_IsPulseGuiding = CameraPropertyTestBoolean(CamPropertyType.IsPulseGuiding, "IsPulseGuiding", false); if (cancellationToken.IsCancellationRequested)
                return;
            if (m_IsPulseGuiding)
                LogMsg("IsPulseGuiding", MessageLevel.msgError, "Camera is showing pulse guiding underway although no PulseGuide command has been issued!");

            m_MaxADU = CameraPropertyTestInteger(CamPropertyType.MaxADU, "MaxADU", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested)
                return;

            m_NumX = CameraPropertyTestInteger(CamPropertyType.NumX, "NumX Read", 1, m_CameraXSize); if (cancellationToken.IsCancellationRequested)
                return;
            CameraPropertyWriteTest(CamPropertyType.NumX, "NumX", System.Convert.ToInt32(m_CameraXSize / (double)2));

            m_NumY = CameraPropertyTestInteger(CamPropertyType.NumY, "NumY Read", 1, m_CameraYSize); if (cancellationToken.IsCancellationRequested)
                return;
            CameraPropertyWriteTest(CamPropertyType.NumY, "NumY", System.Convert.ToInt32(m_CameraYSize / (double)2));

            m_PixelSizeX = CameraPropertyTestDouble(CamPropertyType.PixelSizeX, "PixelSizeX", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;
            m_PixelSizeY = CameraPropertyTestDouble(CamPropertyType.PixelSizeY, "PixelSizeY", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;

            m_SetCCDTemperature = CameraPropertyTestDouble(CamPropertyType.SetCCDTemperature, "SetCCDTemperature Read", MIN_CAMERA_SETPOINT_TEMPERATURE, MAX_CAMERA_SETPOINT_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested)
                return;
            if (m_CanSetCCDTemperature)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SetCCDTemperature");
                    m_Camera.SetCCDTemperature = 0.0; // Try an acceptable value
                    LogMsg("SetCCDTemperature Write", MessageLevel.msgOK, "Successfully wrote 0.0");

                    // Execution only gets here if the CCD temperature can be set successfully
                    bool l_ExceptionGenerated;
                    double l_SetPoint;

                    // Find low setpoint at which an exception is generated, stop at -280 as this is unphysical
                    l_ExceptionGenerated = false;
                    l_SetPoint = -0.0;
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SetCCDTemperature multiple times...");
                    do
                    {
                        try
                        {
                            l_SetPoint -= 5;
                            if (settings.DisplayMethodCalls)
                                m_Camera.SetCCDTemperature = l_SetPoint;
                        }
                        catch (Exception)
                        {
                            l_ExceptionGenerated = true;
                        }
                    }
                    while (!(l_ExceptionGenerated | (l_SetPoint <= MIN_CAMERA_SETPOINT_TEMPERATURE)))// Reached lower limit so exit loop
    ;

                    if (!l_ExceptionGenerated & (l_SetPoint == MIN_CAMERA_SETPOINT_TEMPERATURE))
                    {
                        // Now test whether it is possible to set the temperature just below the minimum setpoint, which should result in an exception if all is well
                        l_ExceptionGenerated = false;
                        try
                        {
                            m_Camera.SetCCDTemperature = MIN_CAMERA_SETPOINT_TEMPERATURE - CAMERA_SETPOINT_TEST_INCREMENT;
                        }
                        catch (Exception)
                        {
                            l_ExceptionGenerated = true;
                        }// Reached lower limit

                        if (l_ExceptionGenerated)
                            LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, $"Setpoint can be set to {MIN_CAMERA_SETPOINT_TEMPERATURE} degrees");
                        else
                            LogMsg("SetCCDTemperature Write", MessageLevel.msgIssue, $"Setpoint can be set below {MIN_CAMERA_SETPOINT_TEMPERATURE} degrees, which is below absolute zero!");
                    }
                    else
                        LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, $"Setpoint lower limit found in the range {l_SetPoint + 5.0} to {l_SetPoint + 0.001} degrees");

                    // Find high setpoint at which an exception is generated, stop at MAX_CAMERA_SETPOINT_TEMPERATURE as this is a suitably high value
                    l_ExceptionGenerated = false;
                    l_SetPoint = 0.0; // Start at 0.0C

                    // Loop upward in 5 degree temperature steps to find the maximum temperature that can be set
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SetCCDTemperature multiple times...");
                    do
                    {
                        try
                        {
                            l_SetPoint += 5.0;
                            m_Camera.SetCCDTemperature = l_SetPoint;
                        }
                        catch (Exception)
                        {
                            l_ExceptionGenerated = true;
                        }
                    }
                    while (!(l_ExceptionGenerated | (l_SetPoint >= MAX_CAMERA_SETPOINT_TEMPERATURE)))// Reached upper limit so exit loop
    ;

                    if (!l_ExceptionGenerated & (l_SetPoint == MAX_CAMERA_SETPOINT_TEMPERATURE))
                    {
                        // Now test whether it is possible to set the temperature just above the maximum setpoint, which should result in an exception if all is well
                        l_ExceptionGenerated = false;
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SetCCDTemperature");
                            m_Camera.SetCCDTemperature = MAX_CAMERA_SETPOINT_TEMPERATURE + CAMERA_SETPOINT_TEST_INCREMENT;
                        }
                        catch (Exception)
                        {
                            l_ExceptionGenerated = true;
                        }// Reached upper limit

                        if (l_ExceptionGenerated)
                            LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, $"Setpoint can be set to {MAX_CAMERA_SETPOINT_TEMPERATURE} degrees");
                        else
                            LogMsg("SetCCDTemperature Write", MessageLevel.msgIssue, $"Setpoint can be set in excess of {MAX_CAMERA_SETPOINT_TEMPERATURE} degrees");
                    }
                    else
                        LogMsg("SetCCDTemperature Write", MessageLevel.msgInfo, $"Setpoint upper limit found in the range {l_SetPoint - 5.0} to {l_SetPoint - 0.001} degrees");
                }
                catch (Exception ex)
                {
                    HandleException("SetCCDTemperature Write", MemberType.Property, Required.MustBeImplemented, ex, "Setting a legal value 0");
                }

                // Restore original value
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SetCCDTemperature");
                try
                {
                    m_Camera.SetCCDTemperature = m_SetCCDTemperature;
                }
                catch
                {
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SetCCDTemperature");
                    m_Camera.SetCCDTemperature = 0;
                    LogMsg("SetCCDTemperature Write", MessageLevel.msgError, "CanSetCCDTemperature is false but no exception generated on write");
                }
                catch (Exception ex)
                {
                    HandleException("SetCCDTemperature Write", MemberType.Property, Required.Optional, ex, "");
                }

            m_StartX = CameraPropertyTestInteger(CamPropertyType.StartX, "StartX Read", 0, m_CameraXSize - 1); if (cancellationToken.IsCancellationRequested)
                return;
            CameraPropertyWriteTest(CamPropertyType.StartX, "StartX", System.Convert.ToInt32(m_CameraXSize / (double)2));
            m_StartY = CameraPropertyTestInteger(CamPropertyType.StartY, "StartY Read", 0, m_CameraYSize - 1); if (cancellationToken.IsCancellationRequested)
                return;
            CameraPropertyWriteTest(CamPropertyType.StartY, "StartY", System.Convert.ToInt32(m_CameraYSize / (double)2));



            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get InterfaceVersion");
            if (m_Camera.InterfaceVersion > 1)
            {
                // SensorType - Mandatory
                // This must be tested before BayerOffset because BayerOffset is mandatory for colour and optional for monochrome cameras
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Sensortype");
                    m_SensorType = m_Camera.SensorType;
                    m_CanReadSensorType = true; // Set a flag to indicate that we have got a valid SensorType value
                                                // Successfully retrieved a value
                    LogMsg("SensorType Read", MessageLevel.msgOK, m_SensorType.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("SensorType Read", MemberType.Property, Required.Mandatory, ex, "");
                }

                // BayerOffset Read
                if (m_CanReadSensorType)
                {
                    if (m_SensorType == SensorType.Monochrome)
                    {
                        // Monochrome so both BayerOffset properties should throw not implemented exceptions
                        CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read");
                        CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read");
                    }
                    else
                    {
                        m_BayerOffsetX = CameraPropertyTestShort(CamPropertyType.BayerOffsetX, "BayerOffsetX Read", 0, 10000, true);
                        m_BayerOffsetY = CameraPropertyTestShort(CamPropertyType.BayerOffsetY, "BayerOffsetY Read", 0, 10000, true);
                    }
                }
                else
                {
                    LogMsg("BayerOffset Read", MessageLevel.msgInfo, "Unable to read SensorType value; assuming that the sensor is Monochrome");
                    // Monochrome so both BayerOffset properties should throw not implemented exceptions
                    CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read");
                    CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read");
                }

                // ExposureMin and ExpoosureMax Read
                m_ExposureMax = CameraPropertyTestDouble(CamPropertyType.ExposureMax, "ExposureMax Read", 0.0001, double.MaxValue, true);
                m_ExposureMin = CameraPropertyTestDouble(CamPropertyType.ExposureMin, "ExposureMin Read", 0.0, double.MaxValue, true);
                if (m_ExposureMin <= m_ExposureMax)
                    LogMsg("ExposureMin", MessageLevel.msgOK, "ExposureMin is less than or equal to ExposureMax");
                else
                    LogMsg("ExposureMin", MessageLevel.msgError, "ExposureMin is greater than ExposureMax");

                // ExposureResolution Read
                m_ExposureResolution = CameraPropertyTestDouble(CamPropertyType.ExposureResolution, "ExposureResolution Read", 0.0, double.MaxValue, true);
                if (m_ExposureResolution <= m_ExposureMax)
                    LogMsg("ExposureResolution", MessageLevel.msgOK, "ExposureResolution is less than or equal to ExposureMax");
                else
                    LogMsg("ExposureResolution", MessageLevel.msgError, "ExposureResolution is greater than ExposureMax");

                // FastReadout Read Optional
                if (m_CanFastReadout)
                    m_FastReadout = CameraPropertyTestBoolean(CamPropertyType.FastReadout, "FastReadout Read", true);
                else
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get FastReadout");
                        m_FastReadout = m_Camera.FastReadout;
                        LogMsg("FastReadout Read", MessageLevel.msgError, "CanFastReadout is False but a PropertyNotImplementedException was not thrown");
                    }
                    catch (Exception ex)
                    {
                        HandleException("FastReadout Read", MemberType.Property, Required.Optional, ex, "");
                    }

                // FastReadout Write Optional
                if (m_CanFastReadout)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set FastReadout");
                        m_Camera.FastReadout = !m_FastReadout;
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set FastReadout");
                        m_Camera.FastReadout = m_FastReadout;
                        LogMsg("FastReadout Write", MessageLevel.msgOK, "Able to change the FastReadout state OK");
                    }
                    catch (Exception ex)
                    {
                        HandleException("FastReadout Write", MemberType.Property, Required.Mandatory, ex, "");
                    }
                }
                else
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set FastReadout");
                        m_Camera.FastReadout = true;
                        LogMsg("FastReadout Write", MessageLevel.msgError, "CanFastReadout is False but a PropertyNotImplementedException was not thrown");
                    }
                    catch (Exception ex)
                    {
                        HandleException("FastReadout Write", MemberType.Property, Required.Optional, ex, "");
                    }

                // GainMin Read - Optional
                try
                {
                    m_CanReadGainMin = false;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get GainMin");
                    m_GainMin = m_Camera.GainMin;
                    // Successfully retrieved a value
                    m_CanReadGainMin = true;
                    LogMsg("GainMin Read", MessageLevel.msgOK, m_GainMin.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("GainMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                // GainMax Read - Optional
                try
                {
                    m_CanReadGainMax = false;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get GainMax");
                    m_GainMax = m_Camera.GainMax;
                    // Successfully retrieved a value
                    m_CanReadGainMax = true;
                    LogMsg("GainMax Read", MessageLevel.msgOK, m_GainMax.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("GainMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                // Gains Read - Optional
                try
                {
                    m_CanReadGains = false;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Gains");
                    m_Gains = m_Camera.Gains;
                    // Successfully retrieved a value
                    m_CanReadGains = true;
                    foreach (string Gain in m_Gains)
                        LogMsg("Gains Read", MessageLevel.msgOK, Gain.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("Gains Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                if (m_CanReadGainMax & m_CanReadGainMin & m_CanReadGains)
                    LogMsg("Gains", MessageLevel.msgError, "GainMin, GainMax and Gains are all readable. Only one of GainMin/Max as a pair or Gains should be used, the other should throw a PropertyNotImplementedException");
                else
                {
                }

                // Gain Read - Optional 
                try
                {
                    m_CanReadGain = false; // Set default value to indicate can't read gain
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Gain");
                    m_Gain = m_Camera.Gain;
                    m_CanReadGain = true; // Flag that we can read Gain OK
                    if (m_CanReadGains)
                        LogMsg("Gain Read", MessageLevel.msgOK, m_Gain + " " + m_Gains[0].ToString());
                    else
                        LogMsg("Gain Read", MessageLevel.msgOK, m_Gain.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("Gain Read", MemberType.Property, Required.Optional, ex, "");
                }

                // Now check that gain property groups are implemented to handle the three gain modes: NotImplemented, Gain Index (Gain + Gains) and Gain Value (Gain + GainMin + GainMax)
                if (!m_CanReadGain & !m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax)
                    LogMsgOK("Gain Read", "All four gain properties throw exceptions - the driver is in \"Gain Not Implemented\" mode.");
                else if (m_CanReadGain)
                {
                    // Test for Gain Index mode
                    if ((m_CanReadGain & m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax))
                    {
                        m_GainMode = GainOffsetMode.IndexMode;
                        LogMsgOK("Gain Read", "Gain and Gains can be read while GainMin and GainMax throw exceptions - the driver is in \"Gain Index\" mode.");
                    }
                    else if ((m_CanReadGain & !m_CanReadGains & m_CanReadGainMin & m_CanReadGainMax))
                    {
                        m_GainMode = GainOffsetMode.ValueMode;
                        LogMsgOK("Gain Read", "Gain, GainMin and GainMax can be read OK while Gains throws an exception - the driver is in \"Gain Value\" mode.");
                    }
                    else
                    {
                        LogMsgError("Gain Read", $"Unable to determine whether the driver is in \"Gain Not Implemented\", \"Gain Index\" or \"Gain Value\" mode. Please check the interface specification.");
                        LogMsgInfo("Gain Read", $"Gain threw an exception: {m_CanReadGain}, Gains threw an exception: {m_CanReadGains}, GainMin threw an exception: {m_CanReadGainMin}, GainMax threw an exception: {m_CanReadGainMax}.");
                        LogMsgInfo("Gain Read", $"\"Gain Not Implemented\" mode: Gain, Gains, GainMin and GainMax must all throw exceptions.");
                        LogMsgInfo("Gain Read", $"\"Gain Index\" mode: Gain and Gains must work while GainMin and GainMax must throw exceptions.");
                        LogMsgInfo("Gain Read", $"\"Gain Value\" mode: Gain, GainMin and GainMax must work while Gains must throw an exception.");
                    }
                }
                else
                {
                    LogMsgError("Gain Read", $"Gain Read threw an exception but at least one of Gains, GainMin Or GainMax did not throw an exception. If Gain throws an exception, all the other gain properties should do likewise.");
                    LogMsgInfo("Gain Read", $"Gains threw an exception: {m_CanReadGains}, GainMin threw an exception: {m_CanReadGainMin}, GainMax threw an exception: {m_CanReadGainMax}.");
                }

                // Gain write - Optional when neither gain index nor gain value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither gain mode is supported
                if (!m_CanReadGain & !m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Gain");
                        m_Camera.Gain = 0;
                        LogMsgIssue("Gain Write", "Writing to Gain did not throw a PropertyNotImplementedException when reading Gain did.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("Gain Write", MemberType.Property, Required.MustNotBeImplemented, ex, "PropertyNotImplementedException is expected");
                    }
                }
                else
                    switch (m_GainMode)
                    {
                        case GainOffsetMode.Unknown:
                            {
                                LogMsgIssue("Gain Write", "Cannot test Gain Write because of issues with other gain properties - skipping test");
                                break;
                            }

                        case GainOffsetMode.IndexMode:
                            {
                                break;
                            }

                        case GainOffsetMode.ValueMode:
                            {
                                // Test writing the minimum valid value
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Gain");
                                    m_Camera.Gain = m_GainMin;
                                    LogMsgOK("Gain Write", $"Successfully set gain minimum value {m_GainMin}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing the maximum valid value
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Gain");
                                    m_Camera.Gain = m_GainMax;
                                    LogMsgOK("Gain Write", $"Successfully set gain maximum value {m_GainMax}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Gain Write", MemberType.Property, Required.MustNotBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Gain");
                                    m_Camera.Gain = (short)(m_GainMin - 1);
                                    LogMsgIssue("Gain Write", $"Successfully set an gain below the minimum value ({m_GainMin - 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for gain {m_GainMin - 1}, which is lower than the minimum value.");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Gain");
                                    m_Camera.Gain = (short)(m_GainMax + 1);
                                    LogMsgIssue("Gain Write", $"Successfully set an gain above the maximum value({m_GainMax + 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for gain {m_GainMax + 1} which is higher than the maximum value.");
                                }

                                break;
                            }

                        default:
                            {
                                LogMsgError("Gain Write", $"UNEXPECTED VALUE FOR OFFSETMODE: {m_GainMode}");
                                break;
                            }
                    }

                // PercentCompleted Read - Optional - corrected to match the specification
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get PercentCompleted");
                    m_PercentCompleted = m_Camera.PercentCompleted;
                    switch (m_PercentCompleted)
                    {
                        case object _ when m_PercentCompleted < 0 // Lower than minimum value
                       :
                            {
                                LogMsg("PercentCompleted Read", MessageLevel.msgError, "Invalid value: " + m_PercentCompleted.ToString());
                                break;
                            }

                        case object _ when m_PercentCompleted > 100 // Higher than maximum value
                 :
                            {
                                LogMsg("PercentCompleted Read", MessageLevel.msgError, "Invalid value: " + m_PercentCompleted.ToString());
                                break;
                            }

                        default:
                            {
                                LogMsg("PercentCompleted Read", MessageLevel.msgOK, m_PercentCompleted.ToString());
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOK("PercentCompleted Read", MemberType.Property, Required.Optional, ex, "", "Operation is invalid when camera is not imaging or downloading");
                }

                // ReadoutModes - Mandatory
                try
                {
                    m_CanReadReadoutModes = false;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ReadoutModes");
                    m_ReadoutModes = m_Camera.ReadoutModes;
                    // Successfully retrieved a value
                    m_CanReadReadoutModes = true;
                    foreach (string ReadoutMode in m_ReadoutModes)
                        LogMsg("ReadoutModes Read", MessageLevel.msgOK, ReadoutMode.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("ReadoutModes Read", MemberType.Property, Required.Mandatory, ex, "");
                }

                // ReadoutMode Read - Mandatory
                m_ReadoutMode = CameraPropertyTestShort(CamPropertyType.ReadoutMode, "ReadoutMode Read", 0, short.MaxValue, true);
                if (m_CanReadReadoutModes)
                {
                    try
                    {
                        if (m_ReadoutMode < m_ReadoutModes.Count)
                        {
                            LogMsg("ReadoutMode Index", MessageLevel.msgOK, "ReadReadoutMode is within the bounds of the ReadoutModes ArrayList");
                            LogMsg("ReadoutMode Index", MessageLevel.msgInfo, "Current value: " + m_ReadoutModes[m_ReadoutMode].ToString());
                        }
                        else
                            LogMsg("ReadoutMode Index", MessageLevel.msgError, "ReadReadoutMode is outside the bounds of the ReadoutModes ArrayList");
                    }
                    catch (Exception ex)
                    {
                        LogMsg("ReadoutMode Index", MessageLevel.msgError, "Exception: " + ex.Message);
                        LogMsg("ReadoutMode Index", MessageLevel.msgDebug, ex.ToString());
                    }
                }
                else
                    LogMsg("ReadoutMode Index", MessageLevel.msgInfo, "Skipping ReadReadoutMode index test because ReadoutModes is unavailable");

                // SensorName
                m_SensorName = CameraPropertyTestString(CamPropertyType.SensorName, "SensorName Read", 250, true);
            }



            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get InterfaceVersion");
            if (m_Camera.InterfaceVersion > 2)
            {
                // OffsetMin Read - Optional
                try
                {
                    m_CanReadOffsetMin = false;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get OffsetMin");
                    m_OffsetMin = m_Camera.OffsetMin;
                    // Successfully retrieved a value
                    m_CanReadOffsetMin = true;
                    LogMsg("OffsetMin Read", MessageLevel.msgOK, m_OffsetMin.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("OffsetMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                // OffsetMax Read - Optional
                try
                {
                    m_CanReadOffsetMax = false;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get OffsetMax");
                    m_OffsetMax = m_Camera.OffsetMax;
                    // Successfully retrieved a value
                    m_CanReadOffsetMax = true;
                    LogMsg("OffsetMax Read", MessageLevel.msgOK, m_OffsetMax.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("OffsetMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                // Offsets Read - Optional
                try
                {
                    m_CanReadOffsets = false;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Offsets");
                    m_Offsets = m_Camera.Offsets;
                    // Successfully retrieved a value
                    m_CanReadOffsets = true;
                    foreach (string Offset in m_Offsets)
                        LogMsg("Offsets Read", MessageLevel.msgOK, Offset.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("Offsets Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                if (m_CanReadOffsetMax & m_CanReadOffsetMin & m_CanReadOffsets)
                    LogMsg("Offsets", MessageLevel.msgError, "OffsetMin, OffsetMax and Offsets are all readable. Only one of OffsetMin/Max as a pair or Offsets should be used, the other should throw a PropertyNotImplementedException");
                else
                {
                }

                // Offset Read - Optional 
                try
                {
                    m_CanReadOffset = false; // Set default value to indicate can't read offset
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Offset");
                    m_Offset = m_Camera.Offset;
                    m_CanReadOffset = true; // Flag that we can read Offset OK
                    if (m_CanReadOffsets)
                        LogMsg("Offset Read", MessageLevel.msgOK, m_Offset + " " + m_Offsets[0].ToString());
                    else
                        LogMsg("Offset Read", MessageLevel.msgOK, m_Offset.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("Offset Read", MemberType.Property, Required.Optional, ex, "");
                }

                // Now check that offset property groups are implemented to handle the three offset modes: NotImplemented, Offset Index (Offset + Offsets) and Offset Value (Offset + OffsetMin + OffsetMax)
                if (!m_CanReadOffset & !m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax)
                    LogMsgOK("Offset Read", "All four offset properties throw exceptions - the driver is in \"Offset Not Implemented\" mode.");
                else if (m_CanReadOffset)
                {
                    // Test for Offset Index mode
                    if ((m_CanReadOffset & m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax))
                    {
                        m_OffsetMode = GainOffsetMode.IndexMode;
                        LogMsgOK("Offset Read", "Offset and Offsets can be read while OffsetMin and OffsetMax throw exceptions - the driver is in \"Offset Index\" mode.");
                    }
                    else if ((m_CanReadOffset & !m_CanReadOffsets & m_CanReadOffsetMin & m_CanReadOffsetMax))
                    {
                        m_OffsetMode = GainOffsetMode.ValueMode;
                        LogMsgOK("Offset Read", "Offset, OffsetMin and OffsetMax can be read OK while Offsets throws an exception - the driver is in \"Offset Value\" mode.");
                    }
                    else
                    {
                        m_OffsetMode = GainOffsetMode.Unknown;
                        LogMsgIssue("Offset Read", $"Unable to determine whether the driver is in \"Offset Not Implemented\", \"Offset Index\" or \"Offset Value\" mode. Please check the interface specification.");
                        LogMsgInfo("Offset Read", $"Offset threw an exception: {m_CanReadOffset}, Offsets threw an exception: {m_CanReadOffsets}, OffsetMin threw an exception: {m_CanReadOffsetMin}, OffsetMax threw an exception: {m_CanReadOffsetMax}.");
                        LogMsgInfo("Offset Read", $"\"Offset Not Implemented\" mode: Offset, Offsets, OffsetMin and OffsetMax must all throw exceptions.");
                        LogMsgInfo("Offset Read", $"\"Offset Index\" mode: Offset and Offsets must work while OffsetMin and OffsetMax must throw exceptions.");
                        LogMsgInfo("Offset Read", $"\"Offset Value\" mode: Offset, OffsetMin and OffsetMax must work while Offsets must throw an exception.");
                    }
                }
                else
                {
                    LogMsgError("Offset Read", $"Offset Read threw an exception but at least one of Offsets, OffsetMin Or OffsetMax did not throw an exception. If Offset throws an exception, all the other offset properties should do likewise.");
                    LogMsgInfo("Offset Read", $"Offsets threw an exception: {m_CanReadOffsets}, OffsetMin threw an exception: {m_CanReadOffsetMin}, OffsetMax threw an exception: {m_CanReadOffsetMax}.");
                }

                // Offset write - Optional when neither offset index nor offset value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither offset mode is supported
                if (!m_CanReadOffset & !m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Offset");
                        m_Camera.Offset = 0;
                        LogMsgIssue("Offset Write", "Writing to Offset did not throw a PropertyNotImplementedException when reading Offset did.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("Offset Write", MemberType.Property, Required.MustNotBeImplemented, ex, "PropertyNotImplementedException is expected");
                    }
                }
                else
                    switch (m_OffsetMode)
                    {
                        case GainOffsetMode.Unknown:
                            {
                                LogMsgIssue("Offset Write", "Cannot test Offset Write because of issues with other offset properties - skipping test");
                                break;
                            }

                        case GainOffsetMode.IndexMode:
                            {
                                break;
                            }

                        case GainOffsetMode.ValueMode:
                            {
                                // Test writing the minimum valid value
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Offset");
                                    m_Camera.Offset = m_OffsetMin;
                                    LogMsgOK("Offset Write", $"Successfully set offset minimum value {m_OffsetMin}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing the maximum valid value
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Offset");
                                    m_Camera.Offset = m_OffsetMax;
                                    LogMsgOK("Offset Write", $"Successfully set offset maximum value {m_OffsetMax}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Offset Write", MemberType.Property, Required.MustNotBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Offset");
                                    m_Camera.Offset = m_OffsetMin - 1;
                                    LogMsgIssue("Offset Write", $"Successfully set an offset below the minimum value ({m_OffsetMin - 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for offset {m_OffsetMin - 1}, which is lower than the minimum value.");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set Offset");
                                    m_Camera.Offset = m_OffsetMax + 1;
                                    LogMsgIssue("Offset Write", $"Successfully set an offset above the maximum value({m_OffsetMax + 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for offset {m_OffsetMax + 1} which is higher than the maximum value.");
                                }

                                break;
                            }

                        default:
                            {
                                LogMsgError("Offset Write", $"UNEXPECTED VALUE FOR OFFSETMODE: {m_OffsetMode}");
                                break;
                            }
                    }

                // SubExposureDuration Read - Optional 
                m_SubExposureDuration = CameraPropertyTestDouble(CamPropertyType.SubExposureDuration, "SubExposureDuration", double.Epsilon, double.MaxValue, false); if (cancellationToken.IsCancellationRequested)
                    return;

                // SubExposureDuration Write - Optional 
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SubExposureDuration");
                    m_Camera.SubExposureDuration = m_SubExposureDuration;
                    LogMsg("SubExposureDuration write", MessageLevel.msgOK, $"Successfully wrote {m_SubExposureDuration}");
                }
                catch (Exception ex)
                {
                    HandleException("SubExposureDuration write", MemberType.Property, Required.Optional, ex, "");
                }
            }
        }
        private CameraState CameraPropertyTestCameraState(CamPropertyType p_Type, string p_Name)
        {
            CameraState returnValue = CameraState.Idle;

            try
            {
                switch (p_Type)
                {
                    case CamPropertyType.CameraState:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState");
                            returnValue = m_Camera.CameraState;
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                LogMsg(p_Name, MessageLevel.msgOK, returnValue.ToString());
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.Optional, ex, "");
            }

            return returnValue;
        }
        private short CameraPropertyTestShort(CamPropertyType p_Type, string p_Name, short p_Min, short p_Max, bool p_Mandatory)
        {
            short returnValue = 0;

            try
            {
                switch (p_Type)
                {
                    case CamPropertyType.BayerOffsetX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get BayerOffsetX");
                            returnValue = m_Camera.BayerOffsetX;
                            break;
                        }

                    case CamPropertyType.BayerOffsetY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get BayerOffsetY");
                            returnValue = m_Camera.BayerOffsetY;
                            break;
                        }

                    case CamPropertyType.PercentCompleted:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get PercentCompleted");
                            returnValue = m_Camera.PercentCompleted;
                            break;
                        }

                    case CamPropertyType.ReadoutMode:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ReadoutMode");
                            returnValue = m_Camera.ReadoutMode;
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgOK, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }
        private bool CameraPropertyMustNotImplemented(CamPropertyType p_Type, string p_Name)
        {
            short TestShort; // Dummy variable to hold value that should never be returned
            bool returnValue = true;

            try
            {
                switch (p_Type)
                {
                    case CamPropertyType.BayerOffsetX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get BayerOffsetX");
                            TestShort = m_Camera.BayerOffsetX;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogMsg(p_Name, MessageLevel.msgError, "Sensor type is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value");
                            break;
                        }

                    case CamPropertyType.BayerOffsetY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get BayerOffsetY");
                            TestShort = m_Camera.BayerOffsetY;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogMsg(p_Name, MessageLevel.msgError, "Sensor type is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value");
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.MustNotBeImplemented, ex, "Sensor type is Monochrome");
            }
            return returnValue; // Return success indicator, True means property did thrown the exception, False means that it did not
        }
        private int CameraPropertyTestInteger(CamPropertyType p_Type, string p_Name, int p_Min, int p_Max)
        {
            int returnValue = 0;

            try
            {
                switch (p_Type)
                {
                    case CamPropertyType.BinX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get BinX");
                            returnValue = m_Camera.BinX;
                            break;
                        }

                    case CamPropertyType.BinY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get BinY");
                            returnValue = m_Camera.BinY;
                            break;
                        }

                    case CamPropertyType.CameraState:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState");
                            returnValue = (int)m_Camera.CameraState;
                            break;
                        }

                    case CamPropertyType.CameraXSize:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraXSize");
                            returnValue = m_Camera.CameraXSize;
                            break;
                        }

                    case CamPropertyType.CameraYSize:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraYSize");
                            returnValue = m_Camera.CameraYSize;
                            break;
                        }

                    case CamPropertyType.MaxADU:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get MaxADU");
                            returnValue = m_Camera.MaxADU;
                            break;
                        }

                    case CamPropertyType.MaxBinX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get MaxBinX");
                            returnValue = m_Camera.MaxBinX;
                            break;
                        }

                    case CamPropertyType.MaxBinY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get MaxBinY");
                            returnValue = m_Camera.MaxBinY;
                            break;
                        }

                    case CamPropertyType.NumX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get NumX");
                            returnValue = m_Camera.NumX;
                            break;
                        }

                    case CamPropertyType.NumY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get NumY");
                            returnValue = m_Camera.NumY;
                            break;
                        }

                    case CamPropertyType.StartX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get StartX");
                            returnValue = m_Camera.StartX;
                            break;
                        }

                    case CamPropertyType.StartY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get StartY");
                            returnValue = m_Camera.StartY;
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            switch (p_Type) // Provide the required message depending on the property being tested
                            {
                                case CamPropertyType.MaxBinX // Informational message for MaxBinX
                               :
                                    {
                                        LogMsg(p_Name, MessageLevel.msgInfo, $"{returnValue}. This is higher than Conform's test criterion: {MAX_BIN_X}. Is this intended?");
                                        break;
                                    }

                                case CamPropertyType.MaxBinY // Informational message for MaxBinY
                         :
                                    {
                                        LogMsg(p_Name, MessageLevel.msgInfo, $"{returnValue}. This is higher than Conform's test criterion: {MAX_BIN_Y}. Is this intended?");
                                        break;
                                    }

                                default:
                                    {
                                        LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + returnValue.ToString());
                                        break;
                                    }
                            }

                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgOK, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.Optional, ex, "");
            }
            return returnValue;
        }
        private double CameraPropertyTestDouble(CamPropertyType p_Type, string p_Name, double p_Min, double p_Max, bool p_Mandatory)
        {
            double returnValue = 0.0;

            try
            {
                switch (p_Type)
                {
                    case CamPropertyType.CCDTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CCDTemperature");
                            returnValue = m_Camera.CCDTemperature;
                            break;
                        }

                    case CamPropertyType.CoolerPower:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CoolerPower");
                            returnValue = m_Camera.CoolerPower;
                            break;
                        }

                    case CamPropertyType.ElectronsPerADU:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ElectronsPerADU");
                            returnValue = m_Camera.ElectronsPerADU;
                            break;
                        }

                    case CamPropertyType.FullWellCapacity:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get FullWellCapacity");
                            returnValue = m_Camera.FullWellCapacity;
                            break;
                        }

                    case CamPropertyType.HeatSinkTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get HeatSinkTemperature");
                            returnValue = m_Camera.HeatSinkTemperature;
                            break;
                        }

                    case CamPropertyType.PixelSizeX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get PixelSizeX");
                            returnValue = m_Camera.PixelSizeX;
                            break;
                        }

                    case CamPropertyType.PixelSizeY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get PixelSizeY");
                            returnValue = m_Camera.PixelSizeY;
                            break;
                        }

                    case CamPropertyType.SetCCDTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get SetCCDTemperature");
                            returnValue = m_Camera.SetCCDTemperature;
                            break;
                        }

                    case CamPropertyType.ExposureMax:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ExposureMax");
                            returnValue = m_Camera.ExposureMax;
                            break;
                        }

                    case CamPropertyType.ExposureMin:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ExposureMin");
                            returnValue = m_Camera.ExposureMin;
                            break;
                        }

                    case CamPropertyType.ExposureResolution:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ExposureResolution");
                            returnValue = m_Camera.ExposureResolution;
                            break;
                        }

                    case CamPropertyType.SubExposureDuration:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get SubExposureDuration");
                            returnValue = m_Camera.SubExposureDuration;
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case double _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case double _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgOK, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }
        private bool CameraPropertyTestBoolean(CamPropertyType p_Type, string p_Name, bool p_Mandatory)
        {
            bool returnValue = false;

            try
            {
                returnValue = false;
                switch (p_Type)
                {
                    case CamPropertyType.CoolerOn:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CoolerOn");
                            returnValue = m_Camera.CoolerOn;
                            break;
                        }

                    case CamPropertyType.HasShutter:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get HasShutter");
                            returnValue = m_Camera.HasShutter;
                            break;
                        }

                    case CamPropertyType.ImageReady:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageReady");
                            returnValue = m_Camera.ImageReady;
                            break;
                        }

                    case CamPropertyType.IsPulseGuiding:
                        {
                            m_IsPulseGuidingSupported = false;
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get IsPulseGuiding");
                            returnValue = m_Camera.IsPulseGuiding;
                            m_IsPulseGuidingSupported = true; // Command works properly and doesn't cause a not implemented exception
                            break;
                        }

                    case CamPropertyType.FastReadout:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get FastReadout");
                            returnValue = m_Camera.FastReadout;
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                LogMsg(p_Name, MessageLevel.msgOK, returnValue.ToString());
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }
        private string CameraPropertyTestString(CamPropertyType p_Type, string p_Name, int p_MaxLength, bool p_Mandatory)
        {
            string returnValue = "";
            try
            {
                switch (p_Type)
                {
                    case CamPropertyType.Description:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get Description");
                            returnValue = m_Camera.Description;
                            break;
                        }

                    case CamPropertyType.SensorName:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get SensorName");
                            returnValue = m_Camera.SensorName;
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue == "":
                        {
                            LogMsg(p_Name, MessageLevel.msgOK, "The driver returned an empty string");
                            break;
                        }

                    default:
                        {
                            if (Strings.Len(returnValue) <= p_MaxLength)
                                LogMsg(p_Name, MessageLevel.msgOK, returnValue);
                            else
                                LogMsg(p_Name, MessageLevel.msgError, "String exceeds " + p_MaxLength + " characters maximum length - " + returnValue);
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }
        private void CameraPropertyWriteTest(CamPropertyType p_Type, string p_Property, int p_TestOK)
        {

            // NOTE: Out of range values should not be tested here but later when the exposure is actually taken

            try // OK value first
            {
                switch (p_Type)
                {
                    case CamPropertyType.NumX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set NumX");
                            m_Camera.NumX = p_TestOK;
                            break;
                        }

                    case CamPropertyType.NumY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set NumY");
                            m_Camera.NumY = p_TestOK;
                            break;
                        }

                    case CamPropertyType.StartX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set StartX");
                            m_Camera.StartX = p_TestOK;
                            break;
                        }

                    case CamPropertyType.StartY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set StartY");
                            m_Camera.StartY = p_TestOK;
                            break;
                        }
                }
                LogMsg(p_Property + " write", MessageLevel.msgOK, "Successfully wrote " + p_TestOK);
            }
            catch (Exception ex)
            {
                HandleException(p_Property + " write", MemberType.Property, Required.MustBeImplemented, ex, "Can't set legal value: " + p_TestOK.ToString());
            }
        }

        public override void CheckMethods()
        {
            int l_i, l_j, l_MaxBinX, l_MaxBinY;
            // AbortExposure - Mandatory
            try
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState");
                m_CameraState = m_Camera.CameraState;
                switch (m_CameraState)
                {
                    case CameraState.Idle // All is OK so test a call to AbortExposure
                   :
                        {
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call AbortExposure");
                                m_Camera.AbortExposure();
                                if (m_CanAbortExposure)
                                    LogMsg("AbortExposure", MessageLevel.msgOK, "No exception generated when camera is already idle");
                                else
                                    LogMsg("AbortExposure", MessageLevel.msgError, "CanAbortExposure is false but no exception is generated when AbortExposure is called");
                            }
                            catch (COMException)
                            {
                                if (m_CanAbortExposure)
                                    LogMsg("AbortExposure", MessageLevel.msgError, EX_COM + "exception incorrectly generated when camera is idle");
                                else
                                    LogMsg("AbortExposure", MessageLevel.msgOK, "CanAbortExposure is false and COM exception correctly generated");
                            }
                            catch (Exception)
                            {
                                if (m_CanAbortExposure)
                                    LogMsg("AbortExposure", MessageLevel.msgError, EX_NET + "exception incorrectly generated when camera is idle");
                                else
                                    LogMsg("AbortExposure", MessageLevel.msgOK, "CanAbortExposure is false and .NET exception correctly generated");
                            }

                            break;
                        }

                    default:
                        {
                            LogMsg("AbortExposure", MessageLevel.msgError, EX_COM + "camera is not idle, further AbortExposure tests skipped: " + m_CameraState.ToString());
                            break;
                        }
                }
            }
            catch (COMException)
            {
                LogMsg("AbortExposure", MessageLevel.msgError, EX_COM + "exception generated when reading camera state, further AbortExposure tests skipped");
            }
            catch (Exception)
            {
                LogMsg("AbortExposure", MessageLevel.msgError, EX_NET + "exception generated when reading camera state, further AbortExposure tests skipped");
            }
            // PulseGuide
            if (m_CanPulseGuide)
            {
                try
                {
                    CameraPulseGuideTest(GuideDirection.North); if (cancellationToken.IsCancellationRequested)
                        return;
                    CameraPulseGuideTest(GuideDirection.South); if (cancellationToken.IsCancellationRequested)
                        return;
                    CameraPulseGuideTest(GuideDirection.East); if (cancellationToken.IsCancellationRequested)
                        return;
                    CameraPulseGuideTest(GuideDirection.West); if (cancellationToken.IsCancellationRequested)
                        return;
                }
                catch (COMException ex)
                {
                    LogMsg("PulseGuide", MessageLevel.msgError, EX_COM + "CanPulseGuide is true but exception generated when calling method - " + ex.ToString());
                }
                catch (Exception ex)
                {
                    LogMsg("PulseGuide", MessageLevel.msgError, EX_NET + "CanPulseGuide is true but exception generated when calling method - " + ex.ToString());
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call PulseGuide - North");
                    m_Camera.PulseGuide(GuideDirection.North, 0);
                    LogMsg("PulseGuide", MessageLevel.msgError, "CanPulseGuide is false but no exception generated when calling method");
                }
                catch (COMException)
                {
                    LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and exception correctly generated when calling method");
                }
                catch (MethodNotImplementedException)
                {
                    LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and PulseGuide is not implemented in this driver");
                }
                catch (Exception)
                {
                    LogMsg("PulseGuide", MessageLevel.msgOK, "CanPulseGuide is false and exception correctly generated when calling method");
                }

            // StopExposure
            try
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState");
                m_CameraState = m_Camera.CameraState;
                switch (m_CameraState)
                {
                    case CameraState.Idle // all is OK so test that a call to StopExposure doesn't generate an exception
                   :
                        {
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call StopExposure");
                                m_Camera.StopExposure();
                                if (m_CanStopExposure)
                                    LogMsg("StopExposure", MessageLevel.msgOK, "No exception generated when camera is already idle");
                                else
                                    LogMsg("StopExposure", MessageLevel.msgError, "CanStopExposure is false but no exception is generated when StopExposure is called");
                            }
                            catch (COMException)
                            {
                                if (m_CanStopExposure)
                                    LogMsg("StopExposure", MessageLevel.msgError, EX_COM + "exception incorrectly generated when camera is idle");
                                else
                                    LogMsg("StopExposure", MessageLevel.msgOK, "CanStopExposure is false and COM exception correctly generated");
                            }
                            catch (Exception)
                            {
                                if (m_CanStopExposure)
                                    LogMsg("StopExposure", MessageLevel.msgError, EX_NET + "exception incorrectly generated when camera is idle");
                                else
                                    LogMsg("StopExposure", MessageLevel.msgOK, "CanStopExposure is false and .NET exception correctly generated");
                            }

                            break;
                        }

                    default:
                        {
                            LogMsg("StopExposure", MessageLevel.msgError, EX_COM + "camera is not idle, further StopExposure tests skipped: " + m_CameraState.ToString());
                            break;
                        }
                }
            }
            catch (COMException)
            {
                LogMsg("StopExposure", MessageLevel.msgError, EX_COM + "exception generated when reading camera state, further StopExposure tests skipped");
            }
            catch (Exception)
            {
                LogMsg("StopExposure", MessageLevel.msgError, EX_NET + "exception generated when reading camera state, further StopExposure tests skipped");
            }

            // Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogMsg("", MessageLevel.msgAlways, ""); // Insert a blank line if required
            if (settings.CameraMaxBinX > 0)
            {
                l_MaxBinX = settings.CameraMaxBinX;
                LogMsg("StartExposure", MessageLevel.msgComment, string.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", l_MaxBinX, m_MaxBinX));
            }
            else
                l_MaxBinX = m_MaxBinX;
            if (settings.CameraMaxBinY > 0)
            {
                l_MaxBinY = settings.CameraMaxBinY;
                LogMsg("StartExposure", MessageLevel.msgComment, string.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", l_MaxBinY, m_MaxBinY));
            }
            else
                l_MaxBinY = m_MaxBinY;

            if ((settings.CameraMaxBinX > m_MaxBinX) | (settings.CameraMaxBinY > m_MaxBinY))
                LogMsg("", MessageLevel.msgAlways, ""); // Insert a blank line if required
            if (settings.CameraMaxBinX > m_MaxBinX)
                LogMsg("StartExposure", MessageLevel.msgComment, string.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", l_MaxBinX, m_MaxBinX));
            if (settings.CameraMaxBinY > m_MaxBinY)
                LogMsg("StartExposure", MessageLevel.msgComment, string.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", l_MaxBinY, m_MaxBinY));

            // StartExposure - Confirm that correct operation occurs
            int l_BinX, l_BinY;
            if (m_CanAsymmetricBin)
            {
                for (l_BinY = 1; l_BinY <= l_MaxBinY; l_BinY++)
                {
                    for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                    {
                        CameraExposure("Take image full frame " + l_BinX + " x " + l_BinY + " bin", l_BinX, l_BinY, 0, 0, m_CameraXSize / l_BinX, m_CameraYSize / l_BinY, 2.0, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
            }
            else
                for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                {
                    CameraExposure("Take image full frame " + l_BinX + " x " + l_BinX + " bin", l_BinX, l_BinX, 0, 0, m_CameraXSize / l_BinX, m_CameraYSize / l_BinX, 2.0, "");
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }

            // StartExposure - Confirm error cases
            LogMsg("", MessageLevel.msgAlways, "");
            LogMsg("StartExposure error cases", MessageLevel.msgAlways, "");

            // StartExposure - Negative time
            CameraExposure("", 1, 1, 0, 0, m_CameraXSize, m_CameraYSize, -1.0, "negative duration"); if (cancellationToken.IsCancellationRequested)
                return; // Test that negative duration generates an error

            // StartExposure - Invalid Bin values
            for (l_i = 1; l_i <= l_MaxBinX; l_i++)
            {
                for (l_j = 1; l_j <= l_MaxBinY; l_j++)
                {
                    if (m_CanAsymmetricBin)
                    {
                        CameraExposure("", l_i, l_j, 0, 0, System.Convert.ToInt32((m_CameraXSize / (double)l_i) + 1), System.Convert.ToInt32(m_CameraYSize / (double)l_j), 0.1, "X size larger than binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // X size too large for binned size
                        CameraExposure("", l_i, l_j, 0, 0, System.Convert.ToInt32(m_CameraXSize / (double)l_i), System.Convert.ToInt32((m_CameraYSize / (double)l_j) + 1), 0.1, "Y size larger than binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // Y size too large for binned size
                        CameraExposure("", l_i, l_j, System.Convert.ToInt32((m_CameraXSize / (double)l_i) + 1), 0, System.Convert.ToInt32(m_CameraXSize / (double)l_i), System.Convert.ToInt32(m_CameraYSize / (double)l_j), 0.1, "X start outside binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // X start outside binned chip dimensions
                        CameraExposure("", l_i, l_j, 0, System.Convert.ToInt32((m_CameraYSize / (double)l_j) + 1), System.Convert.ToInt32(m_CameraXSize / (double)l_i), System.Convert.ToInt32(m_CameraYSize / (double)l_j), 0.1, "Y start outside binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // Y start outside binned chip dimensions
                    }
                    else if (l_i == l_j)
                    {
                        CameraExposure("", l_i, l_j, 0, 0, System.Convert.ToInt32((m_CameraXSize / (double)l_i) + 1), System.Convert.ToInt32(m_CameraYSize / (double)l_j), 0.1, "X size larger than binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // X size too large for binned size
                        CameraExposure("", l_i, l_j, 0, 0, System.Convert.ToInt32(m_CameraXSize / (double)l_i), System.Convert.ToInt32((m_CameraYSize / (double)l_j) + 1), 0.1, "Y size larger than binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // Y size too large for binned size
                        CameraExposure("", l_i, l_j, System.Convert.ToInt32((m_CameraXSize / (double)l_i) + 1), 0, System.Convert.ToInt32(m_CameraXSize / (double)l_i), System.Convert.ToInt32(m_CameraYSize / (double)l_j), 0.1, "X start outside binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // X start outside binned chip dimensions
                        CameraExposure("", l_i, l_j, 0, System.Convert.ToInt32((m_CameraYSize / (double)l_j) + 1), System.Convert.ToInt32(m_CameraXSize / (double)l_i), System.Convert.ToInt32(m_CameraYSize / (double)l_j), 0.1, "Y start outside binned chip size, Bin " + l_i + "x" + l_j); if (cancellationToken.IsCancellationRequested)
                            return; // Y start outside binned chip dimensions
                    }
                }
            }
        }
        private void CameraExposure(string p_Description, int p_BinX, int p_BinY, int p_StartX, int p_StartY, int p_NumX, int p_NumY, double p_Duration, string p_ExpectedErrorMessage)
        {
            string l_NumPlanes, l_VariantType, l_PercentCompletedMessage;
            bool l_ExposeOK, imageReadyTooEarly = false; // Flag to determine whether we were successful or something failed
            DateTime l_StartTime, l_StartTimeUTC, l_EndTime;
            short l_PercentCompleted;

            if (p_Description != "")
            {
                LogMsg("", MessageLevel.msgAlways, ""); // Blank Line
                LogMsg(p_Description, MessageLevel.msgAlways, "");
            }
            try
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinX");
                m_Camera.BinX = (short)p_BinX;
                if (settings.DisplayMethodCalls)
                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinY");
                m_Camera.BinY = (short)p_BinY;
                try
                {
                    l_ExposeOK = false; // Start off by assuming the worst
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set StartX");
                    m_Camera.StartX = p_StartX;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set StartY");
                    m_Camera.StartY = p_StartY;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set NumX");
                    m_Camera.NumX = p_NumX;
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set NumY");
                    m_Camera.NumY = p_NumY;
                    try
                    {
                        Status(StatusType.staAction, "Start " + p_Duration.ToString() + " second synchronous exposure");

                        // Initiate exposure
                        l_StartTime = DateTime.Now;
                        l_StartTimeUTC = DateTime.UtcNow;
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call StartExposure");
                        m_Camera.StartExposure(p_Duration, true);

                        if (p_ExpectedErrorMessage == "")
                        {
                            l_EndTime = DateTime.Now;

                            // Test whether we have a synchronous or asynchronous camera
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageReady and CameraState");
                            if (m_Camera.ImageReady & (m_Camera.CameraState == CameraState.Idle))
                            {
                                if (l_EndTime.Subtract(l_StartTime).TotalSeconds >= p_Duration)
                                {
                                    LogMsg("StartExposure", MessageLevel.msgOK, "Synchronous exposure found OK: " + p_Duration + " seconds");
                                    CameraTestLast(p_Duration, l_StartTimeUTC);
                                }
                                else
                                    LogMsg("StartExposure", MessageLevel.msgError, "Synchronous exposure found but image was returned before exposure time was complete");
                            }
                            else
                            {
                                Status(StatusType.staAction, "Waiting for exposure to start");

                                // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageReady");
                                imageReadyTooEarly = System.Convert.ToBoolean(m_Camera.ImageReady);

                                // Wait for exposing state
                                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState multiple times");
                                do
                                {
                                    WaitFor(CAMERA_SLEEP_TIME);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                }
                                while (!((m_Camera.CameraState == CameraState.Exposing) | (m_Camera.CameraState == CameraState.Error)));

                                // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageReady");
                                imageReadyTooEarly = m_Camera.ImageReady;

                                // Wait for the exposing state to finish
                                l_StartTime = DateTime.Now;
                                l_StartTimeUTC = DateTime.UtcNow;
                                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState, InterfaceVersion and PercentCompleted multiple times...");
                                do
                                {
                                    l_PercentCompletedMessage = "Not present in a V1 driver";
                                    if (m_Camera.InterfaceVersion > 1)
                                    {
                                        try
                                        {
                                            l_PercentCompleted = m_Camera.PercentCompleted;
                                            l_PercentCompletedMessage = l_PercentCompleted.ToString();
                                        }
                                        catch (COMException ex)
                                        {
                                            switch (ex.ErrorCode)
                                            {
                                                case int i when i == ErrorCodes.NotImplemented:
                                                    {
                                                        l_PercentCompletedMessage = "COM - Not implemented";
                                                        break;
                                                    }

                                                case int j when j == ErrorCodes.InvalidOperationException:
                                                    {
                                                        l_PercentCompletedMessage = "COM - Invalid operation";
                                                        break;
                                                    }

                                                default:
                                                    {
                                                        l_PercentCompletedMessage = "COM - Exception: " + ex.Message + " 0x" + ex.ErrorCode.ToString("X8");
                                                        break;
                                                    }
                                            }
                                        }
                                        catch (ASCOM.InvalidOperationException)
                                        {
                                            l_PercentCompletedMessage = ".NET - Invalid operation";
                                        }
                                        catch (PropertyNotImplementedException)
                                        {
                                            l_PercentCompletedMessage = "Not implemented";
                                        }
                                        catch (Exception ex)
                                        {
                                            l_PercentCompletedMessage = "Exception: " + ex.Message;
                                        }
                                    }

                                    Status(StatusType.staAction, "Waiting for " + p_Duration.ToString() + " second exposure to complete: " + Conversion.Int(DateTime.Now.Subtract(l_StartTime).TotalSeconds) + ",   PercentComplete: " + l_PercentCompletedMessage);
                                    WaitFor(CAMERA_SLEEP_TIME);
                                    if (cancellationToken.IsCancellationRequested) return;
                                }
                                while (m_Camera.CameraState == CameraState.Exposing)// Initialise PercentCompleted message// Operation completed OK// Not implemented// Not valid at this time// Something bad happened!// Not valid at this time// Not implemented// Something bad happened!
    ;

                                // Wait for camera to become idle
                                l_EndTime = DateTime.Now;
                                Status(StatusType.staAction, "Waiting for camera idle state, reading/downloading image");
                                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState multiple times");
                                do
                                {
                                    WaitFor(CAMERA_SLEEP_TIME);
                                    if (cancellationToken.IsCancellationRequested) return;
                                }
                                while (!((m_Camera.CameraState == CameraState.Idle) | (m_Camera.CameraState == CameraState.Error)));

                                // Wait for image to become ready
                                Status(StatusType.staAction, "Waiting for image ready");
                                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get CameraState multiple times");
                                do
                                {
                                    WaitFor(CAMERA_SLEEP_TIME);
                                    if (cancellationToken.IsCancellationRequested) return;
                                }
                                while (!((m_Camera.ImageReady) | (m_Camera.CameraState == CameraState.Error)));

                                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageReady");
                                if (m_Camera.ImageReady)
                                {
                                    LogMsg("StartExposure", MessageLevel.msgOK, "Asynchronous exposure found OK: " + p_Duration + " seconds");
                                    CameraTestLast(p_Duration, l_StartTimeUTC);
                                }
                                else
                                    LogMsg("StartExposure", MessageLevel.msgError, "Camera state is CameraError");
                            }

                            // Display a warning if ImageReady was set too early
                            if (imageReadyTooEarly)
                            {
                                LogMsg("StartExposure", MessageLevel.msgIssue, "ImageReady was set True before the camera completed its exposure.");
                            }
                            // Camera exposed OK and didn't generate an exception
                            else
                            {
                                l_ExposeOK = true;
                            }
                            LogMsg("StartExposure", MessageLevel.msgDebug, $"Camera exposed image OK: {l_ExposeOK}");

                            // Check image array dimensions
                            try
                            {
                                // Retrieve the image array
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageArray");
                                sw.Restart();
                                m_ImageArray = (Array)m_Camera.ImageArray;
                                sw.Stop();
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "Get ImageArray completed in " + sw.ElapsedMilliseconds + "ms");

                                // Examine the returned array
                                if ((m_ImageArray.GetLength(0) == p_NumX) & (m_ImageArray.GetLength(1) == p_NumY))
                                {
                                    if (m_ImageArray.GetType().ToString() == "System.Int32[,]" | m_ImageArray.GetType().ToString() == "System.Int32[,,]")
                                    {
                                        if (m_ImageArray.Rank == 2)
                                            l_NumPlanes = "1 plane";
                                        else
                                        {
                                            l_NumPlanes = "1 plane";
                                            if (m_ImageArray.GetUpperBound(2) > 0)
                                                l_NumPlanes = System.Convert.ToString(m_ImageArray.GetUpperBound(2) + 1) + " planes";
                                        }
                                        LogMsg("ImageArray", MessageLevel.msgOK, "Successfully read 32 bit integer array (" + l_NumPlanes + ") " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1) + " pixels");
                                    }
                                    else
                                        LogMsg("ImageArray", MessageLevel.msgError, "Expected 32 bit integer array, actually got: " + m_ImageArray.GetType().ToString());
                                }
                                else if ((m_ImageArray.GetLength(0) == p_NumY) & (m_ImageArray.GetLength(1) == p_NumX))
                                    LogMsg("ImageArray", MessageLevel.msgError, "Camera image dimensions swapped, expected values: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
                                else
                                    LogMsg("ImageArray", MessageLevel.msgError, "Camera image does not have the expected dimensions of: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
                            }
                            catch (COMException ex)
                            {
                                LogMsg("StartExposure", MessageLevel.msgError, EX_COM + "exception when reading ImageArray" + ex.ToString());
                            }
                            catch (Exception ex)
                            {
                                LogMsg("StartExposure", MessageLevel.msgError, EX_NET + "exception when reading ImageArray" + ex.ToString());
                            }

                            m_ImageArray = null;
                            m_ImageArrayVariant = null;
                            GC.Collect();

                            // Check image array variant dimensions
                            Array imageArrayObject;
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get ImageArrayVariant");
                                sw.Restart();
                                imageArrayObject = (Array)m_Camera.ImageArrayVariant;
                                sw.Stop();
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "Get ImageArrayVariant completed in " + sw.ElapsedMilliseconds + "ms");
                                sw.Restart();
                                m_ImageArrayVariant = (Array)imageArrayObject;
                                sw.Stop();
                                if (settings.DisplayMethodCalls)
                                    LogMsg("ConformanceCheck", MessageLevel.msgComment, "Conversion to Array completed in " + sw.ElapsedMilliseconds + "ms");
                                if ((m_ImageArrayVariant.GetLength(0) == p_NumX) & (m_ImageArrayVariant.GetLength(1) == p_NumY))
                                {
                                    if (m_ImageArrayVariant.GetType().ToString() == "System.Object[,]" | m_ImageArrayVariant.GetType().ToString() == "System.Object[,,]")
                                    {
                                        if (m_ImageArrayVariant.Rank == 2)
                                        {
                                            l_NumPlanes = "1 plane";
                                            l_VariantType = ((object[,])m_ImageArrayVariant)[0, 0].GetType().ToString();
                                        }
                                        else
                                        {
                                            l_NumPlanes = "1 plane";
                                            if (m_ImageArrayVariant.GetUpperBound(2) > 0)
                                            {
                                                l_NumPlanes = System.Convert.ToString(m_ImageArrayVariant.GetUpperBound(2) + 1) + " planes";
                                                l_VariantType = ((object[,,])m_ImageArrayVariant)[0, 0, 0].GetType().ToString();
                                            }
                                            else
                                                l_VariantType = ((object[,])m_ImageArrayVariant)[0, 0].GetType().ToString();
                                        }
                                        LogMsg("ImageArrayVariant", MessageLevel.msgOK, "Successfully read variant array (" + l_NumPlanes + ") with " + l_VariantType + " elements " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1) + " pixels");
                                    }
                                    else
                                        LogMsg("ImageArrayVariant", MessageLevel.msgError, "Expected variant array, actually got: " + m_ImageArrayVariant.GetType().ToString());
                                }
                                else if ((m_ImageArrayVariant.GetLength(0) == p_NumY) & (m_ImageArrayVariant.GetLength(1) == p_NumX))
                                    LogMsg("ImageArrayVariant", MessageLevel.msgError, "Camera image dimensions swapped, expected values: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1));
                                else
                                    LogMsg("ImageArrayVariant", MessageLevel.msgError, "Camera image does not have the expected dimensions of: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1));
                            }
                            catch (COMException ex)
                            {
                                LogMsg("ImageArrayVariant", MessageLevel.msgError, EX_COM + "exception when reading ImageArrayVariant" + ex.ToString());
                            }
                            catch (Exception ex)
                            {
                                LogMsg("ImageArrayVariant", MessageLevel.msgError, EX_NET + "exception when reading ImageArrayVariant" + ex.ToString());
                            }

                            // Release large image objects from memory
                            m_ImageArrayVariant = null;
                            imageArrayObject = null;
                            GC.Collect();
                        }
                        else
                        {
                            LogMsg("StartExposure", MessageLevel.msgComment, "Test: " + p_ExpectedErrorMessage);
                            LogMsg("StartExposure", MessageLevel.msgError, "Expected an exception and didn't get one - BinX:" + p_BinX + " BinY:" + p_BinY + " StartX:" + p_StartX + " StartY:" + p_StartY + " NumX:" + p_NumX + " NumY:" + p_NumY);
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call StopExposure");
                            try
                            {
                                m_Camera.StopExposure();
                            }
                            catch (Exception)
                            {
                            } // Try and do some clean up
                            if (settings.DisplayMethodCalls)
                                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call AbortExposure");
                            try
                            {
                                m_Camera.AbortExposure();
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                    catch (COMException ex)
                    {
                        if (p_ExpectedErrorMessage != "")
                            LogMsg("StartExposure", MessageLevel.msgOK, "Exception correctly generated for " + p_ExpectedErrorMessage);
                        else
                        {
                            LogMsg("", MessageLevel.msgInfo, p_ExpectedErrorMessage);
                            LogMsg("StartExposure", MessageLevel.msgError, EX_COM + "exception generated when exposing, further StartExposure tests skipped - " + ex.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        if (p_ExpectedErrorMessage != "")
                            LogMsg("StartExposure", MessageLevel.msgOK, "Exception correctly generated for " + p_ExpectedErrorMessage);
                        else
                            LogMsg("StartExposure", MessageLevel.msgError, EX_NET + "exception generated when exposing, further StartExposure tests skipped - " + ex.ToString());
                    }
                }
                catch (COMException ex)
                {
                    LogMsg("StartExposure", MessageLevel.msgInfo, "COM Exception: " + ex.Message);
                    LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");
                }
                catch (Exception ex)
                {
                    LogMsg("StartExposure", MessageLevel.msgInfo, ".NET Exception: " + ex.Message);
                    LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");
                }
            }
            catch (COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    case int i when i == ErrorCodes.InvalidValue:
                    case int j when j == g_ExInvalidValue1:
                    case int k when k == g_ExInvalidValue2:
                    case int l when l == g_ExInvalidValue3:
                    case int m when m == g_ExInvalidValue4:
                    case int n when n == g_ExInvalidValue5:
                    case int o when o == g_ExInvalidValue6:
                        {
                            LogMsg("StartExposure", MessageLevel.msgInfo, ex.Message);
                            break;
                        }

                    default:
                        {
                            LogMsg("StartExposure", MessageLevel.msgInfo, "COM Exception: " + ex.Message);
                            LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");
                            break;
                        }
                }
            }
            catch (InvalidValueException ex)
            {
                LogMsg("BinXY Write", MessageLevel.msgInfo, "Exposure skipped because BinX or BinY cannot be set. Exception message:");
                LogMsg("BinXY Write", MessageLevel.msgComment, ex.Message);
            }
            catch (Exception ex)
            {
                LogMsg("StartExposure", MessageLevel.msgInfo, ".NET Exception: " + ex.Message);
                LogMsg("StartExposure", MessageLevel.msgError, "Exception generated when setting camera properties, further StartExposure tests skipped");
            }

            Status(StatusType.staAction, "");
        }
        private void CameraTestLast(double p_Duration, DateTime p_Start)
        {
            DateTime l_StartTime;

            // LastExposureDuration
            try
            {
                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get LastExposureDuration");
                m_LastExposureDuration = m_Camera.LastExposureDuration;
                if ((Math.Abs(m_LastExposureDuration - p_Duration) / p_Duration) < 0.02)
                    LogMsg("LastExposureDuration", MessageLevel.msgOK, "LastExposureDuration is: " + m_LastExposureDuration + " seconds");
                else
                    LogMsg("LastExposureDuration", MessageLevel.msgIssue, "LastExposureDuration is not the same as image duration: " + m_LastExposureDuration + " " + p_Duration);
            }
            catch (Exception ex)
            {
                HandleException("LastExposureDuration", MemberType.Property, Required.Optional, ex, "");
            }

            // LastExposurestartTime
            try // Confirm that it can be read
            {
                if (settings.DisplayMethodCalls) LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get LastExposureStartTime");
                m_LastExposureStartTime = m_Camera.LastExposureStartTime;
                int l_i;
                // Confirm that the format is as expected
                bool l_FormatOK;
                l_FormatOK = true;
                if (Strings.Len(m_LastExposureStartTime) >= 19)
                {
                    for (l_i = 1; l_i <= 19; l_i++)
                    {
                        switch (l_i)
                        {
                            case 5:
                            case 8 // "-"
                           :
                                {
                                    if (Strings.Mid(m_LastExposureStartTime, l_i, 1) != "-")
                                        l_FormatOK = false;
                                    break;
                                }

                            case 11 // "T"
                     :
                                {
                                    if (Strings.Mid(m_LastExposureStartTime, l_i, 1) != "T")
                                        l_FormatOK = false;
                                    break;
                                }

                            case 14:
                            case 17 // ":"
                     :
                                {
                                    if (Strings.Mid(m_LastExposureStartTime, l_i, 1) != ":")
                                        l_FormatOK = false;
                                    break;
                                }

                            default:
                                {
                                    if (!Information.IsNumeric(Strings.Mid(m_LastExposureStartTime, l_i, 1)))
                                        l_FormatOK = false;
                                    break;
                                }
                        }
                    }
                    if (l_FormatOK)
                    {
                        try // Confirm that it parses as a valid date and check for correct value
                        {
                            l_StartTime = DateTime.Parse(m_LastExposureStartTime);
                            if (p_Start.Subtract(l_StartTime).TotalSeconds < 2.0)
                                LogMsg("LastExposureStartTime", MessageLevel.msgOK, "LastExposureStartTime is correct to within 2 seconds: " + m_LastExposureStartTime + " UTC");
                            else
                                LogMsg("LastExposureStartTime", MessageLevel.msgIssue, "LastExposureStartTime is more than 2 seconds inaccurate : " + m_LastExposureStartTime + ", expected: " + Strings.Format(p_Start, "yyyy-MM-ddTHH:mm:ss") + " UTC");
                        }
                        catch (COMException ex)
                        {
                            LogMsg("LastExposureStartTime", MessageLevel.msgError, EX_COM + "parsing LastExposureStartTime - " + ex.Message + " " + ex.ErrorCode.ToString("X8"));
                        }
                        catch (Exception ex)
                        {
                            LogMsg("LastExposureStartTime", MessageLevel.msgError, EX_NET + "parsing LastExposureStartTime - " + ex.Message + " " + m_LastExposureStartTime);
                        }
                    }
                    else
                        LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime not in the expected format yyyy-mm-ddThh:mm:ss - " + m_LastExposureStartTime);
                }
                else if (m_LastExposureStartTime == "")
                    LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime has returned an empty string - expected yyyy-mm-ddThh:mm:ss");
                else
                    LogMsg("LastExposureStartTime", MessageLevel.msgError, "LastExposureStartTime is less than 19 characters - expected yyyy-mm-ddThh:mm:ss - " + m_LastExposureStartTime);
            }
            catch (Exception ex)
            {
                HandleException("LastExposureStartTime", MemberType.Property, Required.Optional, ex, "");
            }
        }
        private void CameraPulseGuideTest(GuideDirection p_Direction)
        {
            DateTime l_StartTime, l_EndTime;
            // Dim pulseGuideStatus As PulseGuideState

            l_StartTime = DateTime.Now;
            Status(StatusType.staAction, "Start " + CAMERA_PULSE_DURATION / (double)1000 + " second pulse guide " + p_Direction.ToString());
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, $"About to call PulseGuide - {p_Direction}");
            m_Camera.PulseGuide(p_Direction, CAMERA_PULSE_DURATION); // Start a 2 second pulse
            l_EndTime = DateTime.Now;

            if (m_IsPulseGuidingSupported)
            {
                if (l_EndTime.Subtract(l_StartTime).TotalMilliseconds < (CAMERA_PULSE_DURATION - 500))
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get IsPulseGuiding");
                    if (m_Camera.IsPulseGuiding)
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get IsPulseGuiding multiple times");
                        do
                        {
                            WaitFor(SLEEP_TIME);
                            if (cancellationToken.IsCancellationRequested)
                                return;
                        }
                        while (!((!m_Camera.IsPulseGuiding) | (DateTime.Now.Subtract(l_StartTime).TotalMilliseconds > 3000))); // Wait for up to 3 seconds

                        if (settings.DisplayMethodCalls)
                            LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get IsPulseGuiding");
                        if (!m_Camera.IsPulseGuiding)
                            LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgOK, "Asynchronous pulse guide found OK");
                        else
                            LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding is TRUE beyond expected time of 2 seconds");
                    }
                    else
                        LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgIssue, "Asynchronous pulse guide expected but IsPulseGuiding has returned FALSE");
                }
                else
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to get IsPulseGuiding");
                    if (!m_Camera.IsPulseGuiding)
                        LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgOK, "Synchronous pulse guide found OK");
                    else
                        LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgIssue, "Synchronous pulse guide expected but IsPulseGuiding has returned TRUE");
                }
            }
            else
                switch (l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION)
                {
                    case object _ when l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION > CAMERA_PULSE_TOLERANCE // Duration was more than 0.5 seconds longer than expected
                   :
                        {
                            LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgIssue, "Synchronous pulse guide longer than expected " + (CAMERA_PULSE_DURATION) / (double)1000 + " seconds: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
                            break;
                        }

                    case object _ when l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION < 20 // Duration was more than 20ms shorter than expected
             :
                        {
                            LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgIssue, "Synchronous pulse guide shorter than expected " + (CAMERA_PULSE_DURATION) / (double)1000 + " seconds: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
                            break;
                        }

                    default:
                        {
                            LogMsg("PulseGuide " + p_Direction.ToString(), MessageLevel.msgOK, "Synchronous pulse guide found OK: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
                            break;
                        }
                }
        }

        public override void CheckPerformance()
        {
            CameraPerformanceTest(CameraPerformance.CameraState, "CameraState");
            CameraPerformanceTest(CameraPerformance.CCDTemperature, "CCDTemperature");
            CameraPerformanceTest(CameraPerformance.CoolerPower, "CoolerPower");
            if (m_CanSetCCDTemperature)
                CameraPerformanceTest(CameraPerformance.HeatSinkTemperature, "HeatSinkTemperature");
            CameraPerformanceTest(CameraPerformance.ImageReady, "ImageReady");
            if (m_CanPulseGuide)
                CameraPerformanceTest(CameraPerformance.IsPulseGuiding, "IsPulseGuiding");
            Status(StatusType.staAction, "Exposure for ImageArray Test");
            Status(StatusType.staStatus, "Start");
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinX");
            m_Camera.BinX = 1;
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set BinY");
            m_Camera.BinY = 1;
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set StartX");
            m_Camera.StartX = 0;
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set StartY");
            m_Camera.StartY = 0;
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set MaxBinX");
            m_Camera.NumX = m_Camera.MaxBinX;
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set MaxBinY");
            m_Camera.NumY = m_Camera.MaxBinY;
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call StartExposure");
            m_Camera.StartExposure(1, true); // 1 second exposure

            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call ImageReady multiple times");
            do
                Status(StatusType.staStatus, "Waiting for ImageReady");
            while (!m_Camera.ImageReady);
            Status(StatusType.staStatus, "Finished");
            CameraPerformanceTest(CameraPerformance.ImageArray, "ImageArray");
            CameraPerformanceTest(CameraPerformance.ImageArrayVariant, "ImageArrayVariant");
        }
        private void CameraPerformanceTest(CameraPerformance p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate;
            Status(StatusType.staAction, p_Name);
            try
            {
                l_StartTime = DateTime.Now;
                l_Count = 0.0;
                l_LastElapsedTime = 0.0;
                do
                {
                    l_Count += 1.0;
                    switch (p_Type)
                    {
                        case CameraPerformance.CameraState:
                            {
                                m_CameraState = m_Camera.CameraState;
                                break;
                            }

                        case CameraPerformance.CCDTemperature:
                            {
                                m_CCDTemperature = m_Camera.CCDTemperature;
                                break;
                            }

                        case CameraPerformance.CoolerPower:
                            {
                                m_CoolerPower = m_Camera.CoolerPower;
                                break;
                            }

                        case CameraPerformance.HeatSinkTemperature:
                            {
                                m_HeatSinkTemperature = m_Camera.HeatSinkTemperature;
                                break;
                            }

                        case CameraPerformance.ImageReady:
                            {
                                m_ImageReady = m_Camera.ImageReady;
                                break;
                            }

                        case CameraPerformance.IsPulseGuiding:
                            {
                                m_IsPulseGuiding = m_Camera.IsPulseGuiding;
                                break;
                            }

                        case CameraPerformance.ImageArray:
                            {
                                m_ImageArray = (Array)m_Camera.ImageArray;
                                break;
                            }

                        case CameraPerformance.ImageArrayVariant:
                            {
                                m_ImageArrayVariant = (Array)m_Camera.ImageArrayVariant;
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "Conform:PerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
                    {
                        Status(StatusType.staStatus, l_Count + " transactions in " + Strings.Format(l_ElapsedTime, "0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (!(l_ElapsedTime > PERF_LOOP_TIME));

                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case object _ when l_Rate > 10.0:
                        {
                            LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= l_Rate && l_Rate <= 10.0:
                        {
                            LogMsg(p_Name, MessageLevel.msgOK, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= l_Rate && l_Rate <= 2.0:
                        {
                            LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " + ex.ToString());
            }
        }

        public override void PostRunCheck()
        {
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call AbortExposure");
            if (m_CanAbortExposure)
            {
                try
                {
                    m_Camera.AbortExposure();
                }
                catch
                {
                }
            }
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to call StopExposure");
            if (m_CanStopExposure)
            {
                try
                {
                    m_Camera.StopExposure();
                }
                catch
                {
                }
            }
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set SetCCDTemperature");
            if (m_CanSetCCDTemperature)
            {
                try
                {
                    m_Camera.SetCCDTemperature = m_SetCCDTemperature;
                }
                catch
                {
                }
            }
            if (settings.DisplayMethodCalls)
                LogMsg("ConformanceCheck", MessageLevel.msgComment, "About to set CoolerOn");
            try
            {
                m_Camera.CoolerOn = m_CoolerOn;
            }
            catch
            {
            }
            LogMsg("PostRunCheck", MessageLevel.msgOK, "Camera returned to initial cooler temperature");
        }
    }

}

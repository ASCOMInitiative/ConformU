using ASCOM;
using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.DeviceInterface;
using Microsoft.Extensions.Logging.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ConformU
{

    internal class CameraTester : DeviceTesterBaseClass
    {
        #region Constants and variables

        const int CAMERA_PULSE_DURATION = 2000; // Duration of camera pulse guide test (ms)
        const int CAMERA_PULSE_TOLERANCE = 300; // Tolerance for acceptable performance (ms)

        const int MAX_BIN_X = 16; // Values of MaxBin above which warnings are given. Implemented to warn developers if they are returning "silly" values
        const int MAX_BIN_Y = 16;
        const double ABSOLUTE_ZERO_TEMPERATURE = -273.15; // Absolute zero (Celsius)
        const double BELOW_ABSOLUTE_ZERO_TEMPERATURE = -273.25; // Value (Celsius) below which CCD temperatures will be flagged as "silly", "un-physical" values.
        const double CAMERA_SETPOINT_INCREMENT = 5.0; // Amount by which the test temperature is decremented or incremented when finding the lowest and highest supported set points.
        const double BOILING_POINT_TEMPERATURE = 100.0; // Value above which CCD set point temperatures will be flagged as "silly" values
        const double MAX_CAMERA_REPORTED_TEMPERATURE = 1000.0; // Value above which the CCD reported temperature will be flagged as a "silly" value. It is higher than the MAX_CAMERA_SETPOINT_TEMPERATURE temperature because this value is not specified in the Interface Standard.
        const double CAMERA_LOW_SETPOINT_START_TEMPERATURE = 15.0; // Start temperature for determining minimum set point value.
        const double CAMERA_HIGH_SETPOINT_START_TEMPERATURE = 0.0; // Start temperature for determining maximum set point value.

        // Camera variables
        private bool m_CanAbortExposure, m_CanAsymmetricBin, m_CanGetCoolerPower, m_CanSetCCDTemperature, m_CanStopExposure, m_CanFastReadout;
        private bool m_CoolerOn, m_ImageReady;
        private int cameraXSize, cameraYSize;
        private short m_MaxBinX, m_MaxBinY, m_BinX, m_BinY;
        private double m_LastExposureDuration;
        private double m_SetCCDTemperature;
        private string m_LastExposureStartTime;
        private CameraState m_CameraState;
        private Array m_ImageArray, m_ImageArrayVariant;
        private bool m_IsPulseGuidingSupported; // Confirm that IsPulseGuiding command will work
        private bool m_CanPulseGuide;
        private bool m_IsPulseGuiding;
        // ICameraV2 properties
        private short m_Gain, m_GainMax, m_GainMin, m_PercentCompleted, m_ReadoutMode;
        private double m_ExposureMax, m_ExposureMin, m_ExposureResolution;
        private bool m_FastReadout, m_CanReadGain, m_CanReadGainMax, m_CanReadGainMin, m_CanReadGains, m_CanReadReadoutModes;
        private IList<string> m_Gains;
        private IList<string> m_ReadoutModes;
        private ASCOM.DeviceInterface.SensorType m_SensorType;
        private bool m_CanReadSensorType = false;
        private readonly Stopwatch sw = new();

        // ICameraV3 properties
        private int m_Offset, m_OffsetMax, m_OffsetMin;
        private bool m_CanReadOffset, m_CanReadOffsetMax, m_CanReadOffsetMin, m_CanReadOffsets;
        private IList<string> m_Offsets;
        private double m_SubExposureDuration;
        private GainOffsetMode m_OffsetMode = GainOffsetMode.Unknown;
        private GainOffsetMode m_GainMode = GainOffsetMode.Unknown;

        private ICameraV3 camera;

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

        public CameraTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, true, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls
        protected override void Dispose(bool disposing)
        {
            LogDebug("CameraTester.Dispose", "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {

                    Task.Run(() =>
                    {
                        LogDebug("CameraTester.Dispose", "About to dispose of camera...");
                        try { camera?.Dispose(); } catch { }
                        LogDebug("CameraTester.Dispose", "Camera disposed");
                    });

                    //try { camera?.Dispose(); } catch { }
                    LogDebug("CameraTester.Dispose", "About to set camera to null...");

                    camera = null;
                    LogDebug("CameraTester.Dispose", "Camera set to null");
                    LogDebug("CameraTester.Dispose", "About to release memory...");
                    try { ReleaseMemory(); } catch { }
                    LogDebug("CameraTester.Dispose", "Memory released");
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            Task.Run(() =>
            {
                LogDebug("CameraTester.Dispose", "About to dispose of Base...");
                base.Dispose(disposing);
                LogDebug("CameraTester.Dispose", "Base disposed");
            });
            disposedValue = true;
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(camera, DeviceTypes.Camera);
        }

        public new void CheckInitialise()
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
            base.CheckInitialise();
        }
        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        LogInfo("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        //m_Camera = new AlpacaCamera(settings.AlpacaConfiguration.AccessServiceType,
                        //    settings.AlpacaDevice.IpAddress,
                        //    settings.AlpacaDevice.IpPort,
                        //    settings.AlpacaDevice.AlpacaDeviceNumber,
                        //    settings.AlpacaConfiguration.StrictCasing,
                        //    settings.TraceAlpacaCalls ? logger : null);


                        camera = new AlpacaCamera(
                                                    settings.AlpacaConfiguration.AccessServiceType,
                                                    settings.AlpacaDevice.IpAddress,
                                                    settings.AlpacaDevice.IpPort,
                                                    settings.AlpacaDevice.AlpacaDeviceNumber,
                                                    settings.AlpacaConfiguration.EstablishConnectionTimeout,
                                                    settings.AlpacaConfiguration.StandardResponseTimeout,
                                                    settings.AlpacaConfiguration.LongResponseTimeout,
                                                    Globals.CLIENT_NUMBER_DEFAULT,
                                                    settings.AlpacaConfiguration.ImageArrayTransferType,
                                                    settings.AlpacaConfiguration.ImageArrayCompression,
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
                                camera = new CameraFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                camera = new Camera(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);

                baseClassDevice = camera; // Assign the driver to the base class

                LogInfo("CreateDevice", "Successfully created driver");
            }
            catch (COMException exCom) when (exCom.ErrorCode == REGDB_E_CLASSNOTREG)
            {
                LogDebug("CreateDevice", $"Error returned: {exCom.Message}\r\n{exCom}");

                throw new Exception($"The driver is not registered as a {(Environment.Is64BitProcess ? "64bit" : "32bit")} driver");
            }
            catch (Exception ex)
            {
                LogDebug("CreateDevice", $"Error returned: {ex.Message}\r\n{ex}");
                throw; // Re throw exception 
            }
        }

        public override bool Connected
        {
            get
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get Connected");
                return camera.Connected;
            }
            set
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set Connected");
                camera.Connected = value;

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
                                LogTestAndMessage("ConformanceCheck", "About to get CanAbortExposure");
                            m_CanAbortExposure = camera.CanAbortExposure;
                            LogOK(p_Name, m_CanAbortExposure.ToString());
                            break;
                        }

                    case CanType.tstCanAsymmetricBin:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanAsymmetricBin");
                            m_CanAsymmetricBin = camera.CanAsymmetricBin;
                            LogOK(p_Name, m_CanAsymmetricBin.ToString());
                            break;
                        }

                    case CanType.tstCanGetCoolerPower:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanGetCoolerPower");
                            m_CanGetCoolerPower = camera.CanGetCoolerPower;
                            LogOK(p_Name, m_CanGetCoolerPower.ToString());
                            break;
                        }

                    case CanType.tstCanPulseGuide:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanPulseGuide");
                            m_CanPulseGuide = camera.CanPulseGuide;
                            LogOK(p_Name, m_CanPulseGuide.ToString());
                            break;
                        }

                    case CanType.tstCanSetCCDTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanSetCCDTemperature");
                            m_CanSetCCDTemperature = camera.CanSetCCDTemperature;
                            LogOK(p_Name, m_CanSetCCDTemperature.ToString());
                            break;
                        }

                    case CanType.tstCanStopExposure:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanStopExposure");
                            m_CanStopExposure = camera.CanStopExposure;
                            LogOK(p_Name, m_CanStopExposure.ToString());
                            break;
                        }

                    case CanType.tstCanFastReadout:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanFastReadout");
                            m_CanFastReadout = camera.CanFastReadout;
                            LogOK(p_Name, m_CanFastReadout.ToString());
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
        public override void PreRunCheck()
        {
            int l_VStringPtr, l_V1, l_V2, l_V3;
            // Add a test for a back level version of the Camera simulator - just abandon this process if any errors occur
            if (settings.ComDevice.ProgId.ToUpper() == "CCDSIMULATOR.CAMERA")
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get Description");
                    l_VStringPtr = camera.Description.ToUpper().IndexOf("VERSION "); // Point at the start of the version string
                    if (l_VStringPtr >= 0)
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get Description");
                        string l_VString = camera.Description.ToUpper().Substring(l_VStringPtr, 8);
                        l_VStringPtr = l_VString.IndexOf(".");
                        if (l_VStringPtr > 0)
                        {
                            l_V1 = Convert.ToInt32(l_VString[..(l_VStringPtr - 1)]); // Extract the number
                            l_VString = l_VString[(l_VStringPtr + 1)..]; // Get the second version number part
                            l_VStringPtr = l_VString.IndexOf(".");
                            if (l_VStringPtr > 1)
                            {
                                l_V2 = Convert.ToInt32(l_VString[..(l_VStringPtr - 1)]); // Extract the number
                                l_VString = l_VString[(l_VStringPtr + 1)..]; // Get the third version number part
                                l_V3 = Convert.ToInt32(l_VString); // Extract the number
                                                                   // Turn the version parts into a whole number
                                l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3;
                                if (l_V1 < 5000008)
                                {
                                    LogIssue("Version Check", "*** This version of the camera simulator has known conformance issues, ***");
                                    LogIssue("Version Check", "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***");
                                    LogNewLine();
                                }
                            }
                        }
                    }
                    else
                    {
                        LogIssue("Version Check", "*** This version of the camera simulator has known conformance issues, ***");
                        LogIssue("Version Check", "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***");
                        LogNewLine();
                    }
                }
                catch (Exception ex)
                {
                    LogIssue("ConformanceCheck", $"Unexpected error in pre-run check: {ex.Message}");
                    LogDebug("ConformanceCheck", $"Exception detail: {ex}");
                }
            }

            // Run camera tests
            if (!cancellationToken.IsCancellationRequested)
            {
                if (settings.CameraFirstUseTests) // Only run these tests if configured to do so
                {
                    LogNewLine();

                    // Check LastExposureDuration
                    LogTestOnly("Last Tests"); try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get LastExposureDuration");
                        m_LastExposureDuration = camera.LastExposureDuration;
                        LogIssue("LastExposureDuration", "LastExposureDuration did not return an error when called before an exposure was made");
                    }
                    catch (Exception ex)
                    {
                        LogOK("LastExposureDuration", $"LastExposureDuration returned an error before an exposure was made: {ex.Message}");
                    }

                    // Check LastExposureStartTime
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get LastExposureStartTime");
                        m_LastExposureStartTime = camera.LastExposureStartTime;
                        LogIssue("LastExposureStartTime", "LastExposureStartTime did not return an error when called before an exposure was made");
                    }
                    catch (Exception ex)
                    {
                        LogOK("LastExposureStartTime", $"LastExposureStartTime returned an error before an exposure was made: {ex.Message}");
                    }

                }
                else
                {
                    LogInfo("LastExposureStartTime", "Conform is configured to omit these \"First use\" tests .");

                }
            }
        }

        public override void CheckProperties()
        {
            int l_BinX, l_BinY, l_MaxBinX, l_MaxBinY;


            // Basic read tests
            m_MaxBinX = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.MaxBinX, "MaxBinX", 1, MAX_BIN_X)); if (cancellationToken.IsCancellationRequested) return;
            m_MaxBinY = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.MaxBinY, "MaxBinY", 1, MAX_BIN_Y)); if (cancellationToken.IsCancellationRequested) return;

            if (!m_CanAsymmetricBin)
            {
                if (m_MaxBinX != m_MaxBinY)
                    LogIssue("CanAsymmetricBin", "CanAsymmetricBin is false but MaxBinX and MaxBinY are not equal!");
            }

            m_BinX = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.BinX, "BinX Read", 1, 1)); if (cancellationToken.IsCancellationRequested) return; // Must default to 1 on start-up
            m_BinY = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.BinY, "BinY Read", 1, 1)); if (cancellationToken.IsCancellationRequested) return; // Must default to 1 on start-up

            if (!m_CanAsymmetricBin)
            {
                if (m_BinX != m_BinY)
                    LogIssue("CanAsymmetricBin", "CanAsymmetricBin is false but BinX and BinY are not equal!");
            }

            // Test writing low and high Bin values outside maximum range
            try // Invalid low value
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                camera.BinX = 0;
                LogIssue("BinX Write", "Invalid value 0 written but no error returned");
            }
            catch (Exception ex)
            {
                LogOK("BinX Write", $"Received error on setting BinX to 0: {ex.Message}");
            }

            try // Invalid high value
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                camera.BinX = (short)(m_MaxBinX + 1);
                LogIssue("BinX Write", "Invalid value " + m_MaxBinX + 1 + " written but no error returned");
            }
            catch (Exception ex)
            {
                LogOK("BinX Write", $"Received error on setting BinX to {m_MaxBinX + 1}: {ex.Message}");
            }

            try // Invalid low value
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinY");
                camera.BinY = 0;
                LogIssue("BinY Write", "Invalid value 0 written but no error returned");
            }
            catch (Exception ex)
            {
                LogOK("BinY Write", $"Received error on setting BinY to 0: {ex.Message}");
            }

            try // Invalid high value
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinY");
                camera.BinY = (short)(m_MaxBinY + 1);
                LogIssue("BinY Write", "Invalid value " + m_MaxBinY + 1 + " written but no error returned");
            }
            catch (Exception ex)
            {
                LogOK("BinY Write", $"Received error on setting BinY to {m_MaxBinY + 1}: {ex.Message}");
            }

            // Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogNewLine(); // Insert a blank line if required

            if (settings.CameraMaxBinX > 0)
            {
                l_MaxBinX = settings.CameraMaxBinX;
                LogTestAndMessage("BinXY Write", string.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", l_MaxBinX, m_MaxBinX));
            }
            else
                l_MaxBinX = m_MaxBinX;

            if (settings.CameraMaxBinY > 0)
            {
                l_MaxBinY = settings.CameraMaxBinY;
                LogTestAndMessage("BinXY Write", string.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", l_MaxBinY, m_MaxBinY));
            }
            else
                l_MaxBinY = m_MaxBinY;

            if ((settings.CameraMaxBinX > m_MaxBinX) | (settings.CameraMaxBinY > m_MaxBinY))
                LogNewLine(); // Insert a blank line if required

            if (settings.CameraMaxBinX > m_MaxBinX)
                LogTestAndMessage("BinXY Write", string.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", l_MaxBinX, m_MaxBinX));

            if (settings.CameraMaxBinY > m_MaxBinY)
                LogTestAndMessage("BinXY Write", string.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", l_MaxBinY, m_MaxBinY));

            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogNewLine(); // Insert a blank line if required

            // Write BinX and BinY
            if (m_CanAsymmetricBin)
            {
                for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                {
                    for (l_BinY = 1; l_BinY <= l_MaxBinY; l_BinY++)
                    {
                        bool binXSetOk = false;
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set BinX");
                            camera.BinX = (short)l_BinX;
                            binXSetOk = true;
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex, "when setting BinX to " + l_BinX, "InvalidValue error received when setting BinX to " + l_BinX);
                        }

                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set BinY");
                            camera.BinY = (short)l_BinY;

                            if (binXSetOk)
                                LogOK("BinXY Write", "Successfully set asymmetric XY binning: " + l_BinX + " x " + l_BinY);
                            else
                                LogOK("BinXY Write", $"Successfully set Y binning to {l_BinY}");
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex, "when setting BinY to " + l_BinY, "InvalidValue error received when setting BinY to " + l_BinY);
                        }
                    }
                }
            }
            else // Can only bin symmetrically
                for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                {
                    bool binXSetOk = false;

                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set BinX");
                        camera.BinX = (short)l_BinX;
                        binXSetOk = true;
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex, "when setting BinX to " + l_BinX, "InvalidValueException received when setting BinX to " + l_BinX);
                    }

                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set BinY");
                        camera.BinY = (short)l_BinX;
                        if (binXSetOk)
                            LogOK("BinXY Write", "Successfully set symmetric XY binning: " + l_BinX + " x " + l_BinX);
                        else
                            LogOK("BinXY Write", $"Successfully set Y binning to {l_BinX}");
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex, "when setting BinY to " + l_BinX, "InvalidValueException received when setting BinY to " + l_BinX);
                    }
                }

            // Reset X and Y binning to 1x1 state
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                camera.BinX = 1;
            }
            catch (Exception)
            {
            }

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set BinY");
            try
            {
                camera.BinY = 1;
            }
            catch (Exception)
            {
            }

            m_CameraState = CameraPropertyTestCameraState(CamPropertyType.CameraState, "CameraState"); if (cancellationToken.IsCancellationRequested) return;
            cameraXSize = CameraPropertyTestInteger(CamPropertyType.CameraXSize, "CameraXSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;
            cameraYSize = CameraPropertyTestInteger(CamPropertyType.CameraYSize, "CameraYSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;

            CameraPropertyTestDouble(CamPropertyType.CCDTemperature, "CCDTemperature", ABSOLUTE_ZERO_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;
            m_CoolerOn = CameraPropertyTestBoolean(CamPropertyType.CoolerOn, "CoolerOn Read", false); if (cancellationToken.IsCancellationRequested) return;

            // Write CoolerOn
            bool l_OriginalCoolerState;
            string l_TargetCoolerState;
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");
                l_OriginalCoolerState = camera.CoolerOn;
                if (l_OriginalCoolerState)
                    l_TargetCoolerState = "off";
                else
                    l_TargetCoolerState = "on";
                try
                {
                    if (l_OriginalCoolerState)
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");
                        camera.CoolerOn = false;
                    }
                    else
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");
                        camera.CoolerOn = true;
                    }
                    LogOK("CoolerOn Write", "Successfully changed CoolerOn state");
                }
                catch (Exception ex)
                {
                    HandleException("CoolerOn Write", MemberType.Property, Required.Optional, ex, "turning Cooler " + l_TargetCoolerState);
                }

                // Restore Cooler state
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");
                    camera.CoolerOn = l_OriginalCoolerState;
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

            CameraPropertyTestDouble(CamPropertyType.CoolerPower, "CoolerPower", 0.0, 100.0, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.ElectronsPerADU, "ElectronsPerADU", 0.00001, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.FullWellCapacity, "FullWellCapacity", 0.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestBoolean(CamPropertyType.HasShutter, "HasShutter", false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.HeatSinkTemperature, "HeatSinkTemperature", ABSOLUTE_ZERO_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;

            m_ImageReady = CameraPropertyTestBoolean(CamPropertyType.ImageReady, "ImageReady", false); if (cancellationToken.IsCancellationRequested) return;
            if (m_ImageReady & settings.CameraFirstUseTests) // Issue this warning if configured to do so
                LogIssue("ImageReady", "Image is flagged as ready but no exposure has been started!");

            // Release memory currently consumed by images
            ReleaseMemory();

            // ImageArray 
            SetFullStatus("ImageArray", "Getting image data from device...", "");

            if (m_ImageReady) // ImageReady is true
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get ImageArray");
                    m_ImageArray = (Array)camera.ImageArray;
                    if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                    {
                        LogIssue("ImageArray", "No image has been taken but ImageArray has not returned an error");
                    }
                    else  // Omit first use tests
                    {
                        LogOK("ImageArray", "ImageArray read OK");
                    }
                }
                catch (Exception ex)
                {
                    LogDebug("ImageArray", $"Exception 1:\r\n{ex}");
                    if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                    {
                        LogOK("ImageArray", $"Received error when camera is idle: {ex.Message}");
                    }
                    else // Omit first use tests
                    {
                        LogIssue("ImageArray", $"Conform is configured to omit \"First use\" tests and ImageReady is true, but ImageArray returned an error: {ex.Message}");
                        LogDebug("ImageArray", $"Exception detail: {ex}");
                    }
                }
            }
            else // ImageReady is false
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get ImageArray");
                    m_ImageArray = (Array)camera.ImageArray;
                    LogIssue("ImageArray", "ImageReady is false and no image has been taken but ImageArray has not returned an error");
                }
                catch (Exception ex)
                {
                    LogOK("ImageArray", $"Received error when ImageReady is false: {ex.Message}");
                    LogDebug("ImageArray", $"Exception 2:\r\n{ex}");
                }
            }

            // Release memory currently consumed by images
            ReleaseMemory();

            // ImageArrayVariant
            // Release memory currently consumed by images
            ReleaseMemory();

            SetFullStatus("ImageArrayVariant", "Getting image data from device...", "");
            if (settings.CameraTestImageArrayVariant) // Test if configured to do so
            {
                if (m_ImageReady)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get ImageArrayVariant");
                        m_ImageArrayVariant = (Array)camera.ImageArrayVariant;

                        if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                        {
                            LogIssue("ImageArrayVariant", "No image has been taken but ImageArrayVariant has not returned an error");
                        }
                        else
                        {
                            LogOK("ImageArrayVariant", "ImageArray read OK");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                        {
                            LogOK("ImageArrayVariant", $"Received error before an image has been taken: {ex.Message}");
                        }
                        else // Omit first use tests
                        {
                            LogIssue("ImageArrayVariant", $"Conform is configured to omit \"First use\" tests and ImageReady is true, but ImageArrayVariant returned an error: {ex.Message}");
                        }
                        LogDebug("ImageArrayVariant", $"Exception detail:\r\n{ex}");
                    }
                }
                else
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get ImageArrayVariant");
                        m_ImageArrayVariant = (Array)camera.ImageArrayVariant;
                        LogIssue("ImageArrayVariant", "ImageReady is false and no image has been taken but ImageArray has not returned an error");
                    }
                    catch (Exception ex)
                    {
                        LogOK("ImageArrayVariant", $"Received error when ImageReady is false: {ex.Message}");
                    }
                }
            }
            else // Log an issue because the ImageArrayVariant test was omitted
            {
                LogIssue("ImageArrayVariant", "Test omitted due to Conform configuration.");
            }

            // Release memory currently consumed by images
            ReleaseMemory();

            ClearStatus();

            m_IsPulseGuiding = CameraPropertyTestBoolean(CamPropertyType.IsPulseGuiding, "IsPulseGuiding", false); if (cancellationToken.IsCancellationRequested) return;
            if (m_IsPulseGuiding)
                LogIssue("IsPulseGuiding", "Camera is showing pulse guiding underway although no PulseGuide command has been issued!");

            CameraPropertyTestInteger(CamPropertyType.MaxADU, "MaxADU", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;

            CameraPropertyTestInteger(CamPropertyType.NumX, "NumX Read", 1, cameraXSize); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.NumX, "NumX", System.Convert.ToInt32(cameraXSize / (double)2));

            CameraPropertyTestInteger(CamPropertyType.NumY, "NumY Read", 1, cameraYSize); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.NumY, "NumY", System.Convert.ToInt32(cameraYSize / (double)2));

            CameraPropertyTestDouble(CamPropertyType.PixelSizeX, "PixelSizeX", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.PixelSizeY, "PixelSizeY", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;

            m_SetCCDTemperature = CameraPropertyTestDouble(CamPropertyType.SetCCDTemperature, "SetCCDTemperature Read", ABSOLUTE_ZERO_TEMPERATURE, BOILING_POINT_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;

            if (m_CanSetCCDTemperature)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
                    camera.SetCCDTemperature = 0.0; // Try an acceptable value
                    LogOK("SetCCDTemperature Write", "Successfully wrote 0.0");

                    // Execution only gets here if the CCD temperature can be set successfully
                    bool exceptionGenerated;
                    double setPoint;

                    // Find low set-point at which an exception is generated, stop at CAMERA_SETPOINT_UNPHYSICAL_TEMPERATURE because this is unphysical
                    exceptionGenerated = false;
                    setPoint = CAMERA_LOW_SETPOINT_START_TEMPERATURE;

                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature multiple times...");

                    // Loop downward in CAMERA_SETPOINT_INCREMENT degree temperature steps to find the maximum temperature that can be set
                    do
                    {
                        try
                        {
                            // Calculate the new test set point ensuring that the value is no lower than absolute zero
                            setPoint -= CAMERA_SETPOINT_INCREMENT;

                            // Set the new temperature set point or a value just below absolute zero on the last cycle
                            if (setPoint >= ABSOLUTE_ZERO_TEMPERATURE) // Normal case so use the calculated value
                            {
                                camera.SetCCDTemperature = setPoint;
                            }
                            else // The new set point is below absolute zero so use the below absolute zero test value instead of the calculated set point value.
                            {
                                camera.SetCCDTemperature = BELOW_ABSOLUTE_ZERO_TEMPERATURE;
                            }
                        }
                        catch (Exception ex)
                        {
                            exceptionGenerated = true;
                            LogDebug("SetCCDTemperature Low", $"Exception: {ex.Message}");
                        }
                    }
                    while (!exceptionGenerated & (setPoint >= ABSOLUTE_ZERO_TEMPERATURE)); // Reached the camera's lower limit or exceeded absolute zero 

                    // Check whether we found a valid set point or whether the lower test limit caused the test to stop
                    if (!exceptionGenerated) // We can set temperature below absolute zero 
                    {
                        LogIssue("SetCCDTemperature Write", $"Set point can be set to {BELOW_ABSOLUTE_ZERO_TEMPERATURE} degrees, which is below absolute zero!");
                    }
                    else // We found the camera's lowest settable temperature
                    {
                        if (setPoint >= ABSOLUTE_ZERO_TEMPERATURE) // Normal case
                            LogInfo("SetCCDTemperature Write", $"Set point lower limit found in the range {setPoint + 5.0:+0.00;-0.00;+0.00} to {setPoint + 0.01:+0.00;-0.00} degrees");
                        else // The new set point is below absolute zero 
                            LogInfo("SetCCDTemperature Write", $"Set point lower limit found in the range {setPoint + 5.0:+0.00;-0.00;+0.00} to {BELOW_ABSOLUTE_ZERO_TEMPERATURE:+0.00;-0.00} degrees");
                    }

                    // Find high set point at which an exception is generated, stop at MAX_CAMERA_SETPOINT_TEMPERATURE as this is a suitably high value
                    exceptionGenerated = false;
                    setPoint = CAMERA_HIGH_SETPOINT_START_TEMPERATURE; // Start at 0.0C

                    // Loop upward in CAMERA_SETPOINT_INCREMENT degree temperature steps to find the maximum temperature that can be set
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature multiple times...");
                    do
                    {
                        try
                        {
                            setPoint += CAMERA_SETPOINT_INCREMENT;
                            camera.SetCCDTemperature = setPoint;
                        }
                        catch (Exception ex)
                        {
                            exceptionGenerated = true;
                            LogDebug("SetCCDTemperature High", $"Exception: {ex.Message}");
                        }
                    }
                    while (!exceptionGenerated & (setPoint < BOILING_POINT_TEMPERATURE));// Reached upper limit so exit loop

                    // Check whether we found a valid set point or whether the higher test limit caused the test to stop
                    if (!exceptionGenerated) // We hit the higher temperature limit
                    {
                        LogIssue("SetCCDTemperature Write", $"The set point can be set to {BOILING_POINT_TEMPERATURE} degrees Celsius!");
                    }
                    else // We found a limit below 100C.
                    {
                        LogInfo("SetCCDTemperature Write", $"Set point upper limit found in the range {setPoint - 5.0:+0.00;-0.00} to {setPoint - 0.01:+0.00;-0.00} degrees");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("SetCCDTemperature Write", MemberType.Property, Required.MustBeImplemented, ex, "Setting a legal value 0");
                }

                // Restore original value
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
                try
                {
                    camera.SetCCDTemperature = m_SetCCDTemperature;
                }
                catch
                {
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
                    camera.SetCCDTemperature = 0;
                    LogIssue("SetCCDTemperature Write", "CanSetCCDTemperature is false but No error returned on write");
                }
                catch (Exception ex)
                {
                    HandleException("SetCCDTemperature Write", MemberType.Property, Required.Optional, ex, "");
                }

            CameraPropertyTestInteger(CamPropertyType.StartX, "StartX Read", 0, cameraXSize - 1); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.StartX, "StartX", System.Convert.ToInt32(cameraXSize / (double)2));
            CameraPropertyTestInteger(CamPropertyType.StartY, "StartY Read", 0, cameraYSize - 1); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.StartY, "StartY", System.Convert.ToInt32(cameraYSize / (double)2));

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to get InterfaceVersion");
            if (camera.InterfaceVersion > 1)
            {
                // SensorType - Mandatory
                // This must be tested before BayerOffset because BayerOffset is mandatory for colour and optional for monochrome cameras
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get SensorType");
                    m_SensorType = (ASCOM.DeviceInterface.SensorType)camera.SensorType;
                    m_CanReadSensorType = true; // Set a flag to indicate that we have got a valid SensorType value
                                                // Successfully retrieved a value
                    LogOK("SensorType Read", m_SensorType.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("SensorType Read", MemberType.Property, Required.Mandatory, ex, "");
                }

                // BayerOffset Read
                if (m_CanReadSensorType)
                {
                    if (m_SensorType == ASCOM.DeviceInterface.SensorType.Monochrome)
                    {
                        // Monochrome so both BayerOffset properties should throw not implemented exceptions
                        CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read");
                        CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read");
                    }
                    else
                    {
                        CameraPropertyTestShort(CamPropertyType.BayerOffsetX, "BayerOffsetX Read", 0, 10000, true);
                        CameraPropertyTestShort(CamPropertyType.BayerOffsetY, "BayerOffsetY Read", 0, 10000, true);
                    }
                }
                else
                {
                    LogInfo("BayerOffset Read", "Unable to read SensorType value; assuming that the sensor is Monochrome");
                    // Monochrome so both BayerOffset properties should throw not implemented exceptions
                    CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetX, "BayerOffsetX Read");
                    CameraPropertyMustNotImplemented(CamPropertyType.BayerOffsetY, "BayerOffsetY Read");
                }

                // ExposureMin and ExpoosureMax Read
                m_ExposureMax = CameraPropertyTestDouble(CamPropertyType.ExposureMax, "ExposureMax Read", 0.0001, double.MaxValue, true);
                m_ExposureMin = CameraPropertyTestDouble(CamPropertyType.ExposureMin, "ExposureMin Read", 0.0, double.MaxValue, true);
                if (m_ExposureMin <= m_ExposureMax)
                    LogOK("ExposureMin", "ExposureMin is less than or equal to ExposureMax");
                else
                    LogIssue("ExposureMin", "ExposureMin is greater than ExposureMax");

                // ExposureResolution Read
                m_ExposureResolution = CameraPropertyTestDouble(CamPropertyType.ExposureResolution, "ExposureResolution Read", 0.0, double.MaxValue, true);
                if (m_ExposureResolution <= m_ExposureMax)
                    LogOK("ExposureResolution", "ExposureResolution is less than or equal to ExposureMax");
                else
                    LogIssue("ExposureResolution", "ExposureResolution is greater than ExposureMax");

                // FastReadout Read Optional
                if (m_CanFastReadout)
                    m_FastReadout = CameraPropertyTestBoolean(CamPropertyType.FastReadout, "FastReadout Read", true);
                else
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get FastReadout");
                        m_FastReadout = camera.FastReadout;
                        LogIssue("FastReadout Read", "CanFastReadout is False but a PropertyNotImplemented error was not returned.");
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
                            LogTestAndMessage("ConformanceCheck", "About to set FastReadout");
                        camera.FastReadout = !m_FastReadout;
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set FastReadout");
                        camera.FastReadout = m_FastReadout;
                        LogOK("FastReadout Write", "Able to change the FastReadout state OK");
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
                            LogTestAndMessage("ConformanceCheck", "About to set FastReadout");
                        camera.FastReadout = true;
                        LogIssue("FastReadout Write", "CanFastReadout is False but a PropertyNotImplemented error was not returned.");
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
                        LogTestAndMessage("ConformanceCheck", "About to get GainMin");
                    m_GainMin = camera.GainMin;
                    // Successfully retrieved a value
                    m_CanReadGainMin = true;
                    LogOK("GainMin Read", m_GainMin.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("GainMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // GainMax Read - Optional
                try
                {
                    m_CanReadGainMax = false;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get GainMax");
                    m_GainMax = camera.GainMax;
                    // Successfully retrieved a value
                    m_CanReadGainMax = true;
                    LogOK("GainMax Read", m_GainMax.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("GainMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // Gains Read - Optional
                try
                {
                    m_CanReadGains = false;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get Gains");
                    m_Gains = camera.Gains;
                    // Successfully retrieved a value
                    m_CanReadGains = true;
                    foreach (string Gain in m_Gains)
                        LogOK("Gains Read", Gain.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("Gains Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                if (m_CanReadGainMax & m_CanReadGainMin & m_CanReadGains)
                    LogIssue("Gains", "GainMin, GainMax and Gains are all readable. Only one of GainMin/Max as a pair or Gains should be used, the other should throw a PropertyNotImplemented error");
                else
                {
                }

                // Gain Read - Optional 
                try
                {
                    m_CanReadGain = false; // Set default value to indicate can't read gain
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get Gain");
                    m_Gain = camera.Gain;
                    m_CanReadGain = true; // Flag that we can read Gain OK
                    if (m_CanReadGains)
                        LogOK("Gain Read", m_Gain + " " + m_Gains[0].ToString());
                    else
                        LogOK("Gain Read", m_Gain.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("Gain Read", MemberType.Property, Required.Optional, ex, "");
                }

                // Now check that gain property groups are implemented to handle the three gain modes: NotImplemented, Gain Index (Gain + Gains) and Gain Value (Gain + GainMin + GainMax)
                if (!m_CanReadGain & !m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax)
                    LogOK("Gain Read", "All four gain properties return errors - the driver is in \"Gain Not Implemented\" mode.");
                else if (m_CanReadGain)
                {
                    // Test for Gain Index mode
                    if ((m_CanReadGain & m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax))
                    {
                        m_GainMode = GainOffsetMode.IndexMode;
                        LogOK("Gain Read", "Gain and Gains can be read while GainMin and GainMax returned errors - the driver is in \"Gain Index\" mode.");
                    }
                    else if ((m_CanReadGain & !m_CanReadGains & m_CanReadGainMin & m_CanReadGainMax))
                    {
                        m_GainMode = GainOffsetMode.ValueMode;
                        LogOK("Gain Read", "Gain, GainMin and GainMax can be read OK while Gains returns an error - the driver is in \"Gain Value\" mode.");
                    }
                    else
                    {
                        LogIssue("Gain Read", $"Unable to determine whether the driver is in \"Gain Not Implemented\", \"Gain Index\" or \"Gain Value\" mode. Please check the interface specification.");
                        LogInfo("Gain Read", $"Gain returned an error: {m_CanReadGain}, Gains returned an error: {m_CanReadGains}, GainMin returned an error: {m_CanReadGainMin}, GainMax returned an error: {m_CanReadGainMax}.");
                        LogInfo("Gain Read", $"\"Gain Not Implemented\" mode: Gain, Gains, GainMin and GainMax must all return errors.");
                        LogInfo("Gain Read", $"\"Gain Index\" mode: Gain and Gains must work while GainMin and GainMax must return errors.");
                        LogInfo("Gain Read", $"\"Gain Value\" mode: Gain, GainMin and GainMax must work while Gains must return an error.");
                    }
                }
                else
                {
                    LogIssue("Gain Read", $"Gain Read returned an error but at least one of Gains, GainMin Or GainMax did not return an error. If Gain returns an error, all the other gain properties should do likewise.");
                    LogInfo("Gain Read", $"Gains returned an error : {m_CanReadGains}, GainMin returned an error : {m_CanReadGainMin}, GainMax returned an error : {m_CanReadGainMax}.");
                }

                // Gain write - Optional when neither gain index nor gain value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither gain mode is supported
                if (!m_CanReadGain & !m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set Gain");
                        camera.Gain = 0;
                        LogIssue("Gain Write", "Writing to Gain did not return a PropertyNotImplemented error whereas this was the case for reading Gain.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("Gain Write", MemberType.Property, Required.MustNotBeImplemented, ex, "PropertyNotImplemented error is expected");
                    }
                }
                else
                {
                    switch (m_GainMode)
                    {
                        case GainOffsetMode.Unknown:
                            {
                                LogIssue("Gain Write", "Cannot test Gain Write because of issues with other gain properties - skipping test");
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
                                        LogTestAndMessage("ConformanceCheck", "About to set Gain");
                                    camera.Gain = m_GainMin;
                                    LogOK("Gain Write", $"Successfully set gain minimum value {m_GainMin}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing the maximum valid value
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Gain");
                                    camera.Gain = m_GainMax;
                                    LogOK("Gain Write", $"Successfully set gain maximum value {m_GainMax}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Gain Write", MemberType.Property, Required.MustNotBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Gain");
                                    camera.Gain = (short)(m_GainMin - 1);
                                    LogIssue("Gain Write", $"Successfully set an gain below the minimum value ({m_GainMin - 1}), this should have resulted in an InvalidValue error.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for gain {m_GainMin - 1}, which is lower than the minimum value.");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Gain");
                                    camera.Gain = (short)(m_GainMax + 1);
                                    LogIssue("Gain Write", $"Successfully set a gain above the maximum value({m_GainMax + 1}), this should have resulted in an InvalidValue error.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for gain {m_GainMax + 1} which is higher than the maximum value.");
                                }

                                break;
                            }

                        default:
                            {
                                LogIssue("Gain Write", $"UNEXPECTED VALUE FOR OFFSETMODE: {m_GainMode}");
                                break;
                            }
                    }
                }
                if (cancellationToken.IsCancellationRequested) return;

                // PercentCompleted Read - Optional - corrected to match the specification
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get PercentCompleted");
                    m_PercentCompleted = camera.PercentCompleted;
                    switch (m_PercentCompleted)
                    {
                        case object _ when m_PercentCompleted < 0 // Lower than minimum value
                       :
                            {
                                LogIssue("PercentCompleted Read", "Invalid value: " + m_PercentCompleted.ToString());
                                break;
                            }

                        case object _ when m_PercentCompleted > 100 // Higher than maximum value
                 :
                            {
                                LogIssue("PercentCompleted Read", "Invalid value: " + m_PercentCompleted.ToString());
                                break;
                            }

                        default:
                            {
                                LogOK("PercentCompleted Read", m_PercentCompleted.ToString());
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
                        LogTestAndMessage("ConformanceCheck", "About to get ReadoutModes");
                    m_ReadoutModes = camera.ReadoutModes;
                    // Successfully retrieved a value
                    m_CanReadReadoutModes = true;
                    foreach (string ReadoutMode in m_ReadoutModes)
                        LogOK("ReadoutModes Read", ReadoutMode.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("ReadoutModes Read", MemberType.Property, Required.Mandatory, ex, "");
                }
                if (cancellationToken.IsCancellationRequested) return;

                // ReadoutMode Read - Mandatory
                m_ReadoutMode = CameraPropertyTestShort(CamPropertyType.ReadoutMode, "ReadoutMode Read", 0, short.MaxValue, true);
                if (m_CanReadReadoutModes)
                {
                    try
                    {
                        if (m_ReadoutMode < m_ReadoutModes.Count)
                        {
                            LogOK("ReadoutMode Index", "ReadReadoutMode is within the bounds of the ReadoutModes ArrayList");
                            LogInfo("ReadoutMode Index", "Current value: " + m_ReadoutModes[m_ReadoutMode].ToString());
                        }
                        else
                            LogIssue("ReadoutMode Index", "ReadReadoutMode is outside the bounds of the ReadoutModes ArrayList");
                    }
                    catch (Exception ex)
                    {
                        LogIssue("ReadoutMode Index", $"Exception: {ex.Message}");
                        LogDebug("ReadoutMode Index", $"Exception detail: {ex}");
                    }
                }
                else
                    LogInfo("ReadoutMode Index", "Skipping ReadReadoutMode index test because ReadoutModes is unavailable");

                // SensorName
                CameraPropertyTestString(CamPropertyType.SensorName, "SensorName Read", 250, true);
            }

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to get InterfaceVersion");
            if (camera.InterfaceVersion > 2)
            {
                // OffsetMin Read - Optional
                try
                {
                    m_CanReadOffsetMin = false;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get OffsetMin");
                    m_OffsetMin = camera.OffsetMin;
                    // Successfully retrieved a value
                    m_CanReadOffsetMin = true;
                    LogOK("OffsetMin Read", m_OffsetMin.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("OffsetMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // OffsetMax Read - Optional
                try
                {
                    m_CanReadOffsetMax = false;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get OffsetMax");
                    m_OffsetMax = camera.OffsetMax;
                    // Successfully retrieved a value
                    m_CanReadOffsetMax = true;
                    LogOK("OffsetMax Read", m_OffsetMax.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("OffsetMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // Offsets Read - Optional
                try
                {
                    m_CanReadOffsets = false;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get Offsets");
                    m_Offsets = camera.Offsets;
                    // Successfully retrieved a value
                    m_CanReadOffsets = true;
                    foreach (string Offset in m_Offsets)
                        LogOK("Offsets Read", Offset.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("Offsets Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                if (m_CanReadOffsetMax & m_CanReadOffsetMin & m_CanReadOffsets)
                    LogIssue("Offsets", "OffsetMin, OffsetMax and Offsets are all readable. Only one of OffsetMin/Max as a pair or Offsets should be used, the other should throw a PropertyNotImplemented error");
                else
                {
                }
                if (cancellationToken.IsCancellationRequested) return;

                // Offset Read - Optional 
                try
                {
                    m_CanReadOffset = false; // Set default value to indicate can't read offset
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get Offset");
                    m_Offset = camera.Offset;
                    m_CanReadOffset = true; // Flag that we can read Offset OK
                    if (m_CanReadOffsets)
                        LogOK("Offset Read", m_Offset + " " + m_Offsets[0].ToString());
                    else
                        LogOK("Offset Read", m_Offset.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("Offset Read", MemberType.Property, Required.Optional, ex, "");
                }

                // Now check that offset property groups are implemented to handle the three offset modes: NotImplemented, Offset Index (Offset + Offsets) and Offset Value (Offset + OffsetMin + OffsetMax)
                if (!m_CanReadOffset & !m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax)
                    LogOK("Offset Read", "All four offset properties return errors - the driver is in \"Offset Not Implemented\" mode.");
                else if (m_CanReadOffset)
                {
                    // Test for Offset Index mode
                    if ((m_CanReadOffset & m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax))
                    {
                        m_OffsetMode = GainOffsetMode.IndexMode;
                        LogOK("Offset Read", "Offset and Offsets can be read while OffsetMin and OffsetMax return errors - the driver is in \"Offset Index\" mode.");
                    }
                    else if ((m_CanReadOffset & !m_CanReadOffsets & m_CanReadOffsetMin & m_CanReadOffsetMax))
                    {
                        m_OffsetMode = GainOffsetMode.ValueMode;
                        LogOK("Offset Read", "Offset, OffsetMin and OffsetMax can be read OK while Offsets returns an error - the driver is in \"Offset Value\" mode.");
                    }
                    else
                    {
                        m_OffsetMode = GainOffsetMode.Unknown;
                        LogIssue("Offset Read", $"Unable to determine whether the driver is in \"Offset Not Implemented\", \"Offset Index\" or \"Offset Value\" mode. Please check the interface specification.");
                        LogInfo("Offset Read", $"Offset returned an error: {m_CanReadOffset}, Offsets returned an error: {m_CanReadOffsets}, OffsetMin returned an error: {m_CanReadOffsetMin}, OffsetMax returned an error: {m_CanReadOffsetMax}.");
                        LogInfo("Offset Read", $"\"Offset Not Implemented\" mode: Offset, Offsets, OffsetMin and OffsetMax must all return errors.");
                        LogInfo("Offset Read", $"\"Offset Index\" mode: Offset and Offsets must work while OffsetMin and OffsetMax must return errors.");
                        LogInfo("Offset Read", $"\"Offset Value\" mode: Offset, OffsetMin and OffsetMax must work while Offsets must throw return an error.");
                    }
                }
                else
                {
                    LogIssue("Offset Read", $"Offset Read returned an error but at least one of Offsets, OffsetMin Or OffsetMax did not return an error. If Offset returns an error, all the other offset properties must do likewise.");
                    LogInfo("Offset Read", $"Offsets returned an error : {m_CanReadOffsets}, OffsetMin returned an error : {m_CanReadOffsetMin}, OffsetMax returned an error : {m_CanReadOffsetMax}.");
                }

                // Offset write - Optional when neither offset index nor offset value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither offset mode is supported
                if (!m_CanReadOffset & !m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set Offset");
                        camera.Offset = 0;
                        LogIssue("Offset Write", "Writing to Offset did not throw a PropertyNotImplemented error when reading Offset did.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("Offset Write", MemberType.Property, Required.MustNotBeImplemented, ex, "PropertyNotImplemented error is expected");
                    }
                }
                else
                    switch (m_OffsetMode)
                    {
                        case GainOffsetMode.Unknown:
                            {
                                LogIssue("Offset Write", "Cannot test Offset Write because of issues with other offset properties - skipping test");
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
                                        LogTestAndMessage("ConformanceCheck", "About to set Offset");
                                    camera.Offset = m_OffsetMin;
                                    LogOK("Offset Write", $"Successfully set offset minimum value {m_OffsetMin}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing the maximum valid value
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Offset");
                                    camera.Offset = m_OffsetMax;
                                    LogOK("Offset Write", $"Successfully set offset maximum value {m_OffsetMax}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Offset Write", MemberType.Property, Required.MustNotBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Offset");
                                    camera.Offset = m_OffsetMin - 1;
                                    LogIssue("Offset Write", $"Successfully set an offset below the minimum value ({m_OffsetMin - 1}), this should have resulted in an InvalidValue error.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for offset {m_OffsetMin - 1}, which is lower than the minimum value.");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Offset");
                                    camera.Offset = m_OffsetMax + 1;
                                    LogIssue("Offset Write", $"Successfully set an offset above the maximum value({m_OffsetMax + 1}), this should have resulted in an InvalidValueerror.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for offset {m_OffsetMax + 1} which is higher than the maximum value.");
                                }
                                break;
                            }

                        default:
                            {
                                LogIssue("Offset Write", $"UNEXPECTED VALUE FOR OFFSETMODE: {m_OffsetMode}");
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
                        LogTestAndMessage("ConformanceCheck", "About to set SubExposureDuration");
                    camera.SubExposureDuration = m_SubExposureDuration;
                    LogOK("SubExposureDuration write", $"Successfully wrote {m_SubExposureDuration}");
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
                                LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                            returnValue = camera.CameraState;
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                LogOK(p_Name, returnValue.ToString());
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
                                LogTestAndMessage("ConformanceCheck", "About to get BayerOffsetX");
                            returnValue = camera.BayerOffsetX;
                            break;
                        }

                    case CamPropertyType.BayerOffsetY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get BayerOffsetY");
                            returnValue = camera.BayerOffsetY;
                            break;
                        }

                    case CamPropertyType.PercentCompleted:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get PercentCompleted");
                            returnValue = camera.PercentCompleted;
                            break;
                        }

                    case CamPropertyType.ReadoutMode:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ReadoutMode");
                            returnValue = camera.ReadoutMode;
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogIssue(p_Name, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogIssue(p_Name, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK(p_Name, returnValue.ToString());
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
                                LogTestAndMessage("ConformanceCheck", "About to get BayerOffsetX");
                            TestShort = camera.BayerOffsetX;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogIssue(p_Name, "Sensor type is Monochrome so this property must throw a PropertyNotImplemented error; it must not return a value");
                            break;
                        }

                    case CamPropertyType.BayerOffsetY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get BayerOffsetY");
                            TestShort = camera.BayerOffsetY;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogIssue(p_Name, "Sensor type is Monochrome so this property must throw a PropertyNotImplemented error; it must not return a value");
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "returnValue: Unknown test type - " + p_Type.ToString());
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
                                LogTestAndMessage("ConformanceCheck", "About to get BinX");
                            returnValue = camera.BinX;
                            break;
                        }

                    case CamPropertyType.BinY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get BinY");
                            returnValue = camera.BinY;
                            break;
                        }

                    case CamPropertyType.CameraState:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                            returnValue = (int)camera.CameraState;
                            break;
                        }

                    case CamPropertyType.CameraXSize:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CameraXSize");
                            returnValue = camera.CameraXSize;
                            break;
                        }

                    case CamPropertyType.CameraYSize:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CameraYSize");
                            returnValue = camera.CameraYSize;
                            break;
                        }

                    case CamPropertyType.MaxADU:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get MaxADU");
                            returnValue = camera.MaxADU;
                            break;
                        }

                    case CamPropertyType.MaxBinX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get MaxBinX");
                            returnValue = camera.MaxBinX;
                            break;
                        }

                    case CamPropertyType.MaxBinY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get MaxBinY");
                            returnValue = camera.MaxBinY;
                            break;
                        }

                    case CamPropertyType.NumX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get NumX");
                            returnValue = camera.NumX;
                            break;
                        }

                    case CamPropertyType.NumY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get NumY");
                            returnValue = camera.NumY;
                            break;
                        }

                    case CamPropertyType.StartX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get StartX");
                            returnValue = camera.StartX;
                            break;
                        }

                    case CamPropertyType.StartY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get StartY");
                            returnValue = camera.StartY;
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogIssue(p_Name, "Invalid value: " + returnValue.ToString());
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
                                        LogInfo(p_Name, $"{returnValue}. This is higher than Conform's test criterion: {MAX_BIN_X}. Is this intended?");
                                        break;
                                    }

                                case CamPropertyType.MaxBinY // Informational message for MaxBinY
                         :
                                    {
                                        LogInfo(p_Name, $"{returnValue}. This is higher than Conform's test criterion: {MAX_BIN_Y}. Is this intended?");
                                        break;
                                    }

                                default:
                                    {
                                        LogIssue(p_Name, "Invalid value: " + returnValue.ToString());
                                        break;
                                    }
                            }

                            break;
                        }

                    default:
                        {
                            LogOK(p_Name, returnValue.ToString());
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
                                LogTestAndMessage("ConformanceCheck", "About to get CCDTemperature");
                            returnValue = camera.CCDTemperature;
                            break;
                        }

                    case CamPropertyType.CoolerPower:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CoolerPower");
                            returnValue = camera.CoolerPower;
                            break;
                        }

                    case CamPropertyType.ElectronsPerADU:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ElectronsPerADU");
                            returnValue = camera.ElectronsPerADU;
                            break;
                        }

                    case CamPropertyType.FullWellCapacity:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get FullWellCapacity");
                            returnValue = camera.FullWellCapacity;
                            break;
                        }

                    case CamPropertyType.HeatSinkTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get HeatSinkTemperature");
                            returnValue = camera.HeatSinkTemperature;
                            break;
                        }

                    case CamPropertyType.PixelSizeX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get PixelSizeX");
                            returnValue = camera.PixelSizeX;
                            break;
                        }

                    case CamPropertyType.PixelSizeY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get PixelSizeY");
                            returnValue = camera.PixelSizeY;
                            break;
                        }

                    case CamPropertyType.SetCCDTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get SetCCDTemperature");
                            returnValue = camera.SetCCDTemperature;
                            break;
                        }

                    case CamPropertyType.ExposureMax:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ExposureMax");
                            returnValue = camera.ExposureMax;
                            break;
                        }

                    case CamPropertyType.ExposureMin:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ExposureMin");
                            returnValue = camera.ExposureMin;
                            break;
                        }

                    case CamPropertyType.ExposureResolution:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ExposureResolution");
                            returnValue = camera.ExposureResolution;
                            break;
                        }

                    case CamPropertyType.SubExposureDuration:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get SubExposureDuration");
                            returnValue = camera.SubExposureDuration;
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case double _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogIssue(p_Name, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case double _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogIssue(p_Name, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK(p_Name, returnValue.ToString());
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
                                LogTestAndMessage("ConformanceCheck", "About to get CoolerOn");
                            returnValue = camera.CoolerOn;
                            break;
                        }

                    case CamPropertyType.HasShutter:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get HasShutter");
                            returnValue = camera.HasShutter;
                            break;
                        }

                    case CamPropertyType.ImageReady:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                            returnValue = camera.ImageReady;
                            break;
                        }

                    case CamPropertyType.IsPulseGuiding:
                        {
                            m_IsPulseGuidingSupported = false;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding");
                            returnValue = camera.IsPulseGuiding;
                            m_IsPulseGuidingSupported = true; // Command works properly and doesn't cause a not implemented exception
                            break;
                        }

                    case CamPropertyType.FastReadout:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get FastReadout");
                            returnValue = camera.FastReadout;
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                LogOK(p_Name, returnValue.ToString());
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
                                LogTestAndMessage("ConformanceCheck", "About to get Description");
                            returnValue = camera.Description;
                            break;
                        }

                    case CamPropertyType.SensorName:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get SensorName");
                            returnValue = camera.SensorName;
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "returnValue: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue == "":
                        {
                            LogOK(p_Name, "The driver returned an empty string");
                            break;
                        }

                    default:
                        {
                            if (returnValue.Length <= p_MaxLength)
                                LogOK(p_Name, returnValue);
                            else
                                LogIssue(p_Name, "String exceeds " + p_MaxLength + " characters maximum length - " + returnValue);
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
                                LogTestAndMessage("ConformanceCheck", "About to set NumX");
                            camera.NumX = p_TestOK;
                            break;
                        }

                    case CamPropertyType.NumY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set NumY");
                            camera.NumY = p_TestOK;
                            break;
                        }

                    case CamPropertyType.StartX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set StartX");
                            camera.StartX = p_TestOK;
                            break;
                        }

                    case CamPropertyType.StartY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set StartY");
                            camera.StartY = p_TestOK;
                            break;
                        }
                }
                LogOK(p_Property + " write", "Successfully wrote " + p_TestOK);
            }
            catch (Exception ex)
            {
                HandleException(p_Property + " write", MemberType.Property, Required.MustBeImplemented, ex, "Can't set legal value: " + p_TestOK.ToString());
            }
        }

        public override void CheckMethods()
        {
            int i, j, maxBinX, maxBinY;
            // AbortExposure - Mandatory
            SetTest("AbortExposure");
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                m_CameraState = camera.CameraState;

                // Test whether the camera is idle, which it should be in a well behaved device
                if (m_CameraState != CameraState.Idle)
                {
                    LogIssue("AbortExposure", $"The camera should be idle but is not: {m_CameraState}");
                }

                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to call AbortExposure");
                    camera.AbortExposure();
                    if (m_CanAbortExposure)
                        LogOK("AbortExposure", "No error returned when camera is already idle");
                    else
                        LogIssue("AbortExposure", "CanAbortExposure is false but no error is returned when AbortExposure is called");
                }
                catch (Exception ex)
                {
                    if (m_CanAbortExposure)
                    {
                        LogIssue("AbortExposure", $"Received error when camera is idle: {ex.Message}");
                        LogDebug("AbortExposure", $"Exception detail: {ex}");
                    }
                    else
                        LogOK("AbortExposure", "CanAbortExposure is false and an error was returned");
                }
            }
            catch (Exception ex)
            {
                LogIssue("AbortExposure", $"Received error when reading camera state: {ex.Message}");
                LogDebug("AbortExposure", $"Exception detail: {ex}");
            }

            // PulseGuide
            SetTest("PulseGuide");
            if (m_CanPulseGuide) // Can pulse guide
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
                catch (Exception ex)
                {
                    LogIssue("PulseGuide", $"CanPulseGuide is true but PulseGuide returned an error when called: {ex.Message}");
                    LogDebug("PulseGuide", $"Exception detail: {ex}");
                }
            }
            else // Can not pulse guide so should return an error
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to call PulseGuide - North");
                    camera.PulseGuide(GuideDirection.North, 0);
                    LogIssue("PulseGuide", "CanPulseGuide is false but no error was returned when calling the method");
                }
                catch (Exception ex)
                {
                    if (IsMethodNotImplementedException(ex))
                    {
                        LogOK("PulseGuide", "CanPulseGuide is false and PulseGuide is not implemented in this driver");
                    }
                    else
                    {
                        LogOK("PulseGuide", $"CanPulseGuide is false and an error was returned when calling the method: {ex.Message}");
                    }
                }

            // StopExposure
            SetTest("StopExposure");
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                m_CameraState = camera.CameraState;

                // Test whether the camera is idle, which it should be in a well behaved device
                if (m_CameraState != CameraState.Idle)
                {
                    LogIssue("AbortExposure", $"The camera should be idle but is not: {m_CameraState}");
                }

                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to call StopExposure");
                    SetAction("Calling StopExposure()");
                    camera.StopExposure();
                    if (m_CanStopExposure)
                        LogOK("StopExposure", "No error returned when camera is already idle");
                    else
                        LogIssue("StopExposure", "CanStopExposure is false but no error is returned when StopExposure is called");
                }
                catch (Exception ex)
                {
                    if (m_CanStopExposure)
                        LogIssue("StopExposure", $"Received error when the camera is idle: {ex.Message}");
                    else
                    {
                        LogOK("StopExposure", $"CanStopExposure is false and an error was returned: {ex.Message}");
                    }

                    LogDebug("StopExposure", $"Exception detail: {ex}");
                }

            }
            catch (Exception ex)
            {
                LogIssue("StopExposure", $"Received error when reading camera state: {ex.Message}");
                LogDebug("StopExposure", $"Exception detail: {ex}");
            }

            // StartExposure
            SetTest("StartExposure");

            // Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogNewLine(); // Insert a blank line if required
            if (settings.CameraMaxBinX > 0)
            {
                maxBinX = settings.CameraMaxBinX;
                LogTestAndMessage("StartExposure", string.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", maxBinX, m_MaxBinX));
            }
            else
                maxBinX = m_MaxBinX;
            if (settings.CameraMaxBinY > 0)
            {
                maxBinY = settings.CameraMaxBinY;
                LogTestAndMessage("StartExposure", string.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", maxBinY, m_MaxBinY));
            }
            else
                maxBinY = m_MaxBinY;

            if ((settings.CameraMaxBinX > m_MaxBinX) | (settings.CameraMaxBinY > m_MaxBinY))
                LogNewLine(); // Insert a blank line if required
            if (settings.CameraMaxBinX > m_MaxBinX)
                LogTestAndMessage("StartExposure", string.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", maxBinX, m_MaxBinX));
            if (settings.CameraMaxBinY > m_MaxBinY)
                LogTestAndMessage("StartExposure", string.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", maxBinY, m_MaxBinY));

            // StartExposure - Confirm that correct operation occurs
            int binX, binY;
            if (m_CanAsymmetricBin)
            {
                for (binY = 1; binY <= maxBinY; binY++)
                {
                    for (binX = 1; binX <= maxBinX; binX++)
                    {
                        // Calculate required image size
                        int xSize = cameraXSize / binX;
                        int ySize = cameraYSize / binY;
                        bool constrainedXSize = false;
                        bool constrainedYSize = false;

                        // Test whether the image size exceeds to configured maximum image dimensions and adjust down if needed
                        if ((xSize > settings.CameraXMax) & (settings.CameraXMax > 0))
                        {
                            constrainedXSize = true;
                            LogTestAndMessage("StartExposure", $"Camera X dimension reduced from {xSize} to {settings.CameraXMax} by Conform configuration.");
                            xSize = settings.CameraXMax;
                        }

                        if ((ySize > settings.CameraYMax) & (settings.CameraYMax > 0))
                        {
                            constrainedYSize = true;
                            LogTestAndMessage("StartExposure", $"Camera Y dimension reduced from {ySize} to {settings.CameraYMax} by Conform configuration.");
                            ySize = settings.CameraYMax;
                        }

                        CameraExposure("StartExposure", $"Taking {(constrainedXSize | constrainedYSize ? "configured size" : "full frame")} image {binX} x {binY} bin " +
                            $"({xSize}{(constrainedXSize ? "*" : "")} x {ySize}{(constrainedYSize ? "*" : "")}) - {Convert.ToDouble(xSize * ySize) / 1000000.0:0.0} MPix",
                            binX, binY, 0, 0, xSize, ySize, settings.CameraExposureDuration, "");

                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
            }
            else
                for (binX = 1; binX <= maxBinX; binX++)
                {
                    // Calculate required image size
                    int xSize = cameraXSize / binX;
                    int ySize = cameraYSize / binX;
                    bool constrainedXSize = false;
                    bool constrainedYSize = false;

                    // Test whether the image size exceeds to configured maximum image dimensions and adjust down if needed
                    if ((xSize > settings.CameraXMax) & (settings.CameraXMax > 0))
                    {
                        constrainedXSize = true;
                        LogTestAndMessage("StartExposure", $"Camera X dimension reduced from {xSize} to {settings.CameraXMax} by Conform configuration.");
                        xSize = settings.CameraXMax;
                    }

                    if ((ySize > settings.CameraYMax) & (settings.CameraYMax > 0))
                    {
                        constrainedYSize = true;
                        LogTestAndMessage("StartExposure", $"Camera Y dimension reduced from {ySize} to {settings.CameraYMax} by Conform configuration.");
                        ySize = settings.CameraYMax;
                    }

                    CameraExposure("StartExposure", $"Taking {(constrainedXSize | constrainedYSize ? "configured size" : "full frame")} image {binX} x {binX} bin " +
                        $"({xSize}{(constrainedXSize ? "*" : "")} x {ySize}{(constrainedYSize ? "*" : "")}) - {Convert.ToDouble(xSize * ySize) / 1000000.0:0.0} MPix",
                        binX, binX, 0, 0, xSize, ySize, settings.CameraExposureDuration, "");

                    if (cancellationToken.IsCancellationRequested)
                        return;
                }

            // StartExposure - Confirm error cases
            LogNewLine();
            LogTestOnly("StartExposure error cases");
            // StartExposure - Negative time
            CameraExposure("Reject Negative Duration", "Testing rejection of negative exposure duration (-1.0)", 1, 1, 0, 0, cameraXSize, cameraYSize, -1.0, "negative duration"); if (cancellationToken.IsCancellationRequested)
                return; // Test that negative duration generates an error

            // StartExposure - Invalid Bin values
            for (i = 1; i <= maxBinX; i++)
            {
                for (j = 1; j <= maxBinY; j++)
                {
                    if (m_CanAsymmetricBin)
                    {
                        // X size too large for binned size
                        CameraExposure($"Reject Bad XSize (bin {i} x {j})", $"Testing rejection of bad X size value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, 0, 0, Convert.ToInt32((cameraXSize / (double)i) + 1), Convert.ToInt32(cameraYSize / (double)j), 0.1, "X size larger than binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y size too large for binned size
                        CameraExposure($"Reject Bad YSize (bin {i} x {j})", $"Testing rejection of bad Y size value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, 0, Convert.ToInt32(cameraXSize / (double)i), Convert.ToInt32((cameraYSize / (double)j) + 1), 0.1, "Y size larger than binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // X start outside binned chip dimensions
                        CameraExposure($"Reject Bad XStart (bin {i} x {j})", $"Testing rejection of bad X start value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, Convert.ToInt32((cameraXSize / (double)i) + 1), 0, Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1, "X start outside binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y start outside binned chip dimensions
                        CameraExposure($"Reject Bad YStart (bin {i} x {j})", $"Testing rejection of bad Y start value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, Convert.ToInt32((cameraYSize / (double)j) + 1), Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1, "Y start outside binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                    else if (i == j)
                    {
                        // X size too large for binned size
                        CameraExposure($"Reject Bad XSize (bin {i} x {j})", $"Testing rejection of bad X size value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, 0, 0, Convert.ToInt32((cameraXSize / (double)i) + 1), Convert.ToInt32(cameraYSize / (double)j), 0.1, "X size larger than binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y size too large for binned size
                        CameraExposure($"Reject Bad YSize (bin {i} x {j})", $"Testing rejection of bad Y size value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, 0, Convert.ToInt32(cameraXSize / (double)i), Convert.ToInt32((cameraYSize / (double)j) + 1), 0.1, "Y size larger than binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // X start outside binned chip dimensions
                        CameraExposure($"Reject Bad XStart (bin {i} x {j})", $"Testing rejection of bad X start value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, Convert.ToInt32((cameraXSize / (double)i) + 1), 0, Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1, "X start outside binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y start outside binned chip dimensions
                        CameraExposure($"Reject Bad YStart (bin {i} x {j})", $"Testing rejection of bad Y start value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, Convert.ToInt32((cameraYSize / (double)j) + 1), Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1, "Y start outside binned chip size, Bin " + i + "x" + j);
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
            }
        }

        private void CameraExposure(string testName, string testDescription, int requiredBinX, int requiredBinY, int requiredStartX, int requiredStartY, int requiredNumX, int requiredNumY, double requiredDuration, string expectedErrorMessage)
        {
            string numPlanes, variantType;
            bool exposedOK, imageReadyTooEarly = false; // Flag to determine whether we were successful or something failed
            DateTime startTime, startTimeUTC, endTime;
            short percentCompleted;

            // Start off by assuming the worst case, this will be set true if the exposure completes OK
            exposedOK = false;

            LogDebug(testName,$"Entered CameraExposure");

            // Log test name for happy path tests
            if (testName.ToUpperInvariant() == "STARTEXPOSURE")
            {
                LogNewLine(); // Blank Line
                LogTestOnly(testDescription);
            }

            #region Validate camera state and input parameters

            // Set camera exposure parameters
            bool parametersSetOk = true; // Flag indicating whether the BinX, BinY, StartX, StartY, NumX and NumY parameters were set correctly

            // Make sure the camera is in idle state before proceeding further
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                CameraState cameraState = camera.CameraState;

                if (cameraState != CameraState.Idle) // Camera is not in idle state ready for an exposure
                {
                    LogIssue(testName, $"Test abandoned because the camera is in state: {cameraState} rather than the expected: {CameraState.Idle}.");
                    parametersSetOk = false;
                }
            }
            catch (Exception ex)
            {
                parametersSetOk = false;
                LogIssue(testName, $"Unable to read camera state: {ex.Message}");
                LogDebug(testName, $"CameraState exception detail:\r\n {ex}");
            }

            // Set BinX value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                camera.BinX = (short)requiredBinX;
            }
            catch (Exception ex)
            {
                parametersSetOk = false;
                LogIssue(testName, $"Unable to set BinX to {requiredBinX}: {ex.Message}");
                LogDebug(testName, $"BinX exception detail:\r\n {ex}");
            }

            // Set BinY value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinY");
                camera.BinY = (short)requiredBinY;
            }
            catch (Exception ex)
            {
                parametersSetOk = false;
                LogIssue(testName, $"Unable to set BinY to {requiredBinY}: {ex.Message}");
                LogDebug(testName, $"BinY exception detail:\r\n {ex}");
            }

            // Set StartX value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set StartX");
                camera.StartX = requiredStartX;
            }
            catch (Exception ex)
            {
                parametersSetOk = false;
                LogIssue(testName, $"Unable to set StartX to {requiredStartX}: {ex.Message}");
                LogDebug(testName, $"StartX exception detail:\r\n {ex}");
            }

            // Set StartY value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set StartY");
                camera.StartY = requiredStartY;
            }
            catch (Exception ex)
            {
                parametersSetOk = false;
                LogIssue(testName, $"Unable to set StartY to {requiredStartY}: {ex.Message}");
                LogDebug(testName, $"StartY exception detail:\r\n {ex}");
            }

            // Set NumX value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set NumX");
                camera.NumX = requiredNumX;
            }
            catch (Exception ex)
            {
                parametersSetOk = false;
                LogIssue(testName, $"Unable to set NumX to {requiredNumX}: {ex.Message}");
                LogDebug(testName, $"NumX exception detail:\r\n {ex}");
            }

            // Set NumY value
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set NumY");
                camera.NumY = requiredNumY;
            }
            catch (Exception ex)
            {
                parametersSetOk = false;
                LogIssue(testName, $"Unable to set NumY to {requiredNumY}: {ex.Message}");
                LogDebug(testName, $"NumY exception detail:\r\n {ex}");
            }

            // Exit if parameters were not set properly
            if (!parametersSetOk)
            {
                // At least one of BinX, BinY, StartX, StartY, NumX or NumY could not be set
                LogInfo(testName, "Exposure test abandoned because the camera is not idle or at least one of BinX, BinY, NumX, NumY, StartX or StartY could not be set.");
                ClearStatus();
            }

            #endregion

            #region Initiate exposure

            // Start exposure because all parameters were set OK

            startTime = DateTime.Now;
            startTimeUTC = DateTime.UtcNow;
            bool initiatedOk = false; // Flag indicating whether the StarteExposure task completed OK
            try
            {
                bool ranToCompletion;

                SetAction($"Exposing for {requiredDuration} seconds");


                // Start  a UI update task to display timing information
                CancellationTokenSource exposeUiTaskTokenSource = new();
                CancellationToken exposeUiTaskCancellationToken = exposeUiTaskTokenSource.Token;

                // Start a task to update the UI while initiating the exposure
                Task.Run(() =>
                {
                    UpdateUI(exposeUiTaskCancellationToken);
                }, exposeUiTaskCancellationToken);

                // Initiate the exposure
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to call StartExposure");

                // Create a cancellation token that we can set if the task times out
                CancellationTokenSource cancellationTokenSource = new();
                CancellationToken cancellationToken = cancellationTokenSource.Token;

                // Create and start the StartExposure task
                Task startExposureTask = Task.Run(() =>
                {
                    try
                    {
                        camera.StartExposure(requiredDuration, true);

                        // Do not respond if the task has been cancelled
                        if (!cancellationToken.IsCancellationRequested) // Task is live and has not been cancelled
                        {
                            LogDebug(testName, "Exposure initiated OK");

                            if (string.IsNullOrEmpty(expectedErrorMessage))
                            {
                                // Flag that the exposure was successfully initiated
                                initiatedOk = true;
                            }
                            else
                            {
                                LogTestAndMessage(testName, $"No error was returned when {char.ToLowerInvariant(expectedErrorMessage[0])}{expectedErrorMessage[1..]}");
                                LogIssue(testName, "Expected an error and didn't get one - BinX:" + requiredBinX + " BinY:" + requiredBinY + " StartX:" + requiredStartX + " StartY:" + requiredStartY + " NumX:" + requiredNumX + " NumY:" + requiredNumY);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Ignore exceptions and do not respond if the task has been cancelled
                        if (!cancellationToken.IsCancellationRequested) // Task is live and has not been cancelled
                        {
                            // Record the error
                            if (expectedErrorMessage != "")
                                LogOK(testName, $"Received error: {ex.Message}");
                            else
                            {
                                LogIssue(testName, $"Error initiating exposure: {ex.Message}");
                                LogDebug(testName, $"Exception detail:\r\n {ex}");
                            }
                        }
                    }
                }, cancellationToken);

                // Wait for the start exposure task to complete or be cancelled
                try
                {
                    ranToCompletion = startExposureTask.Wait(TimeSpan.FromSeconds(settings.CameraWaitTimeout), applicationCancellationToken);
                }
                catch (OperationCanceledException)
                {
                    // The user cancelled the operation
                    LogNewLine();
                    LogError("CONFORMU", "START EXPOSURE WAS INTERRUPTTED LEAVING CONFORMU IN A POTENTIALLY CORRUPTED STATE. RESTART CONFORMU TO ENSURE RELEIABLE OPERATION");
                    LogNewLine();
                    ClearStatus();
                    return;
                }
                finally
                {
                    // Stop the UI update task
                    exposeUiTaskTokenSource.Cancel();
                    ClearStatus();
                }

                // Test the outcome from the task
                if (ranToCompletion) // Completed within the task timeout
                {
                    // Handle the three possible task end states
                    switch (startExposureTask.Status)
                    {
                        case TaskStatus.Canceled: // The application cancel button was pushed
                            initiatedOk = false;
                            LogIssue(testName, "StartExposure was cancelled.");
                            break;

                        case TaskStatus.Faulted: // The StartExposure call failed
                            initiatedOk = false;
                            LogIssue(testName, $"Error initiating exposure: {startExposureTask.Exception?.InnerException.Message}");
                            LogDebug(testName, $"Exception detail:\r\n {startExposureTask.Exception?.InnerException}");
                            break;

                        default: // Some other unexpected state
                            initiatedOk = false;
                            LogError(testName, $"Unexpected task end state from camera.StartExposure: {startExposureTask.Status}");

                            break;

                        case TaskStatus.RanToCompletion: // The StartExposure method completed OK within the specified timeout period.
                                                         // No action here because the initiatedOk flag is set by the code within the task
                            break;
                    }
                }
                else // The task timed out
                {
                    // Cancel the task so that it does not try to write to the log
                    cancellationTokenSource.Cancel();

                    // Log an issue because of the timeout
                    LogIssue(testName, $"StartExposure {(expectedErrorMessage == "" ? "" : $"{expectedErrorMessage} ")}did not return within the timeout period: {requiredDuration + WAITWHILE_EXTRA_WAIT_TIME} seconds.");

                    // Provide a warning about possible application corruption
                    LogNewLine();
                    LogError("CONFORMU", "START EXPOSURE TIMED OUT LEAVING CONFORMU IN A POTENTIALLY CORRUPTED STATE. RESTART CONFORMU TO ENSURE RELEIABLE OPERATION");
                    LogNewLine();

                    // Cancel the task
                    ConformanceTestManager.ConformCancellationTokenSource.Cancel();

                    ClearStatus();
                    return;
                }
            }
            catch (Exception ex) // Something else went wrong...
            {
                initiatedOk = false;
                LogIssue(testName, $"Received error when exposing: {ex.Message}");
                LogDebug(testName, $"Exception: {ex}");
            }

            // Exit if the task did not initiate OK or failed to generate an error when expected to do so
            if (!initiatedOk)
            {
                ClearStatus();
                return;
            }

            #endregion

            #region Wait for the synchronous or asynchronous exposure to complete

            try
            {
                endTime = DateTime.Now;

                // Test whether we have a synchronous or asynchronous camera
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get ImageReady and CameraState");
                if (camera.ImageReady & (camera.CameraState == CameraState.Idle)) // Synchronous exposure
                {
                    #region Check synchronous exposure behaviour

                    try
                    {
                        if (endTime.Subtract(startTime).TotalSeconds >= requiredDuration)
                        {
                            LogOK(testName, "Synchronous exposure found OK: " + requiredDuration + " seconds");
                            CameraTestLast(requiredDuration, startTimeUTC);
                        }
                        else
                            LogIssue(testName, "Synchronous exposure found but image was returned before exposure time was complete");
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while checking initiation: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ClearStatus();
                        return;
                    }

                    #endregion
                }
                else // Asynchronous exposure
                {
                    #region Wait for exposing state to start

                    SetStatus("Waiting for exposure to start");

                    // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                    imageReadyTooEarly = camera.ImageReady;

                    // Wait for exposing state
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get CameraState multiple times");
                    Stopwatch sw = Stopwatch.StartNew();
                    WaitWhile(GetAction(), () =>
                    {
                        // Wait while the camera is exposing and is not in an error state
                        return (camera.CameraState != CameraState.Exposing) & (camera.CameraState != CameraState.Error);
                    }, 500, settings.CameraWaitTimeout, () =>
                    {
                        return $"{sw.Elapsed.TotalSeconds:0.0} / {requiredDuration:0.0}";
                    });

                    if (applicationCancellationToken.IsCancellationRequested) // Exit if required
                    {
                        ClearStatus();
                        return;
                    }

                    // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                    imageReadyTooEarly = camera.ImageReady;

                    #endregion

                    #region Wait for exposing state to end

                    try
                    {
                        // Wait for the exposing state to finish
                        startTime = DateTime.Now;
                        startTimeUTC = DateTime.UtcNow;
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get CameraState, InterfaceVersion and PercentCompleted multiple times...");

                        // Start the loop timing stopwatch
                        sw.Restart();
                        Stopwatch swOverall = Stopwatch.StartNew();

                        WaitWhile($"Waiting for exposure to complete", () =>
                        {
                            // Wait for the camera state to be something other than CamerraStates.Exposing
                            return camera.CameraState == CameraState.Exposing;
                        }, 500, settings.CameraWaitTimeout, () =>
                        {
                            // Create a progress status message
                            bool reportedError = false;
                            string percentCompletedMessage = "Not present in a V1 driver"; // Initialise PercentCompleted message

                            if (camera.InterfaceVersion > 1)
                            {
                                try
                                {
                                    percentCompleted = camera.PercentCompleted;
                                    percentCompletedMessage = $"{percentCompleted}%"; // Operation completed OK
                                }
                                catch (Exception ex)
                                {
                                    if (IsPropertyNotImplementedException(ex))
                                    {
                                        percentCompletedMessage = "Not implemented"; // Not implemented
                                    }
                                    else if (IsInvalidOperationException(testName, ex))
                                    {
                                        percentCompletedMessage = $"Invalid operation: {ex.Message}"; // Not valid at this time
                                    }
                                    else
                                    {
                                        percentCompletedMessage = $"Error: {ex.Message}"; // Something bad happened!

                                        // Only log the error once to avoid unnecessarily filling the log
                                        if (!reportedError)
                                        {
                                            LogDebug("PercentCompleted", $"Exception detail: {ex}");
                                            reportedError = true;
                                        }
                                    }
                                }
                            }

                            return $"Percent complete: {percentCompletedMessage}";
                        });

                        endTime = DateTime.Now;

                        if (applicationCancellationToken.IsCancellationRequested) // Exit if required
                        {
                            ClearStatus();
                            return;
                        }
                    }
                    catch (TimeoutException)
                    {
                        LogIssue(testName, $"Test abandoned, timed out waiting for exposure to complete.");
                        ClearStatus();
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while waiting for exposure to complete: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ClearStatus();
                        return;
                    }

                    #endregion

                    #region Wait for camera to become idle

                    try
                    {
                        // Wait for camera to become idle
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get CameraState multiple times");

                        WaitWhile("Waiting for camera idle state, reading/downloading image", () =>
                        {
                            // Wait until the camera state is idle or error
                            return (camera.CameraState != CameraState.Idle) & (camera.CameraState != CameraState.Error);
                        }, 500, settings.CameraWaitTimeout);

                        if (applicationCancellationToken.IsCancellationRequested) // Exit if required
                        {
                            ClearStatus();
                            return;
                        }
                    }
                    catch (TimeoutException)
                    {
                        LogIssue(testName, $"Test abandoned, timed out waiting for camera to become idle.");
                        ClearStatus();
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while waiting for camera to become idle: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ClearStatus();
                        return;
                    }

                    #endregion

                    #region Wait for image to be ready for retrieval

                    try
                    {
                        // Wait for image to become ready
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get CameraState multiple times");

                        // Wait until ImageReady is true or the camera is in the error state
                        WaitWhile("Waiting for image ready", () =>
                        {
                            return !camera.ImageReady & (camera.CameraState != CameraState.Error);
                        }, 500, settings.CameraWaitTimeout);
                    }
                    catch (TimeoutException)
                    {
                        LogIssue(testName, $"Test abandoned, timed out waiting for ImageReady to become true.");
                        ClearStatus();
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while waiting for ImageReady to become true: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ClearStatus();
                        return;
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {
                LogIssue(testName, $"Unexpected error while waiting for exposure to complete: {ex.Message}");
                LogDebug(testName, $"Exception detail:\r\n{ex}");
                ClearStatus();
                return;
            }

            #endregion

            #region Check exposure outcome

            try
            {
                if (applicationCancellationToken.IsCancellationRequested) // Exit if required
                {
                    ClearStatus();
                    return;
                }

                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                if (camera.ImageReady)
                {
                    LogOK(testName, "Asynchronous exposure found OK: " + requiredDuration + " seconds");
                    CameraTestLast(requiredDuration, startTimeUTC);
                }
                else
                {
                    // Give up because the exposure was not successful
                    LogIssue(testName, "Test abandoned because the camera state is CameraError.");
                    ClearStatus();
                    return;
                }

                // Display a warning if ImageReady was set too early
                if (imageReadyTooEarly)
                {
                    LogIssue(testName, "Test abandoned because ImageReady was set True before the camera completed its exposure.");
                    ClearStatus();
                    return;
                }

                // Camera exposed OK and didn't generate an exception
                else
                {
                    exposedOK = true;
                }
                LogDebug(testName, $"Camera exposed image OK: {exposedOK}");
            }
            catch (Exception ex)
            {
                LogIssue(testName, $"Unexpected error while checking returned image: {ex.Message}");
                LogDebug(testName, $"Exception detail:\r\n{ex}");
                ClearStatus();
                return;
            }

            #endregion

            #region Check ImageArray

            // Check image array dimensions
            try
            {
                // Release memory currently consumed by images
                ReleaseMemory();

                // Retrieve the image array
                SetAction("Retrieving ImageArray");

                // Create a cancellation token that we can set if the task times out
                CancellationTokenSource iaTokenSource = new();
                CancellationToken iaToken = iaTokenSource.Token;

                Task iaTask = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get ImageArray");
                        sw.Restart();
                        m_ImageArray = (Array)camera.ImageArray;
                        sw.Stop();
                        // Do not respond if the task has been cancelled
                        if (!cancellationToken.IsCancellationRequested) // Task is live and has not been cancelled
                        {
                            LogDebug("ImageArray", "Got image array OK");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Ignore exceptions and do not respond if the task has been cancelled
                        if (!cancellationToken.IsCancellationRequested) // Task is live and has not been cancelled
                        {
                            // Record the error
                            LogIssue("ImageArray", $"Error getting ImageArray: {ex.Message}");
                            LogDebug("ImageArray", $"Exception detail:\r\n {ex}");
                        }
                    }
                }, iaToken, TaskCreationOptions.LongRunning, TaskScheduler.Default);

                // Start  a UI update task
                CancellationTokenSource iaUiTaskTokenSource = new();
                CancellationToken iaUiTaskCancellationToken = iaUiTaskTokenSource.Token;

                Task.Run(() =>
                {
                    UpdateUI(iaUiTaskCancellationToken);
                }, iaUiTaskCancellationToken);
                bool ranToCompletion=false;
                try
                {
                    // Wait for the ImageArray task to complete or be cancelled
                    ranToCompletion = iaTask.Wait(TimeSpan.FromSeconds(settings.CameraWaitTimeout), applicationCancellationToken);
                }
                catch (OperationCanceledException)
                {
                    LogNewLine();
                    LogError("CONFORMU", "GET IMAGEARRAY WAS INTERRUPTTED LEAVING CONFORMU IN A POTENTIALLY CORRUPTED STATE. RESTART CONFORMU TO ENSURE RELEIABLE OPERATION");
                    LogNewLine();
                    ClearStatus();
                    return;
                }
                finally
                {
                    // Stop the UI update task
                    iaUiTaskTokenSource.Cancel();
                    ClearStatus();
                }

                // Test the outcome from the task
                if (ranToCompletion) // Completed within the cancellation timeout
                {
                    // Handle the three possible task end states
                    switch (iaTask.Status)
                    {
                        case TaskStatus.Canceled: // The application cancel button was pushed
                            LogIssue("ImageArray", "Retrieve image array was cancelled.");
                            break;

                        case TaskStatus.Faulted: // The StartExposure call failed
                            LogIssue("ImageArray", $"Error getting image array: {iaTask.Exception?.InnerException.Message}");
                            LogDebug("ImageArray", $"Exception detail:\r\n {iaTask.Exception?.InnerException}");
                            break;

                        default: // Some other unexpected state
                            LogError("ImageArray", $"Unexpected task end state from camera.ImageArray: {iaTask.Status}");

                            break;

                        case TaskStatus.RanToCompletion: // The ImageArray method completed OK within the specified timeout period.
                                                         // No action here because the gotImageArrayOk flag is set by the code within the task
                            break;
                    }
                }
                else // Timed out
                {
                    // Cancel the task so that it does not try to write to the log
                    iaTokenSource.Cancel();

                    // Log an issue because of the timeout
                    LogIssue("ImageArray", $"ImageArray did not return within the timeout period: {settings.CameraWaitTimeout} seconds.");

                    // Provide a warning about possible application corruption
                    LogNewLine();
                    LogError("CONFORMU", "IMAGEARRAY TIMED OUT LEAVING CONFORMU IN A POTENTIALLY CORRUPTED STATE. RESTART CONFORMU TO ENSURE RELEIABLE OPERATION");
                    LogNewLine();

                    // Cancel the task
                    ConformanceTestManager.ConformCancellationTokenSource.Cancel();

                    ClearStatus();

                    return;
                }

                // Exit if cancelled
                if (applicationCancellationToken.IsCancellationRequested) // Exit if required
                {
                    ClearStatus();
                    return;
                }

                // Test image array variant array for conformity
                if ((m_ImageArray.GetLength(0) == requiredNumX) & (m_ImageArray.GetLength(1) == requiredNumY))
                {
                    if (m_ImageArray.GetType().ToString() == "System.Int32[,]" | m_ImageArray.GetType().ToString() == "System.Int32[,,]")
                    {
                        if (m_ImageArray.Rank == 2)
                            numPlanes = "1 plane";
                        else
                        {
                            numPlanes = "1 plane";
                            if (m_ImageArray.GetUpperBound(2) > 0)
                                numPlanes = System.Convert.ToString(m_ImageArray.GetUpperBound(2) + 1) + " planes";
                        }
                        LogOK("ImageArray", $"Successfully read 32 bit integer array ({numPlanes}) {m_ImageArray.GetLength(0)} x {m_ImageArray.GetLength(1)} pixels in {sw.ElapsedMilliseconds}ms.");
                    }
                    else
                        LogIssue("ImageArray", "Expected 32 bit integer array, actually got: " + m_ImageArray.GetType().ToString());
                }
                else if ((m_ImageArray.GetLength(0) == requiredNumY) & (m_ImageArray.GetLength(1) == requiredNumX))
                    LogIssue("ImageArray", "Camera image dimensions swapped, expected values: " + requiredNumX + " x " + requiredNumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
                else
                    LogIssue("ImageArray", "Camera image does not have the expected dimensions of: " + requiredNumX + " x " + requiredNumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
            }
            catch (OutOfMemoryException ex)
            {
                // Log an error
                LogError("ImageArray", $"OutOfMemoryException - Conform Universal or the device ran out of memory: {ex.Message}");
                LogDebug("ImageArray", $"Exception detail: {ex}");
                ClearStatus();
                return;
            }
            catch (Exception ex)
            {
                LogIssue("ImageArray", $"Error when reading ImageArray: {ex.Message}");
                LogDebug("ImageArray", $"Exception detail: {ex}");
                ClearStatus();
                return;
            }

            // Release memory currently consumed by images
            ReleaseMemory();

            #endregion

            #region Check ImageArrayVariant

            try
            {
                // Check image array variant dimensions
                if (settings.CameraTestImageArrayVariant) // Test if configured to do so. No need to report an issue because it's already been reported when the ImageArrayVariant property was tested
                {
                    // Release memory currently consumed by images
                    ReleaseMemory();

                    SetAction("Retrieving ImageArrayVariant");

                    // Start  a UI update task
                    CancellationTokenSource iavUiTaskTokenSource = new();
                    CancellationToken iavUiTaskCancellationToken = iavUiTaskTokenSource.Token;

                    Task.Run(() =>
                    {
                        UpdateUI(iavUiTaskCancellationToken);
                    }, iavUiTaskCancellationToken);

                    // Create a cancellation token that we can set if the task times out
                    CancellationTokenSource iavTokenSource = new();
                    CancellationToken iavToken = iavTokenSource.Token;

                    Task iavTask = Task.Factory.StartNew<bool>(() =>
                    {
                        bool gotImageOk = false;
                        try
                        {
                            // Get the variant array image 
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ImageArrayVariant");
                            sw.Restart();
                            object imageObject = camera.ImageArrayVariant;
                            sw.Stop();

                            LogInfo("ImageArrayVariant", $"Received image OK");

                            // Assign the array to the application variable if the task is not cancelled
                            if (!cancellationToken.IsCancellationRequested) // Completed successfully
                            {
                                m_ImageArrayVariant = (Array)imageObject;
                                gotImageOk = true;
                            }
                            else // Operation cancelled
                            {
                                // Release the image memory
                                imageObject = null;

                                // Clean up released memory
                                GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
                                GC.Collect(2, GCCollectionMode.Forced, true, true);
                            }
                        }
                        catch (OutOfMemoryException ex)
                        {
                            if (!cancellationToken.IsCancellationRequested) // Only log the error if the task has not been cancelled
                            {
                                // Log an error
                                LogError("ImageArrayVariant", $"OutOfMemoryException - Conform Universal or the device ran out of memory: {ex.Message}");
                                LogDebug("ImageArrayVariant", $"Exception detail: {ex}");
                            }
                        }
                        catch (Exception ex)
                        {
                            if (!cancellationToken.IsCancellationRequested) // Only log the error if the task has not been cancelled
                            {
                                LogIssue("ImageArrayVariant", $"Error when reading ImageArrayVariant: {ex.Message}");
                                LogDebug("ImageArrayVariant", $"Exception detail: {ex}");
                            }
                        }

                        // Return a result indicating whether we got the image OK
                        return gotImageOk;
                    }, iavToken, TaskCreationOptions.LongRunning, TaskScheduler.Current);

                    bool ranToCompletion=false;
                    // Wait for the ImageArrayVariant task to complete or be cancelled
                    try
                    {
                        ranToCompletion = iavTask.Wait(TimeSpan.FromSeconds(settings.CameraWaitTimeout), applicationCancellationToken);
                    }
                    catch (OperationCanceledException)
                    {
                        LogNewLine();
                        LogError("CONFORMU", "GET IMAGEARRAYVARIANT WAS INTERRUPTTED LEAVING CONFORMU IN A POTENTIALLY CORRUPTED STATE. RESTART CONFORMU TO ENSURE RELEIABLE OPERATION");
                        LogNewLine();
                        ClearStatus();
                        return;
                    }
                    finally
                    {
                        // Stop the UI update task
                        iavUiTaskTokenSource.Cancel();
                        ClearStatus();
                    }

                    // Test the outcome from the task
                    if (ranToCompletion) // Completed within the task timeout
                    {
                        // Handle the three possible task end states
                        switch (iavTask.Status)
                        {
                            case TaskStatus.Canceled: // The application cancel button was pushed
                                LogIssue("ImageArrayVariant", "Retrieve image array was cancelled.");
                                break;

                            case TaskStatus.Faulted: // The StartExposure call failed
                                LogIssue("ImageArrayVariant", $"Error getting image array: {iavTask.Exception?.InnerException.Message}");
                                LogDebug("ImageArrayVariant", $"Exception detail:\r\n {iavTask.Exception?.InnerException}");
                                break;

                            default: // Some other unexpected state
                                LogError("ImageArrayVariant", $"Unexpected task end state from camera.ImageArray: {iavTask.Status}");

                                break;

                            case TaskStatus.RanToCompletion: // The ImageArray method completed OK within the specified timeout period.
                                                             // No action here because the gotImageArrayOk flag is set by the code within the task
                                break;
                        }
                    }
                    else // Timed out
                    {
                        // Cancel the task so that it does not try to write to the log
                        iavTokenSource.Cancel();
                        // Log an issue because of the timeout
                        LogIssue("ImageArrayVariant", $"ImageArrayVariant did not return within the timeout period: {settings.CameraWaitTimeout} seconds.");

                        // Provide a warning about possible application corruption
                        LogNewLine();
                        LogError("CONFORMU", "IMAGEARRAYVARIANT TIMED OUT LEAVING CONFORMU IN A POTENTIALLY CORRUPTED STATE. RESTART CONFORMU TO ENSURE RELEIABLE OPERATION");
                        LogNewLine();

                        // Cancel the task
                        ConformanceTestManager.ConformCancellationTokenSource.Cancel();

                        ClearStatus();
                        return;
                    }

                    // Exit if cancelled
                    if (applicationCancellationToken.IsCancellationRequested)
                    {
                        ClearStatus();
                        return;
                    }

                    // Test image array variant array for conformity
                    if ((m_ImageArrayVariant.GetLength(0) == requiredNumX) & (m_ImageArrayVariant.GetLength(1) == requiredNumY))
                    {
                        if (m_ImageArrayVariant.GetType().ToString() == "System.Object[,]" | m_ImageArrayVariant.GetType().ToString() == "System.Object[,,]")
                        {
                            if (m_ImageArrayVariant.Rank == 2)
                            {
                                numPlanes = "1 plane";
                                variantType = ((object[,])m_ImageArrayVariant)[0, 0].GetType().ToString();
                            }
                            else
                            {
                                numPlanes = "1 plane";
                                if (m_ImageArrayVariant.GetUpperBound(2) > 0)
                                {
                                    numPlanes = System.Convert.ToString(m_ImageArrayVariant.GetUpperBound(2) + 1) + " planes";
                                    variantType = ((object[,,])m_ImageArrayVariant)[0, 0, 0].GetType().ToString();
                                }
                                else
                                    variantType = ((object[,])m_ImageArrayVariant)[0, 0].GetType().ToString();
                            }
                            LogOK("ImageArrayVariant", $"Successfully read variant array ({numPlanes}) with {variantType} elements {m_ImageArrayVariant.GetLength(0)} x {m_ImageArrayVariant.GetLength(1)} pixels in {sw.ElapsedMilliseconds}ms.");
                        }
                        else
                            LogIssue("ImageArrayVariant", "Expected variant array, actually got: " + m_ImageArrayVariant.GetType().ToString());
                    }
                    else if ((m_ImageArrayVariant.GetLength(0) == requiredNumY) & (m_ImageArrayVariant.GetLength(1) == requiredNumX))
                        LogIssue("ImageArrayVariant", "Camera image dimensions swapped, expected values: " + requiredNumX + " x " + requiredNumY + " - actual values: " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1));
                    else
                        LogIssue("ImageArrayVariant", "Camera image does not have the expected dimensions of: " + requiredNumX + " x " + requiredNumY + " - actual values: " + m_ImageArrayVariant.GetLength(0) + " x " + m_ImageArrayVariant.GetLength(1));
                    
                    // Release memory currently consumed by images
                    ReleaseMemory();
                }
            }
            catch (OutOfMemoryException ex)
            {
                // Log an error
                LogError("ImageArrayVariant", $"OutOfMemoryException - Conform Universal or the device ran out of memory: {ex.Message}");
                LogDebug("ImageArrayVariant", $"Exception detail: {ex}");
                ClearStatus();
            }
            catch (Exception ex)
            {
                LogIssue("ImageArrayVariant", $"Unexpected error while processing ImageArrayVariant: {ex.Message}");
                LogDebug("ImageArrayVariant", $"Exception detail:\r\n{ex}");
            }

            // Release memory currently consumed by images
            ReleaseMemory();

            #endregion

            #region Clean up

            // Try and do some clean up
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to call StopExposure");
                camera.StopExposure();
            }
            catch (Exception) { }

            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to call AbortExposure");
                camera.AbortExposure();
            }
            catch (Exception) { }

            ClearStatus();

            #endregion  
        }

        /// <summary>
        /// Updates the UI status field with a count from 0.0 to the operation timeout.
        /// </summary>
        /// <param name="updateUiTaskCancellationToken"></param>
        private void UpdateUI(CancellationToken updateUiTaskCancellationToken)
        {
            // Update the status every 500ms
            const int POLL_INTERVAL = 500;

            // Continually update the status field until the task is cancelled
            do
            {
                // Calculate the current loop number (starts at 0 given that the timer's elapsed time will be zero or very low on the first loop)
                int currentLoopNumber = ((int)(sw.ElapsedMilliseconds) + 50) / POLL_INTERVAL; // Add a small positive offset (50) because integer division always rounds down

                // Calculate the sleep time required to start the next loop at a multiple of the poll interval
                int sleeptime = POLL_INTERVAL * (currentLoopNumber + 1) - (int)sw.ElapsedMilliseconds;

                // Sleep until it is time for the next completion function poll
                Thread.Sleep(sleeptime);

                // Update the status message
                SetStatus($"{sw.Elapsed.TotalSeconds:0.0} / {settings.CameraWaitTimeout:0.0}");

                // Wait for the start of the next loop
                Thread.Sleep(sleeptime);
            } while (!updateUiTaskCancellationToken.IsCancellationRequested & (sw.Elapsed.TotalSeconds < settings.CameraWaitTimeout));
        }

        private void CameraTestLast(double p_Duration, DateTime p_Start)
        {
            DateTime l_StartTime;

            // LastExposureDuration
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get LastExposureDuration");
                m_LastExposureDuration = camera.LastExposureDuration;
                if ((Math.Abs(m_LastExposureDuration - p_Duration) / p_Duration) < 0.02)
                    LogOK("LastExposureDuration", $"Last exposure duration is: {m_LastExposureDuration:0.000} seconds");
                else
                    LogIssue("LastExposureDuration", "LastExposureDuration is not the same as image duration: " + m_LastExposureDuration + " " + p_Duration);
            }
            catch (Exception ex)
            {
                HandleException("LastExposureDuration", MemberType.Property, Required.Optional, ex, "");
            }

            // LastExposurestartTime
            try // Confirm that it can be read
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get LastExposureStartTime");
                m_LastExposureStartTime = camera.LastExposureStartTime;
                int l_i;
                // Confirm that the format is as expected
                bool l_FormatOK;
                l_FormatOK = true;
                if (m_LastExposureStartTime.Length >= 19)
                {
                    for (l_i = 1; l_i <= 19; l_i++)
                    {
                        switch (l_i)
                        {
                            case 5:
                            case 8 // "-"
                           :
                                {
                                    if (m_LastExposureStartTime.Substring(l_i - 1, 1) != "-")
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{m_LastExposureStartTime.Substring(l_i - 1, 1)}' at position {l_i}is not '-'");
                                        l_FormatOK = false;
                                    }
                                    else
                                    {

                                    }
                                    break;
                                }

                            case 11 // "T"
                     :
                                {
                                    if (m_LastExposureStartTime.Substring(l_i - 1, 1) != "T")
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{m_LastExposureStartTime.Substring(l_i - 1, 1)}' at position {l_i}is not 'T'");
                                        l_FormatOK = false;
                                    }
                                    else
                                    {

                                    }
                                    break;
                                }

                            case 14:
                            case 17 // ":"
                     :
                                {
                                    if (m_LastExposureStartTime.Substring(l_i - 1, 1) != ":")
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{m_LastExposureStartTime.Substring(l_i - 1, 1)}' at position {l_i}is not ':'");
                                        l_FormatOK = false;
                                    }
                                    else
                                    {

                                    }
                                    break;
                                }

                            default:
                                {
                                    if (!m_LastExposureStartTime.Substring(l_i - 1, 1).IsNumeric())
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{m_LastExposureStartTime.Substring(l_i - 1, 1)}' at position {l_i}is not 'numeric'");
                                        l_FormatOK = false;
                                    }
                                    else
                                    {

                                    }
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
                                LogOK("LastExposureStartTime", "LastExposureStartTime is correct to within 2 seconds: " + m_LastExposureStartTime + " UTC");
                            else
                                LogIssue("LastExposureStartTime", "LastExposureStartTime is more than 2 seconds inaccurate : " + m_LastExposureStartTime + ", expected: " + p_Start.ToString("yyyy-MM-ddTHH:mm:ss") + " UTC");
                        }
                        catch (Exception ex)
                        {
                            LogIssue("LastExposureStartTime", "Parsing LastExposureStartTime - " + ex.Message + " " + m_LastExposureStartTime);
                            LogDebug("LastExposureStartTime", $"Exception detail:\r\n {ex}");
                        }
                    }
                    else
                        LogIssue("LastExposureStartTime", "LastExposureStartTime not in the expected format yyyy-mm-ddThh:mm:ss - " + m_LastExposureStartTime);
                }
                else if (m_LastExposureStartTime == "")
                    LogIssue("LastExposureStartTime", "LastExposureStartTime has returned an empty string - expected yyyy-mm-ddThh:mm:ss");
                else
                    LogIssue("LastExposureStartTime", "LastExposureStartTime is less than 19 characters - expected yyyy-mm-ddThh:mm:ss - " + m_LastExposureStartTime);
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
            //SetAction("Start " + CAMERA_PULSE_DURATION / (double)1000 + " second pulse guide " + p_Direction.ToString());
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", $"About to call PulseGuide - {p_Direction}");
            camera.PulseGuide(p_Direction, CAMERA_PULSE_DURATION); // Start a 2 second pulse
            l_EndTime = DateTime.Now;
            try
            {
                if (m_IsPulseGuidingSupported)
                {
                    if (l_EndTime.Subtract(l_StartTime).TotalMilliseconds < (CAMERA_PULSE_DURATION - 500))
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding");
                        if (camera.IsPulseGuiding)
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding multiple times");
                            Stopwatch sw = Stopwatch.StartNew();
                            WaitWhile($"Guiding {p_Direction}", () => { return camera.IsPulseGuiding; }, 500, 3, () => { return $"{sw.Elapsed.TotalSeconds:0.0} / {CAMERA_PULSE_DURATION / 1000:0.0} seconds"; });

                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding");
                            if (!camera.IsPulseGuiding)
                                LogOK("PulseGuide " + p_Direction.ToString(), "Asynchronous pulse guide found OK");
                            else
                                LogIssue("PulseGuide " + p_Direction.ToString(), "Asynchronous pulse guide expected but IsPulseGuiding is TRUE beyond expected time of 2 seconds");
                        }
                        else
                            LogIssue("PulseGuide " + p_Direction.ToString(), "Asynchronous pulse guide expected but IsPulseGuiding has returned FALSE");
                    }
                    else
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding");
                        if (!camera.IsPulseGuiding)
                            LogOK("PulseGuide " + p_Direction.ToString(), "Synchronous pulse guide found OK");
                        else
                            LogIssue("PulseGuide " + p_Direction.ToString(), "Synchronous pulse guide expected but IsPulseGuiding has returned TRUE");
                    }
                }
                else
                    switch (l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION)
                    {
                        case object _ when l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION > CAMERA_PULSE_TOLERANCE // Duration was more than 0.5 seconds longer than expected
                       :
                            {
                                LogIssue("PulseGuide " + p_Direction.ToString(), "Synchronous pulse guide longer than expected " + (CAMERA_PULSE_DURATION) / (double)1000 + " seconds: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
                                break;
                            }

                        case object _ when l_EndTime.Subtract(l_StartTime).TotalMilliseconds - CAMERA_PULSE_DURATION < 20 // Duration was more than 20ms shorter than expected
                 :
                            {
                                LogIssue("PulseGuide " + p_Direction.ToString(), "Synchronous pulse guide shorter than expected " + (CAMERA_PULSE_DURATION) / (double)1000 + " seconds: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
                                break;
                            }

                        default:
                            {
                                LogOK("PulseGuide " + p_Direction.ToString(), "Synchronous pulse guide found OK: " + l_EndTime.Subtract(l_StartTime).TotalSeconds + " seconds");
                                break;
                            }
                    }

            }
            catch (TimeoutException ex)
            {
                LogIssue($"PulseGuide {p_Direction}", $"Timed out waiting for IsPulseGuiding to go false. It should have done this in {Convert.ToDouble(CAMERA_PULSE_DURATION) / 1000.0:0.0} seconds");
                LogDebug("PulseGuide", $"Exception detail:\r\n {ex}");
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
            SetAction("Exposure for ImageArray Test");
            SetStatus("Start");
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set BinX");
            camera.BinX = 1;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set BinY");
            camera.BinY = 1;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set StartX");
            camera.StartX = 0;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set StartY");
            camera.StartY = 0;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set MaxBinX");
            camera.NumX = camera.MaxBinX;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set MaxBinY");
            camera.NumY = camera.MaxBinY;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to call StartExposure");
            camera.StartExposure(1, true); // 1 second exposure

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to call ImageReady multiple times");
            do
                SetStatus("Waiting for ImageReady");
            while (!camera.ImageReady);
            SetStatus("Finished");
            CameraPerformanceTest(CameraPerformance.ImageArray, "ImageArray");
            CameraPerformanceTest(CameraPerformance.ImageArrayVariant, "ImageArrayVariant");
        }
        private void CameraPerformanceTest(CameraPerformance p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate;
            SetAction(p_Name);
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
                                m_CameraState = camera.CameraState;
                                break;
                            }

                        case CameraPerformance.CCDTemperature:
                            {
                                _ = camera.CCDTemperature;
                                break;
                            }

                        case CameraPerformance.CoolerPower:
                            {
                                _ = camera.CoolerPower;
                                break;
                            }

                        case CameraPerformance.HeatSinkTemperature:
                            {
                                _ = camera.HeatSinkTemperature;
                                break;
                            }

                        case CameraPerformance.ImageReady:
                            {
                                m_ImageReady = camera.ImageReady;
                                break;
                            }

                        case CameraPerformance.IsPulseGuiding:
                            {
                                m_IsPulseGuiding = camera.IsPulseGuiding;
                                break;
                            }

                        case CameraPerformance.ImageArray:
                            {
                                // Release memory currently consumed by images
                                ReleaseMemory();

                                m_ImageArray = (Array)camera.ImageArray;
                                break;
                            }

                        case CameraPerformance.ImageArrayVariant:
                            {
                                // Release memory currently consumed by images
                                ReleaseMemory();

                                m_ImageArrayVariant = (Array)camera.ImageArrayVariant;
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, "Conform:PerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
                    {
                        SetStatus(l_Count + " transactions in " + l_ElapsedTime.ToString("0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (l_ElapsedTime <= PERF_LOOP_TIME);

                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case object _ when l_Rate > 10.0:
                        {
                            LogInfo(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= l_Rate && l_Rate <= 10.0:
                        {
                            LogOK(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= l_Rate && l_Rate <= 2.0:
                        {
                            LogInfo(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(p_Name, $"Unable to complete test: {ex.Message}");
                LogDebug(p_Name, $"Exception detail:\r\n {ex}");
            }
        }

        public override void PostRunCheck()
        {
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to call AbortExposure");
            if (m_CanAbortExposure)
            {
                try
                {
                    camera.AbortExposure();
                }
                catch
                {
                }
            }
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to call StopExposure");
            if (m_CanStopExposure)
            {
                try
                {
                    camera.StopExposure();
                }
                catch
                {
                }
            }
            if (m_CanSetCCDTemperature)
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
                try
                {
                    camera.SetCCDTemperature = m_SetCCDTemperature;
                }
                catch
                {
                }
            }
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");

            // Reset the camera image parameters to legal values
            try { camera.StartX = 0; } catch { }
            try { camera.StartY = 0; } catch { }
            try { camera.BinX = 1; } catch { }
            try { camera.BinY = 1; } catch { }
            try { camera.NumX = 1; } catch { }
            try { camera.NumY = 1; } catch { }

            try
            {
                camera.CoolerOn = m_CoolerOn;
            }
            catch
            {
            }
            LogOK("PostRunCheck", "Camera returned to initial cooler temperature");
        }

        /// <summary>
        /// Release memory allocated to the large arrays on the large object heap.
        /// </summary>
        private void ReleaseMemory()
        {
            SetAction("Releasing memory");

            // Clear out any previous memory allocations
            m_ImageArray = null;
            m_ImageArrayVariant = null;
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
            GC.Collect(2, GCCollectionMode.Forced, true, true);
        }
    }

}

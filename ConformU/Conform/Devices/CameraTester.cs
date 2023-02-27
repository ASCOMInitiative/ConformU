using ASCOM;
using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Threading;

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
        private bool m_CoolerOn, m_ImageReady;
        private int m_CameraXSize, m_CameraYSize;
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
                    m_Camera?.Dispose();

                    m_Camera = null;
                    m_ImageArray = null;
                    m_ImageArrayVariant = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_Camera, DeviceTypes.Camera);
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


                        m_Camera = new AlpacaCamera(
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
                                m_Camera = new CameraFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                m_Camera = new Camera(settings.ComDevice.ProgId);
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

                baseClassDevice = m_Camera; // Assign the driver to the base class

                LogInfo("CreateDevice", "Successfully created driver");

            }
            catch (Exception ex)
            {
                LogDebug("CreateDevice", "Exception thrown: " + ex.Message);
                throw; // Re throw exception 
            }

        }

        public override bool Connected
        {
            get
            {
                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get Connected");
                return m_Camera.Connected;
            }
            set
            {
                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to set Connected");
                m_Camera.Connected = value;

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
                            m_CanAbortExposure = m_Camera.CanAbortExposure;
                            LogOK(p_Name, m_CanAbortExposure.ToString());
                            break;
                        }

                    case CanType.tstCanAsymmetricBin:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanAsymmetricBin");
                            m_CanAsymmetricBin = m_Camera.CanAsymmetricBin;
                            LogOK(p_Name, m_CanAsymmetricBin.ToString());
                            break;
                        }

                    case CanType.tstCanGetCoolerPower:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanGetCoolerPower");
                            m_CanGetCoolerPower = m_Camera.CanGetCoolerPower;
                            LogOK(p_Name, m_CanGetCoolerPower.ToString());
                            break;
                        }

                    case CanType.tstCanPulseGuide:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanPulseGuide");
                            m_CanPulseGuide = m_Camera.CanPulseGuide;
                            LogOK(p_Name, m_CanPulseGuide.ToString());
                            break;
                        }

                    case CanType.tstCanSetCCDTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanSetCCDTemperature");
                            m_CanSetCCDTemperature = m_Camera.CanSetCCDTemperature;
                            LogOK(p_Name, m_CanSetCCDTemperature.ToString());
                            break;
                        }

                    case CanType.tstCanStopExposure:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanStopExposure");
                            m_CanStopExposure = m_Camera.CanStopExposure;
                            LogOK(p_Name, m_CanStopExposure.ToString());
                            break;
                        }

                    case CanType.tstCanFastReadout:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CanFastReadout");
                            m_CanFastReadout = m_Camera.CanFastReadout;
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
                    l_VStringPtr = m_Camera.Description.ToUpper().IndexOf("VERSION "); // Point at the start of the version string
                    if (l_VStringPtr >= 0)
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get Description");
                        string l_VString = m_Camera.Description.ToUpper().Substring(l_VStringPtr, 8);
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
                    LogIssue("ConformanceCheck", ex.ToString());
                }
            }

            // Run camera tests
            if (!cancellationToken.IsCancellationRequested)
            {
                if (settings.CameraFirstUseTests) // Only run these tests if configured to do so
                {
                    LogNewLine();
                    // Check LastError throws an exception
                    LogTestOnly("Last Tests"); try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get LastExposureDuration");
                        m_LastExposureDuration = m_Camera.LastExposureDuration;
                        LogIssue("LastExposureDuration", "LastExposureDuration did not generate an exception when called before an exposure was made");
                    }
                    catch (Exception)
                    {
                        LogOK("LastExposureDuration", "LastExposureDuration correctly generated an exception before an exposure was made");
                    }

                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get LastExposureStartTime");
                        m_LastExposureStartTime = m_Camera.LastExposureStartTime;
                        LogIssue("LastExposureStartTime", "LastExposureStartTime did not generate an exception when called before an exposure was made");
                    }
                    catch (Exception)
                    {
                        LogOK("LastExposureStartTime", "LastExposureStartTime correctly generated an exception before an exposure was made");
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
                m_Camera.BinX = 0;
                LogIssue("BinX Write", "Invalid value 0 written but no exception generated");
            }
            catch (Exception)
            {
                LogOK("BinX Write", "Exception correctly generated on setting BinX to 0");
            }

            try // Invalid high value
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                m_Camera.BinX = (short)(m_MaxBinX + 1);
                LogIssue("BinX Write", "Invalid value " + m_MaxBinX + 1 + " written but no exception generated");
            }
            catch (Exception)
            {
                LogOK("BinX Write", "Exception correctly generated on setting BinX to " + m_MaxBinX + 1);
            }

            try // Invalid low value
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinY");
                m_Camera.BinY = 0;
                LogIssue("BinY Write", "Invalid value 0 written but no exception generated");
            }
            catch (Exception)
            {
                LogOK("BinY Write", "Exception correctly generated on setting BinY to 0");
            }

            try // Invalid high value
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinY");
                m_Camera.BinY = (short)(m_MaxBinY + 1);
                LogIssue("BinY Write", "Invalid value " + m_MaxBinY + 1 + " written but no exception generated");
            }
            catch (Exception)
            {
                LogOK("BinY Write", "Exception correctly generated on setting BinY to " + m_MaxBinY + 1);
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
                for (l_BinY = 1; l_BinY <= l_MaxBinY; l_BinY++)
                {
                    for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                    {
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set BinY");
                            m_Camera.BinY = (short)l_BinY;
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                                m_Camera.BinX = (short)l_BinX;
                                LogOK("BinXY Write", "Successfully set asymmetric xy binning: " + l_BinX + " x " + l_BinY);
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
                            LogTestAndMessage("ConformanceCheck", "About to set BinY");
                        m_Camera.BinY = (short)l_BinX;
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set BinX");
                            m_Camera.BinX = (short)l_BinX;
                            LogOK("BinXY Write", "Successfully set symmetric xy binning: " + l_BinX + " x " + l_BinX);
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

            // Reset X and Y binning to 1x1 state
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                m_Camera.BinX = 1;
            }
            catch (Exception)
            {
            }

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set BinY");
            try
            {
                m_Camera.BinY = 1;
            }
            catch (Exception)
            {
            }

            m_CameraState = CameraPropertyTestCameraState(CamPropertyType.CameraState, "CameraState"); if (cancellationToken.IsCancellationRequested) return;
            m_CameraXSize = CameraPropertyTestInteger(CamPropertyType.CameraXSize, "CameraXSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;
            m_CameraYSize = CameraPropertyTestInteger(CamPropertyType.CameraYSize, "CameraYSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;

            CameraPropertyTestDouble(CamPropertyType.CCDTemperature, "CCDTemperature", MIN_CAMERA_SETPOINT_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;
            m_CoolerOn = CameraPropertyTestBoolean(CamPropertyType.CoolerOn, "CoolerOn Read", false); if (cancellationToken.IsCancellationRequested) return;

            // Write CoolerOn
            bool l_OriginalCoolerState;
            string l_TargetCoolerState;
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");
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
                            LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");
                        m_Camera.CoolerOn = false;
                    }
                    else
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");
                        m_Camera.CoolerOn = true;
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

            CameraPropertyTestDouble(CamPropertyType.CoolerPower, "CoolerPower", 0.0, 100.0, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.ElectronsPerADU, "ElectronsPerADU", 0.00001, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.FullWellCapacity, "FullWellCapacity", 0.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestBoolean(CamPropertyType.HasShutter, "HasShutter", false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.HeatSinkTemperature, "HeatSinkTemperature", MIN_CAMERA_SETPOINT_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;

            m_ImageReady = CameraPropertyTestBoolean(CamPropertyType.ImageReady, "ImageReady", false); if (cancellationToken.IsCancellationRequested) return;
            if (m_ImageReady & settings.CameraFirstUseTests) // Issue this warning if configured to do so
                LogIssue("ImageReady", "Image is flagged as ready but no exposure has been started!");

            // ImageArray 
            SetFullStatus("ImageArray", "Getting image data from device...", "");

            // Release memory currently consumed by images
            ReleaseMemory();

            if (m_ImageReady) // ImageReady is true
            {
                try
                {
                    if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get ImageArray");
                    m_ImageArray = (Array)m_Camera.ImageArray;
                    if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                    {
                        LogIssue("ImageArray", "No image has been taken but ImageArray has not generated an exception");
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
                        LogOK("ImageArray", "Exception correctly generated before an image has been taken");
                    }
                    else // Omit first use tests
                    {
                        LogIssue("ImageArray", $"Conform is configured to omit \"First use\" tests and ImageReady is true, but ImageArray generated an exception: {ex.Message}");
                    }
                }
            }
            else // ImageReady is false
            {
                try
                {
                    if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get ImageArray");
                    m_ImageArray = (Array)m_Camera.ImageArray;
                    LogIssue("ImageArray", "ImageReady is false and no image has been taken but ImageArray has not generated an exception");
                }
                catch (Exception ex)
                {
                    LogDebug("ImageArray", $"Exception 2:\r\n{ex}");
                    LogOK("ImageArray", "Exception correctly generated when ImageReady is false");
                }
            }

            // Release memory currently consumed by images
            SetAction("Releasing memory");
            ReleaseMemory();

            // ImageArrayVariant
            // Release memory currently consumed by images
            SetAction("Releasing memory");
            ReleaseMemory();

            SetFullStatus("ImageArrayVariant", "Getting image data from device...", "");
            if (settings.CameraTestImageArrayVariant) // Test if configured to do so
            {
                if (m_ImageReady)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get ImageArrayVariant");
                        m_ImageArrayVariant = (Array)m_Camera.ImageArrayVariant;

                        if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                        {
                            LogIssue("ImageArrayVariant", "No image has been taken but ImageArrayVariant has not generated an exception");
                        }
                        else
                        {
                            LogOK("ImageArrayVariant", "ImageArray read OK");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogDebug("ImageArrayVariant", $"Exception:\r\n{ex}");
                        if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                        {
                            LogOK("ImageArrayVariant", "Exception correctly generated before an image has been taken");
                        }
                        else // Omit first use tests
                        {
                            LogIssue("ImageArrayVariant", $"Conform is configured to omit \"First use\" tests and ImageReady is true, but ImageArrayVariant generated an exception: {ex.Message}");
                        }

                    }
                }
                else
                {
                    try
                    {
                        if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get ImageArrayVariant");
                        m_ImageArrayVariant = (Array)m_Camera.ImageArrayVariant;
                        LogIssue("ImageArrayVariant", "ImageReady is false and no image has been taken but ImageArray has not generated an exception");
                    }
                    catch (Exception)
                    {
                        LogOK("ImageArrayVariant", "Exception correctly generated when ImageReady is false");
                    }
                }
            }
            else // Log an issue because the ImageArrayVariant test was omitted
            {
                LogIssue("ImageArrayVariant", "Test omitted due to Conform configuration.");
            }

            // Release memory currently consumed by images
            SetAction("Releasing memory");
            ReleaseMemory();

            ClearStatus();

            m_IsPulseGuiding = CameraPropertyTestBoolean(CamPropertyType.IsPulseGuiding, "IsPulseGuiding", false); if (cancellationToken.IsCancellationRequested) return;
            if (m_IsPulseGuiding)
                LogIssue("IsPulseGuiding", "Camera is showing pulse guiding underway although no PulseGuide command has been issued!");

            CameraPropertyTestInteger(CamPropertyType.MaxADU, "MaxADU", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;

            CameraPropertyTestInteger(CamPropertyType.NumX, "NumX Read", 1, m_CameraXSize); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.NumX, "NumX", System.Convert.ToInt32(m_CameraXSize / (double)2));

            CameraPropertyTestInteger(CamPropertyType.NumY, "NumY Read", 1, m_CameraYSize); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.NumY, "NumY", System.Convert.ToInt32(m_CameraYSize / (double)2));

            CameraPropertyTestDouble(CamPropertyType.PixelSizeX, "PixelSizeX", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.PixelSizeY, "PixelSizeY", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;

            m_SetCCDTemperature = CameraPropertyTestDouble(CamPropertyType.SetCCDTemperature, "SetCCDTemperature Read", MIN_CAMERA_SETPOINT_TEMPERATURE, MAX_CAMERA_SETPOINT_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;

            if (m_CanSetCCDTemperature)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
                    m_Camera.SetCCDTemperature = 0.0; // Try an acceptable value
                    LogOK("SetCCDTemperature Write", "Successfully wrote 0.0");

                    // Execution only gets here if the CCD temperature can be set successfully
                    bool l_ExceptionGenerated;
                    double l_SetPoint;

                    // Find low setpoint at which an exception is generated, stop at -280 as this is unphysical
                    l_ExceptionGenerated = false;
                    l_SetPoint = -0.0;
                    LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature multiple times...");
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
                    while (!l_ExceptionGenerated & (l_SetPoint > MIN_CAMERA_SETPOINT_TEMPERATURE))// Reached lower limit so exit loop
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
                            LogInfo("SetCCDTemperature Write", $"Setpoint can be set to {MIN_CAMERA_SETPOINT_TEMPERATURE} degrees");
                        else
                            LogIssue("SetCCDTemperature Write", $"Setpoint can be set below {MIN_CAMERA_SETPOINT_TEMPERATURE} degrees, which is below absolute zero!");
                    }
                    else
                        LogInfo("SetCCDTemperature Write", $"Setpoint lower limit found in the range {l_SetPoint + 5.0} to {l_SetPoint + 0.001} degrees");

                    // Find high setpoint at which an exception is generated, stop at MAX_CAMERA_SETPOINT_TEMPERATURE as this is a suitably high value
                    l_ExceptionGenerated = false;
                    l_SetPoint = 0.0; // Start at 0.0C

                    // Loop upward in 5 degree temperature steps to find the maximum temperature that can be set
                    LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature multiple times...");
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
                    while (!l_ExceptionGenerated & (l_SetPoint < MAX_CAMERA_SETPOINT_TEMPERATURE));// Reached upper limit so exit loop

                    if (!l_ExceptionGenerated & (l_SetPoint == MAX_CAMERA_SETPOINT_TEMPERATURE))
                    {
                        // Now test whether it is possible to set the temperature just above the maximum setpoint, which should result in an exception if all is well
                        l_ExceptionGenerated = false;
                        try
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
                            m_Camera.SetCCDTemperature = MAX_CAMERA_SETPOINT_TEMPERATURE + CAMERA_SETPOINT_TEST_INCREMENT;
                        }
                        catch (Exception)
                        {
                            l_ExceptionGenerated = true;
                        }// Reached upper limit

                        if (l_ExceptionGenerated)
                            LogInfo("SetCCDTemperature Write", $"Setpoint can be set to {MAX_CAMERA_SETPOINT_TEMPERATURE} degrees");
                        else
                            LogIssue("SetCCDTemperature Write", $"Setpoint can be set in excess of {MAX_CAMERA_SETPOINT_TEMPERATURE} degrees");
                    }
                    else
                        LogInfo("SetCCDTemperature Write", $"Setpoint upper limit found in the range {l_SetPoint - 5.0} to {l_SetPoint - 0.001} degrees");
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
                        LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
                    m_Camera.SetCCDTemperature = 0;
                    LogIssue("SetCCDTemperature Write", "CanSetCCDTemperature is false but no exception generated on write");
                }
                catch (Exception ex)
                {
                    HandleException("SetCCDTemperature Write", MemberType.Property, Required.Optional, ex, "");
                }

            CameraPropertyTestInteger(CamPropertyType.StartX, "StartX Read", 0, m_CameraXSize - 1); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.StartX, "StartX", System.Convert.ToInt32(m_CameraXSize / (double)2));
            CameraPropertyTestInteger(CamPropertyType.StartY, "StartY Read", 0, m_CameraYSize - 1); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.StartY, "StartY", System.Convert.ToInt32(m_CameraYSize / (double)2));

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to get InterfaceVersion");
            if (m_Camera.InterfaceVersion > 1)
            {
                // SensorType - Mandatory
                // This must be tested before BayerOffset because BayerOffset is mandatory for colour and optional for monochrome cameras
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get SensorType");
                    m_SensorType = m_Camera.SensorType;
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
                    if (m_SensorType == SensorType.Monochrome)
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
                        m_FastReadout = m_Camera.FastReadout;
                        LogIssue("FastReadout Read", "CanFastReadout is False but a PropertyNotImplementedException was not thrown");
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
                        m_Camera.FastReadout = !m_FastReadout;
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set FastReadout");
                        m_Camera.FastReadout = m_FastReadout;
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
                        m_Camera.FastReadout = true;
                        LogIssue("FastReadout Write", "CanFastReadout is False but a PropertyNotImplementedException was not thrown");
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
                    m_GainMin = m_Camera.GainMin;
                    // Successfully retrieved a value
                    m_CanReadGainMin = true;
                    LogOK("GainMin Read", m_GainMin.ToString());
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
                        LogTestAndMessage("ConformanceCheck", "About to get GainMax");
                    m_GainMax = m_Camera.GainMax;
                    // Successfully retrieved a value
                    m_CanReadGainMax = true;
                    LogOK("GainMax Read", m_GainMax.ToString());
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
                        LogTestAndMessage("ConformanceCheck", "About to get Gains");
                    m_Gains = m_Camera.Gains;
                    // Successfully retrieved a value
                    m_CanReadGains = true;
                    foreach (string Gain in m_Gains)
                        LogOK("Gains Read", Gain.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("Gains Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                if (m_CanReadGainMax & m_CanReadGainMin & m_CanReadGains)
                    LogIssue("Gains", "GainMin, GainMax and Gains are all readable. Only one of GainMin/Max as a pair or Gains should be used, the other should throw a PropertyNotImplementedException");
                else
                {
                }

                // Gain Read - Optional 
                try
                {
                    m_CanReadGain = false; // Set default value to indicate can't read gain
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get Gain");
                    m_Gain = m_Camera.Gain;
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
                    LogOK("Gain Read", "All four gain properties throw exceptions - the driver is in \"Gain Not Implemented\" mode.");
                else if (m_CanReadGain)
                {
                    // Test for Gain Index mode
                    if ((m_CanReadGain & m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax))
                    {
                        m_GainMode = GainOffsetMode.IndexMode;
                        LogOK("Gain Read", "Gain and Gains can be read while GainMin and GainMax throw exceptions - the driver is in \"Gain Index\" mode.");
                    }
                    else if ((m_CanReadGain & !m_CanReadGains & m_CanReadGainMin & m_CanReadGainMax))
                    {
                        m_GainMode = GainOffsetMode.ValueMode;
                        LogOK("Gain Read", "Gain, GainMin and GainMax can be read OK while Gains throws an exception - the driver is in \"Gain Value\" mode.");
                    }
                    else
                    {
                        LogIssue("Gain Read", $"Unable to determine whether the driver is in \"Gain Not Implemented\", \"Gain Index\" or \"Gain Value\" mode. Please check the interface specification.");
                        LogInfo("Gain Read", $"Gain threw an exception: {m_CanReadGain}, Gains threw an exception: {m_CanReadGains}, GainMin threw an exception: {m_CanReadGainMin}, GainMax threw an exception: {m_CanReadGainMax}.");
                        LogInfo("Gain Read", $"\"Gain Not Implemented\" mode: Gain, Gains, GainMin and GainMax must all throw exceptions.");
                        LogInfo("Gain Read", $"\"Gain Index\" mode: Gain and Gains must work while GainMin and GainMax must throw exceptions.");
                        LogInfo("Gain Read", $"\"Gain Value\" mode: Gain, GainMin and GainMax must work while Gains must throw an exception.");
                    }
                }
                else
                {
                    LogIssue("Gain Read", $"Gain Read threw an exception but at least one of Gains, GainMin Or GainMax did not throw an exception. If Gain throws an exception, all the other gain properties should do likewise.");
                    LogInfo("Gain Read", $"Gains threw an exception: {m_CanReadGains}, GainMin threw an exception: {m_CanReadGainMin}, GainMax threw an exception: {m_CanReadGainMax}.");
                }

                // Gain write - Optional when neither gain index nor gain value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither gain mode is supported
                if (!m_CanReadGain & !m_CanReadGains & !m_CanReadGainMin & !m_CanReadGainMax)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set Gain");
                        m_Camera.Gain = 0;
                        LogIssue("Gain Write", "Writing to Gain did not throw a PropertyNotImplementedException when reading Gain did.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("Gain Write", MemberType.Property, Required.MustNotBeImplemented, ex, "PropertyNotImplementedException is expected");
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
                                    m_Camera.Gain = m_GainMin;
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
                                    m_Camera.Gain = m_GainMax;
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
                                    m_Camera.Gain = (short)(m_GainMin - 1);
                                    LogIssue("Gain Write", $"Successfully set an gain below the minimum value ({m_GainMin - 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for gain {m_GainMin - 1}, which is lower than the minimum value.");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Gain");
                                    m_Camera.Gain = (short)(m_GainMax + 1);
                                    LogIssue("Gain Write", $"Successfully set an gain above the maximum value({m_GainMax + 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for gain {m_GainMax + 1} which is higher than the maximum value.");
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
                    m_PercentCompleted = m_Camera.PercentCompleted;
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
                    m_ReadoutModes = m_Camera.ReadoutModes;
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
                        LogIssue("ReadoutMode Index", "Exception: " + ex.Message);
                        LogDebug("ReadoutMode Index", ex.ToString());
                    }
                }
                else
                    LogInfo("ReadoutMode Index", "Skipping ReadReadoutMode index test because ReadoutModes is unavailable");

                // SensorName
                CameraPropertyTestString(CamPropertyType.SensorName, "SensorName Read", 250, true);
            }

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to get InterfaceVersion");
            if (m_Camera.InterfaceVersion > 2)
            {
                // OffsetMin Read - Optional
                try
                {
                    m_CanReadOffsetMin = false;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get OffsetMin");
                    m_OffsetMin = m_Camera.OffsetMin;
                    // Successfully retrieved a value
                    m_CanReadOffsetMin = true;
                    LogOK("OffsetMin Read", m_OffsetMin.ToString());
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
                        LogTestAndMessage("ConformanceCheck", "About to get OffsetMax");
                    m_OffsetMax = m_Camera.OffsetMax;
                    // Successfully retrieved a value
                    m_CanReadOffsetMax = true;
                    LogOK("OffsetMax Read", m_OffsetMax.ToString());
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
                        LogTestAndMessage("ConformanceCheck", "About to get Offsets");
                    m_Offsets = m_Camera.Offsets;
                    // Successfully retrieved a value
                    m_CanReadOffsets = true;
                    foreach (string Offset in m_Offsets)
                        LogOK("Offsets Read", Offset.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOK("Offsets Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperationException correctly thrown");
                }

                if (m_CanReadOffsetMax & m_CanReadOffsetMin & m_CanReadOffsets)
                    LogIssue("Offsets", "OffsetMin, OffsetMax and Offsets are all readable. Only one of OffsetMin/Max as a pair or Offsets should be used, the other should throw a PropertyNotImplementedException");
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
                    m_Offset = m_Camera.Offset;
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
                    LogOK("Offset Read", "All four offset properties throw exceptions - the driver is in \"Offset Not Implemented\" mode.");
                else if (m_CanReadOffset)
                {
                    // Test for Offset Index mode
                    if ((m_CanReadOffset & m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax))
                    {
                        m_OffsetMode = GainOffsetMode.IndexMode;
                        LogOK("Offset Read", "Offset and Offsets can be read while OffsetMin and OffsetMax throw exceptions - the driver is in \"Offset Index\" mode.");
                    }
                    else if ((m_CanReadOffset & !m_CanReadOffsets & m_CanReadOffsetMin & m_CanReadOffsetMax))
                    {
                        m_OffsetMode = GainOffsetMode.ValueMode;
                        LogOK("Offset Read", "Offset, OffsetMin and OffsetMax can be read OK while Offsets throws an exception - the driver is in \"Offset Value\" mode.");
                    }
                    else
                    {
                        m_OffsetMode = GainOffsetMode.Unknown;
                        LogIssue("Offset Read", $"Unable to determine whether the driver is in \"Offset Not Implemented\", \"Offset Index\" or \"Offset Value\" mode. Please check the interface specification.");
                        LogInfo("Offset Read", $"Offset threw an exception: {m_CanReadOffset}, Offsets threw an exception: {m_CanReadOffsets}, OffsetMin threw an exception: {m_CanReadOffsetMin}, OffsetMax threw an exception: {m_CanReadOffsetMax}.");
                        LogInfo("Offset Read", $"\"Offset Not Implemented\" mode: Offset, Offsets, OffsetMin and OffsetMax must all throw exceptions.");
                        LogInfo("Offset Read", $"\"Offset Index\" mode: Offset and Offsets must work while OffsetMin and OffsetMax must throw exceptions.");
                        LogInfo("Offset Read", $"\"Offset Value\" mode: Offset, OffsetMin and OffsetMax must work while Offsets must throw an exception.");
                    }
                }
                else
                {
                    LogIssue("Offset Read", $"Offset Read threw an exception but at least one of Offsets, OffsetMin Or OffsetMax did not throw an exception. If Offset throws an exception, all the other offset properties should do likewise.");
                    LogInfo("Offset Read", $"Offsets threw an exception: {m_CanReadOffsets}, OffsetMin threw an exception: {m_CanReadOffsetMin}, OffsetMax threw an exception: {m_CanReadOffsetMax}.");
                }

                // Offset write - Optional when neither offset index nor offset value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither offset mode is supported
                if (!m_CanReadOffset & !m_CanReadOffsets & !m_CanReadOffsetMin & !m_CanReadOffsetMax)
                {
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to set Offset");
                        m_Camera.Offset = 0;
                        LogIssue("Offset Write", "Writing to Offset did not throw a PropertyNotImplementedException when reading Offset did.");
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
                                    m_Camera.Offset = m_OffsetMin;
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
                                    m_Camera.Offset = m_OffsetMax;
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
                                    m_Camera.Offset = m_OffsetMin - 1;
                                    LogIssue("Offset Write", $"Successfully set an offset below the minimum value ({m_OffsetMin - 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for offset {m_OffsetMin - 1}, which is lower than the minimum value.");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to set Offset");
                                    m_Camera.Offset = m_OffsetMax + 1;
                                    LogIssue("Offset Write", $"Successfully set an offset above the maximum value({m_OffsetMax + 1}), this should have resulted in an InvalidValueException.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValueException is expected.", $"InvalidValueException correctly generated for offset {m_OffsetMax + 1} which is higher than the maximum value.");
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
                    m_Camera.SubExposureDuration = m_SubExposureDuration;
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
                            returnValue = m_Camera.CameraState;
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
                            returnValue = m_Camera.BayerOffsetX;
                            break;
                        }

                    case CamPropertyType.BayerOffsetY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get BayerOffsetY");
                            returnValue = m_Camera.BayerOffsetY;
                            break;
                        }

                    case CamPropertyType.PercentCompleted:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get PercentCompleted");
                            returnValue = m_Camera.PercentCompleted;
                            break;
                        }

                    case CamPropertyType.ReadoutMode:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ReadoutMode");
                            returnValue = m_Camera.ReadoutMode;
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
                            TestShort = m_Camera.BayerOffsetX;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogIssue(p_Name, "Sensor type is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value");
                            break;
                        }

                    case CamPropertyType.BayerOffsetY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get BayerOffsetY");
                            TestShort = m_Camera.BayerOffsetY;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogIssue(p_Name, "Sensor type is Monochrome so this property must throw a PropertyNotImplementedException; it must not return a value");
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
                            returnValue = m_Camera.BinX;
                            break;
                        }

                    case CamPropertyType.BinY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get BinY");
                            returnValue = m_Camera.BinY;
                            break;
                        }

                    case CamPropertyType.CameraState:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                            returnValue = (int)m_Camera.CameraState;
                            break;
                        }

                    case CamPropertyType.CameraXSize:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CameraXSize");
                            returnValue = m_Camera.CameraXSize;
                            break;
                        }

                    case CamPropertyType.CameraYSize:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CameraYSize");
                            returnValue = m_Camera.CameraYSize;
                            break;
                        }

                    case CamPropertyType.MaxADU:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get MaxADU");
                            returnValue = m_Camera.MaxADU;
                            break;
                        }

                    case CamPropertyType.MaxBinX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get MaxBinX");
                            returnValue = m_Camera.MaxBinX;
                            break;
                        }

                    case CamPropertyType.MaxBinY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get MaxBinY");
                            returnValue = m_Camera.MaxBinY;
                            break;
                        }

                    case CamPropertyType.NumX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get NumX");
                            returnValue = m_Camera.NumX;
                            break;
                        }

                    case CamPropertyType.NumY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get NumY");
                            returnValue = m_Camera.NumY;
                            break;
                        }

                    case CamPropertyType.StartX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get StartX");
                            returnValue = m_Camera.StartX;
                            break;
                        }

                    case CamPropertyType.StartY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get StartY");
                            returnValue = m_Camera.StartY;
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
                            returnValue = m_Camera.CCDTemperature;
                            break;
                        }

                    case CamPropertyType.CoolerPower:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get CoolerPower");
                            returnValue = m_Camera.CoolerPower;
                            break;
                        }

                    case CamPropertyType.ElectronsPerADU:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ElectronsPerADU");
                            returnValue = m_Camera.ElectronsPerADU;
                            break;
                        }

                    case CamPropertyType.FullWellCapacity:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get FullWellCapacity");
                            returnValue = m_Camera.FullWellCapacity;
                            break;
                        }

                    case CamPropertyType.HeatSinkTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get HeatSinkTemperature");
                            returnValue = m_Camera.HeatSinkTemperature;
                            break;
                        }

                    case CamPropertyType.PixelSizeX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get PixelSizeX");
                            returnValue = m_Camera.PixelSizeX;
                            break;
                        }

                    case CamPropertyType.PixelSizeY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get PixelSizeY");
                            returnValue = m_Camera.PixelSizeY;
                            break;
                        }

                    case CamPropertyType.SetCCDTemperature:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get SetCCDTemperature");
                            returnValue = m_Camera.SetCCDTemperature;
                            break;
                        }

                    case CamPropertyType.ExposureMax:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ExposureMax");
                            returnValue = m_Camera.ExposureMax;
                            break;
                        }

                    case CamPropertyType.ExposureMin:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ExposureMin");
                            returnValue = m_Camera.ExposureMin;
                            break;
                        }

                    case CamPropertyType.ExposureResolution:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ExposureResolution");
                            returnValue = m_Camera.ExposureResolution;
                            break;
                        }

                    case CamPropertyType.SubExposureDuration:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get SubExposureDuration");
                            returnValue = m_Camera.SubExposureDuration;
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
                            returnValue = m_Camera.CoolerOn;
                            break;
                        }

                    case CamPropertyType.HasShutter:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get HasShutter");
                            returnValue = m_Camera.HasShutter;
                            break;
                        }

                    case CamPropertyType.ImageReady:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                            returnValue = m_Camera.ImageReady;
                            break;
                        }

                    case CamPropertyType.IsPulseGuiding:
                        {
                            m_IsPulseGuidingSupported = false;
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding");
                            returnValue = m_Camera.IsPulseGuiding;
                            m_IsPulseGuidingSupported = true; // Command works properly and doesn't cause a not implemented exception
                            break;
                        }

                    case CamPropertyType.FastReadout:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get FastReadout");
                            returnValue = m_Camera.FastReadout;
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
                            returnValue = m_Camera.Description;
                            break;
                        }

                    case CamPropertyType.SensorName:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get SensorName");
                            returnValue = m_Camera.SensorName;
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
                            m_Camera.NumX = p_TestOK;
                            break;
                        }

                    case CamPropertyType.NumY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set NumY");
                            m_Camera.NumY = p_TestOK;
                            break;
                        }

                    case CamPropertyType.StartX:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set StartX");
                            m_Camera.StartX = p_TestOK;
                            break;
                        }

                    case CamPropertyType.StartY:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to set StartY");
                            m_Camera.StartY = p_TestOK;
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
            int l_i, l_j, l_MaxBinX, l_MaxBinY;
            // AbortExposure - Mandatory
            SetTest("AbortExposure");
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                m_CameraState = m_Camera.CameraState;
                switch (m_CameraState)
                {
                    case CameraState.Idle // All is OK so test a call to AbortExposure
                   :
                        {
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("ConformanceCheck", "About to call AbortExposure");
                                m_Camera.AbortExposure();
                                if (m_CanAbortExposure)
                                    LogOK("AbortExposure", "No exception generated when camera is already idle");
                                else
                                    LogIssue("AbortExposure", "CanAbortExposure is false but no exception is generated when AbortExposure is called");
                            }
                            catch (Exception)
                            {
                                if (m_CanAbortExposure)
                                    LogIssue("AbortExposure", "Exception incorrectly generated when camera is idle");
                                else
                                    LogOK("AbortExposure", "CanAbortExposure is false and an exception was correctly generated");
                            }

                            break;
                        }

                    default:
                        {
                            LogIssue("AbortExposure", "Camera is not idle, further AbortExposure tests skipped: " + m_CameraState.ToString());
                            break;
                        }
                }
            }
            catch (Exception)
            {
                LogIssue("AbortExposure", "Exception generated when reading camera state, further AbortExposure tests skipped");
            }

            // PulseGuide
            SetTest("PulseGuide");
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
                catch (Exception ex)
                {
                    LogIssue("PulseGuide", "CanPulseGuide is true but exception generated when calling method - " + ex.Message);
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to call PulseGuide - North");
                    m_Camera.PulseGuide(GuideDirection.North, 0);
                    LogIssue("PulseGuide", "CanPulseGuide is false but no exception generated when calling method");
                }
                catch (COMException)
                {
                    LogOK("PulseGuide", "CanPulseGuide is false and exception correctly generated when calling method");
                }
                catch (MethodNotImplementedException)
                {
                    LogOK("PulseGuide", "CanPulseGuide is false and PulseGuide is not implemented in this driver");
                }
                catch (Exception)
                {
                    LogOK("PulseGuide", "CanPulseGuide is false and exception correctly generated when calling method");
                }

            // StopExposure
            SetTest("StopExposure");
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to get CameraState");
                m_CameraState = m_Camera.CameraState;
                switch (m_CameraState)
                {
                    case CameraState.Idle // all is OK so test that a call to StopExposure doesn't generate an exception
                   :
                        {
                            try
                            {
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("ConformanceCheck", "About to call StopExposure");
                                SetAction("Calling StopExposure()");
                                m_Camera.StopExposure();
                                if (m_CanStopExposure)
                                    LogOK("StopExposure", "No exception generated when camera is already idle");
                                else
                                    LogIssue("StopExposure", "CanStopExposure is false but no exception is generated when StopExposure is called");
                            }
                            catch (Exception)
                            {
                                if (m_CanStopExposure)
                                    LogIssue("StopExposure", "Exception incorrectly generated when camera is idle");
                                else
                                    LogOK("StopExposure", "CanStopExposure is false and an exception was correctly generated");
                            }

                            break;
                        }

                    default:
                        {
                            LogIssue("StopExposure", "Camera is not idle, further StopExposure tests skipped: " + m_CameraState.ToString());
                            break;
                        }
                }
            }
            catch (Exception)
            {
                LogIssue("StopExposure", "Exception generated when reading camera state, further StopExposure tests skipped");
            }

            // StartExposure
            SetTest("StartExposure");

            // Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogNewLine(); // Insert a blank line if required
            if (settings.CameraMaxBinX > 0)
            {
                l_MaxBinX = settings.CameraMaxBinX;
                LogTestAndMessage("StartExposure", string.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", l_MaxBinX, m_MaxBinX));
            }
            else
                l_MaxBinX = m_MaxBinX;
            if (settings.CameraMaxBinY > 0)
            {
                l_MaxBinY = settings.CameraMaxBinY;
                LogTestAndMessage("StartExposure", string.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", l_MaxBinY, m_MaxBinY));
            }
            else
                l_MaxBinY = m_MaxBinY;

            if ((settings.CameraMaxBinX > m_MaxBinX) | (settings.CameraMaxBinY > m_MaxBinY))
                LogNewLine(); // Insert a blank line if required
            if (settings.CameraMaxBinX > m_MaxBinX)
                LogTestAndMessage("StartExposure", string.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", l_MaxBinX, m_MaxBinX));
            if (settings.CameraMaxBinY > m_MaxBinY)
                LogTestAndMessage("StartExposure", string.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", l_MaxBinY, m_MaxBinY));

            // StartExposure - Confirm that correct operation occurs
            int l_BinX, l_BinY;
            if (m_CanAsymmetricBin)
            {
                for (l_BinY = 1; l_BinY <= l_MaxBinY; l_BinY++)
                {
                    for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                    {
                        CameraExposure($"Take image full frame {l_BinX} x {l_BinY} bin ({m_CameraXSize / l_BinX}) x {m_CameraYSize / l_BinY})", l_BinX, l_BinY, 0, 0, m_CameraXSize / l_BinX, m_CameraYSize / l_BinY, settings.CameraExposureDuration, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
            }
            else
                for (l_BinX = 1; l_BinX <= l_MaxBinX; l_BinX++)
                {
                    CameraExposure($"Take image full frame {l_BinX} x {l_BinX} bin ({m_CameraXSize / l_BinX}) x {m_CameraYSize / l_BinX})", l_BinX, l_BinX, 0, 0, m_CameraXSize / l_BinX, m_CameraYSize / l_BinX, settings.CameraExposureDuration, "");
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }

            // StartExposure - Confirm error cases
            LogNewLine();
            LogTestOnly("StartExposure error cases");
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
                LogNewLine(); // Blank Line
                LogTestOnly(p_Description);
            }
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinX");
                m_Camera.BinX = (short)p_BinX;
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("ConformanceCheck", "About to set BinY");
                m_Camera.BinY = (short)p_BinY;
                try
                {
                    l_ExposeOK = false; // Start off by assuming the worst
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set StartX");
                    m_Camera.StartX = p_StartX;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set StartY");
                    m_Camera.StartY = p_StartY;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set NumX");
                    m_Camera.NumX = p_NumX;
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to set NumY");
                    m_Camera.NumY = p_NumY;
                    try
                    {
                        SetAction($"Exposing for {p_Duration} seconds");

                        // Initiate exposure
                        l_StartTime = DateTime.Now;
                        l_StartTimeUTC = DateTime.UtcNow;
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to call StartExposure");
                        m_Camera.StartExposure(p_Duration, true);

                        if (p_ExpectedErrorMessage == "")
                        {
                            l_EndTime = DateTime.Now;

                            // Test whether we have a synchronous or asynchronous camera
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to get ImageReady and CameraState");
                            if (m_Camera.ImageReady & (m_Camera.CameraState == CameraState.Idle)) // Synchronous exposure
                            {
                                if (l_EndTime.Subtract(l_StartTime).TotalSeconds >= p_Duration)
                                {
                                    LogOK("StartExposure", "Synchronous exposure found OK: " + p_Duration + " seconds");
                                    CameraTestLast(p_Duration, l_StartTimeUTC);
                                }
                                else
                                    LogIssue("StartExposure", "Synchronous exposure found but image was returned before exposure time was complete");
                            }
                            else
                            {
                                SetStatus("Waiting for exposure to start");

                                // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                                imageReadyTooEarly = Convert.ToBoolean(m_Camera.ImageReady);

                                // Wait for exposing state
                                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get CameraState multiple times");
                                //do
                                //{
                                //    WaitFor(CAMERA_SLEEP_TIME);
                                //    if (cancellationToken.IsCancellationRequested)
                                //        return;
                                //}
                                //while ((m_Camera.CameraState != CameraState.Exposing) & (m_Camera.CameraState != CameraState.Error));
                                Stopwatch sw = Stopwatch.StartNew();
                                WaitWhile("", () => { return (m_Camera.CameraState != CameraState.Exposing) & (m_Camera.CameraState != CameraState.Error); }, 500, Convert.ToInt32(p_Duration + 2), () => { return $"{sw.Elapsed.TotalSeconds:0.0} / {p_Duration:0.0}"; });

                                // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                                imageReadyTooEarly = m_Camera.ImageReady;

                                // Wait for the exposing state to finish
                                l_StartTime = DateTime.Now;
                                l_StartTimeUTC = DateTime.UtcNow;
                                SetAction($"Waiting for exposure to complete");
                                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get CameraState, InterfaceVersion and PercentCompleted multiple times...");

                                // Start the loop timing stopwatch
                                sw.Restart();
                                Stopwatch swOverall = Stopwatch.StartNew();
                                const int POLL_INTERVAL = 500;
                                do
                                {
                                    // Calculate the current loop number (starts at 1 given that the timer's elapsed time will be zero or very low on the first loop)
                                    int currentLoopNumber = (int)sw.ElapsedMilliseconds / POLL_INTERVAL;

                                    // Calculate the sleep time required to start the next loop at a multiple of the poll interval
                                    int sleeptime = POLL_INTERVAL * (currentLoopNumber + 1) - (int)sw.ElapsedMilliseconds;

                                    // Sleep until it is time for the next completion function poll
                                    Thread.Sleep(sleeptime);

                                    l_PercentCompletedMessage = "Not present in a V1 driver"; // Initialise PercentCompleted message
                                    if (m_Camera.InterfaceVersion > 1)
                                    {
                                        try
                                        {
                                            l_PercentCompleted = m_Camera.PercentCompleted;
                                            l_PercentCompletedMessage = $"{l_PercentCompleted}%"; // Operation completed OK
                                        }
                                        catch (COMException ex)
                                        {
                                            switch (ex.ErrorCode)
                                            {
                                                case int i when i == ErrorCodes.NotImplemented:
                                                    {
                                                        l_PercentCompletedMessage = "Not implemented"; // Not implemented
                                                        break;
                                                    }

                                                case int j when j == ErrorCodes.InvalidOperationException:
                                                    {
                                                        l_PercentCompletedMessage = "Invalid operation"; // Not valid at this time
                                                        break;
                                                    }

                                                default:
                                                    {
                                                        l_PercentCompletedMessage = "Exception: " + ex.Message + " 0x" + ((int)ex.ErrorCode).ToString("X8"); // Something bad happened!
                                                        break;
                                                    }
                                            }
                                        }
                                        catch (ASCOM.InvalidOperationException)
                                        {
                                            l_PercentCompletedMessage = ".NET - Invalid operation"; // Not valid at this time
                                        }
                                        catch (PropertyNotImplementedException)
                                        {
                                            l_PercentCompletedMessage = "Not implemented"; // Not implemented
                                        }
                                        catch (Exception ex)
                                        {
                                            l_PercentCompletedMessage = "Exception: " + ex.Message; // Something bad happened!
                                        }
                                    }

                                    SetStatus($"{swOverall.Elapsed.TotalSeconds:0.0} / {p_Duration:0.0},   PercentComplete: {l_PercentCompletedMessage}");
                                    if (cancellationToken.IsCancellationRequested) return;
                                }
                                while (m_Camera.CameraState == CameraState.Exposing);

                                // Wait for camera to become idle
                                l_EndTime = DateTime.Now;
                                SetAction("Waiting for camera idle state, reading/downloading image");
                                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get CameraState multiple times");
                                do
                                {
                                    WaitFor(CAMERA_SLEEP_TIME);
                                    if (cancellationToken.IsCancellationRequested) return;
                                }
                                while ((m_Camera.CameraState != CameraState.Idle) & (m_Camera.CameraState != CameraState.Error));

                                // Wait for image to become ready
                                SetAction("Waiting for image ready");
                                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get CameraState multiple times");
                                do
                                {
                                    WaitFor(CAMERA_SLEEP_TIME);
                                    if (cancellationToken.IsCancellationRequested) return;
                                }
                                while (!m_Camera.ImageReady & (m_Camera.CameraState != CameraState.Error));

                                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get ImageReady");
                                if (m_Camera.ImageReady)
                                {
                                    LogOK("StartExposure", "Asynchronous exposure found OK: " + p_Duration + " seconds");
                                    CameraTestLast(p_Duration, l_StartTimeUTC);
                                }
                                else
                                    LogIssue("StartExposure", "Camera state is CameraError");
                            }

                            // Display a warning if ImageReady was set too early
                            if (imageReadyTooEarly)
                            {
                                LogIssue("StartExposure", "ImageReady was set True before the camera completed its exposure.");
                            }
                            // Camera exposed OK and didn't generate an exception
                            else
                            {
                                l_ExposeOK = true;
                            }
                            LogDebug("StartExposure", $"Camera exposed image OK: {l_ExposeOK}");

                            // Check image array dimensions
                            try
                            {
                                // Release memory currently consumed by images
                                SetAction("Releasing memory");
                                ReleaseMemory();

                                // Retrieve the image array
                                SetAction("Retrieving ImageArray");
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("ConformanceCheck", "About to get ImageArray");
                                sw.Restart();
                                m_ImageArray = (Array)m_Camera.ImageArray;
                                sw.Stop();
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("ConformanceCheck", "Get ImageArray completed in " + sw.ElapsedMilliseconds + "ms");

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
                                        LogOK("ImageArray", $"Successfully read 32 bit integer array ({l_NumPlanes}) {m_ImageArray.GetLength(0)} x {m_ImageArray.GetLength(1)} pixels in {sw.ElapsedMilliseconds}ms.");
                                    }
                                    else
                                        LogIssue("ImageArray", "Expected 32 bit integer array, actually got: " + m_ImageArray.GetType().ToString());
                                }
                                else if ((m_ImageArray.GetLength(0) == p_NumY) & (m_ImageArray.GetLength(1) == p_NumX))
                                    LogIssue("ImageArray", "Camera image dimensions swapped, expected values: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
                                else
                                    LogIssue("ImageArray", "Camera image does not have the expected dimensions of: " + p_NumX + " x " + p_NumY + " - actual values: " + m_ImageArray.GetLength(0) + " x " + m_ImageArray.GetLength(1));
                            }
                            catch (OutOfMemoryException ex)
                            {
                                if (Environment.Is64BitProcess) // Message when running as a 64bit process
                                    LogError("ImageArray", $"InsufficientMemoryException - The application ran out of available memory: {ex.Message}");
                                else // Message when running as a 32bit process
                                    LogError("ImageArray", $"InsufficientMemoryException - The application ran out of available memory.\r\n" +
                                        new string(' ', 58) +
                                        $"***** If your camera device supports this, please re-test with the 64bit version of Conform Universal because it has greater memory headroom.");
                            }
                            catch (Exception ex)
                            {
                                LogIssue("ImageArray", $"Exception when reading ImageArray: {ex.Message}");
                            }

                            // Release memory currently consumed by images
                            SetAction("Releasing memory");
                            ReleaseMemory();

                            // Check image array variant dimensions

                            if (settings.CameraTestImageArrayVariant) // Test if configured to do so. No need to report an issue because it's already been reported when the ImageArrayVariant property was tested
                            {
                                // Release memory currently consumed by images
                                SetAction("Releasing memory");
                                ReleaseMemory();

                                SetAction("Retrieving ImageArrayVariant");

                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "About to get ImageArrayVariant");
                                    sw.Restart();
                                    object imageObject = m_Camera.ImageArrayVariant;
                                    sw.Stop();

                                    GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;

                                    Array imageArr4ayVariant = (Array)imageObject;

                                    GC.Collect(2, GCCollectionMode.Forced, true, true);

                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("ConformanceCheck", "Get ImageArrayVariant completed in " + sw.ElapsedMilliseconds + "ms");

                                    if ((imageArr4ayVariant.GetLength(0) == p_NumX) & (imageArr4ayVariant.GetLength(1) == p_NumY))
                                    {
                                        if (imageArr4ayVariant.GetType().ToString() == "System.Object[,]" | imageArr4ayVariant.GetType().ToString() == "System.Object[,,]")
                                        {
                                            if (imageArr4ayVariant.Rank == 2)
                                            {
                                                l_NumPlanes = "1 plane";
                                                l_VariantType = ((object[,])imageArr4ayVariant)[0, 0].GetType().ToString();
                                            }
                                            else
                                            {
                                                l_NumPlanes = "1 plane";
                                                if (imageArr4ayVariant.GetUpperBound(2) > 0)
                                                {
                                                    l_NumPlanes = System.Convert.ToString(imageArr4ayVariant.GetUpperBound(2) + 1) + " planes";
                                                    l_VariantType = ((object[,,])imageArr4ayVariant)[0, 0, 0].GetType().ToString();
                                                }
                                                else
                                                    l_VariantType = ((object[,])imageArr4ayVariant)[0, 0].GetType().ToString();
                                            }
                                            LogOK("ImageArrayVariant", $"Successfully read variant array ({l_NumPlanes}) with {l_VariantType} elements {imageArr4ayVariant.GetLength(0)} x {imageArr4ayVariant.GetLength(1)} pixels in {sw.ElapsedMilliseconds}ms.");
                                        }
                                        else
                                            LogIssue("ImageArrayVariant", "Expected variant array, actually got: " + imageArr4ayVariant.GetType().ToString());
                                    }
                                    else if ((imageArr4ayVariant.GetLength(0) == p_NumY) & (imageArr4ayVariant.GetLength(1) == p_NumX))
                                        LogIssue("ImageArrayVariant", "Camera image dimensions swapped, expected values: " + p_NumX + " x " + p_NumY + " - actual values: " + imageArr4ayVariant.GetLength(0) + " x " + imageArr4ayVariant.GetLength(1));
                                    else
                                        LogIssue("ImageArrayVariant", "Camera image does not have the expected dimensions of: " + p_NumX + " x " + p_NumY + " - actual values: " + imageArr4ayVariant.GetLength(0) + " x " + imageArr4ayVariant.GetLength(1));
                                }
                                catch (OutOfMemoryException ex)
                                {
                                    if (Environment.Is64BitProcess) // Message when running as a 64bit process
                                        LogError("ImageArrayVariant", $"InsufficientMemoryException - The application ran out of available memory: {ex.Message}");
                                    else // Message when running as a 32bit process
                                        LogError("ImageArrayVariant", $"InsufficientMemoryException - The application ran out of available memory.\r\n" +
                                            new string(' ', 58) +
                                            $"***** If your camera device supports this, please re-test with the 64bit version of Conform Universal because it has greater memory headroom.");
                                }
                                catch (Exception ex)
                                {
                                    LogIssue("ImageArrayVariant", $"Exception when reading ImageArrayVariant: {ex.Message}");
                                }
                            }

                            // Release memory currently consumed by images
                            SetAction("Releasing memory");
                            ReleaseMemory();

                        }
                        else
                        {
                            LogTestAndMessage("StartExposure", "Test: " + p_ExpectedErrorMessage);
                            LogIssue("StartExposure", "Expected an exception and didn't get one - BinX:" + p_BinX + " BinY:" + p_BinY + " StartX:" + p_StartX + " StartY:" + p_StartY + " NumX:" + p_NumX + " NumY:" + p_NumY);
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to call StopExposure");
                            try
                            {
                                m_Camera.StopExposure();
                            }
                            catch (Exception)
                            {
                            } // Try and do some clean up
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("ConformanceCheck", "About to call AbortExposure");
                            try
                            {
                                m_Camera.AbortExposure();
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (p_ExpectedErrorMessage != "")
                            LogOK("StartExposure", "Exception correctly generated for " + p_ExpectedErrorMessage);
                        else
                            LogIssue("StartExposure", "Exception generated when exposing, further StartExposure tests skipped - " + ex.ToString());
                    }
                }
                catch (Exception ex)
                {
                    LogInfo("StartExposure", "Exception: " + ex.Message);
                    LogIssue("StartExposure", "Exception generated when setting camera properties, further StartExposure tests skipped");
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
                            LogInfo("StartExposure", ex.Message);
                            break;
                        }

                    default:
                        {
                            LogInfo("StartExposure", "Exception: " + ex.Message);
                            LogIssue("StartExposure", "Exception generated when setting camera properties, further StartExposure tests skipped");
                            break;
                        }
                }
            }
            catch (InvalidValueException ex)
            {
                LogInfo("BinXY Write", "Exposure skipped because BinX or BinY cannot be set. Exception message:");
                LogTestAndMessage("BinXY Write", ex.Message);
            }
            catch (Exception ex)
            {
                LogInfo("StartExposure", "Exception: " + ex.Message);
                LogIssue("StartExposure", "Exception generated when setting camera properties, further StartExposure tests skipped");
            }

            SetAction("");
        }
        private void CameraTestLast(double p_Duration, DateTime p_Start)
        {
            DateTime l_StartTime;

            // LastExposureDuration
            try
            {
                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get LastExposureDuration");
                m_LastExposureDuration = m_Camera.LastExposureDuration;
                if ((Math.Abs(m_LastExposureDuration - p_Duration) / p_Duration) < 0.02)
                    LogOK("LastExposureDuration", "LastExposureDuration is: " + m_LastExposureDuration + " seconds");
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
                if (settings.DisplayMethodCalls) LogTestAndMessage("ConformanceCheck", "About to get LastExposureStartTime");
                m_LastExposureStartTime = m_Camera.LastExposureStartTime;
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
            m_Camera.PulseGuide(p_Direction, CAMERA_PULSE_DURATION); // Start a 2 second pulse
            l_EndTime = DateTime.Now;

            if (m_IsPulseGuidingSupported)
            {
                if (l_EndTime.Subtract(l_StartTime).TotalMilliseconds < (CAMERA_PULSE_DURATION - 500))
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding");
                    if (m_Camera.IsPulseGuiding)
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding multiple times");
                        Stopwatch sw = Stopwatch.StartNew();
                        WaitWhile($"Guiding {p_Direction}", () => { return m_Camera.IsPulseGuiding; }, 500, 3, () => { return $"{sw.Elapsed.TotalSeconds:0.0} / {CAMERA_PULSE_DURATION / 1000:0.0} seconds"; });

                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("ConformanceCheck", "About to get IsPulseGuiding");
                        if (!m_Camera.IsPulseGuiding)
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
                    if (!m_Camera.IsPulseGuiding)
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
            m_Camera.BinX = 1;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set BinY");
            m_Camera.BinY = 1;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set StartX");
            m_Camera.StartX = 0;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set StartY");
            m_Camera.StartY = 0;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set MaxBinX");
            m_Camera.NumX = m_Camera.MaxBinX;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set MaxBinY");
            m_Camera.NumY = m_Camera.MaxBinY;
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to call StartExposure");
            m_Camera.StartExposure(1, true); // 1 second exposure

            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to call ImageReady multiple times");
            do
                SetStatus("Waiting for ImageReady");
            while (!m_Camera.ImageReady);
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
                                m_CameraState = m_Camera.CameraState;
                                break;
                            }

                        case CameraPerformance.CCDTemperature:
                            {
                                _ = m_Camera.CCDTemperature;
                                break;
                            }

                        case CameraPerformance.CoolerPower:
                            {
                                _ = m_Camera.CoolerPower;
                                break;
                            }

                        case CameraPerformance.HeatSinkTemperature:
                            {
                                _ = m_Camera.HeatSinkTemperature;
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
                                // Release memory currently consumed by images
                                SetAction("Releasing memory");
                                ReleaseMemory();

                                m_ImageArray = (Array)m_Camera.ImageArray;
                                break;
                            }

                        case CameraPerformance.ImageArrayVariant:
                            {
                                // Release memory currently consumed by images
                                SetAction("Releasing memory");
                                ReleaseMemory();

                                m_ImageArrayVariant = (Array)m_Camera.ImageArrayVariant;
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
                LogInfo(p_Name, "Unable to complete test: " + ex.ToString());
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
                    m_Camera.AbortExposure();
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
                    m_Camera.StopExposure();
                }
                catch
                {
                }
            }
            if (settings.DisplayMethodCalls)
                LogTestAndMessage("ConformanceCheck", "About to set SetCCDTemperature");
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
                LogTestAndMessage("ConformanceCheck", "About to set CoolerOn");

            // Reset the camera image parameters to legal values
            try { m_Camera.StartX = 0; } catch { }
            try { m_Camera.StartY = 0; } catch { }
            try { m_Camera.BinX = 1; } catch { }
            try { m_Camera.BinY = 1; } catch { }
            try { m_Camera.NumX = 1; } catch { }
            try { m_Camera.NumY = 1; } catch { }

            try
            {
                m_Camera.CoolerOn = m_CoolerOn;
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
            // Clear out any previous memory allocations
            m_ImageArray = null;
            m_ImageArrayVariant = null;
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
            GC.Collect(2, GCCollectionMode.Forced, true, true);
        }
    }

}

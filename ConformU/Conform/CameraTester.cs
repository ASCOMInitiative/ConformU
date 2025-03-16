using ASCOM;
using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using Blazorise;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace ConformU
{

    internal class CameraTester : DeviceTesterBaseClass
    {
        #region Constants and variables

        private const int CAMERA_PULSE_DURATION_MILLISECONDS = 2000; // Duration of camera pulse guide test (ms)
        private const int CAMERA_PULSE_TOLERANCE_MILLISECONDS = 300; // Tolerance for acceptable performance (ms)

        private const int MAX_BIN_X = 16; // Values of MaxBin above which warnings are given. Implemented to warn developers if they are returning "silly" values
        private const int MAX_BIN_Y = 16;
        private const double ABSOLUTE_ZERO_TEMPERATURE = -273.15; // Absolute zero (Celsius)
        private const double BELOW_ABSOLUTE_ZERO_TEMPERATURE = -273.25; // Value (Celsius) below which CCD temperatures will be flagged as "silly", "un-physical" values.
        private const double CAMERA_SETPOINT_INCREMENT = 5.0; // Amount by which the test temperature is decremented or incremented when finding the lowest and highest supported set points.
        private const double BOILING_POINT_TEMPERATURE = 100.0; // Value above which CCD set point temperatures will be flagged as "silly" values
        private const double MAX_CAMERA_REPORTED_TEMPERATURE = 1000.0; // Value above which the CCD reported temperature will be flagged as a "silly" value. It is higher than the MAX_CAMERA_SETPOINT_TEMPERATURE temperature because this value is not specified in the Interface Standard.
        private const double CAMERA_LOW_SETPOINT_START_TEMPERATURE = 15.0; // Start temperature for determining minimum set point value.
        private const double CAMERA_HIGH_SETPOINT_START_TEMPERATURE = 0.0; // Start temperature for determining maximum set point value.

        // Camera variables
        private bool mCanAbortExposure, mCanAsymmetricBin, mCanGetCoolerPower, mCanSetCcdTemperature, mCanStopExposure, mCanFastReadout;
        private bool mCoolerOn, mImageReady;
        private int cameraXSize, cameraYSize;
        private short mMaxBinX, mMaxBinY, mBinX, mBinY;
        private double mLastExposureDuration;
        private double mSetCcdTemperature;
        private string mLastExposureStartTime;
        private CameraState mCameraState;
        private Array mImageArray, mImageArrayVariant;
        private bool mIsPulseGuidingFunctional; // Confirm that IsPulseGuiding command will work
        private bool mCanPulseGuide;
        private bool mIsPulseGuiding;
        // ICameraV2 properties
        private short mGain, mGainMax, mGainMin, mPercentCompleted, mReadoutMode;
        private double mExposureMax, mExposureMin, mExposureResolution;
        private bool mFastReadout, mCanReadGain, mCanReadGainMax, mCanReadGainMin, mCanReadGains, mCanReadReadoutModes;
        private IList<string> mGains;
        private IList<string> mReadoutModes;
        private SensorType mSensorType;
        private bool mCanReadSensorType = false;
        private readonly Stopwatch sw = new();

        // ICameraV3 properties
        private int mOffset, mOffsetMax, mOffsetMin;
        private bool mCanReadOffset, mCanReadOffsetMax, mCanReadOffsetMin, mCanReadOffsets;
        private IList<string> mOffsets;
        private double mSubExposureDuration;
        private GainOffsetMode mOffsetMode = GainOffsetMode.Unknown;
        private GainOffsetMode mGainMode = GainOffsetMode.Unknown;

        private ICameraV4 camera;

        private enum CanType
        {
            TstCanFindHome = 1,
            TstCanPark = 2,
            TstCanPulseGuide = 3,
            TstCanSetDeclinationRate = 4,
            TstCanSetGuideRates = 5,
            TstCanSetPark = 6,
            TstCanSetPierSide = 7,
            TstCanSetRightAscensionRate = 8,
            TstCanSetTracking = 9,
            TstCanSlew = 10,
            TstCanSlewAsync = 11,
            TstCanSlewAltAz = 12,
            TstCanSlewAltAzAsync = 13,
            TstCanSync = 14,
            TstCanSyncAltAz = 15,
            TstCanUnPark = 16,
            TstCanAbortExposure = 17,
            TstCanAsymmetricBin = 18,
            TstCanGetCoolerPower = 19,
            TstCanSetCcdTemperature = 20,
            TstCanStopExposure = 21,
            // ICameraV2 property
            TstCanFastReadout = 22
        }
        private enum CameraPerformance : int
        {
            CameraState,
            CcdTemperature,
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
            CcdTemperature = 6,
            CoolerOn = 7,
            CoolerPower = 8,
            Description = 9,
            DriverInfo = 10,
            ElectronsPerAdu = 11,
            FullWellCapacity = 12,
            HasShutter = 13,
            HeatSinkTemperature = 14,
            ImageReady = 15,
            IsPulseGuiding = 16,
            MaxAdu = 17,
            MaxBinX = 18,
            MaxBinY = 19,
            NumX = 20,
            NumY = 21,
            PixelSizeX = 22,
            PixelSizeY = 23,
            SetCcdTemperature = 24,
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

        private readonly Settings settings;
        private readonly CancellationToken cancellationToken;
        private readonly ConformLogger logger;

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
            LogDebug("CameraTester.Dispose", $"Disposing of device: {disposing} {disposedValue}");
            if (!disposedValue)
            {
                if (disposing)
                {

                    Task.Run(() =>
                    {
                        LogDebug("CameraTester.Dispose", "About to dispose of camera...");
                        try { camera?.Dispose(); } catch (Exception ex) { LogDebug("CameraTester.Dispose", $"Exception disposing of device:\r\n{ex}"); }
                        LogDebug("CameraTester.Dispose", "Camera disposed");
                    });

                    //try { camera?.Dispose(); } catch { }
                    LogDebug("CameraTester.Dispose", "About to set camera to null...");

                    camera = null;
                    LogDebug("CameraTester.Dispose", "Camera set to null");
                    LogDebug("CameraTester.Dispose", "About to release memory...");
                    mImageArray = null;
                    mImageArrayVariant = null;

                    //SetTest("Cleaning up");
                    //try { ReleaseMemory(); } catch { }
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

        #region Conform Process

        public override void InitialiseTest()
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
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040405;
                            ExInvalidValue2 = (int)0x80040405;
                            ExNotSet1 = (int)0x80040403;
                            break;
                        }
                }
            }
            LogDebug("CheckInitialise-Camera", $"Set GExNotImplemented");
            base.InitialiseTest();
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
                                throw new InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);

                SetDevice(camera, DeviceTypes.Camera); // Assign the driver to the base class

                LogInfo("CreateDevice", "Successfully created driver");

                // Validate the interface version
                ValidateInterfaceVersion();
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

        public override void ReadCanProperties()
        {
            // ICameraV1 properties
            CameraCanTest(CanType.TstCanAbortExposure, "CanAbortExposure");
            CameraCanTest(CanType.TstCanAsymmetricBin, "CanAsymmetricBin");
            CameraCanTest(CanType.TstCanGetCoolerPower, "CanGetCoolerPower");
            CameraCanTest(CanType.TstCanPulseGuide, "CanPulseGuide");
            CameraCanTest(CanType.TstCanSetCcdTemperature, "CanSetCCDTemperature");
            CameraCanTest(CanType.TstCanStopExposure, "CanStopExposure");

            // ICameraV2 properties
            CameraCanTest(CanType.TstCanFastReadout, "CanFastReadout");
        }

        public override void PreRunCheck()
        {
            int lVStringPtr, lV1, lV2, lV3;
            // Add a test for a back level version of the Camera simulator - just abandon this process if any errors occur
            if (settings.ComDevice.ProgId.Equals("CCDSIMULATOR.CAMERA", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to get Description");
                    lVStringPtr = camera.Description.IndexOf("VERSION ", StringComparison.CurrentCultureIgnoreCase); // Point at the start of the version string
                    if (lVStringPtr >= 0)
                    {
                        LogCallToDriver("ConformanceCheck", "About to get Description");
                        string lVString = camera.Description.ToUpper().Substring(lVStringPtr, 8);
                        lVStringPtr = lVString.IndexOf('.');
                        if (lVStringPtr > 0)
                        {
                            lV1 = Convert.ToInt32(lVString[..(lVStringPtr - 1)]); // Extract the number
                            lVString = lVString[(lVStringPtr + 1)..]; // Get the second version number part
                            lVStringPtr = lVString.IndexOf('.');
                            if (lVStringPtr > 1)
                            {
                                lV2 = Convert.ToInt32(lVString[..(lVStringPtr - 1)]); // Extract the number
                                lVString = lVString[(lVStringPtr + 1)..]; // Get the third version number part
                                lV3 = Convert.ToInt32(lVString); // Extract the number
                                                                 // Turn the version parts into a whole number
                                lV1 = lV1 * 1000000 + lV2 * 1000 + lV3;
                                if (lV1 < 5000008)
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
                        LogCallToDriver("ConformanceCheck", "About to get LastExposureDuration");
                        mLastExposureDuration = camera.LastExposureDuration;
                        LogIssue("LastExposureDuration", "LastExposureDuration did not return an error when called before an exposure was made");
                    }
                    catch (Exception ex)
                    {
                        LogOk("LastExposureDuration", $"LastExposureDuration returned an error before an exposure was made: {ex.Message}");
                    }

                    // Check LastExposureStartTime
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to get LastExposureStartTime");
                        mLastExposureStartTime = camera.LastExposureStartTime;
                        LogIssue("LastExposureStartTime", "LastExposureStartTime did not return an error when called before an exposure was made");
                    }
                    catch (Exception ex)
                    {
                        LogOk("LastExposureStartTime", $"LastExposureStartTime returned an error before an exposure was made: {ex.Message}");
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
            int lBinX, lBinY, lMaxBinX, lMaxBinY;

            // Basic read tests
            mMaxBinX = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.MaxBinX, "MaxBinX", 1, MAX_BIN_X));
            if (cancellationToken.IsCancellationRequested)
                return;

            mMaxBinY = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.MaxBinY, "MaxBinY", 1, MAX_BIN_Y));
            if (cancellationToken.IsCancellationRequested)
                return;

            if (!mCanAsymmetricBin)
            {
                if (mMaxBinX != mMaxBinY)
                    LogIssue("CanAsymmetricBin", "CanAsymmetricBin is false but MaxBinX and MaxBinY are not equal!");
            }

            mBinX = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.BinX, "BinX Read", 1, 1)); // Must default to 1 on start-up
            if (cancellationToken.IsCancellationRequested)
                return;

            mBinY = System.Convert.ToInt16(CameraPropertyTestInteger(CamPropertyType.BinY, "BinY Read", 1, 1)); // Must default to 1 on start-up
            if (cancellationToken.IsCancellationRequested)
                return;

            if (!mCanAsymmetricBin)
            {
                if (mBinX != mBinY)
                    LogIssue("CanAsymmetricBin", "CanAsymmetricBin is false but BinX and BinY are not equal!");
            }

            // Test writing low and high Bin values outside maximum range
            try // Invalid low value
            {
                LogCallToDriver("ConformanceCheck", "About to set BinX");
                camera.BinX = 0;
                LogIssue("BinX Write", "Invalid value 0 written but no error returned");
            }
            catch (Exception ex)
            {
                LogOk("BinX Write", $"Received error on setting BinX to 0: {ex.Message}");
            }

            try // Invalid high value
            {
                LogCallToDriver("ConformanceCheck", "About to set BinX");
                camera.BinX = (short)(mMaxBinX + 1);
                LogIssue("BinX Write", $"Invalid value {mMaxBinX}{1} written but no error returned");
            }
            catch (Exception ex)
            {
                LogOk("BinX Write", $"Received error on setting BinX to {mMaxBinX + 1}: {ex.Message}");
            }

            try // Invalid low value
            {
                LogCallToDriver("ConformanceCheck", "About to set BinY");
                camera.BinY = 0;
                LogIssue("BinY Write", "Invalid value 0 written but no error returned");
            }
            catch (Exception ex)
            {
                LogOk("BinY Write", $"Received error on setting BinY to 0: {ex.Message}");
            }

            try // Invalid high value
            {
                LogCallToDriver("ConformanceCheck", "About to set BinY");
                camera.BinY = (short)(mMaxBinY + 1);
                LogIssue("BinY Write", $"Invalid value {mMaxBinY}{1} written but no error returned");
            }
            catch (Exception ex)
            {
                LogOk("BinY Write", $"Received error on setting BinY to {mMaxBinY + 1}: {ex.Message}");
            }

            // Use either the maximum values supplied by the camera driver or the maximum values defined in Conform's setup, if these have been set
            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogNewLine(); // Insert a blank line if required

            if (settings.CameraMaxBinX > 0)
            {
                lMaxBinX = settings.CameraMaxBinX;
                LogTestAndMessage("BinXY Write", string.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", lMaxBinX, mMaxBinX));
            }
            else
                lMaxBinX = mMaxBinX;

            if (settings.CameraMaxBinY > 0)
            {
                lMaxBinY = settings.CameraMaxBinY;
                LogTestAndMessage("BinXY Write", string.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", lMaxBinY, mMaxBinY));
            }
            else
                lMaxBinY = mMaxBinY;

            if ((settings.CameraMaxBinX > mMaxBinX) | (settings.CameraMaxBinY > mMaxBinY))
                LogNewLine(); // Insert a blank line if required

            if (settings.CameraMaxBinX > mMaxBinX)
                LogTestAndMessage("BinXY Write", string.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", lMaxBinX, mMaxBinX));

            if (settings.CameraMaxBinY > mMaxBinY)
                LogTestAndMessage("BinXY Write", string.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", lMaxBinY, mMaxBinY));

            if ((settings.CameraMaxBinX > 0) | (settings.CameraMaxBinY > 0))
                LogNewLine(); // Insert a blank line if required

            // Write BinX and BinY
            if (mCanAsymmetricBin)
            {
                for (lBinX = 1; lBinX <= lMaxBinX; lBinX++)
                {
                    for (lBinY = 1; lBinY <= lMaxBinY; lBinY++)
                    {
                        bool binXSetOk = false;
                        try
                        {
                            LogCallToDriver("ConformanceCheck", "About to set BinX");
                            TimeMethod("BinX Write", () => camera.BinX = (short)lBinX, TargetTime.Standard);
                            binXSetOk = true;
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex,
                                $"when setting BinX to {lBinX}",
                                $"InvalidValue error received when setting BinX to {lBinX}");
                        }

                        try
                        {
                            LogCallToDriver("ConformanceCheck", "About to set BinY");
                            TimeMethod("BinY Write", () => camera.BinY = (short)lBinY, TargetTime.Standard);

                            if (binXSetOk)
                                LogOk("BinXY Write", $"Successfully set asymmetric XY binning: {lBinX} x {lBinY}");
                            else
                                LogOk("BinXY Write", $"Successfully set Y binning to {lBinY}");
                        }
                        catch (Exception ex)
                        {
                            HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex,
                                $"when setting BinY to {lBinY}",
                                $"InvalidValue error received when setting BinY to {lBinY}");
                        }
                    }
                }
            }
            else // Can only bin symmetrically
                for (lBinX = 1; lBinX <= lMaxBinX; lBinX++)
                {
                    bool binXSetOk = false;

                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to set BinX");
                        TimeMethod("BinX Write", () => camera.BinX = (short)lBinX, TargetTime.Standard);
                        binXSetOk = true;
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsInfo("BinX Write", MemberType.Property, Required.Mandatory, ex,
                            $"when setting BinX to {lBinX}",
                            $"InvalidValueException received when setting BinX to {lBinX}");
                    }

                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to set BinY");
                        TimeMethod("BinY Write", () => camera.BinY = (short)lBinX, TargetTime.Standard);
                        if (binXSetOk)
                            LogOk("BinXY Write", $"Successfully set symmetric XY binning: {lBinX} x {lBinX}");
                        else
                            LogOk("BinXY Write", $"Successfully set Y binning to {lBinX}");
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidValueExceptionAsInfo("BinY Write", MemberType.Property, Required.Mandatory, ex,
                            $"when setting BinY to {lBinX}",
                            $"InvalidValueException received when setting BinY to {lBinX}");
                    }
                }

            // Reset X and Y binning to 1x1 state
            try
            {
                LogCallToDriver("ConformanceCheck", "About to set BinX");
                camera.BinX = 1;
            }
            catch (Exception)
            {
            }
            try
            {
                LogCallToDriver("ConformanceCheck", "About to set BinY");
                camera.BinY = 1;
            }
            catch (Exception)
            {
            }

            mCameraState = CameraPropertyTestCameraState(CamPropertyType.CameraState, "CameraState"); if (cancellationToken.IsCancellationRequested) return;
            cameraXSize = CameraPropertyTestInteger(CamPropertyType.CameraXSize, "CameraXSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;
            cameraYSize = CameraPropertyTestInteger(CamPropertyType.CameraYSize, "CameraYSize", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;

            double ccdTemperature = CameraPropertyTestDouble(CamPropertyType.CcdTemperature, "CCDTemperature", ABSOLUTE_ZERO_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;
            mCoolerOn = CameraPropertyTestBoolean(CamPropertyType.CoolerOn, "CoolerOn Read", false); if (cancellationToken.IsCancellationRequested) return;

            // Write CoolerOn
            bool lOriginalCoolerState;
            string lTargetCoolerState;
            try
            {
                LogCallToDriver("ConformanceCheck", "About to set CoolerOn");
                lOriginalCoolerState = camera.CoolerOn;
                if (lOriginalCoolerState)
                    lTargetCoolerState = "off";
                else
                    lTargetCoolerState = "on";

                try
                {
                    if (lOriginalCoolerState)
                    {
                        LogCallToDriver("ConformanceCheck", "About to set CoolerOn");
                        TimeMethod("CoolerOn", () => camera.CoolerOn = false, TargetTime.Standard);
                    }
                    else
                    {
                        LogCallToDriver("ConformanceCheck", "About to set CoolerOn");
                        TimeMethod("CoolerOn", () => camera.CoolerOn = true, TargetTime.Standard);
                    }
                    LogOk("CoolerOn Write", "Successfully changed CoolerOn state");
                }
                catch (Exception ex)
                {
                    HandleException("CoolerOn Write", MemberType.Property, Required.Optional, ex,
                        $"turning Cooler {lTargetCoolerState}");
                }

                // Restore Cooler state
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to set CoolerOn");
                    camera.CoolerOn = lOriginalCoolerState;
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
            CameraPropertyTestDouble(CamPropertyType.ElectronsPerAdu, "ElectronsPerADU", 0.00001, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.FullWellCapacity, "FullWellCapacity", 0.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestBoolean(CamPropertyType.HasShutter, "HasShutter", false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.HeatSinkTemperature, "HeatSinkTemperature", ABSOLUTE_ZERO_TEMPERATURE, MAX_CAMERA_REPORTED_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;

            mImageReady = CameraPropertyTestBoolean(CamPropertyType.ImageReady, "ImageReady", false); if (cancellationToken.IsCancellationRequested) return;
            if (mImageReady & settings.CameraFirstUseTests) // Issue this warning if configured to do so
                LogIssue("ImageReady", "Image is flagged as ready but no exposure has been started!");

            // Release memory currently consumed by images
            ReleaseMemory();

            // ImageArray 
            SetFullStatus("ImageArray", "Getting image data from device...", "");

            if (mImageReady) // ImageReady is true
            {
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to get ImageArray");
                    TimeMethod("ImageArray", () => mImageArray = (Array)camera.ImageArray, TargetTime.Extended);
                    if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                    {
                        LogIssue("ImageArray", "No image has been taken but ImageArray has not returned an error");
                    }
                    else  // Omit first use tests
                    {
                        LogOk("ImageArray", "ImageArray read OK");
                    }
                }
                catch (Exception ex)
                {
                    LogDebug("ImageArray", $"Exception 1:\r\n{ex}");
                    if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                    {
                        LogOk("ImageArray", $"Received error when camera is idle: {ex.Message}");
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
                    LogCallToDriver("ConformanceCheck", "About to get ImageArray");
                    TimeMethod("ImageArray", () => mImageArray = (Array)camera.ImageArray, TargetTime.Extended);
                    LogIssue("ImageArray", "ImageReady is false and no image has been taken but ImageArray has not returned an error");
                }
                catch (Exception ex)
                {
                    LogOk("ImageArray", $"Received error when ImageReady is false: {ex.Message}");
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
                if (mImageReady)
                {
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to get ImageArrayVariant");
                        TimeMethod("ImageArrayVariant", () => mImageArrayVariant = (Array)camera.ImageArrayVariant, TargetTime.Extended);

                        if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                        {
                            LogIssue("ImageArrayVariant", "No image has been taken but ImageArrayVariant has not returned an error");
                        }
                        else
                        {
                            LogOk("ImageArrayVariant", "ImageArray read OK");
                        }
                    }
                    catch (Exception ex)
                    {
                        if (settings.CameraFirstUseTests) // Only perform this test if configured to do so
                        {
                            LogOk("ImageArrayVariant", $"Received error before an image has been taken: {ex.Message}");
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
                        LogCallToDriver("ConformanceCheck", "About to get ImageArrayVariant");
                        TimeMethod("ImageArrayVariant", () => mImageArrayVariant = (Array)camera.ImageArrayVariant, TargetTime.Extended);
                        LogIssue("ImageArrayVariant", "ImageReady is false and no image has been taken but ImageArray has not returned an error");
                    }
                    catch (Exception ex)
                    {
                        LogOk("ImageArrayVariant", $"Received error when ImageReady is false: {ex.Message}");
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

            mIsPulseGuiding = CameraPropertyTestBoolean(CamPropertyType.IsPulseGuiding, "IsPulseGuiding", false); if (cancellationToken.IsCancellationRequested) return;
            if (mIsPulseGuiding)
                LogIssue("IsPulseGuiding", "Camera is showing pulse guiding underway although no PulseGuide command has been issued!");

            CameraPropertyTestInteger(CamPropertyType.MaxAdu, "MaxADU", 1, int.MaxValue); if (cancellationToken.IsCancellationRequested) return;

            CameraPropertyTestInteger(CamPropertyType.NumX, "NumX Read", 1, cameraXSize); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.NumX, "NumX", System.Convert.ToInt32(cameraXSize / (double)2));

            CameraPropertyTestInteger(CamPropertyType.NumY, "NumY Read", 1, cameraYSize); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyWriteTest(CamPropertyType.NumY, "NumY", System.Convert.ToInt32(cameraYSize / (double)2));

            CameraPropertyTestDouble(CamPropertyType.PixelSizeX, "PixelSizeX", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;
            CameraPropertyTestDouble(CamPropertyType.PixelSizeY, "PixelSizeY", 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested) return;

            mSetCcdTemperature = CameraPropertyTestDouble(CamPropertyType.SetCcdTemperature, "SetCCDTemperature Read", ABSOLUTE_ZERO_TEMPERATURE, BOILING_POINT_TEMPERATURE, false); if (cancellationToken.IsCancellationRequested) return;

            if (mCanSetCcdTemperature)
            {
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to set SetCCDTemperature");
                    TimeMethod("SetCCDTemperature Write", () => camera.SetCCDTemperature = 0.0, TargetTime.Standard); // Try an acceptable value
                    LogOk("SetCCDTemperature Write", "Successfully wrote 0.0");

                    // Execution only gets here if the CCD temperature can be set successfully
                    bool exceptionGenerated;
                    double setPoint;

                    // Find low set-point at which an exception is generated, stop at CAMERA_SETPOINT_UNPHYSICAL_TEMPERATURE because this is unphysical
                    exceptionGenerated = false;
                    setPoint = CAMERA_LOW_SETPOINT_START_TEMPERATURE;

                    // Loop downward in CAMERA_SETPOINT_INCREMENT degree temperature steps to find the maximum temperature that can be set
                    LogCallToDriver("ConformanceCheck", "About to set SetCCDTemperature multiple times...");
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
                    LogCallToDriver("ConformanceCheck", "About to set SetCCDTemperature multiple times...");
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
                LogCallToDriver("ConformanceCheck", "About to set SetCCDTemperature");
                try
                {
                    camera.SetCCDTemperature = mSetCcdTemperature;
                }
                catch
                {
                }
            }
            else
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to set SetCCDTemperature");
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
            LogCallToDriver("ConformanceCheck", "About to get InterfaceVersion");
            if (camera.InterfaceVersion > 1)
            {
                // SensorType - Optional
                // This must be tested before BayerOffset because BayerOffset is mandatory for colour and optional for monochrome cameras
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to get SensorType");
                    TimeFunc("SensorType", () => mSensorType = (SensorType)camera.SensorType, TargetTime.Fast);
                    mCanReadSensorType = true; // Set a flag to indicate that we have got a valid SensorType value

                    // Successfully retrieved a value
                    LogOk("SensorType Read", mSensorType.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("SensorType Read", MemberType.Property, Required.Optional, ex, "");
                }

                // BayerOffset Read
                if (mCanReadSensorType)
                {
                    if (mSensorType == SensorType.Monochrome)
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
                mExposureMax = CameraPropertyTestDouble(CamPropertyType.ExposureMax, "ExposureMax Read", 0.0001, double.MaxValue, true);
                mExposureMin = CameraPropertyTestDouble(CamPropertyType.ExposureMin, "ExposureMin Read", 0.0, double.MaxValue, true);
                if (mExposureMin <= mExposureMax)
                    LogOk("ExposureMin", "ExposureMin is less than or equal to ExposureMax");
                else
                    LogIssue("ExposureMin", "ExposureMin is greater than ExposureMax");

                // ExposureResolution Read
                mExposureResolution = CameraPropertyTestDouble(CamPropertyType.ExposureResolution, "ExposureResolution Read", 0.0, double.MaxValue, true);
                if (mExposureResolution <= mExposureMax)
                    LogOk("ExposureResolution", "ExposureResolution is less than or equal to ExposureMax");
                else
                    LogIssue("ExposureResolution", "ExposureResolution is greater than ExposureMax");

                // FastReadout Read Optional
                if (mCanFastReadout)
                    mFastReadout = CameraPropertyTestBoolean(CamPropertyType.FastReadout, "FastReadout Read", true);
                else
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to get FastReadout");
                        mFastReadout = camera.FastReadout;
                        LogIssue("FastReadout Read", "CanFastReadout is False but a PropertyNotImplemented error was not returned.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("FastReadout Read", MemberType.Property, Required.Optional, ex, "");
                    }

                // FastReadout Write Optional
                if (mCanFastReadout)
                {
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to set FastReadout");
                        TimeMethod("FastReadout Write", () => camera.FastReadout = !mFastReadout, TargetTime.Standard);

                        LogCallToDriver("ConformanceCheck", "About to set FastReadout");
                        camera.FastReadout = mFastReadout;

                        LogOk("FastReadout Write", "Able to change the FastReadout state OK");
                    }
                    catch (Exception ex)
                    {
                        HandleException("FastReadout Write", MemberType.Property, Required.Mandatory, ex, "");
                    }
                }
                else
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to set FastReadout");
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
                    mCanReadGainMin = false;
                    LogCallToDriver("ConformanceCheck", "About to get GainMin");
                    TimeFunc<short>("GainMin", () => mGainMin = camera.GainMin, TargetTime.Standard);

                    // Successfully retrieved a value
                    mCanReadGainMin = true;
                    LogOk("GainMin Read", mGainMin.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOk("GainMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // GainMax Read - Optional
                try
                {
                    mCanReadGainMax = false;
                    LogCallToDriver("ConformanceCheck", "About to get GainMax");
                    TimeFunc("GainMax", () => mGainMax = camera.GainMax, TargetTime.Standard);

                    // Successfully retrieved a value
                    mCanReadGainMax = true;
                    LogOk("GainMax Read", mGainMax.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOk("GainMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // Gains Read - Optional
                try
                {
                    mCanReadGains = false;
                    LogCallToDriver("ConformanceCheck", "About to get Gains");
                    TimeFunc("GainMax", () => mGains = camera.Gains, TargetTime.Fast);

                    // Successfully retrieved a value
                    mCanReadGains = true;
                    foreach (string gain in mGains)
                        LogOk("Gains Read", gain.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOk("Gains Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                if (mCanReadGainMax & mCanReadGainMin & mCanReadGains)
                    LogIssue("Gains", "GainMin, GainMax and Gains are all readable. Only one of GainMin/Max as a pair or Gains should be used, the other should throw a PropertyNotImplemented error");
                else
                {
                }

                // Gain Read - Optional 
                try
                {
                    mCanReadGain = false; // Set default value to indicate can't read gain
                    LogCallToDriver("ConformanceCheck", "About to get Gain");
                    TimeFunc("GainMax", () => mGain = camera.Gain, TargetTime.Fast);

                    mCanReadGain = true; // Flag that we can read Gain OK
                    if (mCanReadGains)
                        LogOk("Gain Read", $"{mGain} {mGains[0]}");
                    else
                        LogOk("Gain Read", mGain.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("Gain Read", MemberType.Property, Required.Optional, ex, "");
                }

                // Now check that gain property groups are implemented to handle the three gain modes: NotImplemented, Gain Index (Gain + Gains) and Gain Value (Gain + GainMin + GainMax)
                if (!mCanReadGain & !mCanReadGains & !mCanReadGainMin & !mCanReadGainMax)
                    LogOk("Gain Read", "All four gain properties return errors - the driver is in \"Gain Not Implemented\" mode.");
                else if (mCanReadGain)
                {
                    // Test for Gain Index mode
                    if ((mCanReadGain & mCanReadGains & !mCanReadGainMin & !mCanReadGainMax))
                    {
                        mGainMode = GainOffsetMode.IndexMode;
                        LogOk("Gain Read", "Gain and Gains can be read while GainMin and GainMax returned errors - the driver is in \"Gain Index\" mode.");
                    }
                    else if ((mCanReadGain & !mCanReadGains & mCanReadGainMin & mCanReadGainMax))
                    {
                        mGainMode = GainOffsetMode.ValueMode;
                        LogOk("Gain Read", "Gain, GainMin and GainMax can be read OK while Gains returns an error - the driver is in \"Gain Value\" mode.");
                    }
                    else
                    {
                        LogIssue("Gain Read", $"Unable to determine whether the driver is in \"Gain Not Implemented\", \"Gain Index\" or \"Gain Value\" mode. Please check the interface specification.");
                        LogInfo("Gain Read", $"Gain returned an error: {mCanReadGain}, Gains returned an error: {mCanReadGains}, GainMin returned an error: {mCanReadGainMin}, GainMax returned an error: {mCanReadGainMax}.");
                        LogInfo("Gain Read", $"\"Gain Not Implemented\" mode: Gain, Gains, GainMin and GainMax must all return errors.");
                        LogInfo("Gain Read", $"\"Gain Index\" mode: Gain and Gains must work while GainMin and GainMax must return errors.");
                        LogInfo("Gain Read", $"\"Gain Value\" mode: Gain, GainMin and GainMax must work while Gains must return an error.");
                    }
                }
                else
                {
                    LogIssue("Gain Read", $"Gain Read returned an error but at least one of Gains, GainMin Or GainMax did not return an error. If Gain returns an error, all the other gain properties should do likewise.");
                    LogInfo("Gain Read", $"Gains returned an error : {mCanReadGains}, GainMin returned an error : {mCanReadGainMin}, GainMax returned an error : {mCanReadGainMax}.");
                }

                // Gain write - Optional when neither gain index nor gain value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither gain mode is supported
                if (!mCanReadGain & !mCanReadGains & !mCanReadGainMin & !mCanReadGainMax)
                {
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to set Gain");
                        TimeMethod("Gain Write", () => camera.Gain = 0, TargetTime.Standard);
                        LogIssue("Gain Write", "Writing to Gain did not return a PropertyNotImplemented error whereas this was the case for reading Gain.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("Gain Write", MemberType.Property, Required.MustNotBeImplemented, ex, "PropertyNotImplemented error is expected");
                    }
                }
                else
                {
                    switch (mGainMode)
                    {
                        case GainOffsetMode.Unknown:
                            LogIssue("Gain Write", "Cannot test Gain Write because of issues with other gain properties - skipping test");
                            break;

                        case GainOffsetMode.IndexMode:
                            break;

                        case GainOffsetMode.ValueMode:
                            // Test writing the minimum valid value
                            try
                            {
                                LogCallToDriver("ConformanceCheck", "About to set Gain");
                                TimeMethod("Gain Write", () => camera.Gain = mGainMin, TargetTime.Standard);
                                LogOk("Gain Write", $"Successfully set gain minimum value {mGainMin}.");
                            }
                            catch (Exception ex)
                            {
                                HandleException("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "when writing a legal value");
                            }

                            // Test writing the maximum valid value
                            try
                            {
                                LogCallToDriver("ConformanceCheck", "About to set Gain");
                                camera.Gain = mGainMax;
                                LogOk("Gain Write", $"Successfully set gain maximum value {mGainMax}.");
                            }
                            catch (Exception ex)
                            {
                                HandleException("Gain Write", MemberType.Property, Required.MustNotBeImplemented, ex, "when writing a legal value");
                            }

                            // Test writing a lower than minimum value - this should result in am invalid value exception
                            try
                            {
                                LogCallToDriver("ConformanceCheck", "About to set Gain");
                                camera.Gain = (short)(mGainMin - 1);
                                LogIssue("Gain Write", $"Successfully set an gain below the minimum value ({mGainMin - 1}), this should have resulted in an InvalidValue error.");
                            }
                            catch (Exception ex)
                            {
                                HandleInvalidValueExceptionAsOk("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for gain {mGainMin - 1}, which is lower than the minimum value.");
                            }

                            // Test writing a lower than minimum value - this should result in am invalid value exception
                            try
                            {
                                LogCallToDriver("ConformanceCheck", "About to set Gain");
                                camera.Gain = (short)(mGainMax + 1);
                                LogIssue("Gain Write", $"Successfully set a gain above the maximum value({mGainMax + 1}), this should have resulted in an InvalidValue error.");
                            }
                            catch (Exception ex)
                            {
                                HandleInvalidValueExceptionAsOk("Gain Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for gain {mGainMax + 1} which is higher than the maximum value.");
                            }

                            break;

                        default:
                            LogIssue("Gain Write", $"UNEXPECTED VALUE FOR OFFSETMODE: {mGainMode}");
                            break;
                    }
                }
                if (cancellationToken.IsCancellationRequested) return;

                // PercentCompleted Read - Optional - corrected to match the specification
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to get PercentCompleted");
                    TimeMethod("PercentCompleted", () => mPercentCompleted = camera.PercentCompleted, TargetTime.Fast);
                    switch (mPercentCompleted)
                    {
                        case object _ when mPercentCompleted < 0 // Lower than minimum value
                       :
                            {
                                LogIssue("PercentCompleted Read", $"Invalid value: {mPercentCompleted}");
                                break;
                            }

                        case object _ when mPercentCompleted > 100 // Higher than maximum value
                 :
                            {
                                LogIssue("PercentCompleted Read", $"Invalid value: {mPercentCompleted}");
                                break;
                            }

                        default:
                            {
                                LogOk("PercentCompleted Read", mPercentCompleted.ToString());
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("PercentCompleted Read", MemberType.Property, Required.Optional, ex, "", "Operation is invalid when camera is not imaging or downloading");
                }

                // ReadoutModes - Mandatory
                try
                {
                    mCanReadReadoutModes = false;
                    LogCallToDriver("ConformanceCheck", "About to get ReadoutModes");
                    TimeMethod("ReadoutModes", () => mReadoutModes = camera.ReadoutModes, TargetTime.Fast);

                    // Successfully retrieved a value
                    mCanReadReadoutModes = true;
                    foreach (string readoutMode in mReadoutModes)
                        LogOk("ReadoutModes Read", readoutMode.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("ReadoutModes Read", MemberType.Property, Required.Mandatory, ex, "");
                }
                if (cancellationToken.IsCancellationRequested) return;

                // ReadoutMode Read - Mandatory
                mReadoutMode = CameraPropertyTestShort(CamPropertyType.ReadoutMode, "ReadoutMode Read", 0, short.MaxValue, true);
                if (mCanReadReadoutModes)
                {
                    try
                    {
                        if (mReadoutMode < mReadoutModes.Count)
                        {
                            LogOk("ReadoutMode Index", "ReadReadoutMode is within the bounds of the ReadoutModes ArrayList");
                            LogInfo("ReadoutMode Index", $"Current value: {mReadoutModes[mReadoutMode]}");
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

            LogCallToDriver("ConformanceCheck", "About to get InterfaceVersion");
            if (camera.InterfaceVersion > 2)
            {
                // OffsetMin Read - Optional
                try
                {
                    mCanReadOffsetMin = false;
                    LogCallToDriver("ConformanceCheck", "About to get OffsetMin");
                    TimeMethod("OffsetMin", () => mOffsetMin = camera.OffsetMin, TargetTime.Fast);

                    // Successfully retrieved a value
                    mCanReadOffsetMin = true;
                    LogOk("OffsetMin Read", mOffsetMin.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOk("OffsetMin Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // OffsetMax Read - Optional
                try
                {
                    mCanReadOffsetMax = false;
                    LogCallToDriver("ConformanceCheck", "About to get OffsetMax");
                    TimeMethod("OffsetMax", () => mOffsetMax = camera.OffsetMax, TargetTime.Fast);

                    // Successfully retrieved a value
                    mCanReadOffsetMax = true;
                    LogOk("OffsetMax Read", mOffsetMax.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOk("OffsetMax Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                // Offsets Read - Optional
                try
                {
                    mCanReadOffsets = false;
                    LogCallToDriver("ConformanceCheck", "About to get Offsets");
                    TimeMethod("Offsets", () => mOffsets = camera.Offsets, TargetTime.Fast);

                    // Successfully retrieved a value
                    mCanReadOffsets = true;
                    foreach (string offset in mOffsets)
                        LogOk("Offsets Read", offset.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidOperationExceptionAsOk("Offsets Read", MemberType.Property, Required.Optional, ex, "", "InvalidOperation error returned");
                }

                if (mCanReadOffsetMax & mCanReadOffsetMin & mCanReadOffsets)
                    LogIssue("Offsets", "OffsetMin, OffsetMax and Offsets are all readable. Only one of OffsetMin/Max as a pair or Offsets should be used, the other should throw a PropertyNotImplemented error");
                else
                {
                }
                if (cancellationToken.IsCancellationRequested) return;

                // Offset Read - Optional 
                try
                {
                    mCanReadOffset = false; // Set default value to indicate can't read offset
                    LogCallToDriver("ConformanceCheck", "About to get Offset");
                    TimeMethod("Offset", () => mOffset = camera.Offset, TargetTime.Fast);

                    mCanReadOffset = true; // Flag that we can read Offset OK
                    if (mCanReadOffsets)
                        LogOk("Offset Read", $"{mOffset} {mOffsets[0]}");
                    else
                        LogOk("Offset Read", mOffset.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("Offset Read", MemberType.Property, Required.Optional, ex, "");
                }

                // Now check that offset property groups are implemented to handle the three offset modes: NotImplemented, Offset Index (Offset + Offsets) and Offset Value (Offset + OffsetMin + OffsetMax)
                if (!mCanReadOffset & !mCanReadOffsets & !mCanReadOffsetMin & !mCanReadOffsetMax)
                    LogOk("Offset Read", "All four offset properties return errors - the driver is in \"Offset Not Implemented\" mode.");
                else if (mCanReadOffset)
                {
                    // Test for Offset Index mode
                    if ((mCanReadOffset & mCanReadOffsets & !mCanReadOffsetMin & !mCanReadOffsetMax))
                    {
                        mOffsetMode = GainOffsetMode.IndexMode;
                        LogOk("Offset Read", "Offset and Offsets can be read while OffsetMin and OffsetMax return errors - the driver is in \"Offset Index\" mode.");
                    }
                    else if ((mCanReadOffset & !mCanReadOffsets & mCanReadOffsetMin & mCanReadOffsetMax))
                    {
                        mOffsetMode = GainOffsetMode.ValueMode;
                        LogOk("Offset Read", "Offset, OffsetMin and OffsetMax can be read OK while Offsets returns an error - the driver is in \"Offset Value\" mode.");
                    }
                    else
                    {
                        mOffsetMode = GainOffsetMode.Unknown;
                        LogIssue("Offset Read", $"Unable to determine whether the driver is in \"Offset Not Implemented\", \"Offset Index\" or \"Offset Value\" mode. Please check the interface specification.");
                        LogInfo("Offset Read", $"Offset returned an error: {mCanReadOffset}, Offsets returned an error: {mCanReadOffsets}, OffsetMin returned an error: {mCanReadOffsetMin}, OffsetMax returned an error: {mCanReadOffsetMax}.");
                        LogInfo("Offset Read", $"\"Offset Not Implemented\" mode: Offset, Offsets, OffsetMin and OffsetMax must all return errors.");
                        LogInfo("Offset Read", $"\"Offset Index\" mode: Offset and Offsets must work while OffsetMin and OffsetMax must return errors.");
                        LogInfo("Offset Read", $"\"Offset Value\" mode: Offset, OffsetMin and OffsetMax must work while Offsets must throw return an error.");
                    }
                }
                else
                {
                    LogIssue("Offset Read", $"Offset Read returned an error but at least one of Offsets, OffsetMin Or OffsetMax did not return an error. If Offset returns an error, all the other offset properties must do likewise.");
                    LogInfo("Offset Read", $"Offsets returned an error : {mCanReadOffsets}, OffsetMin returned an error : {mCanReadOffsetMin}, OffsetMax returned an error : {mCanReadOffsetMax}.");
                }

                // Offset write - Optional when neither offset index nor offset value mode is supported; must be implemented if either mode is supported

                // First test for the only valid not implemented state when neither offset mode is supported
                if (!mCanReadOffset & !mCanReadOffsets & !mCanReadOffsetMin & !mCanReadOffsetMax)
                {
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to set Offset");
                        camera.Offset = 0;
                        LogIssue("Offset Write", "Writing to Offset did not throw a PropertyNotImplemented error when reading Offset did.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("Offset Write", MemberType.Property, Required.MustNotBeImplemented, ex, "PropertyNotImplemented error is expected");
                    }
                }
                else
                    switch (mOffsetMode)
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
                                    LogCallToDriver("ConformanceCheck", "About to set Offset");
                                    TimeMethod("Offset Write", () => camera.Offset = mOffsetMin, TargetTime.Standard);
                                    LogOk("Offset Write", $"Successfully set offset minimum value {mOffsetMin}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing the maximum valid value
                                try
                                {
                                    LogCallToDriver("ConformanceCheck", "About to set Offset");
                                    camera.Offset = mOffsetMax;
                                    LogOk("Offset Write", $"Successfully set offset maximum value {mOffsetMax}.");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("Offset Write", MemberType.Property, Required.MustNotBeImplemented, ex, "when writing a legal value");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    LogCallToDriver("ConformanceCheck", "About to set Offset");
                                    camera.Offset = mOffsetMin - 1;
                                    LogIssue("Offset Write", $"Successfully set an offset below the minimum value ({mOffsetMin - 1}), this should have resulted in an InvalidValue error.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for offset {mOffsetMin - 1}, which is lower than the minimum value.");
                                }

                                // Test writing a lower than minimum value - this should result in am invalid value exception
                                try
                                {
                                    LogCallToDriver("ConformanceCheck", "About to set Offset");
                                    camera.Offset = mOffsetMax + 1;
                                    LogIssue("Offset Write", $"Successfully set an offset above the maximum value({mOffsetMax + 1}), this should have resulted in an InvalidValueerror.");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk("Offset Write", MemberType.Property, Required.MustBeImplemented, ex, "an InvalidValue error is expected.", $"InvalidValue Received error for offset {mOffsetMax + 1} which is higher than the maximum value.");
                                }
                                break;
                            }

                        default:
                            {
                                LogIssue("Offset Write", $"UNEXPECTED VALUE FOR OFFSETMODE: {mOffsetMode}");
                                break;
                            }
                    }

                // SubExposureDuration Read - Optional 
                mSubExposureDuration = CameraPropertyTestDouble(CamPropertyType.SubExposureDuration, "SubExposureDuration", double.Epsilon, double.MaxValue, false);
                if (cancellationToken.IsCancellationRequested)
                    return;

                // SubExposureDuration Write - Optional 
                try
                {
                    LogCallToDriver("ConformanceCheck", "About to set SubExposureDuration");
                    TimeMethod("SubExposureDuration Write", () => camera.SubExposureDuration = mSubExposureDuration, TargetTime.Standard);
                    LogOk("SubExposureDuration write", $"Successfully wrote {mSubExposureDuration}");
                }
                catch (Exception ex)
                {
                    HandleException("SubExposureDuration write", MemberType.Property, Required.Optional, ex, "");
                }
            }
        }

        public override void CheckMethods()
        {
            int i, j, maxBinX, maxBinY;
            // AbortExposure - Mandatory
            SetTest("AbortExposure");
            try
            {
                LogCallToDriver("ConformanceCheck", "About to get CameraState");
                mCameraState = camera.CameraState;

                // Test whether the camera is idle, which it should be in a well behaved device
                if (mCameraState != CameraState.Idle)
                {
                    LogIssue("AbortExposure", $"The camera should be idle but is not: {mCameraState}");
                }

                try
                {
                    LogCallToDriver("ConformanceCheck", "About to call AbortExposure");
                    TimeMethod("AbortExposure", camera.AbortExposure, TargetTime.Standard);
                    if (mCanAbortExposure)
                        LogOk("AbortExposure", "No error returned when camera is already idle");
                    else
                        LogIssue("AbortExposure", "CanAbortExposure is false but no error is returned when AbortExposure is called");
                }
                catch (Exception ex)
                {
                    if (mCanAbortExposure)
                    {
                        LogIssue("AbortExposure", $"Received error when camera is idle: {ex.Message}");
                        LogDebug("AbortExposure", $"Exception detail: {ex}");
                    }
                    else
                        LogOk("AbortExposure", "CanAbortExposure is false and an error was returned");
                }
            }
            catch (Exception ex)
            {
                LogIssue("AbortExposure", $"Received error when reading camera state: {ex.Message}");
                LogDebug("AbortExposure", $"Exception detail: {ex}");
            }

            // PulseGuide
            SetTest("PulseGuide");
            if (mCanPulseGuide) // Can pulse guide
            {
                try
                {
                    CameraPulseGuideTest(GuideDirection.North);
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    CameraPulseGuideTest(GuideDirection.South);
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    CameraPulseGuideTest(GuideDirection.East);
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    CameraPulseGuideTest(GuideDirection.West);
                    if (cancellationToken.IsCancellationRequested)
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
                    LogCallToDriver("ConformanceCheck", "About to call PulseGuide - North");
                    camera.PulseGuide(GuideDirection.North, 0);
                    LogIssue("PulseGuide", "CanPulseGuide is false but no error was returned when calling the method");
                }
                catch (Exception ex)
                {
                    if (IsMethodNotImplementedException(ex))
                    {
                        LogOk("PulseGuide", "CanPulseGuide is false and PulseGuide is not implemented in this driver");
                    }
                    else
                    {
                        LogOk("PulseGuide", $"CanPulseGuide is false and an error was returned when calling the method: {ex.Message}");
                    }
                }

            // StopExposure
            SetTest("StopExposure");
            try
            {
                LogCallToDriver("ConformanceCheck", "About to get CameraState");
                mCameraState = camera.CameraState;

                // Test whether the camera is idle, which it should be in a well behaved device
                if (mCameraState != CameraState.Idle)
                {
                    LogIssue("StopExposure", $"The camera should be idle but is not: {mCameraState}");
                }

                try
                {
                    LogCallToDriver("ConformanceCheck", "About to call StopExposure");
                    SetAction("Calling StopExposure()");
                    TimeMethod("StopExposure", camera.StopExposure, TargetTime.Standard);

                    LogCallToDriver("ConformanceCheck", "About to get ImageReady repeatedly...");
                    WaitWhile($"Waiting for exposure to stop", () => camera.ImageReady, 100, settings.CameraWaitTimeout);

                    if (mCanStopExposure)
                        LogOk("StopExposure", "No error returned when camera is already idle");
                    else
                        LogIssue("StopExposure", "CanStopExposure is false but no error is returned when StopExposure is called");
                }
                catch (Exception ex)
                {
                    if (mCanStopExposure)
                        LogIssue("StopExposure", $"Received error when the camera is idle: {ex.Message}");
                    else
                    {
                        LogOk("StopExposure", $"CanStopExposure is false and an error was returned: {ex.Message}");
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
                LogTestAndMessage("StartExposure", string.Format("Test range set to MaxBinX = {0} by Conform configuration, camera MaxBinX = {1}", maxBinX, mMaxBinX));
            }
            else
                maxBinX = mMaxBinX;
            if (settings.CameraMaxBinY > 0)
            {
                maxBinY = settings.CameraMaxBinY;
                LogTestAndMessage("StartExposure", string.Format("Test range set to MaxBinY = {0} by Conform configuration, camera MaxBinY = {1}", maxBinY, mMaxBinY));
            }
            else
                maxBinY = mMaxBinY;

            if ((settings.CameraMaxBinX > mMaxBinX) | (settings.CameraMaxBinY > mMaxBinY))
                LogNewLine(); // Insert a blank line if required
            if (settings.CameraMaxBinX > mMaxBinX)
                LogTestAndMessage("StartExposure", string.Format("WARNING - Conform's configured MaxBinX: {0} is greater than the camera's reported MaxBinX: {1}!", maxBinX, mMaxBinX));
            if (settings.CameraMaxBinY > mMaxBinY)
                LogTestAndMessage("StartExposure", string.Format("WARNING - Conform's configured MaxBinY: {0} is greater than the camera's reported MaxBinY: {1}!", maxBinY, mMaxBinY));

            // StartExposure - Confirm that correct operation occurs
            int binX, binY;
            if (mCanAsymmetricBin)
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

                        CameraExposure("StartExposure",
                            $"Taking {(constrainedXSize | constrainedYSize ? "configured size" : "full frame")} image {binX} x {binY} bin ({xSize}{(constrainedXSize ? "*" : "")} x {ySize}{(constrainedYSize ? "*" : "")}) - {Convert.ToDouble(xSize * ySize) / 1000000.0:0.0} MPix",
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

                    CameraExposure("StartExposure",
                        $"Taking {(constrainedXSize | constrainedYSize ? "configured size" : "full frame")} image {binX} x {binX} bin ({xSize}{(constrainedXSize ? "*" : "")} x {ySize}{(constrainedYSize ? "*" : "")}) - {Convert.ToDouble(xSize * ySize) / 1000000.0:0.0} MPix",
                        binX, binX, 0, 0, xSize, ySize, settings.CameraExposureDuration, "");

                    if (cancellationToken.IsCancellationRequested)
                        return;
                }

            // StartExposure - Confirm error cases
            LogNewLine();
            LogTestOnly("StartExposure error cases");

            // StartExposure - Negative time
            CameraExposure("Reject Negative Duration", "Testing rejection of negative exposure duration (-1.0)", 1, 1, 0, 0, cameraXSize, cameraYSize, -1.0, "negative duration");
            if (cancellationToken.IsCancellationRequested)
                return;

            // StartExposure - Invalid Bin values
            for (i = 1; i <= maxBinX; i++)
            {
                for (j = 1; j <= maxBinY; j++)
                {
                    if (mCanAsymmetricBin)
                    {
                        // X size too large for binned size
                        CameraExposure($"Reject Bad XSize (bin {i} x {j})", $"Testing rejection of bad X size value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, 0, 0, Convert.ToInt32((cameraXSize / (double)i) + 1), Convert.ToInt32(cameraYSize / (double)j), 0.1,
                            $"X size larger than binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y size too large for binned size
                        CameraExposure($"Reject Bad YSize (bin {i} x {j})", $"Testing rejection of bad Y size value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, 0, Convert.ToInt32(cameraXSize / (double)i), Convert.ToInt32((cameraYSize / (double)j) + 1), 0.1,
                            $"Y size larger than binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // X start outside binned chip dimensions
                        CameraExposure($"Reject Bad XStart (bin {i} x {j})", $"Testing rejection of bad X start value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, Convert.ToInt32((cameraXSize / (double)i) + 1), 0, Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1,
                            $"X start outside binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y start outside binned chip dimensions
                        CameraExposure($"Reject Bad YStart (bin {i} x {j})", $"Testing rejection of bad Y start value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, Convert.ToInt32((cameraYSize / (double)j) + 1), Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1,
                            $"Y start outside binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                    else if (i == j)
                    {
                        // X size too large for binned size
                        CameraExposure($"Reject Bad XSize (bin {i} x {j})", $"Testing rejection of bad X size value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, 0, 0, Convert.ToInt32((cameraXSize / (double)i) + 1), Convert.ToInt32(cameraYSize / (double)j), 0.1,
                            $"X size larger than binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y size too large for binned size
                        CameraExposure($"Reject Bad YSize (bin {i} x {j})", $"Testing rejection of bad Y size value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, 0, Convert.ToInt32(cameraXSize / (double)i), Convert.ToInt32((cameraYSize / (double)j) + 1), 0.1,
                            $"Y size larger than binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // X start outside binned chip dimensions
                        CameraExposure($"Reject Bad XStart (bin {i} x {j})", $"Testing rejection of bad X start value for bin {i} x {j} ({Convert.ToInt32(cameraXSize / (double)i) + 1})", i, j, Convert.ToInt32((cameraXSize / (double)i) + 1), 0, Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1,
                            $"X start outside binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Y start outside binned chip dimensions
                        CameraExposure($"Reject Bad YStart (bin {i} x {j})", $"Testing rejection of bad Y start value for bin {i} x {j} ({Convert.ToInt32(cameraYSize / (double)j) + 1})", i, j, 0, Convert.ToInt32((cameraYSize / (double)j) + 1), Convert.ToInt32(cameraXSize / (double)i), System.Convert.ToInt32(cameraYSize / (double)j), 0.1,
                            $"Y start outside binned chip size, Bin {i}x{j}");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
            }
        }

        public override void CheckPerformance()
        {
            CameraPerformanceTest(CameraPerformance.CameraState, "CameraState");
            CameraPerformanceTest(CameraPerformance.CcdTemperature, "CCDTemperature");
            CameraPerformanceTest(CameraPerformance.CoolerPower, "CoolerPower");
            if (mCanSetCcdTemperature)
                CameraPerformanceTest(CameraPerformance.HeatSinkTemperature, "HeatSinkTemperature");
            CameraPerformanceTest(CameraPerformance.ImageReady, "ImageReady");
            if (mCanPulseGuide)
                CameraPerformanceTest(CameraPerformance.IsPulseGuiding, "IsPulseGuiding");
            SetAction("Exposure for ImageArray Test");
            SetStatus("Start");
            LogCallToDriver("ConformanceCheck", "About to set BinX");
            camera.BinX = 1;
            LogCallToDriver("ConformanceCheck", "About to set BinY");
            camera.BinY = 1;
            LogCallToDriver("ConformanceCheck", "About to set StartX");
            camera.StartX = 0;
            LogCallToDriver("ConformanceCheck", "About to set StartY");
            camera.StartY = 0;
            LogCallToDriver("ConformanceCheck", "About to set MaxBinX");
            camera.NumX = camera.MaxBinX;
            LogCallToDriver("ConformanceCheck", "About to set MaxBinY");
            camera.NumY = camera.MaxBinY;
            LogCallToDriver("ConformanceCheck", "About to call StartExposure");
            camera.StartExposure(1, true); // 1 second exposure
            LogCallToDriver("ConformanceCheck", "About to call ImageReady multiple times");
            do
                SetStatus("Waiting for ImageReady");
            while (!camera.ImageReady);
            SetStatus("Finished");
            CameraPerformanceTest(CameraPerformance.ImageArray, "ImageArray");
            CameraPerformanceTest(CameraPerformance.ImageArrayVariant, "ImageArrayVariant");
        }

        public override void PostRunCheck()
        {
            if (mCanAbortExposure)
            {
                LogCallToDriver("ConformanceCheck", "About to call AbortExposure");
                try { camera.AbortExposure(); } catch { }
            }

            if (mCanStopExposure)
            {
                LogCallToDriver("ConformanceCheck", "About to call StopExposure");
                try { camera.StopExposure(); } catch { }
            }

            if (mCanSetCcdTemperature)
            {
                LogCallToDriver("ConformanceCheck", "About to set SetCCDTemperature");
                try
                {
                    camera.SetCCDTemperature = mSetCcdTemperature;
                    LogOk("PostRunCheck", "Camera returned to initial cooler temperature");
                }
                catch { }
            }
            LogCallToDriver("ConformanceCheck", "About to set CoolerOn");
            try { camera.CoolerOn = mCoolerOn; } catch { }

            // Reset the camera image parameters to legal values
            LogCallToDriver("ConformanceCheck", "About to set StartX");
            try { camera.StartX = 0; } catch { }
            LogCallToDriver("ConformanceCheck", "About to set StartY");
            try { camera.StartY = 0; } catch { }
            LogCallToDriver("ConformanceCheck", "About to set BinX");
            try { camera.BinX = 1; } catch { }
            LogCallToDriver("ConformanceCheck", "About to set BinY");
            try { camera.BinY = 1; } catch { }
            LogCallToDriver("ConformanceCheck", "About to set NumX");
            try { camera.NumX = 1; } catch { }
            LogCallToDriver("ConformanceCheck", "About to set NumY");
            try { camera.NumY = 1; } catch { }
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

                // Miscellaneous configuration
                if (!settings.CameraFirstUseTests)
                    LogConfigurationAlert("First use tests were omitted due to Conform configuration.");
                if (!settings.CameraTestImageArrayVariant)
                    LogConfigurationAlert("ImageArrayVariant tests were omitted due to Conform configuration.");

            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n:{ex}");
            }
        }

        #endregion

        #region Support Code

        private void CameraCanTest(CanType pType, string pName)
        {
            try
            {
                switch (pType)
                {
                    case CanType.TstCanAbortExposure:
                        LogCallToDriver("ConformanceCheck", "About to get CanAbortExposure");
                        mCanAbortExposure = TimeFunc<bool>(pName, () => camera.CanAbortExposure, TargetTime.Fast);
                        LogOk(pName, mCanAbortExposure.ToString());
                        break;

                    case CanType.TstCanAsymmetricBin:
                        LogCallToDriver("ConformanceCheck", "About to get CanAsymmetricBin");
                        mCanAsymmetricBin = TimeFunc<bool>(pName, () => camera.CanAsymmetricBin, TargetTime.Fast);
                        LogOk(pName, mCanAsymmetricBin.ToString());
                        break;

                    case CanType.TstCanGetCoolerPower:
                        LogCallToDriver("ConformanceCheck", "About to get CanGetCoolerPower");
                        mCanGetCoolerPower = TimeFunc<bool>(pName, () => camera.CanGetCoolerPower, TargetTime.Fast);
                        LogOk(pName, mCanGetCoolerPower.ToString());
                        break;

                    case CanType.TstCanPulseGuide:
                        LogCallToDriver("ConformanceCheck", "About to get CanPulseGuide");
                        mCanPulseGuide = TimeFunc<bool>(pName, () => camera.CanPulseGuide, TargetTime.Fast);
                        LogOk(pName, mCanPulseGuide.ToString());
                        break;

                    case CanType.TstCanSetCcdTemperature:
                        LogCallToDriver("ConformanceCheck", "About to get CanSetCCDTemperature");
                        mCanSetCcdTemperature = TimeFunc<bool>(pName, () => camera.CanSetCCDTemperature, TargetTime.Fast);
                        LogOk(pName, mCanSetCcdTemperature.ToString());
                        break;

                    case CanType.TstCanStopExposure:
                        LogCallToDriver("ConformanceCheck", "About to get CanStopExposure");
                        mCanStopExposure = TimeFunc<bool>(pName, () => camera.CanStopExposure, TargetTime.Fast);
                        LogOk(pName, mCanStopExposure.ToString());
                        break;

                    case CanType.TstCanFastReadout:
                        LogCallToDriver("ConformanceCheck", "About to get CanFastReadout");
                        mCanFastReadout = TimeFunc<bool>(pName, () => camera.CanFastReadout, TargetTime.Fast);
                        LogOk(pName, mCanFastReadout.ToString());
                        break;

                    default:
                        LogIssue(pName, $"Conform:CanTest: Unknown test type {pType}");
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.Mandatory, ex, "");
            }
        }

        private CameraState CameraPropertyTestCameraState(CamPropertyType pType, string pName)
        {
            CameraState returnValue = CameraState.Idle;

            try
            {
                switch (pType)
                {
                    case CamPropertyType.CameraState:
                        {
                            LogCallToDriver("ConformanceCheck", "About to get CameraState");
                            returnValue = TimeFunc(pName, () => camera.CameraState, TargetTime.Fast);
                            break;
                        }

                    default:
                        {
                            LogIssue(pName, $"returnValue: Unknown test type - {pType}");
                            break;
                        }
                }
                LogOk(pName, returnValue.ToString());
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.Optional, ex, "");
            }

            return returnValue;
        }

        private short CameraPropertyTestShort(CamPropertyType pType, string pName, short pMin, short pMax, bool pMandatory)
        {
            short returnValue = 0;

            try
            {
                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case CamPropertyType.BayerOffsetX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get BayerOffsetX");
                                returnValue = camera.BayerOffsetX;
                                break;
                            }

                        case CamPropertyType.BayerOffsetY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get BayerOffsetY");
                                returnValue = camera.BayerOffsetY;
                                break;
                            }

                        case CamPropertyType.PercentCompleted:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get PercentCompleted");
                                returnValue = camera.PercentCompleted;
                                break;
                            }

                        case CamPropertyType.ReadoutMode:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get ReadoutMode");
                                returnValue = camera.ReadoutMode;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"returnValue: Unknown test type - {pType}");
                                break;
                            }
                    }
                }, TargetTime.Fast);

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < pMin // Lower than minimum value
                   :
                        {
                            LogIssue(pName, $"Invalid value: {returnValue}");
                            break;
                        }

                    case object _ when returnValue > pMax // Higher than maximum value
             :
                        {
                            LogIssue(pName, $"Invalid value: {returnValue}");
                            break;
                        }

                    default:
                        {
                            LogOk(pName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }

        private bool CameraPropertyMustNotImplemented(CamPropertyType pType, string pName)
        {
            short testShort; // Dummy variable to hold value that should never be returned
            bool returnValue = true;

            try
            {
                switch (pType)
                {
                    case CamPropertyType.BayerOffsetX:
                        {
                            LogCallToDriver("ConformanceCheck", "About to get BayerOffsetX");
                            testShort = camera.BayerOffsetX;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogIssue(pName, "Sensor type is Monochrome so this property must throw a PropertyNotImplemented error; it must not return a value");
                            break;
                        }

                    case CamPropertyType.BayerOffsetY:
                        {
                            LogCallToDriver("ConformanceCheck", "About to get BayerOffsetY");
                            testShort = camera.BayerOffsetY;
                            returnValue = false; // Property should throw an exception but did not so record that fact
                            LogIssue(pName, "Sensor type is Monochrome so this property must throw a PropertyNotImplemented error; it must not return a value");
                            break;
                        }

                    default:
                        {
                            LogIssue(pName, $"returnValue: Unknown test type - {pType}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.MustNotBeImplemented, ex, "Sensor type is Monochrome");
            }
            return returnValue; // Return success indicator, True means property did thrown the exception, False means that it did not
        }

        private int CameraPropertyTestInteger(CamPropertyType pType, string pName, int pMin, int pMax)
        {
            int returnValue = 0;

            try
            {
                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case CamPropertyType.BinX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get BinX");
                                returnValue = Convert.ToInt32(camera.BinX);
                                break;
                            }

                        case CamPropertyType.BinY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get BinY");
                                returnValue = Convert.ToInt32(camera.BinY);
                                break;
                            }

                        case CamPropertyType.CameraState:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get CameraState");
                                returnValue = (int)camera.CameraState;
                                break;
                            }

                        case CamPropertyType.CameraXSize:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get CameraXSize");
                                returnValue = camera.CameraXSize;
                                break;
                            }

                        case CamPropertyType.CameraYSize:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get CameraYSize");
                                returnValue = camera.CameraYSize;
                                break;
                            }

                        case CamPropertyType.MaxAdu:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get MaxADU");
                                returnValue = camera.MaxADU;
                                break;
                            }

                        case CamPropertyType.MaxBinX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get MaxBinX");
                                returnValue = camera.MaxBinX;
                                break;
                            }

                        case CamPropertyType.MaxBinY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get MaxBinY");
                                returnValue = camera.MaxBinY;
                                break;
                            }

                        case CamPropertyType.NumX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get NumX");
                                returnValue = camera.NumX;
                                break;
                            }

                        case CamPropertyType.NumY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get NumY");
                                returnValue = camera.NumY;
                                break;
                            }

                        case CamPropertyType.StartX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get StartX");
                                returnValue = camera.StartX;
                                break;
                            }

                        case CamPropertyType.StartY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get StartY");
                                returnValue = camera.StartY;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"returnValue: Unknown test type - {pType}");
                                break;
                            }
                    }
                }, TargetTime.Fast);

                // Successfully retrieved a value so test it
                if (returnValue < pMin)
                {
                    LogIssue(pName, $"Invalid value below expected minimum ({pMin}): {returnValue}");
                }
                else if (returnValue > pMax)
                {
                    switch (pType) // Provide the required message depending on the property being tested
                    {
                        case CamPropertyType.MaxBinX // Informational message for MaxBinX
                       :
                            {
                                LogInfo(pName, $"{returnValue}. This is higher than Conform's test criterion: {MAX_BIN_X}. Is this intended?");
                                break;
                            }

                        case CamPropertyType.MaxBinY // Informational message for MaxBinY
                 :
                            {
                                LogInfo(pName, $"{returnValue}. This is higher than Conform's test criterion: {MAX_BIN_Y}. Is this intended?");
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"Invalid value (expected range: {pMin} - {pMax}): {returnValue}");
                                break;
                            }
                    }
                }
                else
                {
                    LogOk(pName, returnValue.ToString());
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.Optional, ex, "");
            }
            return returnValue;
        }

        private double CameraPropertyTestDouble(CamPropertyType pType, string pName, double pMin, double pMax, bool pMandatory)
        {
            double returnValue = 0.0;

            try
            {
                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case CamPropertyType.CcdTemperature:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get CCDTemperature");
                                returnValue = camera.CCDTemperature;
                                break;
                            }

                        case CamPropertyType.CoolerPower:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get CoolerPower");
                                returnValue = camera.CoolerPower;
                                break;
                            }

                        case CamPropertyType.ElectronsPerAdu:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get ElectronsPerADU");
                                returnValue = camera.ElectronsPerADU;
                                break;
                            }

                        case CamPropertyType.FullWellCapacity:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get FullWellCapacity");
                                returnValue = camera.FullWellCapacity;
                                break;
                            }

                        case CamPropertyType.HeatSinkTemperature:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get HeatSinkTemperature");
                                returnValue = camera.HeatSinkTemperature;
                                break;
                            }

                        case CamPropertyType.PixelSizeX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get PixelSizeX");
                                returnValue = camera.PixelSizeX;
                                break;
                            }

                        case CamPropertyType.PixelSizeY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get PixelSizeY");
                                returnValue = camera.PixelSizeY;
                                break;
                            }

                        case CamPropertyType.SetCcdTemperature:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get SetCCDTemperature");
                                returnValue = camera.SetCCDTemperature;
                                break;
                            }

                        case CamPropertyType.ExposureMax:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get ExposureMax");
                                returnValue = camera.ExposureMax;
                                break;
                            }

                        case CamPropertyType.ExposureMin:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get ExposureMin");
                                returnValue = camera.ExposureMin;
                                break;
                            }

                        case CamPropertyType.ExposureResolution:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get ExposureResolution");
                                returnValue = camera.ExposureResolution;
                                break;
                            }

                        case CamPropertyType.SubExposureDuration:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get SubExposureDuration");
                                returnValue = camera.SubExposureDuration;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"returnValue: Unknown test type - {pType}");
                                break;
                            }
                    }
                }, TargetTime.Fast);
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case double _ when returnValue < pMin // Lower than minimum value
                   :
                        {
                            LogIssue(pName, $"Invalid value: {returnValue}");
                            break;
                        }

                    case double _ when returnValue > pMax // Higher than maximum value
             :
                        {
                            LogIssue(pName, $"Invalid value: {returnValue}");
                            break;
                        }

                    default:
                        {
                            LogOk(pName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        private bool CameraPropertyTestBoolean(CamPropertyType pType, string pName, bool pMandatory)
        {
            bool returnValue = false;

            try
            {
                returnValue = false;
                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case CamPropertyType.CoolerOn:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get CoolerOn");
                                returnValue = camera.CoolerOn;
                                break;
                            }

                        case CamPropertyType.HasShutter:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get HasShutter");
                                returnValue = camera.HasShutter;
                                break;
                            }

                        case CamPropertyType.ImageReady:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get ImageReady");
                                returnValue = camera.ImageReady;
                                break;
                            }

                        case CamPropertyType.IsPulseGuiding:
                            {
                                mIsPulseGuidingFunctional = false;
                                LogCallToDriver("ConformanceCheck", "About to get IsPulseGuiding");
                                returnValue = camera.IsPulseGuiding;
                                mIsPulseGuidingFunctional = true; // Command works properly and doesn't cause a not implemented exception
                                break;
                            }

                        case CamPropertyType.FastReadout:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get FastReadout");
                                returnValue = camera.FastReadout;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"returnValue: Unknown test type - {pType}");
                                break;
                            }
                    }
                }, TargetTime.Fast);

                // Successfully retrieved a value
                LogOk(pName, returnValue.ToString());
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }

        private string CameraPropertyTestString(CamPropertyType pType, string pName, int pMaxLength, bool pMandatory)
        {
            string returnValue = "";
            try
            {
                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case CamPropertyType.Description:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get Description");
                                returnValue = camera.Description;
                                break;
                            }

                        case CamPropertyType.SensorName:
                            {
                                LogCallToDriver("ConformanceCheck", "About to get SensorName");
                                returnValue = camera.SensorName;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"returnValue: Unknown test type - {pType}");
                                break;
                            }
                    }
                }, TargetTime.Fast);

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue == "":
                        {
                            LogOk(pName, "The driver returned an empty string");
                            break;
                        }

                    default:
                        {
                            if (returnValue.Length <= pMaxLength)
                                LogOk(pName, returnValue);
                            else
                                LogIssue(pName,
                                    $"String exceeds {pMaxLength} characters maximum length - {returnValue}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }

        private void CameraPropertyWriteTest(CamPropertyType pType, string pProperty, int pTestOk)
        {

            // NOTE: Out of range values should not be tested here but later when the exposure is actually taken

            try // OK value first
            {
                TimeMethod($"{pProperty} write", () =>
                {

                    switch (pType)
                    {
                        case CamPropertyType.NumX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to set NumX");
                                camera.NumX = pTestOk;
                                break;
                            }

                        case CamPropertyType.NumY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to set NumY");
                                camera.NumY = pTestOk;
                                break;
                            }

                        case CamPropertyType.StartX:
                            {
                                LogCallToDriver("ConformanceCheck", "About to set StartX");
                                camera.StartX = pTestOk;
                                break;
                            }

                        case CamPropertyType.StartY:
                            {
                                LogCallToDriver("ConformanceCheck", "About to set StartY");
                                camera.StartY = pTestOk;
                                break;
                            }
                    }
                }, TargetTime.Standard);

                LogOk($"{pProperty} write", $"Successfully wrote {pTestOk}");
            }
            catch (Exception ex)
            {
                HandleException($"{pProperty} write", MemberType.Property, Required.MustBeImplemented, ex,
                    $"Can't set legal value: {pTestOk}");
            }
        }

        private void CameraExposure(string testName, string testDescription, int requiredBinX, int requiredBinY, int requiredStartX, int requiredStartY, int requiredNumX, int requiredNumY, double requiredDuration, string expectedErrorMessage)
        {
            string numPlanes, variantType;
            bool exposedOk, imageReadyTooEarly = false; // Flag to determine whether we were successful or something failed
            DateTime startTime, startTimeUtc, endTime;
            short percentCompleted;

            // Start off by assuming the worst case, this will be set true if the exposure completes OK
            exposedOk = false;

            LogDebug(testName, $"Entered CameraExposure - Test description: {testDescription}, Expected error message: {expectedErrorMessage}");
            LogDebug(testName, $"BinX: {requiredBinX}, BinY: {requiredBinY}, StartX: {requiredStartX}, StartY: {requiredStartY}, NumX: {requiredNumX}, NumY: {requiredNumY}, Duration: {requiredDuration}");

            // Set the name of this test for status update purposes
            SetTest(testName);

            // Log test name for happy path tests
            if (testName.Equals("STARTEXPOSURE", StringComparison.InvariantCultureIgnoreCase))
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
                LogCallToDriver("ConformanceCheck", "About to get CameraState");
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
                LogCallToDriver("ConformanceCheck", "About to set BinX");
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
                LogCallToDriver("ConformanceCheck", "About to set BinY");
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
                LogCallToDriver("ConformanceCheck", "About to set StartX");
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
                LogCallToDriver("ConformanceCheck", "About to set StartY");
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
                LogCallToDriver("ConformanceCheck", "About to set NumX");
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
                LogCallToDriver("ConformanceCheck", "About to set NumY");
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
                ResetTestActionStatus();
                return;
            }

            #endregion

            #region Initiate exposure

            // Start exposure because all parameters were set OK

            startTime = DateTime.Now;
            startTimeUtc = DateTime.UtcNow;
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
                    UpdateUi(exposeUiTaskCancellationToken);
                }, exposeUiTaskCancellationToken);

                // Initiate the exposure
                LogCallToDriver("ConformanceCheck", "About to call StartExposure");

                // Create a cancellation token that we can set if the task times out
                CancellationTokenSource cancellationTokenSource = new();
                CancellationToken cancellationToken = cancellationTokenSource.Token;

                // Create and start the StartExposure task
                Task startExposureTask = Task.Run(() =>
                {
                    try
                    {
                        TimeMethodTwoParams("StartExposure", camera.StartExposure, requiredDuration, true, TargetTime.Standard);

                        // Do not respond if the task has been cancelled
                        if (!cancellationToken.IsCancellationRequested) // Task is live and has not been cancelled
                        {
                            LogDebug(testName, "Exposure initiated OK");

                            if (string.IsNullOrEmpty(expectedErrorMessage))
                            {
                                // Flag that the exposure was successfully initiated
                                initiatedOk = true;
                                LogOk("StartExposure", "Completed successfully.");
                            }
                            else
                            {
                                LogTestAndMessage(testName, $"No error was returned when {char.ToLowerInvariant(expectedErrorMessage[0])}{expectedErrorMessage[1..]}");
                                LogIssue(testName, $"Expected an error and didn't get one - BinX:{requiredBinX} BinY:{requiredBinY} StartX:{requiredStartX} StartY:{requiredStartY} NumX:{requiredNumX} NumY:{requiredNumY}");
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
                                LogOk(testName, $"Received error: {ex.Message}");
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
                    ranToCompletion = startExposureTask.Wait(TimeSpan.FromSeconds(settings.CameraWaitTimeout), ApplicationCancellationToken);
                }
                catch (OperationCanceledException)
                {
                    // The user cancelled the operation
                    LogNewLine();
                    LogError("INSTABILITY WARNING", "Start exposure was interrupted, which may leave Conform Universal in a corrupted state. Please restart to ensure reliable operation");
                    LogNewLine();
                    ResetTestActionStatus();
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
                    LogIssue(testName, $"StartExposure {(expectedErrorMessage == "" ? "" : $"{expectedErrorMessage} ")}did not return within the timeout period: {settings.CameraWaitTimeout} seconds.");

                    // Provide a warning about possible application corruption
                    LogNewLine();
                    LogError("INSTABILITY WARNING", "Start exposure timed out, which may leave Conform Universal in a corrupted state. Please restart to ensure reliable operation");
                    LogNewLine();

                    // Cancel the task
                    ConformanceTestManager.ConformCancellationTokenSource.Cancel();

                    ResetTestActionStatus();
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
                ResetTestActionStatus();
                return;
            }

            #endregion

            #region Wait for the synchronous or asynchronous exposure to complete

            try
            {
                endTime = DateTime.Now;

                // Test whether we have a synchronous or asynchronous camera
                LogCallToDriver("ConformanceCheck", "About to get ImageReady and CameraState");
                if (camera.ImageReady & (camera.CameraState == CameraState.Idle)) // Synchronous exposure
                {
                    #region Check synchronous exposure behaviour

                    try
                    {
                        // Test whether the required exposure time has passed
                        if (endTime.Subtract(startTime).TotalSeconds >= requiredDuration) // Required exposure time has elapsed
                        {
                            if (IsPlatform7OrLater) // Platform 7 or later device
                            {
                                LogIssue(testName, $"StartExposure operated synchronously: ImageReady was True and CameraState was Idle when StartExposure returned after the full exposure time: {requiredDuration} seconds.");
                                SynchronousBehaviourInformation(testName, "ImageReady False", "ImageReady True");
                            }
                            else // Platform 6 or earlier device
                                LogOk(testName, $"Synchronous exposure found OK: {requiredDuration} seconds");

                            CameraTestLast(requiredDuration, startTimeUtc);
                        }
                        else // Camera returned early before the required exposure time elapsed.
                            LogIssue(testName, "The camera reported CameraState = CameraIdle and ImageReady = true before the requested exposure time had elapsed.");
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while checking initiation: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ResetTestActionStatus();
                        return;
                    }

                    #endregion
                }
                else // Asynchronous exposure
                {
                    #region Wait for exposing state to start

                    SetStatus("Waiting for exposure to start");

                    // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                    LogCallToDriver("ConformanceCheck", "About to get ImageReady");
                    imageReadyTooEarly = camera.ImageReady;

                    // Wait for exposing state
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to get CameraState multiple times");
                        Stopwatch sw = Stopwatch.StartNew();
                        WaitWhile(GetAction(), () => (camera.CameraState != CameraState.Exposing) & (camera.CameraState != CameraState.Error), 500, settings.CameraWaitTimeout);
                    }
                    catch (TimeoutException ex)
                    {
                        LogIssue(testName, $"The expected exposure time was {requiredDuration} seconds but the camera did not enter the 'Exposing' state within the configured timeout of {settings.CameraWaitTimeout} seconds.");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ResetTestActionStatus();
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while waiting for the camera to enter the 'Exposing' state: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ResetTestActionStatus();
                        return;
                    }

                    if (ApplicationCancellationToken.IsCancellationRequested) // Exit if required
                    {
                        ResetTestActionStatus();
                        return;
                    }

                    // Test whether ImageReady is being set too early i.e. before the camera has returned to idle
                    LogCallToDriver("ConformanceCheck", "About to get ImageReady");
                    imageReadyTooEarly = camera.ImageReady;

                    #endregion

                    #region Wait for exposing state to end

                    try
                    {
                        // Wait for the exposing state to finish
                        startTime = DateTime.Now;
                        startTimeUtc = DateTime.UtcNow;
                        LogCallToDriver("ConformanceCheck", "About to get CameraState, InterfaceVersion and PercentCompleted multiple times...");

                        // Start the loop timing stopwatch
                        sw.Restart();
                        Stopwatch swOverall = Stopwatch.StartNew();

                        WaitWhile($"Waiting for exposure to complete", () => camera.CameraState == CameraState.Exposing, 500, settings.CameraWaitTimeout, () =>
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

                        if (ApplicationCancellationToken.IsCancellationRequested) // Exit if required
                        {
                            ResetTestActionStatus();
                            return;
                        }
                    }
                    catch (TimeoutException)
                    {
                        LogIssue(testName, $"Test abandoned, timed out waiting for camera to leave the 'Exposing' state.");
                        ResetTestActionStatus();
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while waiting for exposure to complete: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ResetTestActionStatus();
                        return;
                    }

                    #endregion

                    #region Wait for camera to become idle

                    try
                    {
                        // Wait for camera to become idle
                        LogCallToDriver("ConformanceCheck", "About to get CameraState multiple times");

                        WaitWhile("Waiting for camera idle state, reading/downloading image", () => (camera.CameraState != CameraState.Idle) & (camera.CameraState != CameraState.Error), 500, settings.CameraWaitTimeout);

                        if (ApplicationCancellationToken.IsCancellationRequested) // Exit if required
                        {
                            ResetTestActionStatus();
                            return;
                        }
                    }
                    catch (TimeoutException)
                    {
                        LogIssue(testName, $"Test abandoned, timed out waiting for camera to become idle.");
                        ResetTestActionStatus();
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while waiting for camera to become idle: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ResetTestActionStatus();
                        return;
                    }

                    #endregion

                    #region Wait for image to be ready for retrieval

                    try
                    {
                        // Wait for image to become ready
                        LogCallToDriver("ConformanceCheck", "About to get CameraState multiple times");

                        // Wait until ImageReady is true or the camera is in the error state
                        WaitWhile("Waiting for image ready", () => !camera.ImageReady & (camera.CameraState != CameraState.Error), 500, settings.CameraWaitTimeout);
                    }
                    catch (TimeoutException)
                    {
                        LogIssue(testName, $"Test abandoned, timed out waiting for ImageReady to become true.");
                        ResetTestActionStatus();
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogIssue(testName, $"Unexpected error while waiting for ImageReady to become true: {ex.Message}");
                        LogDebug(testName, $"Exception detail:\r\n{ex}");
                        ResetTestActionStatus();
                        return;
                    }

                    #endregion
                }
            }
            catch (Exception ex)
            {
                LogIssue(testName, $"Unexpected error while waiting for exposure to complete: {ex.Message}");
                LogDebug(testName, $"Exception detail:\r\n{ex}");
                ResetTestActionStatus();
                return;
            }

            #endregion

            #region Check exposure outcome

            try
            {
                // Exit if required
                if (ApplicationCancellationToken.IsCancellationRequested)
                {
                    ResetTestActionStatus();
                    return;
                }

                // Test whether or not an image was successfully recorded
                LogCallToDriver("ConformanceCheck", "About to get ImageReady");
                if (camera.ImageReady) // An image was successfully recorded
                {
                    LogOk("ImageReady", $"ImageReady is True after a successful exposure of {requiredDuration} seconds.");
                    CameraTestLast(requiredDuration, startTimeUtc);
                }
                else // An image was not successfully recorded
                {
                    // Give up because the exposure was not successful
                    LogIssue("ImageReady", "ImageReady is False - Test abandoned because the camera reports that no image is available.");
                    ResetTestActionStatus();
                    return;
                }

                // Display a warning if ImageReady was set too early
                if (imageReadyTooEarly) // ImageReady was set too early
                {
                    LogIssue("ImageReady", "Test abandoned because ImageReady was set True before the exposure time had elapsed.");
                    ResetTestActionStatus();
                    return;
                }
                else // Camera exposed OK and didn't generate an exception
                {
                    exposedOk = true;
                }
                LogDebug(testName, $"Camera exposed image OK: {exposedOk}");

                //Now check the camera state
                CameraState cameraState = camera.CameraState;
                switch (cameraState)
                {
                    case CameraState.Idle:
                        LogOk("CameraState", $"The camera returned the camera state as Camera.Idle after exposure completed.");
                        break;

                    case CameraState.Waiting:
                    case CameraState.Exposing:
                    case CameraState.Reading:
                    case CameraState.Download:
                    case CameraState.Error:
                        LogIssue("CameraState", $"The CameraState property returned an unexpected camera state after exposure completed: {cameraState}.");
                        break;

                    default:
                        throw new InvalidValueException($"The camera returned an unknown camera state: {cameraState}");
                }
            }
            catch (Exception ex)
            {
                LogIssue(testName, $"Unexpected error while checking returned image: {ex.Message}");
                LogDebug(testName, $"Exception detail:\r\n{ex}");
                ResetTestActionStatus();
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

                // Start  a UI update task
                CancellationTokenSource iaUiTaskTokenSource = new();
                CancellationToken iaUiTaskCancellationToken = iaUiTaskTokenSource.Token;
                Task.Run(() =>
                {
                    UpdateUi(iaUiTaskCancellationToken);
                }, iaUiTaskCancellationToken);

                // Create a cancellation token that we can set if the task times out
                CancellationTokenSource iaTokenSource = new();
                CancellationToken iaToken = iaTokenSource.Token;

                // Start a task to retrieve the image array
                Task iaTask = Task.Factory.StartNew(() =>
                {
                    try
                    {
                        LogCallToDriver("ConformanceCheck", "About to get ImageArray");
                        sw.Restart();
                        TimeMethod("ImageArray", () => mImageArray = (Array)camera.ImageArray, TargetTime.Extended);
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
                            throw;
                        }
                    }
                }, iaToken, TaskCreationOptions.LongRunning, TaskScheduler.Default);


                // Wait for the ImageArray task to complete or be cancelled
                bool ranToCompletion = false;
                try
                {
                    ranToCompletion = iaTask.Wait(TimeSpan.FromSeconds(settings.CameraWaitTimeout), ApplicationCancellationToken);
                }
                catch (OperationCanceledException)
                {
                    LogNewLine();
                    LogError("INSTABILITY WARNING", "Get ImageArray was interrupted, which may leave Conform Universal in a corrupted state. Please restart to ensure reliable operation");
                    LogNewLine();
                    ResetTestActionStatus();
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
                            // Test image array variant array for conformity
                            if ((mImageArray.GetLength(0) == requiredNumX) & (mImageArray.GetLength(1) == requiredNumY))
                            {
                                if (mImageArray.GetType().ToString() == "System.Int32[,]" | mImageArray.GetType().ToString() == "System.Int32[,,]")
                                {
                                    if (mImageArray.Rank == 2)
                                        numPlanes = "1 plane";
                                    else
                                    {
                                        numPlanes = "1 plane";
                                        if (mImageArray.GetUpperBound(2) > 0)
                                            numPlanes =
                                                $"{System.Convert.ToString(mImageArray.GetUpperBound(2) + 1)} planes";
                                    }
                                    LogOk("ImageArray", $"Successfully read 32 bit integer array ({numPlanes}) {mImageArray.GetLength(0)} x {mImageArray.GetLength(1)} pixels in {sw.ElapsedMilliseconds}ms.");
                                }
                                else
                                    LogIssue("ImageArray",
                                        $"Expected 32 bit integer array, actually got: {mImageArray.GetType()}");
                            }
                            else if ((mImageArray.GetLength(0) == requiredNumY) & (mImageArray.GetLength(1) == requiredNumX))
                                LogIssue("ImageArray",
                                    $"Camera image dimensions swapped, expected values: {requiredNumX} x {requiredNumY} - actual values: {mImageArray.GetLength(0)} x {mImageArray.GetLength(1)}");
                            else
                                LogIssue("ImageArray",
                                    $"Camera image does not have the expected dimensions of: {requiredNumX} x {requiredNumY} - actual values: {mImageArray.GetLength(0)} x {mImageArray.GetLength(1)}");
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
                    LogError("INSTABILITY WARNING", "Get ImageArray timed out, which may leave Conform Universal in a corrupted state. Please restart to ensure reliable operation");
                    LogNewLine();

                    // Cancel the task
                    ConformanceTestManager.ConformCancellationTokenSource.Cancel();

                    ResetTestActionStatus();

                    return;
                }

                // Exit if cancelled
                if (ApplicationCancellationToken.IsCancellationRequested) // Exit if required
                {
                    ResetTestActionStatus();
                    return;
                }

            }
            catch (OutOfMemoryException ex)
            {
                // Log an error
                LogError("ImageArray", $"OutOfMemoryException - Conform Universal or the device ran out of memory: {ex.Message}");
                LogDebug("ImageArray", $"Exception detail: {ex}");
                ResetTestActionStatus();
                return;
            }
            catch (Exception ex)
            {
                LogIssue("ImageArray", $"Error when reading ImageArray: {ex.Message}");
                LogDebug("ImageArray", $"Exception detail: {ex}");
                ResetTestActionStatus();
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
                        UpdateUi(iavUiTaskCancellationToken);
                    }, iavUiTaskCancellationToken);

                    // Create a cancellation token that we can set if the task times out
                    CancellationTokenSource iavTokenSource = new();
                    CancellationToken iavToken = iavTokenSource.Token;

                    Task iavTask = Task.Factory.StartNew<bool>(() =>
                    {
                        bool gotImageOk = false;
                        try
                        {
                            object imageObject = null;

                            // Get the variant array image
                            LogCallToDriver("ConformanceCheck", "About to get ImageArrayVariant");
                            sw.Restart();
                            TimeMethod("ImageArrayVariant", () => imageObject = camera.ImageArrayVariant, TargetTime.Extended);

                            sw.Stop();

                            // Assign the array to the application variable if the task is not cancelled
                            if (!cancellationToken.IsCancellationRequested) // Completed successfully
                            {
                                mImageArrayVariant = (Array)imageObject;
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

                    bool ranToCompletion = false;
                    // Wait for the ImageArrayVariant task to complete or be cancelled
                    try
                    {
                        ranToCompletion = iavTask.Wait(TimeSpan.FromSeconds(settings.CameraWaitTimeout), ApplicationCancellationToken);
                    }
                    catch (OperationCanceledException)
                    {
                        LogNewLine();
                        LogError("INSTABILITY WARNING", "Get ImageArrayVariant was interrupted, which may leave Conform Universal in a corrupted state. Please restart to ensure reliable operation");
                        LogNewLine();
                        ResetTestActionStatus();
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
                        LogError("INSTABILITY WARNING", "Get ImageArrayVariant timed out, which may leave Conform Universal in a corrupted state. Please restart to ensure reliable operation");
                        LogNewLine();

                        // Cancel the task
                        ConformanceTestManager.ConformCancellationTokenSource.Cancel();

                        ResetTestActionStatus();
                        return;
                    }

                    // Exit if cancelled
                    if (ApplicationCancellationToken.IsCancellationRequested)
                    {
                        ResetTestActionStatus();
                        return;
                    }

                    // Test image array variant array for conformity
                    if ((mImageArrayVariant.GetLength(0) == requiredNumX) & (mImageArrayVariant.GetLength(1) == requiredNumY))
                    {
                        if (mImageArrayVariant.GetType().ToString() == "System.Object[,]" | mImageArrayVariant.GetType().ToString() == "System.Object[,,]")
                        {
                            if (mImageArrayVariant.Rank == 2)
                            {
                                numPlanes = "1 plane";
                                variantType = ((object[,])mImageArrayVariant)[0, 0].GetType().ToString();
                            }
                            else
                            {
                                numPlanes = "1 plane";
                                if (mImageArrayVariant.GetUpperBound(2) > 0)
                                {
                                    numPlanes =
                                        $"{System.Convert.ToString(mImageArrayVariant.GetUpperBound(2) + 1)} planes";
                                    variantType = ((object[,,])mImageArrayVariant)[0, 0, 0].GetType().ToString();
                                }
                                else
                                    variantType = ((object[,])mImageArrayVariant)[0, 0].GetType().ToString();
                            }
                            LogOk("ImageArrayVariant", $"Successfully read variant array ({numPlanes}) with {variantType} elements {mImageArrayVariant.GetLength(0)} x {mImageArrayVariant.GetLength(1)} pixels in {sw.ElapsedMilliseconds}ms.");
                        }
                        else
                            LogIssue("ImageArrayVariant",
                                $"Expected variant array, actually got: {mImageArrayVariant.GetType()}");
                    }
                    else if ((mImageArrayVariant.GetLength(0) == requiredNumY) & (mImageArrayVariant.GetLength(1) == requiredNumX))
                        LogIssue("ImageArrayVariant",
                            $"Camera image dimensions swapped, expected values: {requiredNumX} x {requiredNumY} - actual values: {mImageArrayVariant.GetLength(0)} x {mImageArrayVariant.GetLength(1)}");
                    else
                        LogIssue("ImageArrayVariant",
                            $"Camera image does not have the expected dimensions of: {requiredNumX} x {requiredNumY} - actual values: {mImageArrayVariant.GetLength(0)} x {mImageArrayVariant.GetLength(1)}");

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
                LogCallToDriver("ConformanceCheck", "About to call StopExposure");
                camera.StopExposure();
            }
            catch (Exception) { }

            try
            {
                LogCallToDriver("ConformanceCheck", "About to call AbortExposure");
                camera.AbortExposure();
            }
            catch (Exception) { }

            ResetTestActionStatus();

            #endregion  
        }

        /// <summary>
        /// Updates the UI status field with a count from 0.0 to the operation timeout.
        /// </summary>
        /// <param name="updateUiTaskCancellationToken"></param>
        private void UpdateUi(CancellationToken updateUiTaskCancellationToken)
        {
            // Update the status every 500ms
            const int pollInterval = 500;

            // Continually update the status field until the task is cancelled
            do
            {
                // Calculate the current loop number (starts at 0 given that the timer's elapsed time will be zero or very low on the first loop)
                int currentLoopNumber = ((int)(sw.ElapsedMilliseconds) + 50) / pollInterval; // Add a small positive offset (50) because integer division always rounds down

                // Calculate the sleep time required to start the next loop at a multiple of the poll interval
                int sleeptime = pollInterval * (currentLoopNumber + 1) - (int)sw.ElapsedMilliseconds;

                // Sleep until it is time for the next completion function poll
                Thread.Sleep(sleeptime);

                // Update the status message if the task has not been cancelled
                if (!updateUiTaskCancellationToken.IsCancellationRequested)
                {
                    SetStatus($"{sw.Elapsed.TotalSeconds:0.0} / {settings.CameraWaitTimeout:0.0}");
                }

                // Wait for the start of the next loop
                Thread.Sleep(sleeptime);
            } while (!updateUiTaskCancellationToken.IsCancellationRequested & (sw.Elapsed.TotalSeconds < settings.CameraWaitTimeout));
        }

        private void CameraTestLast(double pDuration, DateTime pStart)
        {
            DateTime lStartTime;

            // LastExposureDuration
            try
            {
                LogCallToDriver("ConformanceCheck", "About to get LastExposureDuration");
                mLastExposureDuration = camera.LastExposureDuration;
                if ((Math.Abs(mLastExposureDuration - pDuration) / pDuration) < 0.02)
                    LogOk("LastExposureDuration", $"Last exposure duration is: {mLastExposureDuration:0.000} seconds");
                else
                    LogIssue("LastExposureDuration",
                        $"LastExposureDuration is not the same as image duration: {mLastExposureDuration} {pDuration}");
            }
            catch (Exception ex)
            {
                HandleException("LastExposureDuration", MemberType.Property, Required.Optional, ex, "");
            }

            // LastExposurestartTime
            try // Confirm that it can be read
            {
                LogCallToDriver("ConformanceCheck", "About to get LastExposureStartTime");
                mLastExposureStartTime = camera.LastExposureStartTime;
                int lI;
                // Confirm that the format is as expected
                bool lFormatOk;
                lFormatOk = true;
                if (mLastExposureStartTime.Length >= 19)
                {
                    for (lI = 1; lI <= 19; lI++)
                    {
                        switch (lI)
                        {
                            case 5:
                            case 8 // "-"
                           :
                                {
                                    if (mLastExposureStartTime.Substring(lI - 1, 1) != "-")
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{mLastExposureStartTime.Substring(lI - 1, 1)}' at position {lI}is not '-'");
                                        lFormatOk = false;
                                    }
                                    else
                                    {

                                    }
                                    break;
                                }

                            case 11 // "T"
                     :
                                {
                                    if (mLastExposureStartTime.Substring(lI - 1, 1) != "T")
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{mLastExposureStartTime.Substring(lI - 1, 1)}' at position {lI}is not 'T'");
                                        lFormatOk = false;
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
                                    if (mLastExposureStartTime.Substring(lI - 1, 1) != ":")
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{mLastExposureStartTime.Substring(lI - 1, 1)}' at position {lI}is not ':'");
                                        lFormatOk = false;
                                    }
                                    else
                                    {

                                    }
                                    break;
                                }

                            default:
                                {
                                    if (!mLastExposureStartTime.Substring(lI - 1, 1).IsNumeric())
                                    {
                                        LogInfo("LastExposureDuration", $"Character '{mLastExposureStartTime.Substring(lI - 1, 1)}' at position {lI}is not 'numeric'");
                                        lFormatOk = false;
                                    }
                                    else
                                    {

                                    }
                                    break;
                                }
                        }
                    }
                    if (lFormatOk)
                    {
                        try // Confirm that it parses as a valid date and check for correct value
                        {
                            lStartTime = DateTime.Parse(mLastExposureStartTime);
                            if (pStart.Subtract(lStartTime).TotalSeconds < 2.0)
                                LogOk("LastExposureStartTime",
                                    $"LastExposureStartTime is correct to within 2 seconds: {mLastExposureStartTime} UTC");
                            else
                                LogIssue("LastExposureStartTime",
                                    $"LastExposureStartTime is more than 2 seconds inaccurate : {mLastExposureStartTime}, expected: {pStart:yyyy-MM-ddTHH:mm:ss} UTC");
                        }
                        catch (Exception ex)
                        {
                            LogIssue("LastExposureStartTime",
                                $"Parsing LastExposureStartTime - {ex.Message} {mLastExposureStartTime}");
                            LogDebug("LastExposureStartTime", $"Exception detail:\r\n {ex}");
                        }
                    }
                    else
                        LogIssue("LastExposureStartTime",
                            $"LastExposureStartTime not in the expected format yyyy-mm-ddThh:mm:ss - {mLastExposureStartTime}");
                }
                else if (mLastExposureStartTime == "")
                    LogIssue("LastExposureStartTime", "LastExposureStartTime has returned an empty string - expected yyyy-mm-ddThh:mm:ss");
                else
                    LogIssue("LastExposureStartTime",
                        $"LastExposureStartTime is less than 19 characters - expected yyyy-mm-ddThh:mm:ss - {mLastExposureStartTime}");
            }
            catch (Exception ex)
            {
                HandleException("LastExposureStartTime", MemberType.Property, Required.Optional, ex, "");
            }
        }

        private void CameraPulseGuideTest(GuideDirection pDirection)
        {
            // If this is an ICameraV4 or later device pulse guiding cannot be tested without the IsPulseGuiding property being functional to support async operation
            if (IsPlatform7OrLater)
            {
                if (!mIsPulseGuidingFunctional)
                {
                    LogIssue($"PulseGuide {pDirection}", $"Skipping this test because IsPulseGuiding returned an error or threw an exception when tested.");
                    return;
                }
            }

            // Start a timer to measure the duration of the PulseGuide method
            Stopwatch duration = Stopwatch.StartNew();

            LogCallToDriver("ConformanceCheck", $"About to call PulseGuide - {pDirection}");
            TimeMethodTwoParams($"PulseGuide {pDirection}", camera.PulseGuide, pDirection, CAMERA_PULSE_DURATION_MILLISECONDS, IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

            // Stop the duration timer
            duration.Stop();

            try
            {
                // Check whether IsPulseGuiding is functional
                if (mIsPulseGuidingFunctional) // IsPulseGuiding is functional - ICameraV4 and later interfaces always use this path
                {
                    // Test how long the PulseGuide method took to complete
                    if (duration.Elapsed.TotalSeconds < (IsPlatform7OrLater ? standardTargetResponseTime : (CAMERA_PULSE_DURATION_MILLISECONDS - 500) / 1000.0)) // Completed before the required pulse guide duration (so assume asynchronous operation)
                    {
                        // Check whether the camera is still pulse guiding 
                        LogCallToDriver("ConformanceCheck", "About to get IsPulseGuiding");
                        if (camera.IsPulseGuiding) // Camera is still pulse guiding
                        {
                            // Wait for the pulse guide to complete
                            Stopwatch sw = Stopwatch.StartNew();
                            LogCallToDriver("ConformanceCheck", "About to get IsPulseGuiding multiple times");
                            WaitWhile($"Guiding {pDirection}", () => camera.IsPulseGuiding, 500, 3, () => $"{sw.Elapsed.TotalSeconds:0.0} / {CAMERA_PULSE_DURATION_MILLISECONDS / 1000:0.0} seconds");

                            // Check whether the camera has now finished pulse guiding
                            LogCallToDriver("ConformanceCheck", "About to get IsPulseGuiding");
                            if (!camera.IsPulseGuiding) // The camera has finished the pulse guide
                                LogOk($"PulseGuide {pDirection}", "Asynchronous pulse guide found OK");
                            else // The wait timed out and the camera is still reporting that it is pulse guiding
                                LogIssue($"PulseGuide {pDirection}", "Asynchronous pulse guide expected but IsPulseGuiding is TRUE beyond expected time of 2 seconds");
                        }
                        else // Camera is not pulse guiding so the pulse guide appears to have finished early
                            LogIssue($"PulseGuide {pDirection}", "PulseGuide returned quickly indicating an asynchronous pulse guide, but IsPulseGuiding returned FALSE before the pulse guide duration had elapsed.");
                    }
                    else // Completed in close to the required pulse guide time (so assume synchronous operation)
                    {
                        // Test whether the camera is still pulse guiding
                        LogCallToDriver("ConformanceCheck", "About to get IsPulseGuiding");
                        if (!camera.IsPulseGuiding) // Camera is NOT pulse guiding
                            LogOk($"PulseGuide {pDirection}", "Synchronous pulse guide found OK");
                        else // Camera is still pulse guiding
                            LogIssue($"PulseGuide {pDirection}", "Synchronous pulse guide expected but IsPulseGuiding returned TRUE");
                    }
                }
                else // IsPulseGuiding is NOT functional - Never executed for ICameraV4 and later interfaces
                {
                    switch (duration.Elapsed.TotalMilliseconds - CAMERA_PULSE_DURATION_MILLISECONDS)
                    {
                        // Duration was more than 0.5 seconds longer than expected
                        case object _ when duration.Elapsed.TotalMilliseconds - CAMERA_PULSE_DURATION_MILLISECONDS > CAMERA_PULSE_TOLERANCE_MILLISECONDS:
                            LogIssue($"PulseGuide {pDirection}", $"Synchronous pulse guide longer than expected {(CAMERA_PULSE_DURATION_MILLISECONDS) / (double)1000} seconds: {duration.Elapsed.TotalSeconds} seconds");
                            break;

                        // Duration was more than 20ms shorter than expected
                        case object _ when duration.Elapsed.TotalMilliseconds - CAMERA_PULSE_DURATION_MILLISECONDS < 20:
                            LogIssue($"PulseGuide {pDirection}", $"Synchronous pulse guide shorter than expected {(CAMERA_PULSE_DURATION_MILLISECONDS) / (double)1000} seconds: {duration.Elapsed.TotalSeconds} seconds");
                            break;

                        // All other cases
                        default:
                            LogOk($"PulseGuide {pDirection}", $"Synchronous pulse guide found OK: {duration.Elapsed.TotalSeconds} seconds");
                            break;
                    }
                }
            }
            catch (TimeoutException ex)
            {
                LogIssue($"PulseGuide {pDirection}", $"Timed out waiting for IsPulseGuiding to go false. It should have done this in {Convert.ToDouble(CAMERA_PULSE_DURATION_MILLISECONDS) / 1000.0:0.0} seconds");
                LogDebug("PulseGuide", $"Exception detail:\r\n {ex}");
            }
        }

        private void CameraPerformanceTest(CameraPerformance pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime, lRate;
            SetAction(pName);
            try
            {
                lStartTime = DateTime.Now;
                lCount = 0.0;
                lLastElapsedTime = 0.0;
                do
                {
                    lCount += 1.0;
                    switch (pType)
                    {
                        case CameraPerformance.CameraState:
                            {
                                mCameraState = camera.CameraState;
                                break;
                            }

                        case CameraPerformance.CcdTemperature:
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
                                mImageReady = camera.ImageReady;
                                break;
                            }

                        case CameraPerformance.IsPulseGuiding:
                            {
                                mIsPulseGuiding = camera.IsPulseGuiding;
                                break;
                            }

                        case CameraPerformance.ImageArray:
                            {
                                // Release memory currently consumed by images
                                ReleaseMemory();

                                mImageArray = (Array)camera.ImageArray;
                                break;
                            }

                        case CameraPerformance.ImageArrayVariant:
                            {
                                // Release memory currently consumed by images
                                ReleaseMemory();

                                mImageArrayVariant = (Array)camera.ImageArrayVariant;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"Conform:PerformanceTest: Unknown test type {pType}");
                                break;
                            }
                    }

                    lElapsedTime = DateTime.Now.Subtract(lStartTime).TotalSeconds;
                    if (lElapsedTime > lLastElapsedTime + 1.0)
                    {
                        SetStatus($"{lCount} transactions in {lElapsedTime:0} seconds");
                        lLastElapsedTime = lElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (lElapsedTime <= PERF_LOOP_TIME);

                lRate = lCount / lElapsedTime;
                switch (lRate)
                {
                    case object _ when lRate > 10.0:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case object _ when 2.0 <= lRate && lRate <= 10.0:
                        {
                            LogOk(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case object _ when 1.0 <= lRate && lRate <= 2.0:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(pName, $"Unable to complete test: {ex.Message}");
                LogDebug(pName, $"Exception detail:\r\n {ex}");
            }
        }

        /// <summary>
        /// Release memory allocated to the large arrays on the large object heap.
        /// </summary>
        private void ReleaseMemory()
        {
            SetAction("Releasing memory");

            // Clear out any previous memory allocations
            mImageArray = null;
            mImageArrayVariant = null;
            GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
            GC.Collect(2, GCCollectionMode.Forced, true, true);
        }

        #endregion

    }
}

using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConformU
{
    internal class CoverCalibratorTester : DeviceTesterBaseClass
    {
        private enum RequiredProperty
        {
            CalibratorState,
            CoverState,
            CalibratorChanging,
            CoverMoving
        }

        private enum PerformanceProperty
        {
            CalibratorState,
            CoverState
        }

        private CoverStatus coverState;
        private CalibratorStatus calibratorState;
        private bool calibratorChanging;
        private bool coverMoving;
        private bool canAsynchronousOpen;
        private double coverOpenTime = 0.0;
        private int maxBrightness;
        private bool calibratorStateOk;
        private bool coverStateOk;
        private bool brightnessOk;
        private bool maxBrightnessOk;
        private bool calibratorChangingOk;
        private bool coverMovingOk;

        // Helper variables
        private ICoverCalibratorV2 coverCalibratorDevice;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose
        public CoverCalibratorTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, true, true, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
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
                    coverCalibratorDevice?.Dispose();
                    coverCalibratorDevice = null;
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
            // Set the error type numbers according to the standards adopted by individual authors.
            // Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
            // messages to driver authors!
            unchecked
            {
                switch (settings.ComDevice.ProgId)
                {
                    default:
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040405;
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
                        LogInfo("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        coverCalibratorDevice = new AlpacaCoverCalibrator(
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

#if WINDOWS
                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComAccessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                coverCalibratorDevice = new CoverCalibratorFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                coverCalibratorDevice = new CoverCalibrator(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;
#endif
                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(coverCalibratorDevice, DeviceTypes.CoverCalibrator); // Assign the driver to the base class

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

        public override void PreConnectChecks()
        {
        }

        public override void CheckProperties()
        {
            int brightness = 0;

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test CalibratroState
            calibratorStateOk = RequiredPropertiesTest(RequiredProperty.CalibratorState, "CalibratorState");

            // Test CoverState
            coverStateOk = RequiredPropertiesTest(RequiredProperty.CoverState, "CoverState");

            // Test MaxBrightness
            if (calibratorStateOk)
            {
                maxBrightnessOk = false; // Assume a bad value
                try
                {
                    maxBrightness = 0; // Initialise to a 'bad' value
                    LogCallToDriver("MaxBrightness", "About to call MaxBrightness property");
                    TimeFunc<int>("MaxBrightness", () => maxBrightness = coverCalibratorDevice.MaxBrightness, TargetTime.Fast);

                    LogCallToDriver("MaxBrightness", "About to call CalibratorState property");
                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent))
                    {
                        if (maxBrightness >= 1)
                        {
                            LogOk("MaxBrightness", maxBrightness.ToString());
                            maxBrightnessOk = true;
                        }
                        else
                            LogIssue("MaxBrightness", $"The returned MaxBrightness value {maxBrightness} is invalid, it must be >= 1");
                    }
                    else
                        LogIssue("MaxBrightness", $" CalibratorStatus is 'NotPresent' but MaxBrightness did not throw a PropertyNotImplementedException. It returned: {maxBrightness}.");
                }
                catch (Exception ex)
                {
                    LogCallToDriver("MaxBrightness", "About to call CalibratorState property");
                    if (coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent)
                        HandleException("MaxBrightness", MemberType.Property, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                    else
                        HandleException("MaxBrightness", MemberType.Property, Required.MustBeImplemented, ex, "CalibratorStatus indicates that he device is a calibrator");
                }
            }
            else
                LogIssue("MaxBrightness", $"Test skipped because CalibratorState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test Brightness
            if (calibratorStateOk)
            {
                brightnessOk = false; // Assume a bad value

                try
                {
                    LogCallToDriver("Brightness", "About to call Brightness property");
                    TimeFunc<int>("Brightness", () => brightness = coverCalibratorDevice.Brightness, TargetTime.Fast);
                    brightnessOk = true;

                    LogCallToDriver("MaxBrightness", "About to call CalibratorState property");
                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent))
                    {
                        if (brightness >= 0)
                        {
                            if (maxBrightnessOk)
                            {
                                if (brightness <= maxBrightness)
                                    LogOk("Brightness", brightness.ToString());
                                else
                                    LogIssue("Brightness", $"The returned Brightness {brightness} is greater than MaxBrightness {maxBrightness}");
                            }
                            else
                                LogIssue("Brightness", $"Skipping Brightness test because MaxBrightness did not return a value.");
                        }
                        else
                            LogIssue("Brightness", $"The returned Brightness value {brightness} is invalid, it must be >= 0");
                    }
                    else
                        LogIssue("Brightness", $" CalibratorStatus is 'NotPresent' but Brightness did not throw a PropertyNotImplementedException. It returned: {brightness}.");
                }
                catch (Exception ex)
                {
                    LogCallToDriver("MaxBrightness", "About to call CalibratorState property");
                    if (coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent)
                        HandleException("Brightness", MemberType.Property, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                    else
                        HandleException("Brightness", MemberType.Property, Required.MustBeImplemented, ex, "CalibratorStatus indicates that he device is a calibrator");
                }
            }
            else
                LogIssue("Brightness", $"Test skipped because CalibratorState returned an exception");

            // If this is an ICoverCalibratorV2 interface or later device make sure we can read CalibratorChanging
            if (IsPlatform7OrLater)
            {
                calibratorChangingOk = RequiredPropertiesTest(RequiredProperty.CalibratorChanging, "CalibratorChanging");
            }

            // If this is an ICoverCalibratorV2 interface or later device make sure we can read CoverMoving
            if (IsPlatform7OrLater)
            {
                coverMovingOk = RequiredPropertiesTest(RequiredProperty.CoverMoving, "CoverMoving");
            }

            // Check that the device implements at least one of the cover or calibrator functions
            if (coverStateOk & calibratorStateOk) // Both cover and calibrator statuses can be read successfully
            {
                if (coverState == CoverStatus.NotPresent & calibratorState == CalibratorStatus.NotPresent) // Both implementations are absent so log an issue
                    LogIssue("DeviceCapabilities", "Both CoverStatus and CalibratorStatus are set to 'NotPresent' - this device won't do a lot!");
            }
        }

        public override void CheckMethods()
        {
            Stopwatch sw = new();
            ClearStatus();  // Clear status messages

            if (cancellationToken.IsCancellationRequested)
                return;

            // Check whether CoverState can be read OK
            if (coverStateOk & (IsPlatform7OrLater ? coverMovingOk : true)) // CoverState and CoverMoving can be read OK
            {
                LogCallToDriver("CoverCalibrator", "About to get CoverState property");
                switch (coverCalibratorDevice.CoverState)
                {
                    // Cover is not implemented so makes sure that not implemented exceptions are thrown
                    case CoverStatus.NotPresent:
                        TestOpenCover();
                        TestCloseCover();
                        TestHaltCover();
                        break;

                    // Status is closed so test OpenCover first then CloseCover
                    case CoverStatus.Closed:
                        TestOpenCover();
                        TestCloseCover();
                        TestHaltCover();
                        break;

                    // Status is moving so skip tests because this should not be the case
                    case CoverStatus.Error:
                    case CoverStatus.Moving:
                        LogIssue("CoverCalibrator", $"The OpenCover, CloseCover and HaltCover tests have been skipped because CoverState reports {coverState} before the tests start.");
                        break;

                    // Status is open so test CloseCover first then OpenCover
                    case CoverStatus.Open:
                        TestCloseCover();
                        TestHaltCover();
                        TestOpenCover();
                        break;

                    // Status is unknown so try to get to the closed state and then test or abandon tests as appropriate
                    case CoverStatus.Unknown:
                        // Try to get to the closed state
                        LogInfo("CoverCalibrator", $"The device reports CoverState as CoverStatus.Unknown. Conform will now test CoverClose to try to get to a known state.");
                        TestCloseCover();

                        // Test whether the cover now reports closed
                        if (coverCalibratorDevice.CoverState == CoverStatus.Closed) // Cover reports a successful close so now test both open and close
                        {
                            TestOpenCover();
                            TestCloseCover();
                            TestHaltCover();
                        }
                        else // Cover does not report closed after the close test completes
                            LogIssue("CoverCalibrator", $"The CloseCover operation completed but the device does not report CoverState as CoverStatus.Closed. The reported state is: {coverState}.");
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CheckMethods - Unknown cover status value: {coverState}.");
                }
            }
            else
            {
                LogIssue("CoverCalibrator", $"The OpenCover, CloseCover and HaltCover tests have been skipped because the CoverState " +
                    $"{(IsPlatform7OrLater ? "or CoverMoving" : "")} property returned an invalid value or threw an exception.");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test CalibratorOn
            SetTest("CalibratorOn");

            // Check whether CalibratorState can be read OK
            if (calibratorStateOk & (IsPlatform7OrLater ? calibratorChangingOk : true)) // CalibratorState and CalibratorChanging did return a values
            {
                // Check whether the calibrator is ready to accept changes
                if (CalibratorReady("CalibratorOn")) // Calibrator is ready to accept changes
                {
                    // Check whether the calibrator is implemented
                    if ((calibratorState != CalibratorStatus.NotPresent)) // Calibrator is implemented
                    {
                        // Check whether MaxBrightness and Brightness could be read OK
                        if (maxBrightnessOk & brightnessOk) // MaxBrightness and Brightness can be read OK
                        {
                            // Test for invalid value -1
                            TestCalibratorOn(-1);
                            if (cancellationToken.IsCancellationRequested)
                                return;

                            // Test for zero brightness
                            TestCalibratorOn(0);
                            if (cancellationToken.IsCancellationRequested)
                                return;

                            // Test intermediate and maximum brightness values
                            switch (maxBrightness)
                            {
                                case 1: // Simple on/ off device
                                    TestCalibratorOn(1);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    break;

                                case 2: // Two brightness level device
                                    TestCalibratorOn(1);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(2);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    break;

                                case 3: // Three brightness level device
                                    TestCalibratorOn(1);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(2);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(3);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    break;

                                case 4: // Four brightness level device
                                    TestCalibratorOn(1);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(2);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(3);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(4);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    break;

                                default:
                                    TestCalibratorOn((int)Math.Ceiling(((maxBrightness + 1.0) / 4.0) - 1.0)); // Round up to ensure that this value is least 1 so there is some level of brightness
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn((int)Math.Floor(((maxBrightness + 1.0) / 2.0) - 1.0));
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn((int)Math.Floor(((maxBrightness + 1.0) * 3.0 / 4.0) - 1.0)); // Round down to ensure that this value is different to the maxBrightness value
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(maxBrightness);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    break;
                            }

                            // Test for rejection of invalid value: MaxBrightness + 1
                            if (maxBrightness < int.MaxValue)
                            {
                                TestCalibratorOn(maxBrightness + 1);
                            }
                            else
                            {
                                LogInfo("CalibratorOn", $"Test of a high invalid brightness value skipped because MaxBrightness is set to the largest positive integer value.");
                            }
                        }
                        else // MaxBrightness and Brightness can not be read OK
                        {
                            LogIssue("CalibratorOn", $"Brightness tests skipped because one of the Brightness or MaxBrightness properties returned an invalid value or threw an exception.");
                        }
                    }
                    else // Calibrator is not implemented so test to ensure that a MethodNotImplementedException is thrown
                    {
                        TestCalibratorOn(1);
                    }
                }
                else // Calibrator is in the not ready state and cannot accept changes
                {
                    LogIssue("CalibratorOn", $"Tests skipped because CalibratorState says the calibrator is not ready - CalibratorState: {calibratorState}");
                }
            }
            else // CalibratorState or CalibratorChanging did not return a valid value
            {
                LogIssue("CalibratorOn", $"Brightness tests skipped because the CalibratorState " +
                    $"{(IsPlatform7OrLater ? "or CalibratorChanging" : "")} property returned an invalid value or threw an exception.");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test CalibratorOff
            SetTest("CalibratorOff");

            // Check whether CalibratorState can be read OK
            if (calibratorStateOk & (IsPlatform7OrLater ? calibratorChangingOk : true))
            {
                TestCalibratorOff();
            }
            else
                LogIssue("CalibratorOff", $"Test skipped because the CoverState property returned an invalid value or threw an exception.");

            ClearStatus();
        }

        public override void PostRunCheck()
        {
            SetTest("Post run check");

            try
            {
                // Attempt to close the cover if present
                if (coverCalibratorDevice.CoverState != CoverStatus.NotPresent) // Cover is present
                {
                    LogInfo("PostRunCheck", "Closing the cover...");
                    coverCalibratorDevice.CloseCover();

                    // if this is a Platform 7 or later device test CoverMoving
                    if (IsPlatform7OrLater) // Platform 7 or later interface ==> use CoverMoving
                    {
                        // Wait until the cover is no longer moving
                        LogCallToDriver("PostRunCheck", "About to call CoverMoving property repeatedly...");
                        WaitWhile("Cover closing", () => coverCalibratorDevice.CoverMoving, 500, 60);
                    }
                    else // Platform 6 interface ==> use CoverState
                    {
                        // Wait until the cover is no longer moving
                        LogCallToDriver("PostRunCheck", "About to call CoverState property repeatedly...");
                        WaitWhile("Cover closing", () => coverCalibratorDevice.CoverState == CoverStatus.Moving, 500, 60);
                    }
                    LogInfo("PostRunCheck", "Cover closed.");
                }
            }
            catch (Exception ex)
            {
                LogIssue("PostRunCheck", $"Exception closing cover: {ex.Message}");
                LogDebug("PostRunCheck", ex.ToString());
            }

            try
            {
                // Attempt to turn the calibrator off if present
                if (coverCalibratorDevice.CalibratorState != CalibratorStatus.NotPresent) // Calibrator is present
                {
                    LogInfo("PostRunCheck", "Turning off the calibrator...");
                    coverCalibratorDevice.CalibratorOff();

                    // if this is a Platform 7 or later device test CoverMoving
                    if (IsPlatform7OrLater) // Platform 7 or later interface ==> use CoverMoving
                    {
                        // Wait until the calibrator is off
                        LogCallToDriver("PostRunCheck", "About to call CalibratorChanging property repeatedly...");
                        WaitWhile("Turning calibrator off", () => coverCalibratorDevice.CalibratorChanging, 500, 60);
                    }
                    else // Platform 6 interface ==> use CoverState
                    {
                        // Wait until the calibrator is ready
                        LogCallToDriver("PostRunCheck", "About to call CoverState property repeatedly...");
                        WaitWhile("Turning calibrator off", () => coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady, 500, 60);
                    }
                    LogInfo("PostRunCheck", "Calibrator off.");
                }
            }
            catch (Exception ex)
            {
                LogIssue("PostRunCheck", $"Exception turning calibrator off: {ex.Message}");
                LogDebug("PostRunCheck", ex.ToString());
            }
        }

        public override void CheckPerformance()
        {
            SetTest("Performance");

            PerformanceTest(PerformanceProperty.CalibratorState, "CalibratorState");
            PerformanceTest(PerformanceProperty.CoverState, "CoverState");

            SetTest("");
            SetAction("");
            SetStatus("");
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

            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n:{ex}");
            }
        }

        #endregion

        #region Support Code

        private void TestOpenCover()
        {
            SetTest("OpenCover");

            try
            {
                Stopwatch sw = Stopwatch.StartNew();

                // Open the cover with the appropriate target time: Platform 6 = Extended, Platform 7 and later = Standard
                LogCallToDriver("OpenCover", "About to call OpenCover method");
                TimeMethod("OpenCover", coverCalibratorDevice.OpenCover, IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

                // Save the longest open time as the cover opening time
                if (sw.Elapsed.TotalSeconds > coverOpenTime)
                    coverOpenTime = sw.Elapsed.TotalSeconds;

                // Check whether the cover is moving after OpenCover returned
                if (CoverIsMoving("OpenCover")) //  Cover is moving after OpenCover returned ==> Asynchronous open
                {
                    canAsynchronousOpen = true;

                    // if this is a Platform 7 or later device test CoverMoving otherwise use CoverStatus
                    if (IsPlatform7OrLater) // Platform 7 or later interface ==> use CoverMoving
                    {
                        // Wait until the cover is no longer moving
                        LogCallToDriver("OpenCover", "About to call CoverMoving property repeatedly...");
                        WaitWhile("Opening", () => coverCalibratorDevice.CoverMoving, 500, 60);
                    }
                    else // Platform 6 device ==> use CoverState
                    {
                        // Wait until the cover is no longer moving
                        LogCallToDriver("OpenCover", "About to call CoverState property repeatedly...");
                        WaitWhile("Opening", () => coverCalibratorDevice.CoverState == CoverStatus.Moving, 500, 60);
                    }
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    // Check the operation outcome
                    LogCallToDriver("OpenCover", "About to call CoverState property");
                    if (coverCalibratorDevice.CoverState == CoverStatus.Open) // Cover reports the expected open state
                    {
                        coverOpenTime = sw.Elapsed.TotalSeconds;
                        LogOk("OpenCover", $"OpenCover was successful. The asynchronous open took {coverOpenTime:0.0} seconds");
                    }
                    else // Cover did not report the expected open state
                    {
                        LogIssue("OpenCover", $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The asynchronous open took {sw.Elapsed.TotalSeconds:0.0} seconds");
                    }
                }
                else // Cover is not moving after OpenCover returned ==> Synchronous open
                {
                    canAsynchronousOpen = false;

                    // Check whether the cover reports that it is open
                    LogCallToDriver("OpenCover", "About to call CoverState property");
                    if (coverCalibratorDevice.CoverState == CoverStatus.Open) // Cover reports that it opened OK
                        if (sw.Elapsed.TotalSeconds <= (IsPlatform7OrLater ? Globals.STANDARD_TARGET_RESPONSE_TIME : Globals.EXTENDED_TARGET_RESPONSE_TIME)) // Cover opened synchronously within the appropriate time
                            LogOk("OpenCover", $"Synchronous open: CoverState was 'Open' when the method completed and the call returned within the target time. The call took {sw.Elapsed.TotalSeconds:0.0} seconds");
                        else // Cover opened synchronously but took longer than the appropriate command time
                        {
                            LogIssue("OpenCover", $"Synchronous open: CoverState was 'Open' when the OpenCover method completed but the operation took {sw.Elapsed.TotalSeconds:0.0} seconds, " +
                                $"which is longer than the target response time: {(IsPlatform7OrLater ? Globals.STANDARD_TARGET_RESPONSE_TIME : Globals.EXTENDED_TARGET_RESPONSE_TIME):0.0} seconds");
                            LogInfo("OpenCover", $"Please implement this method asynchronously: return quickly from OpenCover, set CoverState to Moving and set CoverMoving to true (ICoverCalibratorV2 and later only), resetting these when movement is complete.");
                        }
                    else // Cover is no longer moving but reports some state other than Open
                        LogIssue("OpenCover", $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The synchronous open took {sw.Elapsed.TotalSeconds:0.0} seconds");
                }
            }
            catch (TimeoutException)
            {
                LogIssue("OpenCover", "The open cover operation timed out.");

            }
            catch (Exception ex)
            {
                if (coverState == CoverStatus.NotPresent)
                    HandleException("OpenCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                else
                    HandleException("OpenCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates the device has cover capability");
            }
        }

        private void TestCloseCover()
        {
            SetTest("CloseCover");

            try
            {
                Stopwatch sw = Stopwatch.StartNew();

                // Close the cover with the appropriate target time: Platform 6 = Extended, Platform 7 and later = Standard
                LogCallToDriver("CloseCover", "About to call CloseCover method");
                TimeMethod("CloseCover", coverCalibratorDevice.CloseCover, IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

                // Check whether the cover is moving after CloseCover returned
                LogCallToDriver("CloseCover", "About to call CoverState property");
                if (!CoverIsMoving("CloseCover")) // Cover is not moving after CloseCover returned ==> Synchronous close
                {
                    // Check whether the cover reports that it is closed
                    LogCallToDriver("CloseCover", "About to call CoverState property");
                    if (coverCalibratorDevice.CoverState == CoverStatus.Closed) // Cover reports that it closed OK
                    {
                        if (sw.Elapsed.TotalSeconds <= (IsPlatform7OrLater ? Globals.STANDARD_TARGET_RESPONSE_TIME : Globals.EXTENDED_TARGET_RESPONSE_TIME)) // Cover closed synchronously within the appropriate time
                            LogOk("CloseCover", $"Synchronous close: CoverState was 'Closed' when the method completed: the call took {sw.Elapsed.TotalSeconds:0.0} seconds");
                        else // Cover closed synchronously but took longer than the appropriate command time
                        {
                            LogIssue("CloseCover", $"Synchronous close: CoverState was 'Closed' when the CloseCover method completed but the operation took {sw.Elapsed.TotalSeconds:0.0} seconds, which is longer than the target response time: {(IsPlatform7OrLater ? Globals.EXTENDED_TARGET_RESPONSE_TIME : Globals.EXTENDED_TARGET_RESPONSE_TIME):0.0} seconds");
                            LogInfo("CloseCover", $"Please implement this method asynchronously: return quickly from CloseCover, set CoverState to Moving and set CoverMoving to true (ICoverCalibratorV2 and later only), resetting these when movement is complete.");
                        }
                    }
                    else// Cover is no longer moving but reports some state other than Closed
                        LogIssue("CloseCover", $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The synchronous close took {sw.Elapsed.TotalSeconds:0.0} seconds");
                }
                else  // Cover is moving after CloseCover returned ==> Asynchronous close
                {
                    // if this is a Platform 7 or later device test CoverMoving
                    if (IsPlatform7OrLater) // Platform 7 or later interface ==> use CoverMoving
                    {
                        // Wait until the cover is no longer moving
                        LogCallToDriver("CloseCover", "About to call CoverMoving property repeatedly...");
                        WaitWhile("CloseCover", () => coverCalibratorDevice.CoverMoving, 500, 60);
                    }
                    else
                    {
                        // Wait until the cover is no longer moving
                        LogCallToDriver("CloseCover", "About to call CoverState property repeatedly...");
                        WaitWhile("Closing", () => coverCalibratorDevice.CoverState == CoverStatus.Moving, 500, 60);
                    }
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    // Check the operation outcome
                    LogCallToDriver("CloseCover", "About to call CoverState property");
                    if (coverCalibratorDevice.CoverState == CoverStatus.Closed) // Cover reports the expected closed state
                    {
                        LogOk("CloseCover", $"CloseCover was successful. The asynchronous close took {sw.Elapsed.TotalSeconds:0.0} seconds");
                    }
                    else // Cover did not report the expected closed state
                    {
                        LogIssue("CloseCover", $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The asynchronous close took {sw.Elapsed.TotalSeconds:0.0} seconds");
                    }
                }
            }
            catch (Exception ex)
            {
                if (coverState == CoverStatus.NotPresent)
                    HandleException("CloseCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                else
                    HandleException("CloseCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates the device has cover capability");
            }
        }

        private void TestHaltCover()
        {
            SetTest("HaltCover");

            // Check whether a cover is implemented
            if (!(coverState == CoverStatus.NotPresent)) // Cover is implemented
            {
                // Check whether the cover opens asynchronously
                if (canAsynchronousOpen) // Does open asynchronously
                {
                    // Test whether the cover opened OK or malfunctioned
                    if (coverOpenTime > 0.0) // Cover opened OK
                    {
                        // Initiate a cover open first
                        LogCallToDriver("HaltCover", "About to call OpenCover method");
                        SetAction("Opening cover so that it can be halted");
                        coverCalibratorDevice.OpenCover();

                        // Wait for half of the expected cover open time
                        WaitFor((int)(coverOpenTime * 1000.0 / 2.0));

                        // Confirm that the cover is still moving
                        if (coverCalibratorDevice.CoverState == CoverStatus.Moving) // Cover is still moving
                        {
                            // Issue a halt command
                            try
                            {
                                SetAction("Halting cover");
                                SetStatus("Waiting for Halt to complete");

                                LogCallToDriver("HaltCover", "About to call HaltCover method");
                                TimeMethod("HaltCover", coverCalibratorDevice.HaltCover, TargetTime.Standard);
                                SetStatus("HaltCover command completed");

                                // Confirm that the cover is no longer moving
                                if (!CoverIsMoving("HaltCover")) // The cover is now stationary
                                    LogOk("HaltCover", "Cover is no longer moving after issuing the HaltCover command");
                                else // The cover is still moving
                                {
                                    LogIssue("HaltCover", "The cover was still moving after return from the HaltCover command.");
                                    LogInfo("HaltCover", "HaltCover() is expected to be a short-lived, synchronous, method that quickly stops movement.");
                                }
                            }
                            catch (Exception ex)
                            {
                                HandleException("HaltCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates that the device has cover capability");
                            }
                        }
                        else // The cover should have been moving after just half of the expected open time, but was not
                            LogIssue("HaltCover", "Cover should have been moving after waiting for half of the previous open time, but it was not. Test abandoned");
                    }
                    else // The cover did not open OK so skip the test
                        LogIssue("HaltCover", $"HaltCover tests skipped because either the cover could not be opened successfully.");
                }
                else // The cover opens synchronously so make sure that HaltCover is not implemented
                {
                    try
                    {
                        // Since the cover opens synchronously the HaltCover method should return a MethodNotImplementedException
                        LogCallToDriver("HaltCover", "About to call HaltCover method");
                        coverCalibratorDevice.HaltCover();
                        LogIssue("HaltCover", "The cover operates synchronously but did not throw a MethodNotImplementedException in response to the HaltCover command");
                    }
                    catch (Exception ex)
                    {
                        if (coverState == CoverStatus.NotPresent)
                            HandleException("HaltCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                        else
                            HandleException("HaltCover", MemberType.Method, Required.Optional, ex, "");
                    }
                }
            }
            else // The cover is not implemented so make sure it throws a MethodNotImplementedException
            {
                try
                {
                    LogCallToDriver("HaltCover", "About to call HaltCover method");
                    coverCalibratorDevice.HaltCover();

                    // Should never get here because a MethodNotImplementedException should be thrown when the cover is not implemented...
                    LogIssue("HaltCover", "CoverStatus is 'NotPresent' but HaltCover did not throw a MethodNotImplementedException");
                }
                catch (Exception ex)
                {
                    HandleException("HaltCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                }
            }
        }

        /// <summary>
        /// Determine whether the cover is moving, ensuring that the CoverMoving and CoverState properties agree in ICoverCalibratorV2 and later interfaces 
        /// </summary>
        /// <param name="operation"></param>
        /// <returns>True if the cover is moving</returns>
        private bool CoverIsMoving(string operation)
        {
            // Get the cover state
            LogCallToDriver(operation, "About to call CoverState property");
            CoverStatus status = coverCalibratorDevice.CoverState;

            // Check whether this is an ICoverCalibratorV2 or later interface
            if (IsPlatform7OrLater) // Fond an ICoverCalibratorV2 or later interface
            {
                // Get the CoverMoving value
                LogCallToDriver(operation, "About to call CoverMoving property");
                bool coverMoving = coverCalibratorDevice.CoverMoving;
                LogDebug("CoverIsMoving", $"Received cover state: {status} and cover moving: {coverMoving}");

                // Test whether the values match
                if (coverMoving & (status == CoverStatus.Moving)) // Both agree that the cover is moving
                    return true;
                else if (!coverMoving & (status != CoverStatus.Moving)) // Both agree that the cover is not moving
                    return false;
                else // The two properties disagree
                {
                    // Log an issue and return the coverMoving value, which may or may not be correct!
                    LogIssue(operation, $"CoverMoving and CoverStatus do not match: CoverMoving: {coverMoving}, CoverState: {status}");
                    return coverMoving;
                }
            }
            else // Found an ICoverCalibratorV1 device
            {
                // Return a boolean derived from CoverStatus response
                LogDebug("CoverIsMoving", $"Received cover state: {status}");
                return status == CoverStatus.Moving;
            }
        }

        /// <summary>
        /// Determine whether the calibrator is changing, ensuring that the CalibratorChanging and CalibratorState properties agree in ICoverCalibratorV2 and later interfaces 
        /// </summary>
        /// <param name="operation"></param>
        /// <returns>True if the calibrator is ready</returns>
        private bool CalibratorReady(string operation)
        {
            // Get the calibrator state
            LogCallToDriver(operation, "About to call CalibratorState property");
            CalibratorStatus status = coverCalibratorDevice.CalibratorState;

            // Check whether this is an ICoverCalibratorV2 or later interface
            if (IsPlatform7OrLater) // Fond an ICoverCalibratorV2 or later interface
            {
                // Get the CalibratorChanging value
                LogCallToDriver(operation, "About to call CalibratorChanging property");
                bool calibratorChanging = coverCalibratorDevice.CalibratorChanging;

                // Test whether the values match
                if (calibratorChanging & (status == CalibratorStatus.NotReady)) // Both agree that the calibrator is not ready
                    return false;
                else if (!calibratorChanging & (status != CalibratorStatus.NotReady)) // Both agree that the calibrator is ready
                    return true;
                else // The two properties disagree
                {
                    // Log an issue and return the calibratorChanging value, which may or may not be correct!
                    LogIssue(operation, $"CalibratorChanging and CalibratorState do not match: CalibratorChanging: {calibratorChanging}, CalibratorState: {status}");
                    return !calibratorChanging;
                }
            }
            else // Found an ICoverCalibratorV1 device
            {
                // Return a boolean derived from CalibratorState response
                return status != CalibratorStatus.NotReady;
            }
        }

        private void TestCalibratorOn(int requestedBrightness)
        {
            int returnedBrightness;
            Stopwatch sw = new();

            // Check whether the calibrator is implemented
            if (!(calibratorState == CalibratorStatus.NotPresent)) // Calibrator is implemented
            {
                try
                {
                    sw.Start();
                    SetAction($"Setting calibrator to brightness {requestedBrightness}");

                    LogCallToDriver("CalibratorOn", $"About to call CalibratorOn method with brightness: {requestedBrightness}");
                    TimeMethodOneParam<int>("CalibratorOn", coverCalibratorDevice.CalibratorOn, requestedBrightness, IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

                    // Check whether the call was synchronous or asynchronous
                    if (CalibratorReady("CalibratorOn")) // Synchronous call
                    {
                        // Get the outcome state
                        LogCallToDriver("CalibratorOn", "About to call CalibratorState property");
                        calibratorState = coverCalibratorDevice.CalibratorState;

                        // Check the outcome
                        if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness)) // An invalid value should have been thrown so this code should never be reached
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.");
                        else if (calibratorState == CalibratorStatus.Ready) // A valid brightness was set and a Ready status was returned
                        {
                            // Check whether the synchronous operation completed within the expected response time
                            if (sw.Elapsed.TotalSeconds <= (IsPlatform7OrLater ? Globals.STANDARD_TARGET_RESPONSE_TIME : Globals.EXTENDED_TARGET_RESPONSE_TIME)) // Completed within the expected response time
                                LogOk("CalibratorOn", $"Synchronous operation: CalibratorState was 'Ready' when the method completed: the call took {sw.Elapsed.TotalSeconds:0.0} seconds");
                            else // Did not complete within the expected response time
                            {
                                LogIssue("CalibratorOn", $"Synchronous operation: CalibratorState was 'Ready' when the CalibratorOn method completed but the operation took {sw.Elapsed.TotalSeconds:0.0} seconds, which is longer than the target response time: {Globals.STANDARD_TARGET_RESPONSE_TIME:0.0} seconds");
                                LogInfo("CalibratorOn", $"Please implement this method asynchronously: return quickly from CalibratorOn, set CalibratorChanging to true and set CalibratorState to NotReady resetting these when the change is complete.");
                            }

                            // if this is a Platform 7 or later device test CalibratorChanging to ensure its state is false
                            if (IsPlatform7OrLater)
                            {
                                // Confirm that CalibratorChanging is False
                                LogCallToDriver("CalibratorOn", $"About to call CalibratorChanging property.");
                                if (!coverCalibratorDevice.CalibratorChanging) // CalibratorChanging is False as expected
                                    LogOk("CalibratorOn", $"The CalibratorChanging property returned false when CalibratorState returned CalibratorStatus.Ready.");
                                else // CalibratorChanging is True, which does not match CalibratorStatus.Ready
                                    LogIssue("CalibratorOn", $"The CalibratorChanging property returned true while CalibratorState returned CalibratorStatus.Ready.");
                            }

                            // Confirm that the brightness value is what was set
                            LogCallToDriver("CalibratorOn", $"About to call Brightness property.");
                            returnedBrightness = coverCalibratorDevice.Brightness;

                            // Check the returned brightness value
                            if (returnedBrightness == requestedBrightness) // Returned Brightness value is what was set
                                LogOk("CalibratorOn", $"The Brightness property does return the value that was set");
                            else // Returned Brightness value is not what was set
                                LogIssue("CalibratorOn", $"The Brightness property value: {returnedBrightness} does not match the value that was set: {requestedBrightness}");
                        }
                        else// A valid brightness was set but a status other than Ready was returned
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CalibratorState was '{calibratorState}' instead of 'Ready'. The synchronous operation took {sw.Elapsed.TotalSeconds:0.0} seconds");
                    }
                    else // Asynchronous call
                    {
                        // Wait for the brightness change to complete using the appropriate completion variable depending on interface version
                        if (IsPlatform7OrLater) // The device is Platform 7 or later
                        {
                            // Wait until the cover is no longer moving
                            LogCallToDriver("CalibratorOn", "About to call the CalibratorChanging property multiple times");
                            WaitWhile("Warming up", () => coverCalibratorDevice.CalibratorChanging, 500, 60);
                        }
                        else // The device is Platform 6
                        {
                            // Wait until the cover is no longer moving
                            LogCallToDriver("CalibratorOn", "About to call the CalibratorState property multiple times");
                            WaitWhile("Warming up", () => coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady, 500, 60);
                        }

                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Check the outcome
                        LogCallToDriver("CalibratorOn", "About to call CalibratorState property");
                        calibratorState = coverCalibratorDevice.CalibratorState;
                        if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness)) // An invalid value should have been thrown so this code should never be reached
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.");
                        else if (calibratorState == CalibratorStatus.Ready) // The calibrator operated as expected
                            LogOk("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} was successful. The asynchronous operation took {sw.Elapsed.TotalSeconds:0.0} seconds");
                        else // The calibrator state indicates an unsuccessful change
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CalibratorState was '{calibratorState}' instead of 'Ready'. The asynchronous operation took {sw.Elapsed.TotalSeconds:0.0} seconds");
                    }
                }
                catch (Exception ex)
                {
                    if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness))
                    {
                        if (IsInvalidValueException("CalibratorOn", ex))
                            LogOk("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} threw an InvalidValueException as expected");
                        else
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} threw an {ex.GetType().Name} exception an ASCOM.InvalidValueException was expected");
                    }
                    else
                    {
                        HandleException("CalibratorOn", MemberType.Method, Required.MustBeImplemented, ex, "CalibratorStatus indicates that the device is a calibrator");
                    }
                }
            }
            else // Calibrator is not implemented
            {
                try
                {
                    LogCallToDriver("CalibratorOn", $"About to call CalibratorOn method with brightness: {requestedBrightness}");
                    coverCalibratorDevice.CalibratorOn(requestedBrightness);

                    // Should never get here because a MethodNotImplementedException should have been thrown...
                    LogIssue("CalibratorOn", $"CalibratorStatus is 'NotPresent'but CalibratorOn did not throw a MethodNotImplementedException.");
                }
                catch (Exception ex)
                {
                    HandleException("CalibratorOn", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                }
            }
        }

        private void TestCalibratorOff()
        {
            Stopwatch sw = new();

            // Check whether the calibrator is implemented
            if (!(calibratorState == CalibratorStatus.NotPresent)) // Calibrator is implemented
            {
                try
                {
                    sw.Start();
                    SetAction($"Turning calibrator off");

                    LogCallToDriver("CalibratorOff", $"About to call CalibratorOff method");
                    TimeMethod("CalibratorOff", coverCalibratorDevice.CalibratorOff, IsPlatform7OrLater ? TargetTime.Standard : TargetTime.Extended);

                    // Check whether the call was synchronous or asynchronous
                    if (CalibratorReady("CalibratorOff")) // Synchronous call
                    {
                        // Check the outcome
                        LogCallToDriver("CalibratorOff", "About to call CalibratorState property");
                        if (calibratorState == CalibratorStatus.Ready) // A valid brightness was set and a Ready status was returned
                        {
                            // Check whether the synchronous operation completed within the expected response time
                            if (sw.Elapsed.TotalSeconds <= (IsPlatform7OrLater ? Globals.STANDARD_TARGET_RESPONSE_TIME : Globals.EXTENDED_TARGET_RESPONSE_TIME)) // Completed within the expected response time
                                LogOk("CalibratorOff", $"Synchronous operation: CalibratorState was 'Ready' when the method completed: the call took {sw.Elapsed.TotalSeconds:0.0} seconds");
                            else // Did not complete within the expected response time
                            {
                                LogIssue("CalibratorOff", $"Synchronous operation: CalibratorState was 'Ready' when the CalibratorOff method completed but the operation took {sw.Elapsed.TotalSeconds:0.0} seconds, which is longer than the target response time: {Globals.STANDARD_TARGET_RESPONSE_TIME:0.0} seconds");
                                LogInfo("CalibratorOff", $"Please implement this method asynchronously: return quickly from CalibratorOff, set CalibratorChanging to true and set CalibratorState to NotReady resetting these when the change is complete.");
                            }

                            // if this is a Platform 7 or later device test CalibratorChanging to ensure its state is false
                            if (IsPlatform7OrLater)
                            {
                                // Confirm that CalibratorChanging is False
                                LogCallToDriver("CalibratorOff", $"About to call CalibratorChanging property.");
                                if (!coverCalibratorDevice.CalibratorChanging) // CalibratorChanging is False as expected
                                    LogOk("CalibratorOff", $"The CalibratorChanging property returned false when CalibratorState returned CalibratorStatus.Ready.");
                                else // CalibratorChanging is True, which does not match CalibratorStatus.Ready
                                    LogIssue("CalibratorOff", $"The CalibratorChanging property returned true while CalibratorState returned CalibratorStatus.Ready.");
                            }
                        }
                        else// The operation completed but a status other than Ready was returned
                            LogIssue("CalibratorOff", $"CalibratorOff was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Ready'. The synchronous operation took {sw.Elapsed.TotalSeconds:0.0} seconds");
                    }
                    else // Asynchronous call
                    {
                        // Wait for the brightness change to complete using the appropriate completion variable depending on interface version
                        if (IsPlatform7OrLater) // The device is Platform 7 or later
                        {
                            // Wait until the cover is no longer moving
                            LogCallToDriver("CalibratorOff", "About to call the CalibratorChanging property multiple times");
                            WaitWhile("Switching off", () => coverCalibratorDevice.CalibratorChanging, 500, 60);
                        }
                        else // The device is Platform 6
                        {
                            // Wait until the cover is no longer moving
                            LogCallToDriver("CalibratorOff", "About to call the CalibratorState property multiple times");
                            WaitWhile("Switching off", () => coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady, 500, 60);
                        }

                        if (cancellationToken.IsCancellationRequested)
                            return;

                        // Check the outcome
                        LogCallToDriver("CalibratorOff", "About to call CalibratorState property");
                        calibratorState = coverCalibratorDevice.CalibratorState;
                        if (calibratorState == CalibratorStatus.Off) // The calibrator operated as expected
                            LogOk("CalibratorOff", $"CalibratorOff was successful. The asynchronous operation took {sw.Elapsed.TotalSeconds:0.0} seconds");
                        else // The calibrator state indicates an unsuccessful change
                            LogIssue("CalibratorOff", $"CalibratorOff was unsuccessful - the returned CalibratorState was '{calibratorState}' instead of 'Off'. The asynchronous operation took {sw.Elapsed.TotalSeconds:0.0} seconds");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("CalibratorOff", MemberType.Method, Required.MustBeImplemented, ex, "CalibratorStatus indicates that the device is a calibrator");
                }
            }
            else // Calibrator is not implemented
            {
                try
                {
                    LogCallToDriver("CalibratorOff", $"About to call CalibratorOff method");
                    coverCalibratorDevice.CalibratorOff();

                    // Should never get here because a MethodNotImplementedException should have been thrown...
                    LogIssue("CalibratorOff", $"CalibratorStatus is 'NotPresent'but CalibratorOff did not throw a MethodNotImplementedException.");
                }
                catch (Exception ex)
                {
                    HandleException("CalibratorOff", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                }
            }
        }

        private bool RequiredPropertiesTest(RequiredProperty propertyToTest, string propertyName)
        {
            bool testWasSuccessful;

            try
            {
                testWasSuccessful = true; // Assume success
                switch (propertyToTest)
                {
                    case RequiredProperty.CalibratorState:
                        LogCallToDriver("CalibratorState", "About to call CalibratorState property");
                        TimeFunc<CalibratorStatus>(propertyName, () => calibratorState = coverCalibratorDevice.CalibratorState, TargetTime.Fast);
                        LogOk(propertyName, calibratorState.ToString());
                        break;

                    case RequiredProperty.CoverState:
                        LogCallToDriver("CoverState", "About to call CoverState property");
                        TimeFunc<CoverStatus>(propertyName, () => coverState = coverCalibratorDevice.CoverState, TargetTime.Fast);
                        LogOk(propertyName, coverState.ToString());
                        break;

                    case RequiredProperty.CalibratorChanging:
                        LogCallToDriver("CalibratorChanging", "About to call CalibratorChanging property");
                        TimeFunc<bool>(propertyName, () => calibratorChanging = coverCalibratorDevice.CalibratorChanging, TargetTime.Fast);
                        LogOk(propertyName, calibratorChanging.ToString());
                        break;

                    case RequiredProperty.CoverMoving:
                        LogCallToDriver("CoverMoving", "About to call CoverMoving property");
                        TimeFunc<bool>(propertyName, () => coverMoving = coverCalibratorDevice.CoverMoving, TargetTime.Fast);
                        LogOk(propertyName, coverMoving.ToString());
                        break;

                    default:
                        testWasSuccessful = false; // Flag that an issue occurred
                        LogIssue(propertyName, $"RequiredPropertiesTest: Unknown test type {propertyToTest}");
                        break;
                }
            }
            catch (Exception ex)
            {
                testWasSuccessful = false; // Flag that an issue occurred
                HandleException(propertyName, MemberType.Property, Required.Mandatory, ex, "");
            }

            return testWasSuccessful;
        }

        private void PerformanceTest(PerformanceProperty propertyToTest, string propertyName)
        {
            DateTime startTime;
            double loopCount, lastElapsedTime, elapsedTime, loopRate;
            CalibratorStatus testCalibratorState;
            CoverStatus testCoverState;

            SetAction(propertyName);
            try
            {
                startTime = DateTime.Now;
                loopCount = 0.0;
                lastElapsedTime = 0.0;
                do
                {
                    loopCount += 1.0;
                    switch (propertyToTest)
                    {
                        case PerformanceProperty.CalibratorState:
                            {
                                testCalibratorState = coverCalibratorDevice.CalibratorState;
                                break;
                            }

                        case PerformanceProperty.CoverState:
                            {
                                testCoverState = coverCalibratorDevice.CoverState;
                                break;
                            }

                        default:
                            {
                                LogIssue(propertyName, $"PerformanceTest: Unknown test type {propertyToTest}");
                                break;
                            }
                    }

                    elapsedTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                    if (elapsedTime > lastElapsedTime + 1.0)
                    {
                        SetStatus($"{loopCount} transactions in {elapsedTime:0} seconds");
                        lastElapsedTime = elapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (elapsedTime <= PERF_LOOP_TIME);

                loopRate = loopCount / elapsedTime;
                switch (loopRate)
                {
                    case object _ when loopRate > 10.0:
                        {
                            LogInfo(propertyName, $"Transaction rate: {loopRate:0.0} per second");
                            break;
                        }

                    case object _ when 2.0 <= loopRate && loopRate <= 10.0:
                        {
                            LogOk(propertyName, $"Transaction rate: {loopRate:0.0} per second");
                            break;
                        }

                    case object _ when 1.0 <= loopRate && loopRate <= 2.0:
                        {
                            LogInfo(propertyName, $"Transaction rate: {loopRate:0.0} per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(propertyName, $"Transaction rate: {loopRate:0.0} per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(propertyName, $"Unable to complete test: {ex}");
            }
        }

        #endregion

    }
}

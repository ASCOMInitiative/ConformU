using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
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
            CoverState
        }

        private enum PerformanceProperty
        {
            CalibratorState,
            CoverState
        }

        private CoverStatus coverState;
        private CalibratorStatus calibratorState;
        private bool canAsynchronousOpen;
        private double asynchronousOpenTime;
        private double asynchronousCloseTime;
        private int maxBrightness;
        private bool calibratorStateOk;
        private bool coverStateOk;
        private bool brightnessOk;
        private bool maxBrightnessOk;

        // Helper variables
        private ICoverCalibratorV1 coverCalibratorDevice;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose
        public CoverCalibratorTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, true, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
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
                    coverCalibratorDevice?.Dispose();
                    coverCalibratorDevice = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        public new void CheckInitialise()
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
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040405;
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

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                baseClassDevice = coverCalibratorDevice; // Assign the driver to the base class

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

        public override void PreConnectChecks()
        {
        }

        public override bool Connected
        {
            get
            {
                LogCallToDriver("ConformanceCheck", "About to get Connected");
                return coverCalibratorDevice.Connected;
            }
            set
            {
                LogCallToDriver("ConformanceCheck", "About to set Connected");
                SetTest("Connected");
                SetAction("Waiting for Connected to become 'true'");
                coverCalibratorDevice.Connected = value;
                ResetTestActionStatus();

                // Make sure that the value set is reflected in Connected GET
                bool connectedState = Connected;
                if (connectedState != value)
                {
                    throw new ASCOM.InvalidOperationException($"Connected was set to {value} but Connected Get returned {connectedState}.");
                }
            }
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(coverCalibratorDevice, DeviceTypes.CoverCalibrator);
        }

        public override void CheckProperties()
        {
            int brightness;

            if (cancellationToken.IsCancellationRequested)
                return;

            calibratorStateOk = RequiredPropertiesTest(RequiredProperty.CalibratorState, "CalibratorState");
            coverStateOk = RequiredPropertiesTest(RequiredProperty.CoverState, "CoverState");

            if (calibratorStateOk)
            {
                maxBrightnessOk = false; // Assume a bad value
                try
                {
                    maxBrightness = 0; // Initialise to a 'bad' value
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("MaxBrightness", "About to call MaxBrightness property");
                    maxBrightness = coverCalibratorDevice.MaxBrightness;

                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent))
                    {
                        if (maxBrightness >= 1)
                        {
                            LogOK("MaxBrightness", maxBrightness.ToString());
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

            if (calibratorStateOk)
            {
                brightnessOk = false; // Assume a bad value

                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("Brightness", "About to call Brightness property");
                    brightness = coverCalibratorDevice.Brightness;
                    brightnessOk = true;

                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent))
                    {
                        if (brightness >= 0)
                        {
                            if (maxBrightnessOk)
                            {
                                if (brightness <= maxBrightness)
                                    LogOK("Brightness", maxBrightness.ToString());
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
                    if (coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent)
                        HandleException("Brightness", MemberType.Property, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                    else
                        HandleException("Brightness", MemberType.Property, Required.MustBeImplemented, ex, "CalibratorStatus indicates that he device is a calibrator");
                }
            }
            else
                LogIssue("Brightness", $"Test skipped because CalibratorState returned an exception");

            if (coverStateOk & calibratorStateOk)
            {
                if (coverState == CoverStatus.NotPresent & calibratorState == CalibratorStatus.NotPresent)
                    LogIssue("DeviceCapabilities", "Both CoverStatus and CalibratorStatus are set to 'NotPresent' - this driver won't do a lot!");
            }
        }

        public override void CheckMethods()
        {
            DateTime startTime;
            ClearStatus();  // Clear status messages

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test OpenCover
            SetTest("OpenCover");
            if (coverStateOk)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("OpenCover", "About to call OpenCover method");
                    startTime = DateTime.Now;

                    coverCalibratorDevice.OpenCover();
                    if (!(coverState == CoverStatus.NotPresent))
                    {
                        if (!(coverCalibratorDevice.CoverState == CoverStatus.Moving)) // Synchronous open
                        {
                            canAsynchronousOpen = false;
                            if (coverCalibratorDevice.CoverState == CoverStatus.Open)
                                LogOK("OpenCover", $"OpenCover was successful. The synchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                            else
                                LogIssue("OpenCover", $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The synchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                        }
                        else // Asynchronous open
                        {
                            canAsynchronousOpen = true;
                            asynchronousOpenTime = 0.0;

                            // Wait until the cover is no longer moving
                            WaitWhile("Opening", () => { return coverCalibratorDevice.CoverState == CoverStatus.Moving; }, 500, 60);
                            if (cancellationToken.IsCancellationRequested)
                                return;

                            if (coverCalibratorDevice.CoverState == CoverStatus.Open)
                            {
                                asynchronousOpenTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                                LogOK("OpenCover", $"OpenCover was successful. The asynchronous open took {asynchronousOpenTime:0.0} seconds");
                            }
                            else
                                LogIssue("OpenCover", $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The asynchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                        }
                    }
                    else
                        LogIssue("OpenCover", $"CoverStatus is 'NotPresent' but OpenCover did not throw a MethodNotImplementedException.");
                }
                catch (Exception ex)
                {
                    if (coverState == CoverStatus.NotPresent)
                        HandleException("OpenCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                    else
                        HandleException("OpenCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates the device has cover capability");
                }
            }
            else
                LogIssue("OpenCover", $"Test skipped because CoverState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test CloseCover
            SetTest("CloseCover");
            if (coverStateOk)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("CloseCover", "About to call CloseCover method");
                    startTime = DateTime.Now;
                    asynchronousCloseTime = 0.0;

                    coverCalibratorDevice.CloseCover();
                    if (!(coverState == CoverStatus.NotPresent)) // Synchronous close
                    {
                        if (!(coverCalibratorDevice.CoverState == CoverStatus.Moving))
                        {
                            canAsynchronousOpen = false;
                            if (coverCalibratorDevice.CoverState == CoverStatus.Closed)
                                LogOK("CloseCover", $"CloseCover was successful. The synchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                            else
                                LogIssue("CloseCover", $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The synchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                        }
                        else // Asynchronous close
                        {
                            canAsynchronousOpen = true;
                            // Wait until the cover is no longer moving
                            WaitWhile("Closing", () => { return coverCalibratorDevice.CoverState == CoverStatus.Moving; }, 500, 60);
                            if (cancellationToken.IsCancellationRequested)
                                return;

                            if (coverCalibratorDevice.CoverState == CoverStatus.Closed)
                            {
                                asynchronousCloseTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                                LogOK("CloseCover", $"CloseCover was successful. The asynchronous close took {asynchronousCloseTime:0.0} seconds");
                            }
                            else
                                LogIssue("CloseCover", $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The asynchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                        }
                    }
                    else
                        LogIssue("CloseCover", $"CoverStatus is 'NotPresent' but CloseCover did not throw a MethodNotImplementedException.");
                }
                catch (Exception ex)
                {
                    if (coverState == CoverStatus.NotPresent)
                        HandleException("CloseCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                    else
                        HandleException("CloseCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates the device has cover capability");
                }
            }
            else
                LogIssue("CloseCover", $"Test skipped because CoverState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test HaltCover
            SetTest("HaltCover");
            if (coverStateOk)
            {
                if (!(coverState == CoverStatus.NotPresent)) // Cover present
                {
                    if (canAsynchronousOpen) // Can open asynchronously
                    {
                        if ((asynchronousOpenTime > 0.0) & (asynchronousCloseTime > 0.0))
                        {

                            // Initiate a cover open first
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("HaltCover", "About to call OpenCover method");
                            SetAction("Opening cover so that it can be halted");
                            coverCalibratorDevice.OpenCover();

                            // Wait for half of the expected cover open time
                            WaitFor((int)(asynchronousOpenTime * 1000.0 / 2.0));

                            // Confirm that he cover is still moving
                            if (coverCalibratorDevice.CoverState == CoverStatus.Moving)
                            {
                                try
                                {
                                    // Issue a halt command
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("HaltCover", "About to call HaltCover method");
                                    SetAction("Halting cover");
                                    SetStatus("Waiting for Halt to complete");
                                    coverCalibratorDevice.HaltCover();
                                    SetStatus("Halt completed");

                                    // Confirm that the cover is no longer moving
                                    if (!(coverCalibratorDevice.CoverState == CoverStatus.Moving))
                                        LogOK("HaltCover", "Cover is no longer moving after issuing the HaltCover command");
                                    else
                                        LogIssue("HaltCover", "Cover is still moving after issuing the HaltCover command");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("HaltCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates that the device has cover capability");
                                }
                            }
                            else
                                LogIssue("HaltCover", "Cover should have been moving after waiting for half of the previous open time, but it was not. Test abandoned");
                        }
                        else
                            LogIssue("HaltCover", $"HaltCover tests skipped because either the cover could not be opened or closed successfully.");
                    }
                    else
                        try
                        {
                            // Since the cover opens synchronously the HaltCover method should return a MethodNotImplementedException
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("HaltCover", "About to call HaltCover method");
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
                else // Cover not present
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("HaltCover", "About to call HaltCover method");
                        coverCalibratorDevice.HaltCover();
                        // Should never get here...
                        LogIssue("HaltCover", "CoverStatus is 'NotPresent' but HaltCover did not throw a MethodNotImplementedException");
                    }
                    catch (Exception ex)
                    {
                        HandleException("HaltCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                    }
            }
            else
                LogIssue("HaltCover", $"Test skipped because CoverState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test CalibratorOn
            SetTest("CalibratorOn");
            if (calibratorStateOk)
            {
                if (!(calibratorState == CalibratorStatus.NotPresent))
                {
                    if (maxBrightnessOk & brightnessOk)
                    {
                        TestCalibratorOn(-1); // Test for invalid value -1
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        TestCalibratorOn(0); // Test for zero brightness
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        switch (maxBrightness)
                        {
                            case 1 // Simple on/ off device
                           :
                                {
                                    TestCalibratorOn(1);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    break;
                                }

                            case 2 // Two brightness level device
                     :
                                {
                                    TestCalibratorOn(1);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;

                                    TestCalibratorOn(2);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    break;
                                }

                            case 3 // Three brightness level device
                     :
                                {
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
                                }

                            case 4 // Four brightness level device
                     :
                                {
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
                                }

                            default:
                                {
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
                        }

                        if (maxBrightness < int.MaxValue)
                        {
                            TestCalibratorOn(maxBrightness + 1); // Test for invalid value of MaxBrightness + 1
                            if (cancellationToken.IsCancellationRequested)
                                return;
                        }
                        else
                            LogInfo("CalibratorOn", $"Test of a high invalid brightness value skipped because MaxBrightness is set to the largest positive integer value.");
                    }
                    else
                        LogIssue("CalibratorOn", $"Brightness tests skipped because one of the Brightness or MaxBrightness properties returned an invalid value or threw an exception.");
                }
                else
                {
                    TestCalibratorOn(1);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
                LogIssue("CalibratorOn", $"Brightness tests skipped because the CoverState property returned an invalid value or threw an exception.");

            // Test CalibratorOff
            SetTest("CalibratorOff");
            if (calibratorStateOk)
            {
                if (!(calibratorState == CalibratorStatus.NotPresent))
                {
                    try
                    {
                        startTime = DateTime.Now;

                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("CalibratorOff", "About to call CalibratorOff method");
                        coverCalibratorDevice.CalibratorOff();

                        if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady)) // Synchronous call
                        {
                            if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Off)
                            {
                                LogOK("CalibratorOff", $"CalibratorOff was successful. The synchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");

                                // Confirm that Brightness returns to zero when calibrator is turned off
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("CalibratorOff", "About to call Brightness property");
                                if (coverCalibratorDevice.Brightness == 0)
                                    LogOK("CalibratorOff", $"Brightness is set to zero when the calibrator is turned off");
                                else
                                    LogIssue("CalibratorOff", $"Brightness is not set to zero when the calibrator is turned off");
                            }
                            else
                                LogIssue("CalibratorOff", $"CalibratorOff was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Off'. The synchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                        }
                        else // Asynchronous call
                        {
                            // Wait until the calibrator is off 
                            WaitWhile("Cooling down", () => { return coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady; }, 500, 60);
                            if (cancellationToken.IsCancellationRequested)
                                return;

                            if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Off)
                            {
                                LogOK("CalibratorOff", $"CalibratorOff was successful. The asynchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");

                                // Confirm that Brightness returns to zero when calibrator is turned off
                                if (settings.DisplayMethodCalls)
                                    LogTestAndMessage("CalibratorOff", "About to call Brightness property");
                                if (coverCalibratorDevice.Brightness == 0)
                                    LogOK("CalibratorOff", $"Brightness is set to zero when the calibrator is turned off");
                                else
                                    LogIssue("CalibratorOff", $"Brightness is not set to zero when the calibrator is turned off");
                            }
                            else
                                LogIssue("CalibratorOff", $"CalibratorOff was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Off'. The asynchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("CalibratorOff", MemberType.Method, Required.MustBeImplemented, ex, "CalibratorStatus indicates that the device is a calibrator");
                    }
                }
                else
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("CalibratorOff", "About to call CalibratorOff method");
                        coverCalibratorDevice.CalibratorOff();
                        // Should never get here...
                        LogIssue("CalibratorOff", $"CalibratorStatus is 'NotPresent'but CalibratorOff did not throw a MethodNotImplementedException.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("CalibratorOff", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                    }
            }
            else
                LogIssue("CalibratorOff", $"Test skipped because the CoverState property returned an invalid value or threw an exception.");

            ClearStatus();
        }

        private void TestCalibratorOn(int requestedBrightness)
        {
            int returnedBrightness;
            DateTime startTime;

            if (!(calibratorState == CalibratorStatus.NotPresent))
            {
                try
                {
                    startTime = DateTime.Now;

                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("CalibratorOn", $"About to call CalibratorOn method with brightness: {requestedBrightness}");
                    SetAction($"Setting calibrator to brightness {requestedBrightness}");
                    coverCalibratorDevice.CalibratorOn(requestedBrightness);

                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady)) // Synchronous call
                    {
                        if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness))
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.");
                        else if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Ready)
                        {
                            LogOK("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} was successful. The synchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");

                            // Confirm that the brightness value is what was set
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("CalibratorOn", $"About to call Brightness property.");
                            returnedBrightness = coverCalibratorDevice.Brightness;

                            if (returnedBrightness == requestedBrightness)
                                LogOK("CalibratorOn", $"The Brightness property does return the value that was set");
                            else
                                LogIssue("CalibratorOn", $"The Brightness property value: {returnedBrightness} does not match the value that was set: {requestedBrightness}");
                        }
                        else
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Ready'. The synchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                    }
                    else // Asynchronous call
                    {
                        // Wait until the cover is no longer moving
                        WaitWhile("Warming up", () => { return coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady; }, 500, 60);
                        if (cancellationToken.IsCancellationRequested)
                            return;

                        if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness))
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.");
                        else if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Ready)
                            LogOK("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} was successful. The asynchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                        else
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Ready'. The asynchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds:0.0} seconds");
                    }
                }
                catch (Exception ex)
                {
                    if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness))
                    {
                        if (IsInvalidValueException("CalibratorOn", ex))
                            LogOK("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} threw an InvalidValueException as expected");
                        else
                            LogIssue("CalibratorOn", $"CalibratorOn with brightness {requestedBrightness} threw an {ex.GetType().Name} exception an InvalidValueException was expected");
                    }
                    else
                        HandleException("CalibratorOn", MemberType.Method, Required.MustBeImplemented, ex, "CalibratorStatus indicates that the device is a calibrator");
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("CalibratorOn", $"About to call CalibratorOn method with brightness: {requestedBrightness}");
                    coverCalibratorDevice.CalibratorOn(requestedBrightness);
                    // Should never get here...
                    LogIssue("CalibratorOn", $"CalibratorStatus is 'NotPresent'but CalibratorOn did not throw a MethodNotImplementedException.");
                }
                catch (Exception ex)
                {
                    HandleException("CalibratorOn", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
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

        private bool RequiredPropertiesTest(RequiredProperty propertyToTest, string propertyName)
        {
            bool testWasSuccessful;

            try
            {
                testWasSuccessful = true; // Assume success
                switch (propertyToTest)
                {
                    case RequiredProperty.CalibratorState:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("CalibratorState", "About to call CalibratorState property");
                            calibratorState = coverCalibratorDevice.CalibratorState;
                            LogOK(propertyName, calibratorState.ToString());
                            break;
                        }

                    case RequiredProperty.CoverState:
                        {
                            if (settings.DisplayMethodCalls)
                                LogTestAndMessage("CoverState", "About to call CoverState property");
                            coverState = coverCalibratorDevice.CoverState;
                            LogOK(propertyName, coverState.ToString());
                            break;
                        }

                    default:
                        {
                            testWasSuccessful = false; // Flag that an issue occurred
                            LogIssue(propertyName, "RequiredPropertiesTest: Unknown test type " + propertyToTest.ToString());
                            break;
                        }
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
                                LogIssue(propertyName, "PerformanceTest: Unknown test type " + propertyToTest.ToString());
                                break;
                            }
                    }

                    elapsedTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                    if (elapsedTime > lastElapsedTime + 1.0)
                    {
                        SetStatus(loopCount + " transactions in " + elapsedTime.ToString("0") + " seconds");
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
                            LogInfo(propertyName, "Transaction rate: " + loopRate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= loopRate && loopRate <= 10.0:
                        {
                            LogOK(propertyName, "Transaction rate: " + loopRate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= loopRate && loopRate <= 2.0:
                        {
                            LogInfo(propertyName, "Transaction rate: " + loopRate.ToString("0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(propertyName, "Transaction rate: " + loopRate.ToString("0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(propertyName, "Unable to complete test: " + ex.ToString());
            }
        }
    }
}

using System;
using ASCOM.Standard.Interfaces;
using System.Threading;
using ASCOM.Standard.COM.DriverAccess;
using ASCOM.Standard.AlpacaClients;

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
        public CoverCalibratorTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, true, true, false, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        protected override void Dispose(bool disposing)
        {
            LogMsg("Dispose", MessageLevel.Debug, "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (coverCalibratorDevice is not null) coverCalibratorDevice.Dispose();
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
                        logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        coverCalibratorDevice = new AlpacaCoverCalibrator(settings.AlpacaConfiguration.AccessServiceType.ToString(),
                            settings.AlpacaDevice.IpAddress,
                            settings.AlpacaDevice.IpPort,
                            settings.AlpacaDevice.AlpacaDeviceNumber,
                            settings.StrictCasing,
                            settings.DisplayMethodCalls ? logger : null);

                        logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                coverCalibratorDevice = new CoverCalibratorFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                coverCalibratorDevice = new CoverCalibrator(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogMsg("CreateDevice", MessageLevel.Debug, "Successfully created driver");
                baseClassDevice = coverCalibratorDevice; // Assign the driver to the base class

                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
                g_Stop = false;
            }
            catch (Exception ex)
            {
                LogMsg("CreateDevice", MessageLevel.Debug, "Exception thrown: " + ex.Message);
                throw; // Re throw exception 
            }

            if (g_Stop) WaitFor(200);
        }

        public override void PreConnectChecks()
        {
        }

        public override bool Connected
        {
            get
            {
                LogCallToDriver("Connected", "About to get Connected property");
                return coverCalibratorDevice.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                coverCalibratorDevice.Connected = value;
            }
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(coverCalibratorDevice, DeviceType.CoverCalibrator);
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
                        LogMsg("MaxBrightness", MessageLevel.Comment, "About to call MaxBrightness property");
                    maxBrightness = coverCalibratorDevice.MaxBrightness;

                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent))
                    {
                        if (maxBrightness >= 1)
                        {
                            LogMsg("MaxBrightness", MessageLevel.OK, maxBrightness.ToString());
                            maxBrightnessOk = true;
                        }
                        else
                            LogMsg("MaxBrightness", MessageLevel.Issue, $"The returned MaxBrightness value {maxBrightness} is invalid, it must be >= 1");
                    }
                    else
                        LogMsg("MaxBrightness", MessageLevel.Error, $" CalibratorStatus is 'NotPresent' but MaxBrightness did not throw a PropertyNotImplementedException. It returned: {maxBrightness}.");
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
                LogMsg("MaxBrightness", MessageLevel.Issue, $"Test skipped because CalibratorState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            if (calibratorStateOk)
            {
                brightnessOk = false; // Assume a bad value

                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("Brightness", MessageLevel.Comment, "About to call Brightness property");
                    brightness = coverCalibratorDevice.Brightness;
                    brightnessOk = true;

                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotPresent))
                    {
                        if (brightness >= 0)
                        {
                            if (maxBrightnessOk)
                            {
                                if (brightness <= maxBrightness)
                                    LogMsg("Brightness", MessageLevel.OK, maxBrightness.ToString());
                                else
                                    LogMsg("Brightness", MessageLevel.Error, $"The returned Brightness {brightness} is greater than MaxBrightness {maxBrightness}");
                            }
                            else
                                LogMsg("Brightness", MessageLevel.Issue, $"Skipping Brightness test because MaxBrightness did not return a value.");
                        }
                        else
                            LogMsg("Brightness", MessageLevel.Issue, $"The returned Brightness value {brightness} is invalid, it must be >= 0");
                    }
                    else
                        LogMsg("Brightness", MessageLevel.Error, $" CalibratorStatus is 'NotPresent' but Brightness did not throw a PropertyNotImplementedException. It returned: {brightness}.");
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
                LogMsg("Brightness", MessageLevel.Issue, $"Test skipped because CalibratorState returned an exception");

            if (coverStateOk & calibratorStateOk)
            {
                if (coverState == CoverStatus.NotPresent & calibratorState == CalibratorStatus.NotPresent)
                    LogMsg("DeviceCapabilities", MessageLevel.Warning, "Both CoverStatus and CalibratorStatus are set to 'NotPresent' - this driver won't do a lot!");
            }
        }

        public override void CheckMethods()
        {
            DateTime startTime;
            SetStatus("", "", "");  // Clear status messages

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test OpenCover
            if (coverStateOk)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("OpenCover", MessageLevel.Comment, "About to call OpenCover method");
                    startTime = DateTime.Now;

                    coverCalibratorDevice.OpenCover();
                    if (!(coverState == CoverStatus.NotPresent))
                    {
                        if (!(coverCalibratorDevice.CoverState == CoverStatus.Moving))
                        {
                            canAsynchronousOpen = false;
                            if (coverCalibratorDevice.CoverState == CoverStatus.Open)
                                LogMsg("OpenCover", MessageLevel.OK, $"OpenCover was successful. The synchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                            else
                                LogMsg("OpenCover", MessageLevel.Error, $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The synchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                        }
                        else
                        {
                            canAsynchronousOpen = true;
                            asynchronousOpenTime = 0.0;

                            // Wait until the cover is no longer moving
                            while (coverCalibratorDevice.CoverState == CoverStatus.Moving)
                            {
                                WaitFor(10);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                            }
                            if (coverCalibratorDevice.CoverState == CoverStatus.Open)
                            {
                                asynchronousOpenTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                                LogMsg("OpenCover", MessageLevel.OK, $"OpenCover was successful. The asynchronous open took {asynchronousOpenTime.ToString("0.0")} seconds");
                            }
                            else
                                LogMsg("OpenCover", MessageLevel.Error, $"OpenCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Open'. The asynchronous open took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                        }
                    }
                    else
                        LogMsg("OpenCover", MessageLevel.Error, $"CoverStatus is 'NotPresent' but OpenCover did not throw a MethodNotImplementedException.");
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
                LogMsg("OpenCover", MessageLevel.Issue, $"Test skipped because CoverState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test CloseCover
            if (coverStateOk)
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("CloseCover", MessageLevel.Comment, "About to call CloseCover method");
                    startTime = DateTime.Now;
                    asynchronousCloseTime = 0.0;

                    coverCalibratorDevice.CloseCover();
                    if (!(coverState == CoverStatus.NotPresent))
                    {
                        if (!(coverCalibratorDevice.CoverState == CoverStatus.Moving))
                        {
                            canAsynchronousOpen = false;
                            if (coverCalibratorDevice.CoverState == CoverStatus.Closed)
                                LogMsg("CloseCover", MessageLevel.OK, $"CloseCover was successful. The synchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                            else
                                LogMsg("CloseCover", MessageLevel.Error, $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The synchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                        }
                        else
                        {
                            canAsynchronousOpen = true;
                            // Wait until the cover is no longer moving
                            while (coverCalibratorDevice.CoverState == CoverStatus.Moving)
                            {
                                WaitFor(10);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                            }
                            if (coverCalibratorDevice.CoverState == CoverStatus.Closed)
                            {
                                asynchronousCloseTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                                LogMsg("CloseCover", MessageLevel.OK, $"CloseCover was successful. The asynchronous close took {asynchronousCloseTime.ToString("0.0")} seconds");
                            }
                            else
                                LogMsg("CloseCover", MessageLevel.Error, $"CloseCover was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Closed'. The asynchronous close took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                        }
                    }
                    else
                        LogMsg("CloseCover", MessageLevel.Error, $"CoverStatus is 'NotPresent' but CloseCover did not throw a MethodNotImplementedException.");
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
                LogMsg("CloseCover", MessageLevel.Issue, $"Test skipped because CoverState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test HaltCover
            if (coverStateOk)
            {
                if (!(coverState == CoverStatus.NotPresent))
                {
                    if (canAsynchronousOpen)
                    {
                        if ((asynchronousOpenTime > 0.0) & (asynchronousCloseTime > 0.0))
                        {

                            // Initiate a cover open first
                            if (settings.DisplayMethodCalls)
                                LogMsg("HaltCover", MessageLevel.Comment, "About to call OpenCover method");
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
                                        LogMsg("HaltCover", MessageLevel.Comment, "About to call HaltCover method");
                                    coverCalibratorDevice.HaltCover();

                                    // Confirm that the cover is no longer moving
                                    if (!(coverCalibratorDevice.CoverState == CoverStatus.Moving))
                                        LogMsg("HaltCover", MessageLevel.OK, "Cover is no longer moving after issuing the HaltCover command");
                                    else
                                        LogMsg("HaltCover", MessageLevel.Error, "Cover is still moving after issuing the HaltCover command");
                                }
                                catch (Exception ex)
                                {
                                    HandleException("HaltCover", MemberType.Method, Required.MustBeImplemented, ex, "CoverStatus indicates that the device has cover capability");
                                }
                            }
                            else
                                LogMsg("HaltCover", MessageLevel.Issue, "Cover should have been moving after waiting for half of the previous open time, but it was not. Test abandoned");
                        }
                        else
                            LogMsg("HaltCover", MessageLevel.Issue, $"HaltCover tests skipped because either the cover could not be opened or closed successfully.");
                    }
                    else
                        try
                        {
                            // Since the cover opens synchronously the HaltCover method should return a MethodNotImplementedException
                            if (settings.DisplayMethodCalls)
                                LogMsg("HaltCover", MessageLevel.Comment, "About to call HaltCover method");
                            coverCalibratorDevice.HaltCover();
                            LogMsg("HaltCover", MessageLevel.Error, "The cover operates synchronously but did not throw a MethodNotImplementedException in response to the HaltCover command");
                        }
                        catch (Exception ex)
                        {
                            if (coverState == CoverStatus.NotPresent)
                                HandleException("HaltCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                            else
                                HandleException("HaltCover", MemberType.Method, Required.Optional, ex, "");
                        }
                }
                else
                    try
                    {
                        if (settings.DisplayMethodCalls)
                            LogMsg("HaltCover", MessageLevel.Comment, "About to call HaltCover method");
                        coverCalibratorDevice.HaltCover();
                        // Should never get here...
                        LogMsg("HaltCover", MessageLevel.Error, "CoverStatus is 'NotPresent' but HaltCover did not throw a MethodNotImplementedException");
                    }
                    catch (Exception ex)
                    {
                        HandleException("HaltCover", MemberType.Method, Required.MustNotBeImplemented, ex, "CoverStatus is 'NotPresent'");
                    }
            }
            else
                LogMsg("HaltCover", MessageLevel.Issue, $"Test skipped because CoverState returned an exception");

            if (cancellationToken.IsCancellationRequested)
                return;

            // Test CalibratorOn
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
                            LogMsg("CalibratorOn", MessageLevel.Info, $"Test of a high invalid brightness value skipped because MaxBrightness is set to the largest positive integer value.");
                    }
                    else
                        LogMsg("CalibratorOn", MessageLevel.Issue, $"Brightness tests skipped because one of the Brightness or MaxBrightness properties returned an invalid value or threw an exception.");
                }
                else
                {
                    TestCalibratorOn(1);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            else
                LogMsg("CalibratorOn", MessageLevel.Issue, $"Brightness tests skipped because the CoverState property returned an invalid value or threw an exception.");

            // Test CalibratorOff
            if (calibratorStateOk)
            {
                if (!(calibratorState == CalibratorStatus.NotPresent))
                {
                    try
                    {
                        startTime = DateTime.Now;

                        if (settings.DisplayMethodCalls)
                            LogMsg("CalibratorOff", MessageLevel.Comment, "About to call CalibratorOff method");
                        coverCalibratorDevice.CalibratorOff();

                        if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady))
                        {
                            if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Off)
                            {
                                LogMsg("CalibratorOff", MessageLevel.OK, $"CalibratorOff was successful. The synchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");

                                // Confirm that Brightness returns to zero when calibrator is turned off
                                if (settings.DisplayMethodCalls)
                                    LogMsg("CalibratorOff", MessageLevel.Comment, "About to call Brightness property");
                                if (coverCalibratorDevice.Brightness == 0)
                                    LogMsg("CalibratorOff", MessageLevel.OK, $"Brightness is set to zero when the calibrator is turned off");
                                else
                                    LogMsg("CalibratorOff", MessageLevel.Issue, $"Brightness is not set to zero when the calibrator is turned off");
                            }
                            else
                                LogMsg("CalibratorOff", MessageLevel.Error, $"CalibratorOff was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Off'. The synchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                        }
                        else
                        {

                            // Wait until the cover is no longer moving
                            while (coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady)
                            {
                                WaitFor(10);
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                            }
                            if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Off)
                            {
                                LogMsg("CalibratorOff", MessageLevel.OK, $"CalibratorOff was successful. The asynchronous action took {asynchronousCloseTime.ToString("0.0")} seconds");

                                // Confirm that Brightness returns to zero when calibrator is turned off
                                if (settings.DisplayMethodCalls)
                                    LogMsg("CalibratorOff", MessageLevel.Comment, "About to call Brightness property");
                                if (coverCalibratorDevice.Brightness == 0)
                                    LogMsg("CalibratorOff", MessageLevel.OK, $"Brightness is set to zero when the calibrator is turned off");
                                else
                                    LogMsg("CalibratorOff", MessageLevel.Issue, $"Brightness is not set to zero when the calibrator is turned off");
                            }
                            else
                                LogMsg("CalibratorOff", MessageLevel.Error, $"CalibratorOff was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Off'. The asynchronous action took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
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
                            LogMsg("CalibratorOff", MessageLevel.Comment, "About to call CalibratorOff method");
                        coverCalibratorDevice.CalibratorOff();
                        // Should never get here...
                        LogMsg("CalibratorOff", MessageLevel.Error, $"CalibratorStatus is 'NotPresent'but CalibratorOff did not throw a MethodNotImplementedException.");
                    }
                    catch (Exception ex)
                    {
                        HandleException("CalibratorOff", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                    }
            }
            else
                LogMsg("CalibratorOff", MessageLevel.Issue, $"Test skipped because the CoverState property returned an invalid value or threw an exception.");
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
                        LogMsg("CalibratorOn", MessageLevel.Comment, $"About to call CalibratorOn method with brightness: {requestedBrightness}");
                    coverCalibratorDevice.CalibratorOn(requestedBrightness);

                    if (!(coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady))
                    {
                        if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness))
                            LogMsg("CalibratorOn", MessageLevel.Issue, $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.");
                        else if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Ready)
                        {
                            LogMsg("CalibratorOn", MessageLevel.OK, $"CalibratorOn with brightness {requestedBrightness} was successful. The synchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");

                            // Confirm that the brightness value is what was set
                            if (settings.DisplayMethodCalls)
                                LogMsg("CalibratorOn", MessageLevel.Comment, $"About to call Brightness property.");
                            returnedBrightness = coverCalibratorDevice.Brightness;

                            if (returnedBrightness == requestedBrightness)
                                LogMsg("CalibratorOn", MessageLevel.OK, $"The Brightness property does return the value that was set");
                            else
                                LogMsg("CalibratorOn", MessageLevel.Issue, $"The Brightness property value: {returnedBrightness} does not match the value that was set: {requestedBrightness}");
                        }
                        else
                            LogMsg("CalibratorOn", MessageLevel.Error, $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CalibratorState was '{coverCalibratorDevice.CalibratorState.ToString().Trim()}' instead of 'Ready'. The synchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                    }
                    else
                    {
                        // Wait until the cover is no longer moving
                        while (coverCalibratorDevice.CalibratorState == CalibratorStatus.NotReady)
                        {
                            WaitFor(10);
                            if (cancellationToken.IsCancellationRequested)
                                return;
                        }

                        if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness))
                            LogMsg("CalibratorOn", MessageLevel.Issue, $"CalibratorOn with brightness {requestedBrightness} should have thrown an InvalidValueException but did not.");
                        else if (coverCalibratorDevice.CalibratorState == CalibratorStatus.Ready)
                            LogMsg("CalibratorOn", MessageLevel.OK, $"CalibratorOn with brightness {requestedBrightness} was successful. The asynchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                        else
                            LogMsg("CalibratorOn", MessageLevel.Error, $"CalibratorOn with brightness {requestedBrightness} was unsuccessful - the returned CoverState was '{coverCalibratorDevice.CoverState.ToString().Trim()}' instead of 'Ready'. The asynchronous operation took {DateTime.Now.Subtract(startTime).TotalSeconds.ToString("0.0")} seconds");
                    }
                }
                catch (Exception ex)
                {
                    if ((requestedBrightness < 0) | (requestedBrightness > maxBrightness))
                    {
                        if (IsInvalidValueException("CalibratorOn", ex))
                            LogMsg("CalibratorOn", MessageLevel.OK, $"CalibratorOn with brightness {requestedBrightness} threw an InvalidValueException as expected");
                        else
                            LogMsg("CalibratorOn", MessageLevel.Error, $"CalibratorOn with brightness {requestedBrightness} threw an {ex.GetType().Name} exception an InvalidValueException was expected");
                    }
                    else
                        HandleException("CalibratorOn", MemberType.Method, Required.MustBeImplemented, ex, "CalibratorStatus indicates that the device is a calibrator");
                }
            }
            else
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("CalibratorOn", MessageLevel.Comment, $"About to call CalibratorOn method with brightness: {requestedBrightness}");
                    coverCalibratorDevice.CalibratorOn(requestedBrightness);
                    // Should never get here...
                    LogMsg("CalibratorOn", MessageLevel.Error, $"CalibratorStatus is 'NotPresent'but CalibratorOn did not throw a MethodNotImplementedException.");
                }
                catch (Exception ex)
                {
                    HandleException("CalibratorOn", MemberType.Method, Required.MustNotBeImplemented, ex, "CalibratorStatus is 'NotPresent'");
                }
        }


        public override void CheckPerformance()
        {
            Status(StatusType.staTest, "Performance");

            PerformanceTest(PerformanceProperty.CalibratorState, "CalibratorState");
            PerformanceTest(PerformanceProperty.CoverState, "CoverState");

            Status(StatusType.staTest, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
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
                                LogMsg("CalibratorState", MessageLevel.Comment, "About to call CalibratorState property");
                            calibratorState = coverCalibratorDevice.CalibratorState;
                            LogMsg(propertyName, MessageLevel.OK, calibratorState.ToString());
                            break;
                        }

                    case RequiredProperty.CoverState:
                        {
                            if (settings.DisplayMethodCalls)
                                LogMsg("CoverState", MessageLevel.Comment, "About to call CoverState property");
                            coverState = coverCalibratorDevice.CoverState;
                            LogMsg(propertyName, MessageLevel.OK, coverState.ToString());
                            break;
                        }

                    default:
                        {
                            testWasSuccessful = false; // Flag that an issue occurred
                            LogMsg(propertyName, MessageLevel.Error, "RequiredPropertiesTest: Unknown test type " + propertyToTest.ToString());
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

            Status(StatusType.staAction, propertyName);
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
                                LogMsg(propertyName, MessageLevel.Error, "PerformanceTest: Unknown test type " + propertyToTest.ToString());
                                break;
                            }
                    }

                    elapsedTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                    if (elapsedTime > lastElapsedTime + 1.0)
                    {
                        Status(StatusType.staStatus, loopCount + " transactions in " + elapsedTime.ToString( "0") + " seconds");
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
                            LogMsg(propertyName, MessageLevel.Info, "Transaction rate: " + loopRate.ToString( "0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= loopRate && loopRate <= 10.0:
                        {
                            LogMsg(propertyName, MessageLevel.OK, "Transaction rate: " + loopRate.ToString( "0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= loopRate && loopRate <= 2.0:
                        {
                            LogMsg(propertyName, MessageLevel.Info, "Transaction rate: " + loopRate.ToString( "0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogMsg(propertyName, MessageLevel.Info, "Transaction rate: " + loopRate.ToString( "0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogMsg(propertyName, MessageLevel.Info, "Unable to complete test: " + ex.ToString());
            }
        }
    }
}

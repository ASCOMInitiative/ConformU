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
    internal class FocuserTester : DeviceTesterBaseClass
    {
        private enum FocuserPropertyMethod
        {
            IsMoving,
            Position,
            Temperature
        }

        // Focuser variables
        private bool mAbsolute, mIsMoving, mTempComp, mTempCompAvailable;
        private int mMaxIncrement, mMaxStep, mPosition, mPositionOrg;
        private double mStepSize, mTemperature;
        private bool mTempCompTrueOk, mTempCompFalseOk; // Variable to confirm that TempComp can be successfully set to True
        private bool mAbsolutePositionOk = false; // Variable to confirm that absolute position can be read OK
        private bool mCanReadIsMoving = false; // Variable to confirm that IsMoving can be read OK
        private bool mCanReadTemperature = false; // Variable to confirm that Temperature can be read OK

        private const int OUT_OF_RANGE_INCREMENT = 10; // For absolute focusers, the position delta, below 0 or above maximum steps, to test that the focuser will not move to the specified position

        // Helper variables
        private IFocuserV4 focuser;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;
        private readonly int focuserMoveTolerance;

        #region New and Dispose
        public FocuserTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
            focuserMoveTolerance = settings.FocuserMoveTolerance;
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
                    focuser?.Dispose();
                    focuser = null;
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
                            ExInvalidValue1 = (int)0x80040404;
                            ExInvalidValue2 = (int)0x80040404;
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
                        focuser = new AlpacaFocuser(
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
                                focuser = new FocuserFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                focuser = new Focuser(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(focuser, DeviceTypes.Focuser); // Assign the driver to the base class

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

        public override void CheckProperties()
        {
            // Absolute - Required
            try
            {
                LogCallToDriver("Absolute", "About to get Absolute property");
                TimeMethod("Absolute", () => mAbsolute = focuser.Absolute, TargetTime.Fast);
                LogOk("Absolute", mAbsolute.ToString());
            }
            catch (Exception ex)
            {
                HandleException("Absolute", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // IsMoving - Required
            try
            {
                LogCallToDriver("IsMoving", "About to get IsMoving property");
                mCanReadIsMoving = false;
                TimeMethod("IsMoving", () => mIsMoving = focuser.IsMoving, TargetTime.Fast);

                if (!mIsMoving)
                {
                    LogOk("IsMoving", mIsMoving.ToString());
                    mCanReadIsMoving = true;
                }
                else
                    LogIssue("IsMoving", "IsMoving is True at start of tests and it should be false");
            }
            catch (Exception ex)
            {
                HandleException("IsMoving", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // MaxStep - Required
            try
            {
                LogCallToDriver("MaxStep", "About to get MaxStep property");
                TimeMethod("MaxStep", () => mMaxStep = focuser.MaxStep, TargetTime.Fast);
                LogOk("MaxStep", mMaxStep.ToString());
            }
            catch (Exception ex)
            {
                HandleException("MaxStep", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // MaxIncrement - Required
            try
            {
                LogCallToDriver("MaxIncrement", "About to get MaxIncrement property");
                TimeMethod("MaxIncrement", () => mMaxIncrement = focuser.MaxIncrement, TargetTime.Fast);

                // Minimum value is 1, 0 or negative must be a bad value, > MaxStep is a bad value
                switch (mMaxIncrement)
                {
                    case object _ when mMaxIncrement < 1:
                        LogIssue("MaxIncrement", $"MaxIncrement must be at least 1, actual value: {mMaxIncrement}");
                        break;

                    case object _ when mMaxIncrement > mMaxStep:
                        LogIssue("MaxIncrement",
                            $"MaxIncrement is greater than MaxStep and shouldn't be: {mMaxIncrement}");
                        break;

                    default:
                        LogOk("MaxIncrement", mMaxIncrement.ToString());
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleException("MaxIncrement", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // Position - Optional
            if (mAbsolute)
            {
                try
                {
                    mAbsolutePositionOk = false;
                    LogCallToDriver("Position", "About to get Position property");
                    TimeMethod("Position", () => mPosition = focuser.Position, TargetTime.Fast);

                    switch (mPosition) // Check that position is a valid value
                    {
                        case object _ when mPosition < 0: // Lower than lowest position
                            LogIssue("Position", $"Position is < 0, actual value: {mPosition}");
                            break;

                        case object _ when mPosition > mMaxStep: // > highest position
                            LogIssue("Position", $"Position is > MaxStep, actual value: {mPosition}");
                            break;

                        default:
                            {
                                LogOk("Position", mPosition.ToString());
                                mAbsolutePositionOk = true;
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Position", MemberType.Property, Required.MustBeImplemented, ex, "Position must be implemented for an absolute focuser");
                }
            }
            else
            {
                try
                {
                    LogCallToDriver("Position", "About to get Position property");
                    mPosition = focuser.Position;
                    LogIssue("Position", "This is a relative focuser but it didn't raise an exception for Focuser.Position");
                }
                catch (Exception ex)
                {
                    HandleException("Position", MemberType.Property, Required.MustNotBeImplemented, ex, "Position must not be implemented for a relative focuser");
                }
            }
            if (cancellationToken.IsCancellationRequested) return;

            // StepSize - Optional
            try
            {
                LogCallToDriver("StepSize", "About to get StepSize property");
                mStepSize = focuser.StepSize;
                switch (mStepSize)
                {
                    case object _ when mStepSize <= 0.0: // Must be >0
                        LogIssue("StepSize", $"StepSize must be > 0.0, actual value: {mStepSize}");
                        break;

                    default:
                        LogOk("StepSize", mStepSize.ToString());
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleException("StepSize", MemberType.Property, Required.Optional, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // TempCompAvailable - Required
            try
            {
                LogCallToDriver("TempCompAvailable", "About to get TempCompAvailable property");
                TimeMethod("TempCompAvailable", () => mTempCompAvailable = focuser.TempCompAvailable, TargetTime.Fast);
                LogOk("TempCompAvailable", mTempCompAvailable.ToString());
            }
            catch (Exception ex)
            {
                HandleException("StepSize", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // TempComp Read - Required
            try
            {
                LogCallToDriver("TempComp Read", "About to get TempComp property");
                TimeMethod("TempComp Read", () => mTempComp = focuser.TempComp, TargetTime.Fast);

                if (mTempComp & !mTempCompAvailable)
                    LogIssue("TempComp Read", "TempComp is True when TempCompAvailable is False - this should not be so");
                else
                    LogOk("TempComp Read", mTempComp.ToString());
            }
            catch (Exception ex)
            {
                HandleException("TempComp Read", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // TempComp Write - Optional
            if (mTempCompAvailable)
            {
                try
                {
                    mTempCompTrueOk = false; // Initialise to false
                    mTempCompFalseOk = false;
                    // Turn compensation on 
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    TimeMethod("TempComp Write", () => focuser.TempComp = true, TargetTime.Standard);

                    LogOk("TempComp Write", "Successfully turned temperature compensation on");
                    mTempCompTrueOk = true; // Set to true to indicate TempComp can be successfully set to True

                    // Turn compensation off
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    focuser.TempComp = false;
                    LogOk("TempComp Write", "Successfully turned temperature compensation off");
                    mTempCompFalseOk = true;
                }
                catch (Exception ex)
                {
                    HandleException("TempComp Write", MemberType.Property, Required.MustBeImplemented, ex, "Temperature compensation is available but");
                }
            }
            else
            {
                try
                {
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    focuser.TempComp = true;
                    LogIssue("TempComp Write", "Temperature compensation is not available but no exception was raised when TempComp was set True");
                }
                catch (Exception ex)
                {
                    HandleException("TempComp Write", MemberType.Property, Required.MustNotBeImplemented, ex, "Temperature compensation is not available");
                }
            }
            if (cancellationToken.IsCancellationRequested) return;

            // Restore original TempComp setting if possible
            LogCallToDriver("TempComp Write", "About to set TempComp property");
            try
            {
                focuser.TempComp = mTempComp;
            }
            catch
            {
            }
            if (cancellationToken.IsCancellationRequested) return;

            // Temperature - Optional
            try
            {
                mCanReadTemperature = false;
                LogCallToDriver("Temperature", "About to get Temperature property");
                TimeMethod("Temperature", () => mTemperature = focuser.Temperature, TargetTime.Fast);

                switch (mTemperature)
                {
                    case object _ when mTemperature <= -50.0: // Probably a bad value
                        LogIssue("Temperature",
                            $"Temperature < -50.0, - possibly an issue, actual value: {mTemperature}");
                        break;

                    case object _ when mTemperature >= 50.0: // Probably a bad value
                        LogIssue("Temperature",
                            $"Temperature > 50.0, - possibly an issue, actual value: {mTemperature}");
                        break;

                    default:
                        LogOk("Temperature", mTemperature.ToString());
                        mCanReadTemperature = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleException("Temperature", MemberType.Property, Required.Optional, ex, "");
            }
        }

        public override void CheckMethods()
        {
            // Halt - optional
            try
            {
                LogCallToDriver("Halt", "About to call Halt method");
                TimeMethod("Halt", () => focuser.Halt(), TargetTime.Standard);
                LogOk("Halt", "Focuser halted OK");
            }
            catch (Exception ex)
            {
                HandleException("Halt", MemberType.Method, Required.Optional, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // Move - Required
            SetTest("Focuser Move");
            try
            {
                // Ensure that TempComp is false so that a move should be possible
                if (mTempCompFalseOk)
                {
                    LogCallToDriver("Move - TempComp False", "About to set TempComp property");
                    focuser.TempComp = false;
                }
                MoveFocuser("Move - TempComp False", true); // Report any deviation from the expected position
            }
            catch (Exception ex)
            {
                HandleException("Move", MemberType.Method, Required.Mandatory, ex, "");
            }

            SetTest("");
            SetAction("");
            SetStatus("");
            if (cancellationToken.IsCancellationRequested) return;

            // Move with TempComp True (if supported) - Should throw an error
            SetTest("Focuser Move");
            if (mTempCompTrueOk)
            {
                // Select the correct test strategy based on interface version
                if (GetInterfaceVersion() < 3)// Original test method for IFocuserV2 and earlier devices
                {
                    try
                    {
                        LogCallToDriver("Move - TempComp True", "About to set TempComp property");
                        focuser.TempComp = true;
                        MoveFocuser("Move - TempComp True", true); // Report any deviation from the expected position
                        LogIssue("Move - TempComp True", "TempComp is True but no exception is thrown by the Move Method - See Focuser.TempComp entry in Platform help file");
                    }
                    catch (Exception ex)
                    {
                        HandleInvalidOperationExceptionAsOk("", MemberType.Method, Required.MustBeImplemented, ex, "TempComp is True but incorrect exception was thrown by the Move Method", "InvalidOperation Exception correctly raised as expected");
                    }

                }
                else // Test method for revised IFocuserV3 and later behaviour introduced in Platform 6.4
                {
                    try
                    {
                        LogCallToDriver("Move - TempComp True V3", "About to set TempComp property");
                        focuser.TempComp = true;
                        MoveFocuser("Move - TempComp True V3", false); // Ignore any focuser position movement because all bets are off when temperature compensation is enabled
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - TempComp True V3", MemberType.Method, Required.Mandatory, ex, "");
                    }

                }
                if (cancellationToken.IsCancellationRequested) return;

                // For absolute focusers, test movement to the 0 and MaxStep limits, also that the focuser will gracefully stop at the limits if commanded to move beyond them
                if (mAbsolute)
                {
                    if (mTempCompFalseOk)
                    {
                        LogCallToDriver("Move - To 0", "About to set TempComp property");
                        focuser.TempComp = false; // Set temperature compensation off
                    }

                    // Test movement to the 0 limit
                    try
                    {
                        MoveFocuserToPosition("Move - To 0", 0);
                        LogCallToDriver("Move - To 0", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(focuserPosition) <= focuserMoveTolerance)
                        {
                            LogOk("Move - To 0", $"Reported position: {focuserPosition}.");
                        }
                        else
                        {
                            LogIssue("Move - To 0", $"Move ended at {focuserPosition}, which is {focuserPosition} steps away from the expected position 0. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - To 0", MemberType.Method, Required.Mandatory, ex, "");
                    }
                    if (cancellationToken.IsCancellationRequested) return;

                    // Test movement below the 0 limit
                    try
                    {
                        MoveFocuserToPosition("Move - Below 0", -OUT_OF_RANGE_INCREMENT);
                        LogCallToDriver("Move - Below 0", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                        {
                            LogOk("Move - Below 0", $"Movement below 0 was not permitted. (Actually moved to {focuser.Position})");
                        }
                        else
                        {
                            LogIssue("Move - Below 0", $"Move was permitted below position 0. Move ended at {focuserPosition}, which is {focuserPosition} steps away from the expected position 0. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - Below 0", MemberType.Method, Required.Mandatory, ex, "Move should fail gracefully by just moving to position 0; it should not throw an exception");
                    }
                    if (cancellationToken.IsCancellationRequested) return;

                    // Test movement to the mid-point
                    try
                    {
                        int midPoint = mMaxStep / 2;
                        MoveFocuserToPosition("Move - To MidPoint", midPoint);
                        LogCallToDriver("Move - To MidPoint", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(midPoint - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                        {
                            LogOk("Move - To MidPoint", $"Reported position: {focuserPosition}.");
                        }
                        else
                        {
                            LogIssue("Move - To MidPoint", $"Move ended at {focuserPosition}, which is {Math.Abs(focuserPosition - midPoint)} steps away from the expected position {midPoint}. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - To MidPoint", MemberType.Method, Required.Mandatory, ex, "");
                    }
                    if (cancellationToken.IsCancellationRequested) return;

                    // Test movement to the MaxSteps limit
                    try
                    {
                        MoveFocuserToPosition("Move - To MaxStep", mMaxStep);
                        LogCallToDriver("Move - To MaxStep", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(mMaxStep - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                        {
                            LogOk("Move - To MaxStep", $"Reported position: {focuserPosition}.");
                        }
                        else
                        {
                            LogIssue("Move - To MaxStep", $"Move ended at {focuserPosition}, which is {Math.Abs(focuserPosition - mMaxStep)} steps away from the expected position {mMaxStep}. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - To MaxStep", MemberType.Method, Required.Mandatory, ex, "");
                    }
                    if (cancellationToken.IsCancellationRequested) return;

                    // Test movement above the MaxStep limit
                    try
                    {
                        MoveFocuserToPosition("Move - Above MaxStep", mMaxStep + OUT_OF_RANGE_INCREMENT);
                        LogCallToDriver("Move - Above MaxStep", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(mMaxStep - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                        {
                            LogOk("Move - Above MaxStep", $"Movement above MaxStep was not permitted. (Actually moved to {focuser.Position})");
                        }
                        else
                        {
                            LogIssue("Move - Above MaxStep", $"Move was permitted above position MaxStep. Move ended at {focuserPosition}, which is {Math.Abs(mMaxStep - focuserPosition)} steps away from the expected position {mMaxStep}. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - Above Maxstep", MemberType.Method, Required.Mandatory, ex, "Move should fail gracefully by just moving to position MaxStep; it should not throw an exception");
                    }
                    if (cancellationToken.IsCancellationRequested) return;
                }

                // Restore original TempComp value
                try
                {
                    focuser.TempComp = mTempComp;
                }
                catch
                {
                }

                SetTest("");
                SetAction("");
                SetStatus("");
            }
        }

        public override void CheckPerformance()
        {
            // Position
            if (mAbsolutePositionOk)
                FocuserPerformanceTest(FocuserPropertyMethod.Position, "Position");
            else
                LogInfo("Position", "Skipping test as property is not supported");

            // IsMoving
            if (mCanReadIsMoving)
                FocuserPerformanceTest(FocuserPropertyMethod.IsMoving, "IsMoving");
            else
                LogInfo("IsMoving", "Skipping test as property is not supported");

            // Temperature
            if (mCanReadTemperature)
                FocuserPerformanceTest(FocuserPropertyMethod.Temperature, "Temperature");
            else
                LogInfo("Temperature", "Skipping test as property is not supported");
            SetAction("");
            SetStatus("");
            SetTest("");
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

        private void MoveFocuser(string testName, bool checkMoveAccuracy)
        {
            if (mAbsolute) // Absolute focuser
            {
                // Save the current absolute position
                LogCallToDriver(testName, "About to get Position property");
                mPositionOrg = focuser.Position;

                // Calculate an acceptable focus position
                mPosition = mPositionOrg + Convert.ToInt32(mMaxStep / 10); // Move by 1/10 of the maximum focus distance outwards
                if (mPosition >= mMaxStep)
                {
                    mPosition = mPositionOrg - Convert.ToInt32(mMaxStep / 10.0);// Move by 1/10 of the maximum focus distance inwards
                }

                if (Math.Abs(mPosition - mPositionOrg) > mMaxIncrement)
                {
                    mPosition = mPositionOrg + mMaxIncrement; // Apply the MaxIncrement check
                }
            }
            else // Relative focuser
            {
                mPosition = Convert.ToInt32(mMaxIncrement / 10.0);
                // Apply the MaxIncrement check
                if (mPosition > mMaxIncrement) mPosition = mMaxIncrement;
            }

            MoveFocuserToPosition(testName, mPosition); // Move the focuser to the new test position within the focuser's movement range

            // Test outcome if absolute
            if (mAbsolute)
            {
                // Test outcome if required
                if (checkMoveAccuracy)
                {
                    LogCallToDriver(testName, "About to get Position property");
                    int focuserPosition = focuser.Position;

                    if (Math.Abs(mPosition - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                    {
                        LogOk(testName, "Absolute move OK");
                    }
                    else
                    {
                        LogIssue(testName, $"Move ended at {focuserPosition}, which is {Math.Abs(focuserPosition - mPosition)} steps away from the expected position {mPosition}. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
                    }
                }
                else
                    LogOk(testName, "Absolute move OK");
            }
            else
                LogOk(testName, "Relative move OK");

            SetAction($"Returning to original position: {mPositionOrg}");
            if (mAbsolute)
            {
                LogCallToDriver(testName, "About to call Move method");
                focuser.Move(mPositionOrg); // Return to original position
                                            // Wait for asynchronous move to finish
                WaitWhile($"Moving back to starting position", () => focuser.IsMoving, 500, settings.FocuserTimeout, () => $"{focuser.Position} / {mPositionOrg}");
            }
            else
            {
                LogCallToDriver(testName, "About to call Move method");
                focuser.Move(-mPosition); // Return to original position
                                          // Wait for asynchronous move to finish
                WaitWhile($"Moving back to starting position", () => focuser.IsMoving, 500, settings.FocuserTimeout);
            }
        }

        public void MoveFocuserToPosition(string testName, int newPosition)
        {
            DateTime startTime;

            LogDebug(testName, $"New position: {newPosition}");

            // Confirm that the focuser is not moving
            LogCallToDriver(testName, "About to get IsMoving property");
            if (focuser.IsMoving) // This is an issue as we are expecting the focuser to be not moving
            {
                LogIssue(testName, "Focuser is already moving before start of Move test, rest of test skipped");
            }
            else // Focuser not moving so proceed with the test
            {
                // Move the focuser
                if (mAbsolute)
                {
                    LogDebug(testName, $"Moving to position: {newPosition}");
                }
                else
                {
                    LogDebug(testName, $"Moving by: {newPosition}");
                }

                SetAction(testName);
                startTime = DateTime.Now;

                LogCallToDriver(testName, "About to call Move method");
                TimeMethod($"Move to {newPosition}", () => focuser.Move(newPosition), TargetTime.Standard); // Move the focuser
                TimeSpan duration = DateTime.Now.Subtract(startTime);

                // Check whether the Move is a synchronous call or an asynchronous call where Move() exceeded the standard response time
                if (duration.TotalSeconds > standardTargetResponseTime) // The Move command duration was more than the standard response time, so could be a synchronous call
                {
                    // Check whether this is a Platform 6 or Platform 7 device
                    if (IsPlatform7OrLater) // Platform 7 or later interface
                    {
                        LogDebug(testName, $"Platform 7 - Synchronous call behaviour - the call returned in {duration.TotalSeconds} seconds.");

                        // Check whether this really is a synchronous move
                        LogCallToDriver(testName, "About to get IsMoving property");
                        if (focuser.IsMoving) // This is an asynchronous move but with a long initiation time
                        {
                            LogIssue(testName, $"The Move method took {duration.TotalSeconds:0.000} seconds to complete, which exceeded Conform's standard response time of {standardTargetResponseTime:0.000} seconds.");

                            // Wait for the move to complete
                            if (mAbsolute)
                            {
                                WaitWhile($"Moving focuser", () => focuser.IsMoving, 500, settings.FocuserTimeout, () => $"{focuser.Position} / {newPosition}"); // Wait for move to complete
                            }
                            else // Relative focuser that doesn't report position
                            {
                                WaitWhile($"Moving focuser", () => focuser.IsMoving, 500, settings.FocuserTimeout); // Wait for move to complete
                            }
                        }
                        else // This is a synchronous move in excess of the standard response time
                        {
                            // Log an issue because of the synchronous move
                            LogIssue(testName, $"The focuser moved synchronously and the Move() method exceeded the standard response time of {standardTargetResponseTime:0.0} seconds.");
                        }
                    }
                    else // Platform 6 or earlier interface
                    {
                        LogDebug(testName, $"Synchronous call behaviour - the call returned in {duration.TotalSeconds} seconds.");

                        // Confirm that IsMoving is false
                        LogCallToDriver(testName, "About to get IsMoving property");
                        if (focuser.IsMoving)
                        {
                            LogIssue(testName, $"The Move method took {duration.TotalSeconds:0.000} seconds to complete and was assumed to be synchronous, but the IsMoving property returned TRUE after the Move completed.");
                            LogInfo(testName, $"The move was assumed to be synchronous because the Move method duration exceeded Conform's standard response time of {standardTargetResponseTime:0.000} seconds.");
                            if (mAbsolute)
                            {
                                WaitWhile($"Moving focuser", () => focuser.IsMoving, 500, settings.FocuserTimeout, () => $"{focuser.Position} / {newPosition}"); // Wait for move to complete
                            }
                            else // Relative focuser that doesn't report position
                            {
                                WaitWhile($"Moving focuser", () => focuser.IsMoving, 500, settings.FocuserTimeout); // Wait for move to complete
                            }
                        }
                        else
                            LogTestAndMessage(testName, "Synchronous move found");
                    }
                }
                else // Move took less than the standard response time so assume an asynchronous call
                {
                    LogDebug(testName, $"Asynchronous call behaviour");
                    SetStatus("Waiting for asynchronous move to complete");
                    LogCallToDriver(testName, "About to get IsMoving and Position properties repeatedly");
                    if (mAbsolute)
                    {
                        WaitWhile($"Moving focuser", () => focuser.IsMoving, 500, settings.FocuserTimeout, () => $"{focuser.Position} / {newPosition}");
                        LogDebug(testName, $"Final position: {focuser.Position}, IsMoving: {focuser.IsMoving}");
                    }
                    else // Relative focuser that doesn't report position
                    {
                        WaitWhile($"Moving focuser", () => focuser.IsMoving, 500, settings.FocuserTimeout); // Wait for move to complete
                    }
                }
            }
        }

        private void FocuserPerformanceTest(FocuserPropertyMethod pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime;
            float lSingle;
            bool lBoolean;
            double lRate;
            SetTest("Performance Test");
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
                        case FocuserPropertyMethod.IsMoving:
                            {
                                lBoolean = focuser.IsMoving;
                                break;
                            }

                        case FocuserPropertyMethod.Position:
                            {
                                lSingle = focuser.Position;
                                break;
                            }

                        case FocuserPropertyMethod.Temperature:
                            {
                                lRate = focuser.Temperature;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"FocuserPerformanceTest: Unknown test type {pType}");
                                break;
                            }
                    }

                    lElapsedTime = DateTime.Now.Subtract(lStartTime).TotalSeconds;
                    if (lElapsedTime > lLastElapsedTime + 1.0)
                    {
                        SetStatus($"{lCount} transactions in {lElapsedTime:0} seconds");
                        lLastElapsedTime = lElapsedTime;
                        if (cancellationToken.IsCancellationRequested) return;
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
            }
        }

        #endregion

    }
}

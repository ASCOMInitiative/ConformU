using ASCOM;
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
        // Focuser moves can be synchronous or asynchronous and the duration of the Move command is used to differentiate the two behaviours.
        // If the Move completes within the SYNC_TEST_TIME the move is considered to be asynchronous and it takes longer it will be assumed to be synchronous
        private const int  MOVE_SYNC_TEST_TIME = 1000; // Duration of a Move command beyond which the move is considered synchronous (Seconds)

        enum FocuserPropertyMethod
        {
            IsMoving,
            Position,
            Temperature
        }

        // Focuser variables
        private bool m_Absolute, m_IsMoving, m_TempComp, m_TempCompAvailable;
        private int m_MaxIncrement, m_MaxStep, m_Position, m_PositionOrg;
        private double m_StepSize, m_Temperature;
        private bool m_TempCompTrueOK, m_TempCompFalseOK; // Variable to confirm that TempComp can be successfully set to True
        private bool m_AbsolutePositionOK = false; // Variable to confirm that absolute position can be read OK
        private bool m_CanReadIsMoving = false; // Variable to confirm that IsMoving can be read OK
        private bool m_CanReadTemperature = false; // Variable to confirm that Temperature can be read OK

        private const int OUT_OF_RANGE_INCREMENT = 10; // For absolute focusers, the position delta, below 0 or above maximum steps, to test that the focuser will not move to the specified position

        // Helper variables
        private IFocuserV3 focuser;
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
            LogDebug("Dispose", "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
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
                            g_ExInvalidValue1 = (int)0x80040404;
                            g_ExInvalidValue2 = (int)0x80040404;
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
                baseClassDevice = focuser; // Assign the driver to the base class

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
                return focuser.Connected;
            }
            set
            {
                LogCallToDriver("ConformanceCheck", "About to set Connected");
                SetTest("Connected");
                SetAction("Waiting for Connected to become 'true'");
                focuser.Connected = value;
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
            base.CheckCommonMethods(focuser, DeviceTypes.Focuser);
        }

        public override void CheckProperties()
        {
            // Absolute - Required
            try
            {
                LogCallToDriver("Absolute", "About to get Absolute property");
                m_Absolute = focuser.Absolute;
                LogOK("Absolute", m_Absolute.ToString());
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
                m_CanReadIsMoving = false;
                m_IsMoving = focuser.IsMoving;
                if (!m_IsMoving)
                {
                    LogOK("IsMoving", m_IsMoving.ToString());
                    m_CanReadIsMoving = true;
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
                m_MaxStep = focuser.MaxStep;
                LogOK("MaxStep", m_MaxStep.ToString());
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
                m_MaxIncrement = focuser.MaxIncrement;
                // Minimum value is 1, 0 or negative must be a bad value, >maxstep is a bad value
                switch (m_MaxIncrement)
                {
                    case object _ when m_MaxIncrement < 1:
                        {
                            LogIssue("MaxIncrement", "MaxIncrement must be at least 1, actual value: " + m_MaxIncrement);
                            break;
                        }

                    case object _ when m_MaxIncrement > m_MaxStep:
                        {
                            LogIssue("MaxIncrement", "MaxIncrement is greater than MaxStep and shouldn't be: " + m_MaxIncrement);
                            break;
                        }

                    default:
                        {
                            LogOK("MaxIncrement", m_MaxIncrement.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("MaxIncrement", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // Position - Optional
            if (m_Absolute)
            {
                try
                {
                    m_AbsolutePositionOK = false;
                    LogCallToDriver("Position", "About to get Position property");
                    m_Position = focuser.Position;
                    switch (m_Position) // Check that position is a valid value
                    {
                        case object _ when m_Position < 0 // Lower than lowest position
                       :
                            {
                                LogIssue("Position", "Position is < 0, actual value: " + m_Position.ToString());
                                break;
                            }

                        case object _ when m_Position > m_MaxStep // > highest position
                 :
                            {
                                LogIssue("Position", "Position is > MaxStep, actual value: " + m_Position.ToString());
                                break;
                            }

                        default:
                            {
                                LogOK("Position", m_Position.ToString());
                                m_AbsolutePositionOK = true;
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
                    m_Position = focuser.Position;
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
                m_StepSize = focuser.StepSize;
                switch (m_StepSize)
                {
                    case object _ when m_StepSize <= 0.0 // Must be >0
                   :
                        {
                            LogIssue("StepSize", "StepSize must be > 0.0, actual value: " + m_StepSize);
                            break;
                        }

                    default:
                        {
                            LogOK("StepSize", m_StepSize.ToString());
                            break;
                        }
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
                m_TempCompAvailable = focuser.TempCompAvailable;
                LogOK("TempCompAvailable", m_TempCompAvailable.ToString());
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
                m_TempComp = focuser.TempComp;
                if (m_TempComp & !m_TempCompAvailable)
                    LogIssue("TempComp Read", "TempComp is True when TempCompAvailable is False - this should not be so");
                else
                    LogOK("TempComp Read", m_TempComp.ToString());
            }
            catch (Exception ex)
            {
                HandleException("TempComp Read", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // TempComp Write - Optional
            if (m_TempCompAvailable)
            {
                try
                {
                    m_TempCompTrueOK = false; // Initialise to false
                    m_TempCompFalseOK = false;
                    // Turn compensation on 
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    focuser.TempComp = true;
                    LogOK("TempComp Write", "Successfully turned temperature compensation on");
                    m_TempCompTrueOK = true; // Set to true to indicate TempComp can be successfully set to True
                                             // Turn compensation off
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    focuser.TempComp = false;
                    LogOK("TempComp Write", "Successfully turned temperature compensation off");
                    m_TempCompFalseOK = true;
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
                focuser.TempComp = m_TempComp;
            }
            catch
            {
            }
            if (cancellationToken.IsCancellationRequested) return;

            // Temperature - Optional
            try
            {
                m_CanReadTemperature = false;
                LogCallToDriver("Temperature", "About to get Temperature property");
                m_Temperature = focuser.Temperature;
                switch (m_Temperature)
                {
                    case object _ when m_Temperature <= -50.0 // Probably a bad value
                   :
                        {
                            LogIssue("Temperature", "Temperature < -50.0, - possibly an issue, actual value: " + m_Temperature.ToString());
                            break;
                        }

                    case object _ when m_Temperature >= 50.0 // Probably a bad value
             :
                        {
                            LogIssue("Temperature", "Temperature > 50.0, - possibly an issue, actual value: " + m_Temperature.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK("Temperature", m_Temperature.ToString());
                            m_CanReadTemperature = true;
                            break;
                        }
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
                focuser.Halt();
                LogOK("Halt", "Focuser halted OK");
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
                if (m_TempCompFalseOK)
                {
                    LogCallToDriver("Move - TempComp False", "About to set TempComp property");
                    focuser.TempComp = false;
                }
                MoveFocuser("Move - TempComp False");
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
            if (m_TempCompTrueOK)
            {
                switch (g_InterfaceVersion)
                {
                    case 0:
                    case 1:
                    case 2 // Original test method for IFocuserV2 and earlier devices
                   :
                        {
                            try
                            {
                                LogCallToDriver("Move - TempComp True", "About to set TempComp property");
                                focuser.TempComp = true;
                                MoveFocuser("Move - TempComp True");
                                LogIssue("Move - TempComp True", "TempComp is True but no exception is thrown by the Move Method - See Focuser.TempComp entry in Platform help file");
                            }
                            catch (Exception ex)
                            {
                                HandleInvalidOperationExceptionAsOK("", MemberType.Method, Required.MustBeImplemented, ex, "TempComp is True but incorrect exception was thrown by the Move Method", "InvalidOperation Exception correctly raised as expected");
                            }

                            break;
                        }

                    case 3 // Test method for revised IFocuserV3 behaviour introduced in Platform 6.4
             :
                        {
                            try
                            {
                                LogCallToDriver("Move - TempComp True V3", "About to set TempComp property");
                                focuser.TempComp = true;
                                MoveFocuser("Move - TempComp True V3");
                            }
                            catch (Exception ex)
                            {
                                HandleException("Move - TempComp True V3", MemberType.Method, Required.Mandatory, ex, "");
                            }

                            break;
                        }

                    default:
                        {
                            LogIssue("Move - TempComp True", string.Format("Unknown interface version returned {0}, Move test with temperature compensation enabled skipped.", g_InterfaceVersion));
                            break;
                        }
                }
                if (cancellationToken.IsCancellationRequested) return;

                // For absolute focusers, test movement to the 0 and MaxStep limits, also that the focuser will gracefully stop at the limits if commanded to move beyond them
                if (m_Absolute)
                {
                    if (m_TempCompFalseOK)
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
                            LogOK("Move - To 0", $"Reported position: {focuserPosition}.");
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
                            LogOK("Move - Below 0", $"Movement below 0 was not permitted. (Actually moved to {focuser.Position})");
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
                        int midPoint = m_MaxStep / 2;
                        MoveFocuserToPosition("Move - To MidPoint", midPoint);
                        LogCallToDriver("Move - To MidPoint", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(midPoint - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                        {
                            LogOK("Move - To MidPoint", $"Reported position: {focuserPosition}.");
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
                        MoveFocuserToPosition("Move - To MaxStep", m_MaxStep);
                        LogCallToDriver("Move - To MaxStep", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(m_MaxStep - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                        {
                            LogOK("Move - To MaxStep", $"Reported position: {focuserPosition}.");
                        }
                        else
                        {
                            LogIssue("Move - To MaxStep", $"Move ended at {focuserPosition}, which is {Math.Abs(focuserPosition - m_MaxStep)} steps away from the expected position {m_MaxStep}. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
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
                        MoveFocuserToPosition("Move - Above MaxStep", m_MaxStep + OUT_OF_RANGE_INCREMENT);
                        LogCallToDriver("Move - Above MaxStep", "About to get Position property");
                        int focuserPosition = focuser.Position;
                        if (Math.Abs(m_MaxStep - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                        {
                            LogOK("Move - Above MaxStep", $"Movement above MaxStep was not permitted. (Actually moved to {focuser.Position})");
                        }
                        else
                        {
                            LogIssue("Move - Above MaxStep", $"Move was permitted above position MaxStep. Move ended at {focuserPosition}, which is {Math.Abs(m_MaxStep - focuserPosition)} steps away from the expected position {m_MaxStep}. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
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
                    focuser.TempComp = m_TempComp;
                }
                catch
                {
                }

                SetTest("");
                SetAction("");
                SetStatus("");
            }
        }

        private void MoveFocuser(string testName)
        {
            if (m_Absolute) // Absolute focuser
            {
                // Save the current absolute position
                LogCallToDriver(testName, "About to get Position property");
                m_PositionOrg = focuser.Position;
                // Calculate an acceptable focus position
                m_Position = m_PositionOrg + Convert.ToInt32(m_MaxStep / 10); // Move by 1/10 of the maximum focus distance outwards
                if (m_Position >= m_MaxStep)
                {
                    m_Position = m_PositionOrg - Convert.ToInt32(m_MaxStep / 10.0);// Move by 1/10 of the maximum focus distance inwards
                }

                if (Math.Abs(m_Position - m_PositionOrg) > m_MaxIncrement)
                {
                    m_Position = m_PositionOrg + m_MaxIncrement; // Apply the MaxIncrement check
                }
            }
            else // Relative focuser
            {
                m_Position = Convert.ToInt32(m_MaxIncrement / 10.0);
                // Apply the MaxIncrement check
                if (m_Position > m_MaxIncrement) m_Position = m_MaxIncrement;
            }

            MoveFocuserToPosition(testName, m_Position); // Move the focuser to the new test position within the focuser's movement range

            // Test outcome if absolute
            if (m_Absolute)
            {
                LogCallToDriver(testName, "About to get Position property");
                int focuserPosition = focuser.Position;

                if (Math.Abs(m_Position - focuserPosition) <= focuserMoveTolerance) // Allow a small tolerance
                {
                    LogOK(testName, "Absolute move OK");
                }
                else
                {
                    LogIssue(testName, $"Move ended at {focuserPosition}, which is {Math.Abs(focuserPosition - m_Position)} steps away from the expected position {m_Position}. This is outside Conform's configured move tolerance: {focuserMoveTolerance}.");
                }
            }
            else
                LogOK(testName, "Relative move OK");

            SetAction($"Returning to original position: {m_PositionOrg}");
            if (m_Absolute)
            {
                LogCallToDriver(testName, "About to call Move method");
                focuser.Move(m_PositionOrg); // Return to original position
                                             // Wait for asynchronous move to finish
                WaitWhile($"Moving back to starting position", () => { return focuser.IsMoving; }, 500, settings.FocuserTimeout, () => { return $"{focuser.Position} / {m_PositionOrg}"; });
            }
            else
            {
                LogCallToDriver(testName, "About to call Move method");
                focuser.Move(-m_Position); // Return to original position
                                           // Wait for asynchronous move to finish
                WaitWhile($"Moving back to starting position", () => { return focuser.IsMoving; }, 500, settings.FocuserTimeout);
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
                if (m_Absolute)
                {
                    LogDebug(testName, "Moving to position: " + newPosition.ToString());
                }
                else
                {
                    LogDebug(testName, "Moving by: " + newPosition.ToString());
                }

                SetAction(testName);
                startTime = DateTime.Now;
                LogCallToDriver(testName, "About to call Move method");
                focuser.Move(newPosition); // Move the focuser
                TimeSpan duration = DateTime.Now.Subtract(startTime);

                // Test whether the Move will be treated as synchronous or asynchronous
                if (duration.TotalMilliseconds > MOVE_SYNC_TEST_TIME) // The Move command duration was more than the configured time, so assume a synchronous call
                {
                    LogDebug(testName, $"Synchronous call behaviour - the call returned in {duration.TotalSeconds} seconds.");

                    // Confirm that IsMoving is false
                    LogCallToDriver(testName, "About to get IsMoving property");
                    if (focuser.IsMoving)
                    {
                        LogIssue(testName, $"The Move method took {duration.TotalSeconds:0.000} seconds to complete and was assumed to be synchronous, but the IsMoving property returned TRUE after the Move completed.");
                        LogInfo(testName, $"The move was assumed to be synchronous because the Move method duration exceeded Conform's built-in sync/async test time of {MOVE_SYNC_TEST_TIME/1000:0.000} seconds.");
                        if (m_Absolute)
                        {
                            WaitWhile($"Moving focuser", () => { return focuser.IsMoving; }, 500, settings.FocuserTimeout, () => { return $"{focuser.Position} / {newPosition}"; }); // Wait for move to complete
                        }
                        else // Relative focuser that doesn't report position
                        {
                            WaitWhile($"Moving focuser", () => { return focuser.IsMoving; }, 500, settings.FocuserTimeout); // Wait for move to complete
                        }
                    }
                    else
                        LogTestAndMessage(testName, "Synchronous move found");
                }
                else // Move took less than 1 second so assume an asynchronous call
                {
                    LogDebug(testName, $"Asynchronous call behaviour");
                    SetStatus("Waiting for asynchronous move to complete");
                    LogCallToDriver(testName, "About to get IsMoving and Position properties repeatedly");
                    if (m_Absolute)
                    {
                        WaitWhile($"Moving focuser", () => { return focuser.IsMoving; }, 500, settings.FocuserTimeout, () => { return $"{focuser.Position} / {newPosition}"; });
                        LogDebug(testName, $"Final position: {focuser.Position}, IsMoving: {focuser.IsMoving}");
                    }
                    else // Relative focuser that doesn't report position
                    {
                        WaitWhile($"Moving focuser", () => { return focuser.IsMoving; }, 500, settings.FocuserTimeout); // Wait for move to complete
                    }


                    // LogTestAndMessage(testName, "Asynchronous move completed");
                }
            }
        }

        public override void CheckPerformance()
        {
            // Position
            if (m_AbsolutePositionOK)
                FocuserPerformanceTest(FocuserPropertyMethod.Position, "Position");
            else
                LogInfo("Position", "Skipping test as property is not supported");

            // IsMoving
            if (m_CanReadIsMoving)
                FocuserPerformanceTest(FocuserPropertyMethod.IsMoving, "IsMoving");
            else
                LogInfo("IsMoving", "Skipping test as property is not supported");

            // Temperature
            if (m_CanReadTemperature)
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

        private void FocuserPerformanceTest(FocuserPropertyMethod p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
            float l_Single;
            bool l_Boolean;
            double l_Rate;
            SetTest("Performance Test");
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
                        case FocuserPropertyMethod.IsMoving:
                            {
                                l_Boolean = focuser.IsMoving;
                                break;
                            }

                        case FocuserPropertyMethod.Position:
                            {
                                l_Single = focuser.Position;
                                break;
                            }

                        case FocuserPropertyMethod.Temperature:
                            {
                                l_Rate = focuser.Temperature;
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, "FocuserPerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
                    {
                        SetStatus(l_Count + " transactions in " + l_ElapsedTime.ToString("0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        if (cancellationToken.IsCancellationRequested) return;
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
                LogInfo(p_Name, "Unable to complete test: " + ex.Message);
            }
        }
    }
}

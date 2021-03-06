using ASCOM;
using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConformU
{
    internal class FocuserTester : DeviceTesterBaseClass
    {

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

        private const int GOOD_MOVE_TOLERANCE = 2; // ± Position tolerance within which a move will be considered to be OK
        private const int OUT_OF_RANGE_INCREMENT = 10; // For absolute focusers, the position delta, below 0 or above maximum steps, to test that the focuser will not move to the specified position

        // Helper variables
        private IFocuserV3 m_Focuser;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;


        #region New and Dispose
        public FocuserTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
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
                    if (m_Focuser is not null) m_Focuser.Dispose();
                    m_Focuser = null;
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
                        m_Focuser = new AlpacaFocuser(settings.AlpacaConfiguration.AccessServiceType,
                            settings.AlpacaDevice.IpAddress,
                            settings.AlpacaDevice.IpPort,
                            settings.AlpacaDevice.AlpacaDeviceNumber,
                            settings.AlpacaConfiguration.StrictCasing,
                            settings.TraceAlpacaCalls ? logger : null);
                        LogInfo("CreateDevice", $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_Focuser = new FocuserFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                m_Focuser = new Focuser(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                baseClassDevice = m_Focuser; // Assign the driver to the base class

                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");

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
                LogCallToDriver("Connected", "About to get Connected property");
                return m_Focuser.Connected;
            }
            set
            {
                try
                {
                    LogDebug("Connected", "Setting connected state to: " + value.ToString());
                    LogCallToDriver("Connected", "About to set Link property");
                    m_Focuser.Connected = value;
                    LogDebug("AccessChecks", "Successfully changed connected state");
                }
                catch (Exception ex)
                {
                    LogIssue("Connected", "Error changing focuser connected state: " + ex.ToString());
                }
            }
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_Focuser, DeviceType.Focuser);
        }

        public override void CheckProperties()
        {
            // Absolute - Required
            try
            {
                LogCallToDriver("Absolute", "About to get Absolute property");
                m_Absolute = m_Focuser.Absolute;
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
                m_IsMoving = m_Focuser.IsMoving;
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
                m_MaxStep = m_Focuser.MaxStep;
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
                m_MaxIncrement = m_Focuser.MaxIncrement;
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
                    m_Position = m_Focuser.Position;
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
                    m_Position = m_Focuser.Position;
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
                m_StepSize = m_Focuser.StepSize;
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
                m_TempCompAvailable = m_Focuser.TempCompAvailable;
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
                m_TempComp = m_Focuser.TempComp;
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
                    m_Focuser.TempComp = true;
                    LogOK("TempComp Write", "Successfully turned temperature compensation on");
                    m_TempCompTrueOK = true; // Set to true to indicate TempComp can be successfully set to True
                                             // Turn compensation off
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    m_Focuser.TempComp = false;
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
                    m_Focuser.TempComp = true;
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
                m_Focuser.TempComp = m_TempComp;
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
                m_Temperature = m_Focuser.Temperature;
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
                m_Focuser.Halt();
                LogOK("Halt", "Focuser halted OK");
            }
            catch (COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    case object _ when ex.ErrorCode == g_ExNotImplemented:
                    case object _ when ex.ErrorCode == ErrorCodes.NotImplemented:
                        {
                            LogOK("Halt", "COM Exception - Halt is not supported by this focuser");
                            break;
                        }

                    default:
                        {
                            LogIssue("Halt", $"{EX_COM}{ex.Message}{ex.ErrorCode: X8}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Halt", MemberType.Method, Required.Optional, ex, "");
            }
            if (cancellationToken.IsCancellationRequested) return;

            // Move - Required
            Status(StatusType.staTest, "Focuser Move");
            try
            {
                // Ensure that TempComp is false so that a move should be possible
                if (m_TempCompFalseOK)
                {
                    LogCallToDriver("Move - TempComp False", "About to set TempComp property");
                    m_Focuser.TempComp = false;
                }
                MoveFocuser("Move - TempComp False");
            }
            catch (Exception ex)
            {
                HandleException("Move", MemberType.Method, Required.Mandatory, ex, "");
            }
            Status(StatusType.staTest, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
            if (cancellationToken.IsCancellationRequested) return;

            // Move with TempComp True (if supported) - Should throw an error
            Status(StatusType.staTest, "Focuser Move");
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
                                m_Focuser.TempComp = true;
                                MoveFocuser("Move - TempComp True");
                                LogIssue("Move - TempComp True", "TempComp is True but no exception is thrown by the Move Method - See Focuser.TempComp entry in Platform help file");
                            }
                            catch (COMException)
                            {
                                LogOK("Move - TempComp True", "COM Exception correctly raised as expected");
                            }
                            catch (ASCOM.InvalidOperationException)
                            {
                                LogOK("Move - TempComp True", ".NET InvalidOperation Exception correctly raised as expected");
                            }
                            catch (System.InvalidOperationException)
                            {
                                LogIssue("Move - TempComp True", "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException");
                            }
                            catch (Exception ex)
                            {
                                LogIssue("Move - TempComp True", "Unexpected .NET Exception: " + ex.Message);
                            }

                            break;
                        }

                    case 3 // Test method for revised IFocuserV3 behaviour introduced in Platform 6.4
             :
                        {
                            try
                            {
                                LogCallToDriver("Move - TempComp True V3", "About to set TempComp property");
                                m_Focuser.TempComp = true;
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
                        m_Focuser.TempComp = false; // Set temperature compensation off
                    }

                    // Test movement to the 0 limit
                    try
                    {
                        MoveFocuserToPosition("Move - To 0", 0);
                        LogCallToDriver("Move - To 0", "About to get Position property");
                        switch (m_Focuser.Position)
                        {
                            case object _ when -GOOD_MOVE_TOLERANCE <= m_Focuser.Position && m_Focuser.Position <= +GOOD_MOVE_TOLERANCE // OK if within a small tolerance of expected value
                           :
                                {
                                    LogCallToDriver("Move - To 0", "About to get Position property");
                                    LogOK("Move - To 0", string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - To 0", "About to get Position property");
                                    LogInfo("Move - To 0", string.Format("Move was within {0} counts of desired position", m_Focuser.Position));
                                    break;
                                }
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
                        switch (m_Focuser.Position)
                        {
                            case object _ when -GOOD_MOVE_TOLERANCE <= m_Focuser.Position && m_Focuser.Position <= +GOOD_MOVE_TOLERANCE // OK if within a small tolerance of expected value
                           :
                                {
                                    LogCallToDriver("Move - Below 0", "About to get Position property");
                                    LogOK("Move - Below 0", string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - Below 0", "About to get Position property");
                                    LogIssue("Move - Below 0", string.Format("Move was permitted below position 0: {0} ", m_Focuser.Position));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - Below 0", MemberType.Method, Required.Mandatory, ex, "Move should fail gracefully by just moving to position 0; it should not throw an exception");
                    }
                    if (cancellationToken.IsCancellationRequested) return;

                    // Test movement to the MaxSteps limit
                    try
                    {
                        MoveFocuserToPosition("Move - To MaxStep", m_MaxStep);
                        LogCallToDriver("Move - To MaxStep", "About to get Position property");
                        switch (m_Focuser.Position)
                        {
                            case object _ when m_MaxStep - GOOD_MOVE_TOLERANCE <= m_Focuser.Position && m_Focuser.Position <= m_MaxStep + GOOD_MOVE_TOLERANCE // OK if within a small tolerance of expected value
                           :
                                {
                                    LogCallToDriver("Move - To MaxStep", "About to get Position property");
                                    LogOK("Move - To MaxStep", string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - To MaxStep", "About to get Position property");
                                    LogInfo("Move - To MaxStep", string.Format("Move position: {0}, within {1} counts of desired position", m_Focuser.Position, m_Focuser.Position - m_MaxStep));
                                    break;
                                }
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
                        MoveFocuserToPosition("Move - Above Maxstep", m_MaxStep + OUT_OF_RANGE_INCREMENT);
                        LogCallToDriver("Move - Above MaxStep", "About to get Position property");
                        switch (m_Focuser.Position)
                        {
                            case object _ when m_MaxStep - GOOD_MOVE_TOLERANCE <= m_Focuser.Position && m_Focuser.Position <= m_MaxStep + GOOD_MOVE_TOLERANCE // OK if within a small tolerance of expected value
                           :
                                {
                                    LogCallToDriver("Move - Above MaxStep", "About to get Position property");
                                    LogOK("Move - Above Maxstep", string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - Above MaxStep", "About to get Position property");
                                    LogIssue("Move - Above Maxstep", string.Format("Moved to {0}, {1} steps from MaxStep ", m_Focuser.Position, m_Focuser.Position - m_MaxStep));
                                    break;
                                }
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
                    m_Focuser.TempComp = m_TempComp;
                }
                catch
                {
                }

                Status(StatusType.staTest, "");
                Status(StatusType.staAction, "");
                Status(StatusType.staStatus, "");
            }
        }

        private void MoveFocuser(string testName)
        {
            if (m_Absolute)
            {
                // Save the current absolute position
                LogCallToDriver(testName, "About to get Position property");
                m_PositionOrg = m_Focuser.Position;
                // Calculate an acceptable focus position
                m_Position = m_PositionOrg + Convert.ToInt32(m_MaxStep / 10); // Move by 1/10 of the maximum focus distance out 
                if (m_Position >= m_MaxStep) m_Position = m_PositionOrg - Convert.ToInt32(m_MaxStep / 10.0);// Move by 1/10 of the maximum focus distance in

                if (Math.Abs(m_Position - m_PositionOrg) > m_MaxIncrement) m_Position = m_PositionOrg + m_MaxIncrement; // Apply the MaxIncrement check
            }
            else
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
                switch (m_Focuser.Position - m_Position)
                {
                    case object _ when -GOOD_MOVE_TOLERANCE <= m_Focuser.Position - m_Position && m_Focuser.Position - m_Position <= +GOOD_MOVE_TOLERANCE // OK if within a small tolerance of expected value
                   :
                        {
                            LogOK(testName, "Absolute move OK");
                            break;
                        }

                    default:
                        {
                            LogCallToDriver(testName, "About to get Position property");
                            LogInfo(testName, $"Move was within {m_Focuser.Position - m_Position} counts of desired position");
                            break;
                        }
                }
            }
            else
                LogOK(testName, "Relative move OK");

            Status(StatusType.staStatus, "");
            Status(StatusType.staAction, "Returning to original position: " + m_PositionOrg);
            LogInfo(testName, "Returning to original position: " + m_PositionOrg);
            if (m_Absolute)
            {
                LogCallToDriver(testName, "About to call Move method");
                m_Focuser.Move(m_PositionOrg); // Return to original position
            }
            else
            {
                LogCallToDriver(testName, "About to call Move method");
                m_Focuser.Move(-m_Position); // Return to original position
            }
            Status(StatusType.staStatus, "Waiting for asynchronous move to complete");

            // Wait for asynchronous move to finish
            LogCallToDriver(testName, "About to get IsMoving and Position properties repeatedly");
            while (m_Focuser.IsMoving & !cancellationToken.IsCancellationRequested)
            {
                if (m_AbsolutePositionOK)
                    Status(StatusType.staStatus, "Waiting for asynchronous move to complete, Position: " + m_Focuser.Position + " / " + m_PositionOrg);
                WaitFor(500);
            }
        }

        public void MoveFocuserToPosition(string testName, int newPosition)
        {
            DateTime l_StartTime, l_EndTime;

            LogDebug(testName, $"New position: {newPosition}");

            // Confirm that the focuser is not moving
            LogCallToDriver(testName, "About to get IsMoving property");
            if (m_Focuser.IsMoving) // This is an issue as we are expecting the focuser to be not moving
            {
                LogIssue(testName, "Focuser is already moving before start of Move test, rest of test skipped");
            }
            else // Focuser not moving so proceed with the test
            {
                // Move the focuser
                if (m_Absolute)
                {
                    LogTestAndMessage(testName, "Moving to position: " + newPosition.ToString());
                }
                else
                {
                    LogTestAndMessage(testName, "Moving by: " + newPosition.ToString());
                }

                Status(StatusType.staAction, "Moving to new position");
                l_StartTime = DateTime.Now;
                LogCallToDriver(testName, "About to call Move method");
                m_Focuser.Move(newPosition); // Move the focuser
                l_EndTime = DateTime.Now;

                if (l_EndTime.Subtract(l_StartTime).TotalMilliseconds > 1000) // Move took more than 1 second so assume a synchronous call
                {
                    LogDebug(testName, $"Synchronous call behaviour");
                    // Confirm that IsMoving is false
                    LogCallToDriver(testName, "About to get IsMoving property");
                    if (m_Focuser.IsMoving) LogIssue(testName, "Synchronous move expected but focuser is moving after return from Focuser.Move");
                    else LogTestAndMessage(testName, "Synchronous move found");
                }
                else // Move took less than 1 second so assume an asynchronous call
                {
                    LogDebug(testName, $"Asynchronous call behaviour");
                    Status(StatusType.staStatus, "Waiting for asynchronous move to complete");
                    LogCallToDriver(testName, "About to get IsMoving and Position properties repeatedly");
                    do
                    {
                        if (m_AbsolutePositionOK)
                        {
                            Status(StatusType.staStatus, $"Waiting for asynchronous move to complete, Position: {m_Focuser.Position} / { newPosition}, IsMoving: {m_Focuser.IsMoving}");
                            LogDebug(testName, $"Waiting for asynchronous move to complete, Position: {m_Focuser.Position} / { newPosition}, IsMoving: {m_Focuser.IsMoving}");
                        }
                        WaitFor(100);
                    }
                    while (m_Focuser.IsMoving & !cancellationToken.IsCancellationRequested);
                    LogDebug(testName, $"Final position: {m_Focuser.Position}, IsMoving: {m_Focuser.IsMoving}");

                    LogTestAndMessage(testName, "Asynchronous move completed");
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
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
            Status(StatusType.staTest, "");
        }

        private void FocuserPerformanceTest(FocuserPropertyMethod p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
            float l_Single;
            bool l_Boolean;
            double l_Rate;
            Status(StatusType.staTest, "Performance Test");
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
                        case FocuserPropertyMethod.IsMoving:
                            {
                                l_Boolean = m_Focuser.IsMoving;
                                break;
                            }

                        case FocuserPropertyMethod.Position:
                            {
                                l_Single = m_Focuser.Position;
                                break;
                            }

                        case FocuserPropertyMethod.Temperature:
                            {
                                l_Rate = m_Focuser.Temperature;
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
                        Status(StatusType.staStatus, l_Count + " transactions in " + l_ElapsedTime.ToString("0") + " seconds");
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

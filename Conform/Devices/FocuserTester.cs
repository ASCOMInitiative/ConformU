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
using ASCOM.Standard.Interfaces;
using System.Threading;
using ASCOM.Standard.COM.DriverAccess;
using ASCOM.Standard.AlpacaClients;
using System.Runtime.InteropServices;
using ASCOM;

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
        private bool m_Absolute, m_IsMoving, m_Link, m_TempComp, m_TempCompAvailable;
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
        public FocuserTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
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
                    if (m_Focuser is not null) m_Focuser.Dispose();
                    m_Focuser = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion



        public override void CheckInitialise()
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
                        m_Focuser = new AlpacaFocuser(settings.AlpacaConfiguration.AccessServiceType.ToString(), settings.AlpacaDevice.IpAddress, settings.AlpacaDevice.IpPort, settings.AlpacaDevice.AlpacaDeviceNumber, settings.AlpacaConfiguration.LogCalls ? logger : null);
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_Focuser = new FocuserFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                m_Focuser = new Focuser(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");
                baseClassDevice = m_Focuser; // Assign the driver to the base class

                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
                g_Stop = false;
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
                LogCallToDriver("Connected", "About to get Connected property");
                return m_Focuser.Connected;
            }
            set
            {
                try
                {
                    LogMsg("Connected", MessageLevel.msgDebug, "Setting connected state to: " + value.ToString());
                    LogCallToDriver("Connected", "About to set Link property");
                    m_Focuser.Link = value;
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully changed connected state");
                }
                catch (Exception ex)
                {
                    LogMsg("Connected", MessageLevel.msgError, "Error changing focuser connected state: " + ex.ToString());
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
                LogMsg("Absolute", MessageLevel.msgOK, m_Absolute.ToString());
            }
            catch (Exception ex)
            {
                HandleException("Absolute", MemberType.Property, Required.Mandatory, ex, "");
            }

            // IsMoving - Required
            try
            {
                LogCallToDriver("IsMoving", "About to get IsMoving property");
                m_CanReadIsMoving = false;
                m_IsMoving = m_Focuser.IsMoving;
                if (!m_IsMoving)
                {
                    LogMsg("IsMoving", MessageLevel.msgOK, m_IsMoving.ToString());
                    m_CanReadIsMoving = true;
                }
                else
                    LogMsg("IsMoving", MessageLevel.msgError, "IsMoving is True at start of tests and it should be false");
            }
            catch (Exception ex)
            {
                HandleException("IsMoving", MemberType.Property, Required.Mandatory, ex, "");
            }

            // MaxStep - Required
            try
            {
                LogCallToDriver("MaxStep", "About to get MaxStep property");
                m_MaxStep = m_Focuser.MaxStep;
                LogMsg("MaxStep", MessageLevel.msgOK, m_MaxStep.ToString());
            }
            catch (Exception ex)
            {
                HandleException("MaxStep", MemberType.Property, Required.Mandatory, ex, "");
            }

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
                            LogMsg("MaxIncrement", MessageLevel.msgIssue, "MaxIncrement must be at least 1, actual value: " + m_MaxIncrement);
                            break;
                        }

                    case object _ when m_MaxIncrement > m_MaxStep:
                        {
                            LogMsg("MaxIncrement", MessageLevel.msgIssue, "MaxIncrement is greater than MaxStep and shouldn't be: " + m_MaxIncrement);
                            break;
                        }

                    default:
                        {
                            LogMsg("MaxIncrement", MessageLevel.msgOK, m_MaxIncrement.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("MaxIncrement", MemberType.Property, Required.Mandatory, ex, "");
            }

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
                                LogMsg("", MessageLevel.msgWarning, "Position is < 0, actual value: " + m_Position.ToString());
                                break;
                            }

                        case object _ when m_Position > m_MaxStep // > highest position
                 :
                            {
                                LogMsg("", MessageLevel.msgWarning, "Position is > MaxStep, actual value: " + m_Position.ToString());
                                break;
                            }

                        default:
                            {
                                LogMsg("Position", MessageLevel.msgOK, m_Position.ToString());
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
                try
                {
                    LogCallToDriver("Position", "About to get Position property");
                    m_Position = m_Focuser.Position;
                    LogMsg("Position", MessageLevel.msgIssue, "This is a relative focuser but it didn't raise an exception for Focuser.Position");
                }
                catch (Exception ex)
                {
                    HandleException("Position", MemberType.Property, Required.MustNotBeImplemented, ex, "Position must not be implemented for a relative focuser");
                }

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
                            LogMsg("StepSize", MessageLevel.msgIssue, "StepSize must be > 0.0, actual value: " + m_StepSize);
                            break;
                        }

                    default:
                        {
                            LogMsg("StepSize", MessageLevel.msgOK, m_StepSize.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("StepSize", MemberType.Property, Required.Optional, ex, "");
            }

            // TempCompAvailable - Required
            try
            {
                LogCallToDriver("TempCompAvailable", "About to get TempCompAvailable property");
                m_TempCompAvailable = m_Focuser.TempCompAvailable;
                LogMsg("TempCompAvailable", MessageLevel.msgOK, m_TempCompAvailable.ToString());
            }
            catch (Exception ex)
            {
                HandleException("StepSize", MemberType.Property, Required.Mandatory, ex, "");
            }

            // TempComp Read - Required
            try
            {
                LogCallToDriver("TempComp Read", "About to get TempComp property");
                m_TempComp = m_Focuser.TempComp;
                if (m_TempComp & !m_TempCompAvailable)
                    LogMsg("TempComp Read", MessageLevel.msgIssue, "TempComp is True when TempCompAvailable is False - this should not be so");
                else
                    LogMsg("TempComp Read", MessageLevel.msgOK, m_TempComp.ToString());
            }
            catch (Exception ex)
            {
                HandleException("TempComp Read", MemberType.Property, Required.Mandatory, ex, "");
            }

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
                    LogMsg("TempComp Write", MessageLevel.msgOK, "Successfully turned temperature compensation on");
                    m_TempCompTrueOK = true; // Set to true to indicate TempComp can be successfully set to True
                                             // Turn compensation off
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    m_Focuser.TempComp = false;
                    LogMsg("TempComp Write", MessageLevel.msgOK, "Successfully turned temperature compensation off");
                    m_TempCompFalseOK = true;
                }
                catch (Exception ex)
                {
                    HandleException("TempComp Write", MemberType.Property, Required.MustBeImplemented, ex, "Temperature compensation is available but");
                }
            }
            else
                try
                {
                    LogCallToDriver("TempComp Write", "About to set TempComp property");
                    m_Focuser.TempComp = true;
                    LogMsg("TempComp Write", MessageLevel.msgIssue, "Temperature compensation is not available but no exception was raised when TempComp was set True");
                }
                catch (Exception ex)
                {
                    HandleException("TempComp Write", MemberType.Property, Required.MustNotBeImplemented, ex, "Temperature compensation is not available");
                }

            // Restore original TempComp setting if possible
            LogCallToDriver("TempComp Write", "About to set TempComp property");
            try
            {
                m_Focuser.TempComp = m_TempComp;
            }
            catch
            {
            }

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
                            LogMsg("Temperature", MessageLevel.msgWarning, "Temperature < -50.0, - possibly an issue, actual value: " + m_Temperature.ToString());
                            break;
                        }

                    case object _ when m_Temperature >= 50.0 // Probably a bad value
             :
                        {
                            LogMsg("Temperature", MessageLevel.msgWarning, "Temperature > 50.0, - possibly an issue, actual value: " + m_Temperature.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg("Temperature", MessageLevel.msgOK, m_Temperature.ToString());
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
                LogMsg("Halt", MessageLevel.msgOK, "Focuser halted OK");
            }
            catch (COMException ex)
            {
                switch (ex.ErrorCode)
                {
                    case object _ when ex.ErrorCode == g_ExNotImplemented:
                    case object _ when ex.ErrorCode == ErrorCodes.NotImplemented:
                        {
                            LogMsg("Halt", MessageLevel.msgOK, "COM Exception - Halt is not supported by this focuser");
                            break;
                        }

                    default:
                        {
                            LogMsg("Halt", MessageLevel.msgError, $"{EX_COM}{ex.Message}{ex.ErrorCode: X8}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException("Halt", MemberType.Method, Required.Optional, ex, "");
            }

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
                                LogMsg("Move - TempComp True", MessageLevel.msgError, "TempComp is True but no exception is thrown by the Move Method - See Focuser.TempComp entry in Platform help file");
                            }
                            catch (COMException)
                            {
                                LogMsg("Move - TempComp True", MessageLevel.msgOK, "COM Exception correctly raised as expected");
                            }
                            catch (ASCOM.InvalidOperationException)
                            {
                                LogMsg("Move - TempComp True", MessageLevel.msgOK, ".NET InvalidOperation Exception correctly raised as expected");
                            }
                            catch (System.InvalidOperationException)
                            {
                                LogMsg("Move - TempComp True", MessageLevel.msgIssue, "Received System.InvalidOperationException instead of expected ASCOM.InvalidOperationException");
                            }
                            catch (Exception ex)
                            {
                                LogMsg("Move - TempComp True", MessageLevel.msgError, "Unexpected .NET Exception: " + ex.Message);
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
                            LogMsgError("Move - TempComp True", string.Format("Unknown interface version returned {0}, Move test with temperature compensation enabled skipped.", g_InterfaceVersion));
                            break;
                        }
                }

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
                                    LogMsg("Move - To 0", MessageLevel.msgOK, string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - To 0", "About to get Position property");
                                    LogMsg("Move - To 0", MessageLevel.msgInfo, string.Format("Move was within {0} counts of desired position", m_Focuser.Position));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - To 0", MemberType.Method, Required.Mandatory, ex, "");
                    }

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
                                    LogMsg("Move - Below 0", MessageLevel.msgOK, string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - Below 0", "About to get Position property");
                                    LogMsg("Move - Below 0", MessageLevel.msgError, string.Format("Move was permitted below position 0: {0} ", m_Focuser.Position));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - Below 0", MemberType.Method, Required.Mandatory, ex, "Move should fail gracefully by just moving to position 0; it should not throw an exception");
                    }

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
                                    LogMsg("Move - To MaxStep", MessageLevel.msgOK, string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - To MaxStep", "About to get Position property");
                                    LogMsg("Move - To MaxStep", MessageLevel.msgInfo, string.Format("Move position: {0}, within {1} counts of desired position", m_Focuser.Position, m_Focuser.Position - m_MaxStep));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - To MaxStep", MemberType.Method, Required.Mandatory, ex, "");
                    }

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
                                    LogMsg("Move - Above Maxstep", MessageLevel.msgOK, string.Format("Moved to {0}", m_Focuser.Position));
                                    break;
                                }

                            default:
                                {
                                    LogCallToDriver("Move - Above MaxStep", "About to get Position property");
                                    LogMsg("Move - Above Maxstep", MessageLevel.msgError, string.Format("Moved to {0}, {1} steps from MaxStep ", m_Focuser.Position, m_Focuser.Position - m_MaxStep));
                                    break;
                                }
                        }
                    }
                    catch (Exception ex)
                    {
                        HandleException("Move - Above Maxstep", MemberType.Method, Required.Mandatory, ex, "Move should fail gracefully by just moving to position MaxStep; it should not throw an exception");
                    }
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
                m_Position = m_PositionOrg + System.Convert.ToInt32(m_MaxStep / (double)10); // Move by 1/10 of the maximum focus distance out 
                if (m_Position >= m_MaxStep)
                    m_Position = m_PositionOrg - System.Convert.ToInt32(m_MaxStep / (double)10);// Move by 1/10 of the maximum focus distance in
                                                                                                // Apply the MaxIncrement check
                if (Math.Abs(m_Position - m_PositionOrg) > m_MaxIncrement)
                    m_Position = m_PositionOrg + m_MaxIncrement;
            }
            else
            {
                m_Position = System.Convert.ToInt32(m_MaxIncrement / (double)10);
                // Apply the MaxIncrement check
                if (m_Position > m_MaxIncrement)
                    m_Position = m_MaxIncrement;
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
                            LogMsg(testName, MessageLevel.msgOK, "Absolute move OK");
                            break;
                        }

                    default:
                        {
                            LogCallToDriver(testName, "About to get Position property");
                            LogMsg(testName, MessageLevel.msgInfo, $"Move was within {m_Focuser.Position - m_Position} counts of desired position");
                            break;
                        }
                }
            }
            else
                LogMsg(testName, MessageLevel.msgOK, "Relative move OK");

            Status(StatusType.staStatus, "");
            Status(StatusType.staAction, "Returning to original position: " + m_PositionOrg);
            LogMsg(testName, MessageLevel.msgInfo, "Returning to original position: " + m_PositionOrg);
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
            while (m_Focuser.IsMoving & (!g_Stop))
            {
                if (m_AbsolutePositionOK)
                    Status(StatusType.staStatus, "Waiting for asynchronous move to complete, Position: " + m_Focuser.Position + " / " + m_PositionOrg);
                WaitFor(500);
            }
        }

        public void MoveFocuserToPosition(string testName, int newPosition)
        {
            DateTime l_StartTime, l_EndTime;

            // Confirm that the focuser is not moving
            LogCallToDriver(testName, "About to get IsMoving property");
            if (m_Focuser.IsMoving)
                LogMsg(testName, MessageLevel.msgIssue, "Focuser is already moving before start of Move test, rest of test skipped");
            else
            {
                // Move the focuser
                if (m_Absolute)
                    LogMsg(testName, MessageLevel.msgComment, "Moving to position: " + newPosition.ToString());
                else
                    LogMsg(testName, MessageLevel.msgComment, "Moving by: " + newPosition.ToString());

                Status(StatusType.staAction, "Moving to new position");
                l_StartTime = DateTime.Now;
                LogCallToDriver(testName, "About to call Move method");
                m_Focuser.Move(newPosition); // Move the focuser
                l_EndTime = DateTime.Now;

                if (l_EndTime.Subtract(l_StartTime).TotalMilliseconds > 1000)
                {
                    // Confirm that IsMoving is false
                    LogCallToDriver(testName, "About to get IsMoving property");
                    if (m_Focuser.IsMoving)
                        LogMsg(testName, MessageLevel.msgIssue, "Synchronous move expected but focuser is moving after return from Focuser.Move");
                    else
                        LogMsg(testName, MessageLevel.msgComment, "Synchronous move found");
                }
                else
                {
                    Status(StatusType.staStatus, "Waiting for asynchronous move to complete");
                    LogCallToDriver(testName, "About to get IsMoving and Position properties repeatedly");
                    while ((m_Focuser.IsMoving & (!g_Stop)))
                    {
                        if (m_AbsolutePositionOK)
                            Status(StatusType.staStatus, "Waiting for asynchronous move to complete, Position: " + m_Focuser.Position + " / " + newPosition);
                        WaitFor(500);
                    }
                    LogMsg(testName, MessageLevel.msgComment, "Asynchronous move found");
                }
            }
        }

        public override void CheckPerformance()
        {
            // Position
            if (m_AbsolutePositionOK)
                FocuserPerformanceTest(FocuserPropertyMethod.Position, "Position");
            else
                LogMsg("Position", MessageLevel.msgInfo, "Skipping test as property is not supported");

            // IsMoving
            if (m_CanReadIsMoving)
                FocuserPerformanceTest(FocuserPropertyMethod.IsMoving, "IsMoving");
            else
                LogMsg("IsMoving", MessageLevel.msgInfo, "Skipping test as property is not supported");

            // Temperature
            if (m_CanReadTemperature)
                FocuserPerformanceTest(FocuserPropertyMethod.Temperature, "Temperature");
            else
                LogMsg("Temperature", MessageLevel.msgInfo, "Skipping test as property is not supported");
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
                                LogMsg(p_Name, MessageLevel.msgError, "FocuserPerformanceTest: Unknown test type " + p_Type.ToString());
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
                LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " + ex.Message);
            }
        }
    }
}

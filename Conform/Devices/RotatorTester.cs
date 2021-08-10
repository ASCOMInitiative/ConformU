using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic; // Install-Package Microsoft.VisualBasic
using Microsoft.VisualBasic.CompilerServices; // Install-Package Microsoft.VisualBasic
using static ConformU.ConformConstants;
using System.Threading;
using ASCOM.Standard.Interfaces;
using ASCOM.Standard.AlpacaClients;

namespace ConformU
{
    internal partial class RotatorTester : DeviceTesterBaseClass
    {

        #region Variables and Constants
        private const double ROTATOR_WAIT_LIMIT = 30.0d;
        private const double ROTATOR_OK_TOLERANCE = 1.0d;
        private const double ROTATOR_INFO_TOLERANCE = 2.0d;
        private const float ROTATOR_POSITION_TOLERANCE = 0.001f; // Degrees

        // Rotator variables
        private bool m_CanReadIsMoving, canReadPosition, m_CanReadTargetPosition, m_CanReadStepSize;
        private bool m_CanReverse, m_IsMoving;
        private float m_TargetPosition, m_RotatorStepSize, m_RotatorPosition, mechanicalPosition;
        private bool m_Reverse;
        private bool m_LastMoveWasAsync;
        private bool canReadMechanicalPosition;

        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        private IRotatorV3 m_Rotator;

        #endregion

        #region Enums
        private enum RotatorPropertyMethod
        {
            Position,
            TargetPosition,
            StepSize,
            Move,
            MoveAbsolute,
            IsMoving,
            MoveMechanical
        }
        #endregion

        #region New and Dispose
        public RotatorTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, true, true, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
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
                    if (m_Rotator is not null) m_Rotator.Dispose();
                    m_Rotator = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion


        #region Code

        public override void CheckInitialise()
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

        }

        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_Rotator = new AlpacaRotator("http", settings.AlpacaDevice.IpAddress, settings.AlpacaDevice.IpPort, settings.AlpacaDevice.AlpacaDeviceNumber, logger);
                        logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        m_Rotator = new ASCOM.Standard.COM.DriverAccess.Rotator(settings.ComDevice.ProgId);
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogMsg("CreateDevice", MessageLevel.Debug, "Successfully created driver");

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

        public override bool Connected
        {
            get
            {
                bool ConnectedRet = default;
                LogCallToDriver("Connected", "About to get Connected property");
                ConnectedRet = Conversions.ToBoolean(m_Rotator.Connected);
                return ConnectedRet;
            }

            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                m_Rotator.Connected = value;
            }
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_Rotator, DeviceType.Rotator);
        }

        public override void ReadCanProperties()
        {
            try
            {
                LogCallToDriver("CanReverse", "About to get CanReverse property");
                m_CanReverse = Conversions.ToBoolean(m_Rotator.CanReverse);
                LogMsg("CanReverse", MessageLevel.OK, m_CanReverse.ToString());
            }
            catch (Exception ex)
            {
                HandleException("CanReverse", MemberType.Property, Required.Mandatory, ex, "");
            }
        }

        public override void PreRunCheck()
        {
            DateTime l_Now;
            // Get the rotator into a standard state
            g_Stop = true;
            LogCallToDriver("PreRunCheck", "About to call Halt method");
            try
            {
                m_Rotator.Halt();
            }
            catch
            {
            } // Stop any movement

            l_Now = DateAndTime.Now;
            try // Confirm that rotator is not moving or wait for it to stop
            {
                Status(StatusType.staAction, "Waiting up to " + ROTATOR_WAIT_LIMIT + " seconds for rotator to stop moving");
                LogCallToDriver("CanReverse", "About to get IsMoving property repeatedly");
                do
                {
                    WaitFor(500);
                    Status(StatusType.staStatus, DateAndTime.Now.Subtract(l_Now).TotalSeconds + "/" + ROTATOR_WAIT_LIMIT);
                }
                while (!(!m_Rotator.IsMoving | DateAndTime.Now.Subtract(l_Now).TotalSeconds > ROTATOR_WAIT_LIMIT));
                if (Conversions.ToBoolean(!m_Rotator.IsMoving)) // Rotator is stopped so OK
                {
                    g_Stop = false; // Clear stop flag to allow other tests to run
                }
                else // Report error message and don't do other tests
                {
                    LogMsg("Pre-run Check", MessageLevel.Error, "Rotator still moving after " + ROTATOR_WAIT_LIMIT + "seconds, IsMoving stuck on?");
                }

                LogMsg("Pre-run Check", MessageLevel.OK, "Rotator is stationary");
            }
            catch 
            {
                // Don't report errors at this point
            }
        }

        public override void CheckProperties()
        {
            // IsMoving - Optional (V1,V2), Mandatory (V3)
            try
            {
                m_CanReadIsMoving = false;
                LogCallToDriver("IsMoving", "About to get IsMoving property");
                m_IsMoving = Conversions.ToBoolean(m_Rotator.IsMoving);
                m_CanReadIsMoving = true; // Can read OK, doesn't generate an exception
                if (m_IsMoving)
                {
                    LogMsg("IsMoving", MessageLevel.Error, "IsMoving is True before any movement has been commanded!");
                    LogMsg("IsMoving", MessageLevel.Info, "Further tests have been skipped");
                    g_Stop = true;
                }
                else
                {
                    LogMsg("IsMoving", MessageLevel.OK, m_IsMoving.ToString());
                }
            }
            catch (Exception ex)
            {
                if (g_InterfaceVersion < 3)
                {
                    HandleException("IsMoving", MemberType.Property, Required.Optional, ex, "");
                }
                else
                {
                    HandleException("IsMoving", MemberType.Property, Required.Mandatory, ex, "");
                }
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Position - Optional (V1,V2), Mandatory (V3)
            m_RotatorPosition = RotatorPropertyTestSingle(RotatorPropertyMethod.Position, "Position", 0.0f, 360.0f, Required.Mandatory);
            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetPosition - Optional (V1,V2), Mandatory (V3)
            m_TargetPosition = RotatorPropertyTestSingle(RotatorPropertyMethod.TargetPosition, "TargetPosition", 0.0f, 360.0f, Required.Mandatory);
            if (cancellationToken.IsCancellationRequested)
                return;

            // StepSize - Optional (V1,V2 and V3)
            m_RotatorStepSize = RotatorPropertyTestSingle(RotatorPropertyMethod.StepSize, "StepSize", 0.0f, 360.0f, Required.Optional);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Reverse Read - Optional if CanReverse is False, Mandatory if CanReverse is True (V1,V2), Mandatory (V3)
            try
            {
                LogCallToDriver("Reverse", "About to get Reverse property");
                m_Reverse = Conversions.ToBoolean(m_Rotator.Reverse);
                if (m_CanReverse)
                {
                    LogMsg("Reverse Read", MessageLevel.OK, m_Reverse.ToString());
                }
                else
                {
                    LogMsg("Reverse Read", MessageLevel.Error, "CanReverse is false but no exception generated");
                }
            }
            catch (Exception ex)
            {
                if (g_InterfaceVersion < 3) // Can be optional in IRotatorV1 and V2
                {
                    if (m_CanReverse)
                    {
                        HandleException("Reverse Read", MemberType.Property, Required.MustBeImplemented, ex, "when CanReverse is True");
                    }
                    else
                    {
                        HandleException("Reverse Read", MemberType.Property, Required.MustNotBeImplemented, ex, "when CanReverse is False");
                    }
                }
                else // Mandatory in IRotatorV3
                {
                    HandleException("Reverse Read", MemberType.Property, Required.Mandatory, ex, "");
                }
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Reverse Write - Optional if CanReverse is False, Mandatory if CanReverse is True (V1,V2), Mandatory (V3)
            try
            {
                if (m_Reverse) // Try and set reverse to the opposite state
                {
                    LogCallToDriver("Reverse", "About to set Reverse property");
                    m_Rotator.Reverse = false;
                }
                else
                {
                    LogCallToDriver("Reverse", "About to set Reverse property");
                    m_Rotator.Reverse = true;
                }

                LogCallToDriver("Reverse", "About to set Reverse property");
                m_Rotator.Reverse = m_Reverse; // Restore original value
                if (m_CanReverse)
                {
                    LogMsg("Reverse Write", MessageLevel.OK, "Reverse state successfully changed and restored");
                }
                else
                {
                    LogMsg("Reverse Write", MessageLevel.Error, "CanReverse is false but no exception generated");
                }
            }
            catch (Exception ex)
            {
                if (g_InterfaceVersion < 3) // Can be optional in IRotatorV1 and V2
                {
                    if (m_CanReverse)
                    {
                        HandleException("Reverse Write", MemberType.Property, Required.MustBeImplemented, ex, "when CanReverse is True");
                    }
                    else
                    {
                        HandleException("Reverse Write", MemberType.Property, Required.MustNotBeImplemented, ex, "when CanReverse is False");
                    }
                }
                else // Mandatory in IRotatorV3
                {
                    HandleException("Reverse Write", MemberType.Property, Required.Mandatory, ex, "");
                }
            }

            // Test MechanicalPosition introduced in IRotatorV3
            if (g_InterfaceVersion >= 3)
            {
                try
                {
                    canReadMechanicalPosition = false;
                    LogCallToDriver("MechanicalPosition", "About to set MechanicalPosition property");
                    mechanicalPosition = Conversions.ToSingle(m_Rotator.MechanicalPosition);
                    canReadMechanicalPosition = true; // Can read mechanical position OK, doesn't generate an exception

                    // Successfully retrieved a value
                    switch (mechanicalPosition)
                    {
                        case var @case when @case < 0.0f: // Lower than minimum value
                            {
                                LogMsg("MechanicalPosition", MessageLevel.Error, "Invalid value: " + mechanicalPosition.ToString());
                                break;
                            }

                        case var case1 when case1 >= 360.0f: // Higher than maximum value
                            {
                                LogMsg("MechanicalPosition", MessageLevel.Error, "Invalid value: " + mechanicalPosition.ToString()); // OK value
                                break;
                            }

                        default:
                            {
                                LogMsg("MechanicalPosition", MessageLevel.OK, mechanicalPosition.ToString());
                                break;
                            }
                    }

                    // For information show the sync offset, if possible, using OFFSET = SKYPOSITION - MECHANICALPOSITION
                    if (canReadPosition) // Can read synced position and mechanical position
                    {
                        LogMsgInfo("MechanicalPosition", $"Rotator sync offset: {m_RotatorPosition - mechanicalPosition}");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("MechanicalPosition", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
        }

        private float RotatorPropertyTestSingle(RotatorPropertyMethod p_Type, string p_Name, float p_Min, float p_Max, Required requiredIRotatorV3State)
        {
            float RotatorPropertyTestSingleRet = default;
            Required requiredState;

            // Handle properties that were optional in IRotatorV1 and IRotoatorV2 but may have become mandatory in IRotatorV3
            if (g_InterfaceVersion < 3)
            {
                requiredState = Required.Optional; // Default optional state for early versions
            }
            else
            {
                requiredState = requiredIRotatorV3State;
            } // Use the supplied required state for this specific property

            try
            {
                RotatorPropertyTestSingleRet = 0.0f;
                LogCallToDriver(p_Name, $"About to get {p_Name} property");
                switch (p_Type)
                {
                    case RotatorPropertyMethod.Position:
                        {
                            canReadPosition = false;
                            RotatorPropertyTestSingleRet = Conversions.ToSingle(m_Rotator.Position);
                            canReadPosition = true; // Can read position OK, doesn't generate an exception
                            break;
                        }

                    case RotatorPropertyMethod.StepSize:
                        {
                            m_CanReadStepSize = false;
                            RotatorPropertyTestSingleRet = Conversions.ToSingle(m_Rotator.StepSize);
                            m_CanReadStepSize = true;
                            break;
                        }

                    case RotatorPropertyMethod.TargetPosition:
                        {
                            m_CanReadTargetPosition = false;
                            RotatorPropertyTestSingleRet = Conversions.ToSingle(m_Rotator.TargetPosition);
                            m_CanReadTargetPosition = true;
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.Error, "RotatorPropertyTestSingle: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (RotatorPropertyTestSingleRet)
                {
                    case var @case when @case < p_Min: // Lower than minimum value
                        {
                            LogMsg(p_Name, MessageLevel.Error, "Invalid value: " + RotatorPropertyTestSingleRet.ToString());
                            break;
                        }

                    case var case1 when case1 >= p_Max: // Higher than maximum value
                        {
                            LogMsg(p_Name, MessageLevel.Error, "Invalid value: " + RotatorPropertyTestSingleRet.ToString()); // OK value
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.OK, RotatorPropertyTestSingleRet.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, requiredState, ex, "");
            }

            return RotatorPropertyTestSingleRet;
        }

        public override void CheckMethods()
        {
            LogCallToDriver("AccessChecks", "About to get Connected property");
            LogMsg("CheckMethods", MessageLevel.Debug, "Rotator is connected: " + m_Rotator.Connected.ToString());

            // Halt - Optional (V1,V2 and V3)
            try
            {
                LogCallToDriver("Halt", $"About to call Halt method");
                m_Rotator.Halt();
                LogMsg("Halt", MessageLevel.OK, "Halt command successful");
            }
            catch (Exception ex)
            {
                HandleException("Halt", MemberType.Method, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // MoveAbsolute - Optional (V1,V2), Mandatory (V3)
            RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 45.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 135.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 225.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 315.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", (float)-405.0d, "Movement to large negative angle -405 degrees");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.MoveAbsolute, "MoveAbsolute", 405.0f, "Movement to large positive angle 405 degrees");
            if (cancellationToken.IsCancellationRequested)
                return;

            // Move - Optional (V1,V2), Mandatory (V3)
            RelativeMoveTest(10.0f);
            if (cancellationToken.IsCancellationRequested)
                return;
            RelativeMoveTest(40.0f);
            if (cancellationToken.IsCancellationRequested)
                return;
            RelativeMoveTest(130.0f);
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.Move, "Move", (float)-375.0d, "Movement to large negative angle -375 degrees");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.Move, "Move", 375.0f, "Movement to large positive angle 375 degrees");
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test the new IRotatorV3 methods
            if (g_InterfaceVersion >= 3)
            {
                // Test the MoveMechanical method
                try
                {
                    if (canReadMechanicalPosition & canReadPosition)
                    {
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 45.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 135.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 225.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 315.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", (float)-405.0d, "Movement to large negative angle -405 degrees");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 405.0f, "Movement to large positive angle 405 degrees");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                    else // Message saying we are skipping tests because we can't read required properties
                    {
                        LogMsgInfo("MoveMechanical", "Skipping tests because either the MechanicalPosition or Position property cannot be read.");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("MoveMechanical", MemberType.Method, Required.Mandatory, ex, "");
                }

                // Test the Sync method
                try
                {
                    if (canReadMechanicalPosition & canReadPosition) // Test new IRotaotrV3 methods
                    {
                        RotatorSynctest(90.0f, 90.0f); // Make sure that the rotator can be synced to its mechanical position
                        RotatorSynctest(120.0f, 90.0f); // Test sync to a positive offset
                        RotatorSynctest(60.0f, 90.0f); // Test sync to a negative offset
                        RotatorSynctest(00.0f, 00.0f); // Test sync to zero
                        RotatorSynctest(30.0f, 00.0f); // Test sync to a positive offset
                        RotatorSynctest(330.0f, 00.0f); // Test sync to a negative offset that is through zero
                    }
                    else // Message saying we are skipping tests because we can't read required properties
                    {
                        LogMsgInfo("Sync", "Skipping tests because either the MechanicalPosition or Position property cannot be read.");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Sync", MemberType.Method, Required.Mandatory, ex, "");
                }
            }
        }

        private void RotatorSynctest(float SyncAngle, float MechanicalAngle)
        {
            float syncAngleDifference;
            RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "Sync", MechanicalAngle, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            try
            {
                LogCallToDriver("Sync", $"About to call Sync method");
                m_Rotator.Sync(SyncAngle);
                LogMsgOK("Sync", "Synced OK");

                // Check that Position and MechanicalPosition are now the same
                syncAngleDifference = Conversions.ToSingle(Operators.SubtractObject(m_Rotator.Position, SyncAngle));
                if (Math.Abs(syncAngleDifference) < ROTATOR_POSITION_TOLERANCE)
                {
                    LogMsgOK("Sync", $"Rotator Position has synced to {SyncAngle} OK.");
                }
                else
                {
                    LogMsgIssue("Sync", $"Rotator Position is more than {ROTATOR_POSITION_TOLERANCE} different from requested position {SyncAngle}.");
                }
            }
            catch (Exception ex)
            {
                HandleException("Sync", MemberType.Method, Required.Mandatory, ex, "");
            }
        }

        private void RotatorMoveTest(RotatorPropertyMethod p_type, string p_Name, float p_Value, string p_ExpectErrorMsg)
        {
            float l_RotatorStartPosition = default, rotatorPosition;
            double l_OKLimit, l_PositionOffset;
            LogCallToDriver(p_Name, $"About to get Position property");
            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Start value, position: " + p_Value.ToString("0.000") + " " + m_Rotator.Position.ToString()[Conversions.ToInteger("0.000")]);
            try
            {
                // Move to requested position
                switch (p_type)
                {
                    case RotatorPropertyMethod.Move:
                        {
                            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Reading rotator start position: " + canReadPosition);
                            if (canReadPosition) // Get us to a starting point of 10 degrees
                            {
                                LogCallToDriver(p_Name, $"About to get Position property");
                                l_RotatorStartPosition = Conversions.ToSingle(m_Rotator.Position);
                            }

                            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Starting relative move");
                            LogCallToDriver(p_Name, $"About to call Move method");
                            m_Rotator.Move(p_Value);
                            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Starting relative move");
                            break;
                        }

                    case RotatorPropertyMethod.MoveAbsolute:
                        {
                            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Starting absolute move");
                            l_RotatorStartPosition = 0.0f;
                            LogCallToDriver(p_Name, $"About to call MoveAbsolute method");
                            m_Rotator.MoveAbsolute(p_Value);
                            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Completed absolute move");
                            break;
                        }

                    case RotatorPropertyMethod.MoveMechanical:
                        {
                            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Starting mechanical move");
                            l_RotatorStartPosition = 0.0f;
                            LogCallToDriver(p_Name, $"About to call MoveMechanical method");
                            m_Rotator.MoveMechanical(p_Value);
                            LogMsg("RotatorMoveTest", MessageLevel.Debug, "Completed mechanical move");
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.Error, "RotatorMoveTest: Unknown test type - " + p_type.ToString());
                            break;
                        }
                }

                RotatorWait(p_type, p_Name, p_Value, l_RotatorStartPosition);
                if (m_LastMoveWasAsync) // Asynchronous move
                {
                    switch (p_type)
                    {
                        case RotatorPropertyMethod.Move:
                            {
                                if (canReadPosition)
                                {
                                    LogCallToDriver(p_Name, $"About to get Position property");
                                    LogMsg(p_Name, MessageLevel.OK, $"Asynchronous move successful - moved by {p_Value} degrees to: {m_Rotator.Position} degrees");
                                }
                                else
                                {
                                    LogMsg(p_Name, MessageLevel.OK, "Asynchronous move successful");
                                }

                                break;
                            }

                        case RotatorPropertyMethod.MoveMechanical:
                            {
                                break;
                            }

                        case RotatorPropertyMethod.MoveAbsolute:
                            {
                                if (canReadPosition)
                                {
                                    LogCallToDriver(p_Name, $"About to get Position property");
                                    LogMsg(p_Name, MessageLevel.OK, $"Asynchronous move successful to: {m_Rotator.Position} degrees");
                                }
                                else
                                {
                                    LogMsg(p_Name, MessageLevel.OK, "Asynchronous move successful");
                                }

                                break;
                            }
                    }
                }
                else if (canReadPosition) // Synchronous move
                {
                    LogCallToDriver(p_Name, $"About to get Position property");
                    LogMsg(p_Name, MessageLevel.OK, $"Synchronous move successful to: {m_Rotator.Position} degrees");
                }
                else
                {
                    LogMsg(p_Name, MessageLevel.OK, "Synchronous move successful");
                }

                // Now test whether we got to where we expected to go
                if (canReadPosition)
                {
                    if (m_CanReadStepSize)
                    {
                        l_OKLimit = 1.1d * m_RotatorStepSize; // Set to 110% of step size to allow tolerance on reporting within 1 step of required location
                    }
                    else
                    {
                        l_OKLimit = ROTATOR_OK_TOLERANCE;
                    }

                    LogCallToDriver(p_Name, $"About to get Position property");
                    LogMsg(p_Name + "1", MessageLevel.Debug, "Position, value, start, tolerance: " + m_Rotator.Position.ToString()[Conversions.ToInteger("0.000")] + " " + p_Value.ToString("0.000") + " " + l_RotatorStartPosition.ToString("0.000") + " " + l_OKLimit.ToString("0.000"));
                    LogCallToDriver(p_Name, $"About to get Position property");
                    rotatorPosition = Conversions.ToSingle(m_Rotator.Position);
                    if (g_InterfaceVersion < 3) // Interface V1 and V2 behaviour
                    {
                        if (rotatorPosition < 0.0d)
                            LogMsg(p_Name, MessageLevel.Info, "Rotator supports angles < 0.0");
                        if (rotatorPosition > 360.0d)
                            LogMsg(p_Name, MessageLevel.Info, "Rotator supports angles > 360.0");
                    }
                    else if (rotatorPosition < 0.0d | rotatorPosition >= 360.0d) // Interface V3 behaviour (Position must be 0..359.99999...)
                        LogMsg(p_Name, MessageLevel.Error, $"Rotator position {rotatorPosition:0.000} is outside the valid range: 0.0 to 359.99999...");

                    // Get the relevant position value
                    if (p_type == RotatorPropertyMethod.MoveMechanical) // Use the MechanicalPosition property
                    {
                        LogCallToDriver(p_Name, $"About to get MechanicalPosition property");
                        rotatorPosition = Conversions.ToSingle(m_Rotator.MechanicalPosition);
                    }
                    else // Use the Position property for all other methods
                    {
                        LogCallToDriver(p_Name, $"About to get Position property");
                        rotatorPosition = Conversions.ToSingle(m_Rotator.Position);
                    }
                    // Calculate the position offset from the required position
                    l_PositionOffset = Math.Abs((720.0d + rotatorPosition - (p_Value + l_RotatorStartPosition)) % 360.0d); // Account for rotator positions that report < 0.0 or > 360.0 degrees
                    if (l_PositionOffset > 180.0d)
                        l_PositionOffset = 360.0d - l_PositionOffset; // Cope with positions that return just under the expected value
                    switch (Math.Abs(l_PositionOffset))
                    {
                        case 0.0d:
                            {
                                LogMsg(p_Name, MessageLevel.OK, $"Rotator is at the expected position: {rotatorPosition}");
                                break;
                            }

                        case var @case when 0.0d <= @case && @case <= l_OKLimit:
                            {
                                LogMsg(p_Name, MessageLevel.OK, $"Rotator is within {l_OKLimit.ToString("0.000")} {Interaction.IIf(l_PositionOffset <= 1.0d, " degree", " degrees").ToString()} of the expected position: {rotatorPosition}");
                                break;
                            }

                        case var case1 when 0.0d <= case1 && case1 <= ROTATOR_INFO_TOLERANCE:
                            {
                                LogMsg(p_Name, MessageLevel.Info, $"Rotator is {l_PositionOffset.ToString("0.000")} degrees from expected position: {rotatorPosition}");
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.Issue, $"Rotator is {l_PositionOffset.ToString("0.000")} degrees from expected position {rotatorPosition}, which is more than the conformance value of {ROTATOR_INFO_TOLERANCE.ToString("0.0")} degrees");
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(p_ExpectErrorMsg)) // Test for normal behaviour
                {
                    if (g_InterfaceVersion < 3)
                    {
                        HandleException(p_Name, MemberType.Method, Required.Optional, ex, "");
                    }
                    else
                    {
                        HandleException(p_Name, MemberType.Method, Required.Mandatory, ex, "");
                    }
                }
                else if (IsInvalidValueException(p_Name, ex)) // This is expected to fail because a bad position was used
                                                              // Test for an InvalidValueException and handle if found
                {
                    HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.Mandatory, ex, "", p_ExpectErrorMsg);
                }
                else if (g_InterfaceVersion < 3) // Some other type of exception occurred
                {
                    HandleException(p_Name, MemberType.Method, Required.Optional, ex, "");
                }
                else
                {
                    HandleException(p_Name, MemberType.Method, Required.Mandatory, ex, "");
                }
            }

            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
            Status(StatusType.staTest, "");
        }

        private void RotatorWait(RotatorPropertyMethod p_type, string p_Name, float p_value, float p_RotatorStartPosition)
        {
            LogMsg("RotatorWait", MessageLevel.Debug, "Entered RotatorWait");
            if (m_CanReadIsMoving) // Can read IsMoving so test for asynchronous and synchronous behaviour
            {
                LogMsg("RotatorWait", MessageLevel.Debug, "Can Read IsMoving OK");
                LogCallToDriver(p_Name, $"About to get Ismoving property");
                if (Conversions.ToBoolean(m_Rotator.IsMoving))
                {
                    LogMsg("RotatorWait", MessageLevel.Debug, "Rotator is moving, waiting for move to complete");
                    Status(StatusType.staTest, p_Name + " test");
                    Status(StatusType.staAction, "Waiting for move to complete");
                    LogCallToDriver(p_Name, $"About to get Position and Ismoving properties repeatedly");
                    do
                    {
                        WaitFor(500);
                        if (canReadPosition) // Only do this if position doesn't generate an exception
                        {
                            switch (p_type)
                            {
                                case RotatorPropertyMethod.Move:
                                    {
                                        Status(StatusType.staStatus, $"{Math.Abs(m_Rotator.Position - p_RotatorStartPosition)} / {p_value} relative");
                                        break;
                                    }

                                case RotatorPropertyMethod.MoveAbsolute:
                                    {
                                        Status(StatusType.staStatus, $"{Math.Abs(m_Rotator.Position - p_RotatorStartPosition)} / {Math.Abs(p_value - p_RotatorStartPosition)} absolute");
                                        break;
                                    }
                            }
                        }

                    }
                    while (m_Rotator.IsMoving);
                    LogMsg("RotatorWait", MessageLevel.Debug, "Rotator has stopped moving");
                    Status(StatusType.staAction, "");
                    m_LastMoveWasAsync = true;
                }
                else
                {
                    m_LastMoveWasAsync = false;
                }
            }
            else // Can only test for synchronous move
            {
                LogMsg("RotatorWait", MessageLevel.Debug, "Cannot Read IsMoving");
                m_LastMoveWasAsync = false;
            }
        }

        private void RelativeMoveTest(float p_RelativeStepSize)
        {
            float l_Target;
            if (canReadPosition)
            {
                LogCallToDriver("Move", $"About to get Position property");
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLess(m_Rotator.Position, p_RelativeStepSize, false))) // Set a value that should succeed OK
                {
                    l_Target = p_RelativeStepSize;
                }
                else
                {
                    l_Target = -p_RelativeStepSize;
                }
            }
            else
            {
                l_Target = p_RelativeStepSize;
            }

            RotatorMoveTest(RotatorPropertyMethod.Move, "Move", l_Target, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.Move, "Move", -l_Target, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            // Should now be back where we started

        }

        public override void CheckPerformance()
        {
            // Position
            if (canReadPosition)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.Position, "Position");
            }
            else
            {
                LogMsg("Position", MessageLevel.Info, "Skipping test as property is not supported");
            }

            // TargetPosition
            if (m_CanReadTargetPosition)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.TargetPosition, "TargetPosition");
            }
            else
            {
                LogMsg("TargetPosition", MessageLevel.Info, "Skipping test as property is not supported");
            }

            // StepSize
            if (m_CanReadStepSize)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.StepSize, "StepSize");
            }
            else
            {
                LogMsg("StepSize", MessageLevel.Info, "Skipping test as property is not supported");
            }

            // IsMoving
            if (m_CanReadIsMoving)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.IsMoving, "IsMoving");
            }
            else
            {
                LogMsg("IsMoving", MessageLevel.Info, "Skipping test as property is not supported");
            }
        }

        private void RotatorPerformanceTest(RotatorPropertyMethod p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
            float l_Single;
            bool l_Boolean;
            double l_Rate;
            Status(StatusType.staAction, p_Name);
            try
            {
                l_StartTime = DateAndTime.Now;
                l_Count = 0.0d;
                l_LastElapsedTime = 0.0d;
                do
                {
                    l_Count += 1.0d;
                    switch (p_Type)
                    {
                        case RotatorPropertyMethod.Position:
                            {
                                l_Single = Conversions.ToSingle(m_Rotator.Position);
                                break;
                            }

                        case RotatorPropertyMethod.TargetPosition:
                            {
                                l_Single = Conversions.ToSingle(m_Rotator.TargetPosition);
                                break;
                            }

                        case RotatorPropertyMethod.StepSize:
                            {
                                l_Single = Conversions.ToSingle(m_Rotator.StepSize);
                                break;
                            }

                        case RotatorPropertyMethod.IsMoving:
                            {
                                l_Boolean = Conversions.ToBoolean(m_Rotator.IsMoving);
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.Error, "RotatorPerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateAndTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0d)
                    {
                        Status(StatusType.staStatus, l_Count + " transactions in " + Strings.Format(l_ElapsedTime, "0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (l_ElapsedTime <= PERF_LOOP_TIME);
                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case var @case when @case > 10.0d:
                        {
                            LogMsg(p_Name, MessageLevel.Info, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    case var case1 when 2.0d <= case1 && case1 <= 10.0d:
                        {
                            LogMsg(p_Name, MessageLevel.OK, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    case var case2 when 1.0d <= case2 && case2 <= 2.0d:
                        {
                            LogMsg(p_Name, MessageLevel.Info, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.Info, "Transaction rate: " + Strings.Format(l_Rate, "0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogMsg(p_Name, MessageLevel.Info, "Unable to complete test: " + ex.Message);
            }
        }

        #endregion

    }
}

using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.Tools;
using System;
using System.Reflection;
using System.Threading;

namespace ConformU
{
    internal partial class RotatorTester : DeviceTesterBaseClass
    {

        #region Variables and Constants
        private const double ROTATOR_OK_TOLERANCE = 1.0d;
        private const double ROTATOR_INFO_TOLERANCE = 2.0d;
        private const float ROTATOR_POSITION_TOLERANCE = 0.001f; // Degrees
        private const float ROTATOR_POSITION_UNKNOWN = float.NaN; // Define a constant to represent position unknown. Used when restoring rotator position after testing.

        // Rotator variables
        private bool m_CanReadIsMoving, canReadPosition, m_CanReadTargetPosition, m_CanReadStepSize;
        private bool m_CanReverse, m_IsMoving;
        private float m_RotatorStepSize, m_RotatorPosition, mechanicalPosition;
        private bool m_Reverse;
        private bool m_LastMoveWasAsync;
        private bool canReadMechanicalPosition;
        private float initialPosiiton = ROTATOR_POSITION_UNKNOWN;
        private float initialMechanicalPosiiton = ROTATOR_POSITION_UNKNOWN;
        private float initialSyncOffset = ROTATOR_POSITION_UNKNOWN;

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
        public RotatorTester(ConformConfigurationService conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, true, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
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
                    m_Rotator?.Dispose();
                    m_Rotator = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        #region Code

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

        }

        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        LogInfo("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_Rotator = new AlpacaRotator(
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
                                                    Assembly.GetExecutingAssembly().GetName().Version.ToString(4));

                        LogInfo("CreateDevice", $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComAccessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating Native COM device: {settings.ComDevice.ProgId}");
                                m_Rotator = new RotatorFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                m_Rotator = new Rotator(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }

                        LogInfo("CreateDevice", "Successfully created driver");
                        baseClassDevice = m_Rotator; // Assign the driver to the base class

                        SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                        WaitFor(1000, 100);
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }


            }
            catch (Exception ex)
            {
                LogInfo("CreateDevice", "Exception thrown: " + ex.Message);
                throw; // Re throw exception 
            }

        }

        public override bool Connected
        {
            get
            {
                LogCallToDriver("Connected", "About to get Connected property");
                bool ConnectedRet = m_Rotator.Connected;
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
            base.CheckCommonMethods(m_Rotator, DeviceTypes.Rotator);
        }

        public override void ReadCanProperties()
        {
            try
            {
                LogCallToDriver("CanReverse", "About to get CanReverse property");
                m_CanReverse = m_Rotator.CanReverse;
                LogOK("CanReverse", m_CanReverse.ToString());
            }
            catch (Exception ex)
            {
                HandleException("CanReverse", MemberType.Property, Required.Mandatory, ex, "");
            }
        }

        public override void PreRunCheck()
        {

            // Initialise to the unknown position value
            initialPosiiton = ROTATOR_POSITION_UNKNOWN;

            // Get the rotator into a standard state
            LogCallToDriver("PreRun Check", "About to call Halt method");
            try
            {
                m_Rotator.Halt();
            }
            catch
            {
            } // Stop any movement

            // Confirm that rotator is not moving or wait for it to stop
            try
            {
                SetTest("Pre-run check");
                SetAction($"Waiting up to {settings.RotatorTimeout} seconds for rotator to stop moving");
                LogCallToDriver("PreRun Check", "About to get IsMoving property repeatedly");
                RotatorWait(RotatorPropertyMethod.Move, "Ensuring that movement is stopped", 0, 0);

                if (!m_Rotator.IsMoving)
                {
                    LogOK("Pre-run Check", "Rotator is stationary");

                    // Try to record the current position of the rotator so that it can be restored after testing. If this fails the initial position will be set to unknown value.
                    try
                    {
                        LogCallToDriver("PreRun Check", "About to get Position property");
                        initialPosiiton = m_Rotator.Position;
                        LogOK("Pre-run Check", $"Rotator initial position: {initialPosiiton}");

                        // Attempt to get the rotator's current mechanical position. If this fails the initial mechanical position will be set to unknown value.
                        try
                        {
                            LogCallToDriver("PreRun Check", "About to get MechanicalPosition property");
                            initialMechanicalPosiiton = m_Rotator.MechanicalPosition;
                            initialSyncOffset = (float)Utilities.Range((double)(initialPosiiton - initialMechanicalPosiiton), -180.0, true, 180.0, true);
                            LogOK("Pre-run Check", $"Rotator initial mechanical position: {initialMechanicalPosiiton}, Initial sync offset: {initialSyncOffset}");
                        }
                        catch (Exception ex)
                        {
                            // Don't report errors at this point
                            LogInfo("Pre-run Check", $"Rotator initial mechanical position could not be read: {ex.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Don't report errors at this point
                        LogInfo("Pre-run Check", $"Rotator initial position could not be read: {ex.Message}");
                    }
                }
                else
                    LogIssue("Pre-run Check", $"Rotator still moving after {settings.RotatorTimeout} seconds, could IsMoving be stuck on?");
            }
            catch (Exception)
            {
            }
        }

        public override void PostRunCheck()
        {
            float currentPosition, currentMechanicalPosition, relativeMovement, syncPosition;

            // Restore the initial position of the rotator if possible
            if (Single.IsNaN(initialPosiiton))
                // The initial position could not be determined so log a message to this effect.
                LogMsg("Post-run Check", MessageLevel.TestAndMessage, "The rotator's initial position could not be determined so it is not possible to restore it's initial position.");
            else
                // We have a valid initial position so attempt to reset the rotator to this position.
                try
                {
                    // Get the current position
                    LogCallToDriver("Post-run Check", $"About to get Position property");
                    currentPosition = m_Rotator.Position;
                    LogOK("Post-run Check", $"Current position: {currentPosition}");

                    // Restore the original sync offset, if possible 
                    if (!Single.IsNaN(initialMechanicalPosiiton))
                    {
                        // Get the current mechanical position
                        LogCallToDriver("Post-run Check", $"About to get MechanicalPosition property");
                        currentMechanicalPosition = m_Rotator.MechanicalPosition;
                        LogOK("Post-run Check", $"Current mechanical position: {currentMechanicalPosition}");

                        syncPosition = (float)Utilities.Range((double)(currentMechanicalPosition + initialSyncOffset), 0.0, true, 360.0, false);
                        LogOK("Post-run Check", $"New sync position: {syncPosition}");

                        LogCallToDriver("Post-run Check", $"About to call Sync method. Position: {syncPosition}");
                        m_Rotator.Sync(syncPosition);
                        LogOK("Post-run Check", $"Completed Sync ({initialSyncOffset} degrees) from position: {currentPosition} to {syncPosition}");
                    }

                    // Re-get the current position because the sync will have changed it
                    LogCallToDriver("Post-run Check", $"About to get Position property");
                    currentPosition = m_Rotator.Position;
                    LogOK("Post-run Check", $"New current position: {currentPosition}");

                    // Calculate the smallest relative movement required to get to the initial position
                    relativeMovement = (float)Utilities.Range((double)(initialPosiiton - currentPosition), -180.0, true, 180.0, true);

                    // Move to the starting position
                    LogCallToDriver("Post-run Check", $"About to move by {relativeMovement} to {initialPosiiton}");
                    m_Rotator.Move(relativeMovement);

                    // Wait for the move to complete
                    RotatorWait(RotatorPropertyMethod.Move, "Restoring original position", relativeMovement, currentPosition);

                    if (!m_Rotator.IsMoving)
                        LogOK("Post-run Check", $"Rotator starting position successfully restored to {initialPosiiton}");
                    else
                        LogError("Post-run Check", $"Unable to restore rotator starting position, the rotator is still moving after {settings.RotatorTimeout} seconds. Could IsMoving be stuck on?");
                }
                catch (Exception ex)
                {
                    LogError("Post-run Check", $"Exception: {ex}");
                }
        }

        public override void CheckProperties()
        {
            // IsMoving - Optional (V1,V2), Mandatory (V3)
            try
            {
                m_CanReadIsMoving = false;
                LogCallToDriver("IsMoving", "About to get IsMoving property");
                m_IsMoving = m_Rotator.IsMoving;
                m_CanReadIsMoving = true; // Can read OK, doesn't generate an exception
                if (m_IsMoving)
                {
                    LogIssue("IsMoving", "IsMoving is True before any movement has been commanded!");
                    LogInfo("IsMoving", "Further tests have been skipped");
                }
                else
                {
                    LogOK("IsMoving", m_IsMoving.ToString());
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
            RotatorPropertyTestSingle(RotatorPropertyMethod.TargetPosition, "TargetPosition", 0.0f, 360.0f, Required.Mandatory);
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
                m_Reverse = m_Rotator.Reverse;
                if (m_CanReverse)
                {
                    LogOK("Reverse Read", m_Reverse.ToString());
                }
                else
                {
                    LogIssue("Reverse Read", "CanReverse is false but no exception generated");
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
                    LogOK("Reverse Write", "Reverse state successfully changed and restored");
                }
                else
                {
                    LogIssue("Reverse Write", "CanReverse is false but no exception generated");
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
                    mechanicalPosition = m_Rotator.MechanicalPosition;
                    canReadMechanicalPosition = true; // Can read mechanical position OK, doesn't generate an exception

                    // Successfully retrieved a value
                    switch (mechanicalPosition)
                    {
                        case var @case when @case < 0.0f: // Lower than minimum value
                            {
                                LogIssue("MechanicalPosition", "Invalid value: " + mechanicalPosition.ToString());
                                break;
                            }

                        case var case1 when case1 >= 360.0f: // Higher than maximum value
                            {
                                LogIssue("MechanicalPosition", "Invalid value: " + mechanicalPosition.ToString()); // OK value
                                break;
                            }

                        default:
                            {
                                LogOK("MechanicalPosition", mechanicalPosition.ToString());
                                break;
                            }
                    }

                    // For information show the sync offset, if possible, using OFFSET = SKYPOSITION - MECHANICALPOSITION
                    if (canReadPosition) // Can read synced position and mechanical position
                    {
                        LogInfo("MechanicalPosition", $"Rotator sync offset: {m_RotatorPosition - mechanicalPosition}");
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
                            RotatorPropertyTestSingleRet = m_Rotator.Position;
                            canReadPosition = true; // Can read position OK, doesn't generate an exception
                            break;
                        }

                    case RotatorPropertyMethod.StepSize:
                        {
                            m_CanReadStepSize = false;
                            RotatorPropertyTestSingleRet = m_Rotator.StepSize;
                            m_CanReadStepSize = true;
                            break;
                        }

                    case RotatorPropertyMethod.TargetPosition:
                        {
                            m_CanReadTargetPosition = false;
                            RotatorPropertyTestSingleRet = m_Rotator.TargetPosition;
                            m_CanReadTargetPosition = true;
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "RotatorPropertyTestSingle: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (RotatorPropertyTestSingleRet)
                {
                    case var @case when @case < p_Min: // Lower than minimum value
                        {
                            LogIssue(p_Name, "Invalid value: " + RotatorPropertyTestSingleRet.ToString());
                            break;
                        }

                    case var case1 when case1 >= p_Max: // Higher than maximum value
                        {
                            LogIssue(p_Name, "Invalid value: " + RotatorPropertyTestSingleRet.ToString()); // OK value
                            break;
                        }

                    default:
                        {
                            LogOK(p_Name, RotatorPropertyTestSingleRet.ToString());
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
            LogDebug("CheckMethods", "Rotator is connected: " + m_Rotator.Connected.ToString());

            // Halt - Optional (V1,V2 and V3)
            try
            {
                LogCallToDriver("Halt", $"About to call Halt method");
                m_Rotator.Halt();
                LogOK("Halt", "Halt command successful");
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
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", -405.0f, "Movement to large negative angle -405 degrees");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "MoveMechanical", 405.0f, "Movement to large positive angle 405 degrees");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                    else // Message saying we are skipping tests because we can't read required properties
                    {
                        LogInfo("MoveMechanical", "Skipping tests because either the MechanicalPosition or Position property cannot be read.");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("MoveMechanical", MemberType.Method, Required.Mandatory, ex, "");
                }
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Test the Sync method
                try
                {
                    if (canReadMechanicalPosition & canReadPosition) // Test new IRotaotrV3 methods
                    {
                        RotatorSynctest(90.0f, 90.0f); // Make sure that the rotator can be synced to its mechanical position
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorSynctest(120.0f, 90.0f); // Test sync to a positive offset
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorSynctest(60.0f, 90.0f); // Test sync to a negative offset
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorSynctest(00.0f, 00.0f); // Test sync to zero
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorSynctest(30.0f, 00.0f); // Test sync to a positive offset
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorSynctest(330.0f, 00.0f); // Test sync to a negative offset that is through zero
                    }
                    else // Message saying we are skipping tests because we can't read required properties
                    {
                        LogInfo("Sync", "Skipping tests because either the MechanicalPosition or Position property cannot be read.");
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
                LogOK("Sync", "Synced OK");

                // Check that Position and MechanicalPosition are now the same
                syncAngleDifference = m_Rotator.Position - SyncAngle;
                if (Math.Abs(syncAngleDifference) < ROTATOR_POSITION_TOLERANCE)
                {
                    LogOK("Sync", $"Rotator Position has synced to {SyncAngle} OK.");
                }
                else
                {
                    LogIssue("Sync", $"Rotator Position is {syncAngleDifference} degrees from the requested position {SyncAngle}. Alert tolerance is {ROTATOR_POSITION_TOLERANCE} degrees.");
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
            LogDebug("RotatorMoveTest", "Start value, position: " + p_Value.ToString("0.000") + " " + m_Rotator.Position.ToString("0.000"));

            SetTest(p_Name);
            SetAction("Setting position");
            try
            {
                // Move to requested position
                switch (p_type)
                {
                    case RotatorPropertyMethod.Move:
                        {
                            LogDebug("RotatorMoveTest", "Reading rotator start position: " + canReadPosition);
                            if (canReadPosition) // Get us to a starting point of 10 degrees
                            {
                                LogCallToDriver(p_Name, $"About to get Position property");
                                l_RotatorStartPosition = m_Rotator.Position;
                            }

                            LogDebug("RotatorMoveTest", "Starting relative move");
                            LogCallToDriver(p_Name, $"About to call Move method");
                            m_Rotator.Move(p_Value);
                            LogDebug("RotatorMoveTest", "Starting relative move");
                            break;
                        }

                    case RotatorPropertyMethod.MoveAbsolute:
                        {
                            LogDebug("RotatorMoveTest", "Starting absolute move");
                            l_RotatorStartPosition = 0.0f;
                            LogCallToDriver(p_Name, $"About to call MoveAbsolute method");
                            m_Rotator.MoveAbsolute(p_Value);
                            LogDebug("RotatorMoveTest", "Completed absolute move");
                            break;
                        }

                    case RotatorPropertyMethod.MoveMechanical:
                        {
                            LogDebug("RotatorMoveTest", "Starting mechanical move");
                            l_RotatorStartPosition = 0.0f;
                            LogCallToDriver(p_Name, $"About to call MoveMechanical method");
                            m_Rotator.MoveMechanical(p_Value);
                            LogDebug("RotatorMoveTest", "Completed mechanical move");
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "RotatorMoveTest: Unknown test type - " + p_type.ToString());
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
                                    LogOK(p_Name, $"Asynchronous move successful - moved by {p_Value} degrees to: {m_Rotator.Position} degrees");
                                }
                                else
                                {
                                    LogOK(p_Name, "Asynchronous move successful");
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
                                    LogOK(p_Name, $"Asynchronous move successful to: {m_Rotator.Position} degrees");
                                }
                                else
                                {
                                    LogOK(p_Name, "Asynchronous move successful");
                                }

                                break;
                            }
                    }
                }
                else if (canReadPosition) // Synchronous move
                {
                    LogCallToDriver(p_Name, $"About to get Position property");
                    LogOK(p_Name, $"Synchronous move successful to: {m_Rotator.Position} degrees");
                }
                else
                {
                    LogOK(p_Name, "Synchronous move successful");
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
                    LogDebug(p_Name + "1", "Position, value, start, tolerance: " + m_Rotator.Position.ToString("0.000") + " " + p_Value.ToString("0.000") + " " + l_RotatorStartPosition.ToString("0.000") + " " + l_OKLimit.ToString("0.000"));
                    LogCallToDriver(p_Name, $"About to get Position property");
                    rotatorPosition = m_Rotator.Position;
                    if (g_InterfaceVersion < 3) // Interface V1 and V2 behaviour
                    {
                        if (rotatorPosition < 0.0d)
                            LogInfo(p_Name, "Rotator supports angles < 0.0");
                        if (rotatorPosition > 360.0d)
                            LogInfo(p_Name, "Rotator supports angles > 360.0");
                    }
                    else if (rotatorPosition < 0.0d | rotatorPosition >= 360.0d) // Interface V3 behaviour (Position must be 0..359.99999...)
                        LogIssue(p_Name, $"Rotator position {rotatorPosition:0.000} is outside the valid range: 0.0 to 359.99999...");

                    // Get the relevant position value
                    if (p_type == RotatorPropertyMethod.MoveMechanical) // Use the MechanicalPosition property
                    {
                        LogCallToDriver(p_Name, $"About to get MechanicalPosition property");
                        rotatorPosition = m_Rotator.MechanicalPosition;
                    }
                    else // Use the Position property for all other methods
                    {
                        LogCallToDriver(p_Name, $"About to get Position property");
                        rotatorPosition = m_Rotator.Position;
                    }
                    // Calculate the position offset from the required position
                    l_PositionOffset = Math.Abs((720.0d + rotatorPosition - (p_Value + l_RotatorStartPosition)) % 360.0d); // Account for rotator positions that report < 0.0 or > 360.0 degrees
                    if (l_PositionOffset > 180.0d)
                        l_PositionOffset = 360.0d - l_PositionOffset; // Cope with positions that return just under the expected value
                    switch (Math.Abs(l_PositionOffset))
                    {
                        case 0.0d:
                            {
                                LogOK(p_Name, $"Rotator is at the expected position: {rotatorPosition}");
                                break;
                            }

                        case var @case when 0.0d <= @case && @case <= l_OKLimit:
                            {
                                LogOK(p_Name, $"Rotator is within {l_OKLimit:0.000)} {((l_PositionOffset <= 1.0d) ? " degree" : " degrees")} of the expected position: {rotatorPosition}");
                                break;
                            }

                        case var case1 when 0.0d <= case1 && case1 <= ROTATOR_INFO_TOLERANCE:
                            {
                                LogInfo(p_Name, $"Rotator is {l_PositionOffset:0.000} degrees from expected position: {rotatorPosition}");
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, $"Rotator is {l_PositionOffset:0.000} degrees from expected position {rotatorPosition}, which is more than the conformance value of {ROTATOR_INFO_TOLERANCE:0.0} degrees");
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

            SetAction("");
            SetStatus("");
            SetTest("");
        }

        private void RotatorWait(RotatorPropertyMethod p_type, string p_Name, float p_value, float p_RotatorStartPosition)
        {
            LogDebug("RotatorWait", "Entered RotatorWait");
            if (m_CanReadIsMoving) // Can read IsMoving so test for asynchronous and synchronous behaviour
            {
                LogDebug("RotatorWait", "Can Read IsMoving OK");
                LogCallToDriver(p_Name, $"About to get IsMoving property");
                if (m_Rotator.IsMoving)
                {
                    LogDebug("RotatorWait", "Rotator is moving, waiting for move to complete");
                    SetTest(p_Name);
                    SetAction("Waiting for move to complete");
                    LogCallToDriver(p_Name, $"About to get Position and IsMoving properties repeatedly");
                    switch (p_type)
                    {
                        case RotatorPropertyMethod.Move:
                            {
                                WaitWhile($"Moving by {p_value:000} degrees", () => { return m_Rotator.IsMoving; }, 500, settings.RotatorTimeout, () => { return $"{Math.Abs(m_Rotator.Position - p_RotatorStartPosition):000} / {Math.Abs(p_value % 360.0):000} relative"; });
                                break;
                            }

                        case RotatorPropertyMethod.MoveMechanical:
                        case RotatorPropertyMethod.MoveAbsolute:
                            {
                                WaitWhile($"Moving to {p_value:000} degrees", () => { return m_Rotator.IsMoving; }, 500, settings.RotatorTimeout, () => { return $"{Math.Abs(m_Rotator.Position - p_RotatorStartPosition):000} / {Math.Abs((p_value - p_RotatorStartPosition) % 360.0):000} absolute"; });
                                break;
                            }

                        default:
                            {
                                WaitWhile("Waiting for move to complete", () => { return m_Rotator.IsMoving; }, 500, settings.RotatorTimeout, null);
                                break;
                            }
                    }

                    LogDebug("RotatorWait", "Rotator has stopped moving");
                    SetAction("");
                    m_LastMoveWasAsync = true;
                }
                else
                {
                    m_LastMoveWasAsync = false;
                }
            }
            else // Can only test for synchronous move
            {
                LogDebug("RotatorWait", "Cannot Read IsMoving");
                m_LastMoveWasAsync = false;
            }
        }

        private void RelativeMoveTest(float p_RelativeStepSize)
        {
            float l_Target;
            if (canReadPosition)
            {
                LogCallToDriver("Move", $"About to get Position property");
                if (m_Rotator.Position < p_RelativeStepSize) // Set a value that should succeed OK
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
                LogInfo("Position", "Skipping test as property is not supported");
            }

            // TargetPosition
            if (m_CanReadTargetPosition)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.TargetPosition, "TargetPosition");
            }
            else
            {
                LogInfo("TargetPosition", "Skipping test as property is not supported");
            }

            // StepSize
            if (m_CanReadStepSize)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.StepSize, "StepSize");
            }
            else
            {
                LogInfo("StepSize", "Skipping test as property is not supported");
            }

            // IsMoving
            if (m_CanReadIsMoving)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.IsMoving, "IsMoving");
            }
            else
            {
                LogInfo("IsMoving", "Skipping test as property is not supported");
            }
        }

        private void RotatorPerformanceTest(RotatorPropertyMethod p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
            float l_Single;
            bool l_Boolean;
            double l_Rate;
            SetAction(p_Name);
            try
            {
                l_StartTime = DateTime.Now;
                l_Count = 0.0d;
                l_LastElapsedTime = 0.0d;
                do
                {
                    l_Count += 1.0d;
                    switch (p_Type)
                    {
                        case RotatorPropertyMethod.Position:
                            {
                                l_Single = m_Rotator.Position;
                                break;
                            }

                        case RotatorPropertyMethod.TargetPosition:
                            {
                                l_Single = m_Rotator.TargetPosition;
                                break;
                            }

                        case RotatorPropertyMethod.StepSize:
                            {
                                l_Single = m_Rotator.StepSize;
                                break;
                            }

                        case RotatorPropertyMethod.IsMoving:
                            {
                                l_Boolean = m_Rotator.IsMoving;
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, "RotatorPerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0d)
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
                    case var @case when @case > 10.0d:
                        {
                            LogInfo(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case var case1 when 2.0d <= case1 && case1 <= 10.0d:
                        {
                            LogOK(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case var case2 when 1.0d <= case2 && case2 <= 2.0d:
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

        #endregion

    }
}

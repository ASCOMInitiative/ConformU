using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.Tools;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
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
        private bool canReadIsMoving, canReadPosition, canReadTargetPosition, canReadStepSize;
        private bool canReverse, isMoving;
        private float rotatorStepSize, rotatorPosition, mechanicalPosition;
        private bool reverse;
        private bool lastMoveWasAsync;
        private bool canReadMechanicalPosition;
        private float initialPosiiton = ROTATOR_POSITION_UNKNOWN;
        private float initialMechanicalPosiiton = ROTATOR_POSITION_UNKNOWN;
        private float initialSyncOffset = ROTATOR_POSITION_UNKNOWN;

        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        private IRotatorV4 rotator;

        #endregion

        #region Enums
        private enum RotatorMember
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
        public RotatorTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, true, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
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
                    rotator?.Dispose();
                    rotator = null;
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
                        rotator = new AlpacaRotator(
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
                                LogInfo("CreateDevice", $"Creating Native COM device: {settings.ComDevice.ProgId}");
                                rotator = new RotatorFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                rotator = new Rotator(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }

                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(rotator, DeviceTypes.Rotator); // Assign the driver to the base class

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

        public override void ReadCanProperties()
        {
            try
            {
                LogCallToDriver("CanReverse", "About to get CanReverse property");
                TimeMethod("CanReverse", () => canReverse = rotator.CanReverse, TargetTime.Fast);
                LogOk("CanReverse", canReverse.ToString());
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
                rotator.Halt();
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
                RotatorWait(RotatorMember.Move, "Ensuring that movement is stopped", 0, 0);

                if (!rotator.IsMoving)
                {
                    LogOk("Pre-run Check", "Rotator is stationary");

                    // Try to record the current position of the rotator so that it can be restored after testing. If this fails the initial position will be set to unknown value.
                    try
                    {
                        LogCallToDriver("PreRun Check", "About to get Position property");
                        initialPosiiton = rotator.Position;
                        LogOk("Pre-run Check", $"Rotator initial position: {initialPosiiton}");

                        // Attempt to get the rotator's current mechanical position. If this fails the initial mechanical position will be set to unknown value.
                        try
                        {
                            LogCallToDriver("PreRun Check", "About to get MechanicalPosition property");
                            initialMechanicalPosiiton = rotator.MechanicalPosition;
                            initialSyncOffset = (float)Utilities.Range((double)(initialPosiiton - initialMechanicalPosiiton), -180.0, true, 180.0, true);
                            LogOk("Pre-run Check", $"Rotator initial mechanical position: {initialMechanicalPosiiton}, Initial sync offset: {initialSyncOffset}");
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
                    currentPosition = rotator.Position;
                    LogOk("Post-run Check", $"Current position: {currentPosition}");

                    // Restore the original sync offset, if possible, for IRotatorV3 and later devices
                    if (GetInterfaceVersion() >= 3)
                    {

                        if (!Single.IsNaN(initialMechanicalPosiiton))
                        {
                            // Get the current mechanical position
                            LogCallToDriver("Post-run Check", $"About to get MechanicalPosition property");
                            currentMechanicalPosition = rotator.MechanicalPosition;
                            LogOk("Post-run Check", $"Current mechanical position: {currentMechanicalPosition}");

                            syncPosition = (float)Utilities.Range((double)(currentMechanicalPosition + initialSyncOffset), 0.0, true, 360.0, false);
                            LogOk("Post-run Check", $"New sync position: {syncPosition}");

                            LogCallToDriver("Post-run Check", $"About to call Sync method. Position: {syncPosition}");
                            rotator.Sync(syncPosition);
                            LogOk("Post-run Check", $"Completed Sync ({initialSyncOffset} degrees) from position: {currentPosition} to {syncPosition}");

                            // Re-get the current position because the sync will have changed it
                            LogCallToDriver("Post-run Check", $"About to get Position property");
                            currentPosition = rotator.Position;
                            LogOk("Post-run Check", $"New current position: {currentPosition}");
                        }
                    }

                    // Calculate the smallest relative movement required to get to the initial position
                    relativeMovement = (float)Utilities.Range((double)(initialPosiiton - currentPosition), -180.0, true, 180.0, true);

                    // Move to the starting position
                    LogCallToDriver("Post-run Check", $"About to move by {relativeMovement} to {initialPosiiton}");
                    rotator.Move(relativeMovement);

                    // Wait for the move to complete
                    RotatorWait(RotatorMember.Move, "Restoring original position", relativeMovement, currentPosition);

                    if (!rotator.IsMoving)
                        LogOk("Post-run Check", $"Rotator starting position successfully restored to {initialPosiiton}");
                    else
                        LogError("Post-run Check", $"Unable to restore rotator starting position, the rotator is still moving after {settings.RotatorTimeout} seconds. Could IsMoving be stuck on?");
                }
                catch (Exception ex)
                {
                    LogError("Post-run Check", $"Exception: {ex.Message}");
                    LogDebug("Post-run Check", $"Exception detail:\r\n{ex.Message}");
                }
        }

        public override void CheckProperties()
        {
            // IsMoving - Optional (V1,V2), Mandatory (V3)
            try
            {
                canReadIsMoving = false;
                LogCallToDriver("IsMoving", "About to get IsMoving property");
                isMoving = rotator.IsMoving;
                canReadIsMoving = true; // Can read OK, doesn't generate an exception
                if (isMoving)
                {
                    LogIssue("IsMoving", "IsMoving is True before any movement has been commanded!");
                    LogInfo("IsMoving", "Further tests have been skipped");
                }
                else
                {
                    LogOk("IsMoving", isMoving.ToString());
                }
            }
            catch (Exception ex)
            {
                // Check whether this is a Platform 7 or later interface
                if (GetInterfaceVersion() < 3) // Platform 6 or earlier interface
                {
                    HandleException("IsMoving", MemberType.Property, Required.Optional, ex, ""); // Handle as optional
                }
                else // Platform 7 or later interface
                {
                    HandleException("IsMoving", MemberType.Property, Required.Mandatory, ex, "");  // Handle as mandatory
                }
            }
            if (cancellationToken.IsCancellationRequested)
                return;

            // Position - Optional (V1,V2), Mandatory (V3)
            rotatorPosition = RotatorPropertyTestSingle(RotatorMember.Position, "Position", 0.0f, 360.0f, Required.Mandatory);
            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetPosition - Optional (V1,V2), Mandatory (V3)
            RotatorPropertyTestSingle(RotatorMember.TargetPosition, "TargetPosition", 0.0f, 360.0f, Required.Mandatory);
            if (cancellationToken.IsCancellationRequested)
                return;

            // StepSize - Optional (V1,V2 and V3)
            rotatorStepSize = RotatorPropertyTestSingle(RotatorMember.StepSize, "StepSize", 0.0f, 360.0f, Required.Optional);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Reverse Read - Optional if CanReverse is False, Mandatory if CanReverse is True (V1,V2), Mandatory (V3)
            try
            {
                LogCallToDriver("Reverse", "About to get Reverse property");
                TimeMethod("Reverse", () => reverse = rotator.Reverse, TargetTime.Fast);
                if (canReverse)
                {
                    LogOk("Reverse Read", reverse.ToString());
                }
                else
                {
                    LogIssue("Reverse Read", "CanReverse is false but no exception generated");
                }
            }
            catch (Exception ex)
            {
                if (GetInterfaceVersion() < 3) // Can be optional in IRotatorV1 and V2
                {
                    if (canReverse)
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
                if (reverse) // Try and set reverse to the opposite state
                {
                    LogCallToDriver("Reverse", "About to set Reverse property");
                    TimeMethod("Reverse", () => rotator.Reverse = false, TargetTime.Standard);
                }
                else
                {
                    LogCallToDriver("Reverse", "About to set Reverse property");
                    rotator.Reverse = true;
                }

                LogCallToDriver("Reverse", "About to set Reverse property");
                rotator.Reverse = reverse; // Restore original value
                if (canReverse)
                {
                    LogOk("Reverse Write", "Reverse state successfully changed and restored");
                }
                else
                {
                    LogIssue("Reverse Write", "CanReverse is false but no exception generated");
                }
            }
            catch (Exception ex)
            {
                if (GetInterfaceVersion() < 3) // Can be optional in IRotatorV1 and V2
                {
                    if (canReverse)
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
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test MechanicalPosition introduced in IRotatorV3
            if (GetInterfaceVersion() >= 3)
            {
                try
                {
                    canReadMechanicalPosition = false;
                    LogCallToDriver("MechanicalPosition", "About to set MechanicalPosition property");
                    TimeMethod("Reverse", () => mechanicalPosition = rotator.MechanicalPosition, TargetTime.Standard);
                    canReadMechanicalPosition = true; // Can read mechanical position OK, doesn't generate an exception

                    // Successfully retrieved a value
                    switch (mechanicalPosition)
                    {
                        case var @case when @case < 0.0f: // Lower than minimum value
                            {
                                LogIssue("MechanicalPosition", $"Invalid value: {mechanicalPosition}");
                                break;
                            }

                        case var case1 when case1 >= 360.0f: // Higher than maximum value
                            {
                                LogIssue("MechanicalPosition", $"Invalid value: {mechanicalPosition}"); // OK value
                                break;
                            }

                        default:
                            {
                                LogOk("MechanicalPosition", mechanicalPosition.ToString());
                                break;
                            }
                    }

                    // For information show the sync offset, if possible, using OFFSET = SKYPOSITION - MECHANICALPOSITION
                    if (canReadPosition) // Can read synced position and mechanical position
                    {
                        LogInfo("MechanicalPosition", $"Rotator sync offset: {rotatorPosition - mechanicalPosition}");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("MechanicalPosition", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
        }

        public override void CheckMethods()
        {
            LogCallToDriver("AccessChecks", "About to get Connected property");
            LogDebug("CheckMethods", $"Rotator is connected: {rotator.Connected}");

            // Halt - Optional (V1,V2 and V3)
            try
            {
                LogCallToDriver("Halt", $"About to call Halt method");
                rotator.Halt();
                LogOk("Halt", "Halt command successful");
            }
            catch (Exception ex)
            {
                HandleException("Halt", MemberType.Method, Required.Optional, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // MoveAbsolute - Optional (V1,V2), Mandatory (V3)
            RotatorMoveTest(RotatorMember.MoveAbsolute, "MoveAbsolute", 45.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;

            RotatorMoveTest(RotatorMember.MoveAbsolute, "MoveAbsolute", 135.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;

            RotatorMoveTest(RotatorMember.MoveAbsolute, "MoveAbsolute", 225.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;

            RotatorMoveTest(RotatorMember.MoveAbsolute, "MoveAbsolute", 315.0f, "");
            if (cancellationToken.IsCancellationRequested)
                return;

            RotatorMoveTest(RotatorMember.MoveAbsolute, "MoveAbsolute", (float)-405.0d, "Movement to large negative angle -405 degrees");
            if (cancellationToken.IsCancellationRequested)
                return;

            RotatorMoveTest(RotatorMember.MoveAbsolute, "MoveAbsolute", 405.0f, "Movement to large positive angle 405 degrees");
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

            RotatorMoveTest(RotatorMember.Move, "Move", (float)-375.0d, "Movement to large negative angle -375 degrees");
            if (cancellationToken.IsCancellationRequested)
                return;

            RotatorMoveTest(RotatorMember.Move, "Move", 375.0f, "Movement to large positive angle 375 degrees");
            if (cancellationToken.IsCancellationRequested)
                return;

            // Test the new IRotatorV3 methods
            if (GetInterfaceVersion() >= 3)
            {
                // Test the MoveMechanical method
                try
                {
                    if (canReadMechanicalPosition & canReadPosition)
                    {
                        RotatorMoveTest(RotatorMember.MoveMechanical, "MoveMechanical", 45.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorMember.MoveMechanical, "MoveMechanical", 135.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorMember.MoveMechanical, "MoveMechanical", 225.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorMember.MoveMechanical, "MoveMechanical", 315.0f, "");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorMember.MoveMechanical, "MoveMechanical", -405.0f, "Movement to large negative angle -405 degrees");
                        if (cancellationToken.IsCancellationRequested)
                            return;
                        RotatorMoveTest(RotatorMember.MoveMechanical, "MoveMechanical", 405.0f, "Movement to large positive angle 405 degrees");
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

        public override void CheckPerformance()
        {
            // Position
            if (canReadPosition)
            {
                RotatorPerformanceTest(RotatorMember.Position, "Position");
            }
            else
            {
                LogInfo("Position", "Skipping test as property is not supported");
            }

            // TargetPosition
            if (canReadTargetPosition)
            {
                RotatorPerformanceTest(RotatorMember.TargetPosition, "TargetPosition");
            }
            else
            {
                LogInfo("TargetPosition", "Skipping test as property is not supported");
            }

            // StepSize
            if (canReadStepSize)
            {
                RotatorPerformanceTest(RotatorMember.StepSize, "StepSize");
            }
            else
            {
                LogInfo("StepSize", "Skipping test as property is not supported");
            }

            // IsMoving
            if (canReadIsMoving)
            {
                RotatorPerformanceTest(RotatorMember.IsMoving, "IsMoving");
            }
            else
            {
                LogInfo("IsMoving", "Skipping test as property is not supported");
            }
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

        private float RotatorPropertyTestSingle(RotatorMember pType, string pName, float pMin, float pMax, Required requiredIRotatorV3State)
        {
            float rotatorPropertyTestSingleRet = default;
            Required requiredState;

            // Handle properties that were optional in IRotatorV1 and IRotoatorV2 but may have become mandatory in IRotatorV3
            if (GetInterfaceVersion() < 3)
            {
                requiredState = Required.Optional; // Default optional state for early versions
            }
            else
            {
                requiredState = requiredIRotatorV3State;
            } // Use the supplied required state for this specific property

            try
            {
                rotatorPropertyTestSingleRet = 0.0f;
                LogCallToDriver(pName, $"About to get {pName} property");
                switch (pType)
                {
                    case RotatorMember.Position:
                        {
                            canReadPosition = false;
                            TimeMethod(pName, () => rotatorPropertyTestSingleRet = rotator.Position, TargetTime.Fast);
                            canReadPosition = true; // Can read position OK, doesn't generate an exception
                            break;
                        }

                    case RotatorMember.StepSize:
                        {
                            canReadStepSize = false;
                            TimeMethod(pName, () => rotatorPropertyTestSingleRet = rotator.StepSize, TargetTime.Fast);
                            canReadStepSize = true;
                            break;
                        }

                    case RotatorMember.TargetPosition:
                        {
                            canReadTargetPosition = false;
                            TimeMethod(pName, () => rotatorPropertyTestSingleRet = rotator.TargetPosition, TargetTime.Fast);
                            canReadTargetPosition = true;
                            break;
                        }

                    default:
                        {
                            LogIssue(pName, $"RotatorPropertyTestSingle: Unknown test type - {pType}");
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (rotatorPropertyTestSingleRet)
                {
                    case var @case when @case < pMin: // Lower than minimum value
                        {
                            LogIssue(pName, $"Invalid value: {rotatorPropertyTestSingleRet}");
                            break;
                        }

                    case var case1 when case1 >= pMax: // Higher than maximum value
                        {
                            LogIssue(pName, $"Invalid value: {rotatorPropertyTestSingleRet}"); // OK value
                            break;
                        }

                    default:
                        {
                            LogOk(pName, rotatorPropertyTestSingleRet.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, requiredState, ex, "");
            }

            return rotatorPropertyTestSingleRet;
        }

        private void RotatorPerformanceTest(RotatorMember pType, string pName)
        {
            DateTime startTime;
            double count, lastElapsedTime, elapsedTime;
            float single;
            bool boolean;
            double rate;

            SetAction(pName);

            try
            {
                startTime = DateTime.Now;
                count = 0.0d;
                lastElapsedTime = 0.0d;

                do
                {
                    count += 1.0d;
                    switch (pType)
                    {
                        case RotatorMember.Position:
                            {
                                single = rotator.Position;
                                break;
                            }

                        case RotatorMember.TargetPosition:
                            {
                                single = rotator.TargetPosition;
                                break;
                            }

                        case RotatorMember.StepSize:
                            {
                                single = rotator.StepSize;
                                break;
                            }

                        case RotatorMember.IsMoving:
                            {
                                boolean = rotator.IsMoving;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"RotatorPerformanceTest: Unknown test type {pType}");
                                break;
                            }
                    }

                    elapsedTime = DateTime.Now.Subtract(startTime).TotalSeconds;
                    if (elapsedTime > lastElapsedTime + 1.0d)
                    {
                        SetStatus($"{count} transactions in {elapsedTime:0} seconds");
                        lastElapsedTime = elapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (elapsedTime <= PERF_LOOP_TIME);

                rate = count / elapsedTime;
                switch (rate)
                {
                    case var @case when @case > 10.0d:
                        {
                            LogInfo(pName, $"Transaction rate: {rate:0.0} per second");
                            break;
                        }

                    case var case1 when 2.0d <= case1 && case1 <= 10.0d:
                        {
                            LogOk(pName, $"Transaction rate: {rate:0.0} per second");
                            break;
                        }

                    case var case2 when 1.0d <= case2 && case2 <= 2.0d:
                        {
                            LogInfo(pName, $"Transaction rate: {rate:0.0} per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(pName, $"Transaction rate: {rate:0.0} per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(pName, $"Unable to complete test: {ex.Message}");
            }
        }

        private void RotatorSynctest(float syncAngle, float mechanicalAngle)
        {
            float syncAngleDifference;
            RotatorMoveTest(RotatorMember.MoveMechanical, "Sync", mechanicalAngle, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            try
            {
                LogCallToDriver("Sync", $"About to call Sync method");
                rotator.Sync(syncAngle);
                LogOk("Sync", "Synced OK");

                // Check that Position and MechanicalPosition are now the same
                syncAngleDifference = rotator.Position - syncAngle;
                if (Math.Abs(syncAngleDifference) < ROTATOR_POSITION_TOLERANCE)
                {
                    LogOk("Sync", $"Rotator Position has synced to {syncAngle} OK.");
                }
                else
                {
                    LogIssue("Sync", $"Rotator Position is {syncAngleDifference} degrees from the requested position {syncAngle}. Alert tolerance is {ROTATOR_POSITION_TOLERANCE} degrees.");
                }
            }
            catch (Exception ex)
            {
                HandleException("Sync", MemberType.Method, Required.Mandatory, ex, "");
            }
        }

        private void RotatorMoveTest(RotatorMember moveMethod, string memberName, float requiredPosition, string expectedErrorMsg)
        {
            float rotatorStartPosition = 0.0f, rotatorPositionAfterMove = ROTATOR_POSITION_UNKNOWN;
            double okLimit, positionOffset;

            SetTest(memberName);
            SetAction("Setting position");

            memberName = $"{memberName} {requiredPosition}";
            if (settings.Debug)
                LogNewLine();

            try
            {
                // Move to requested position
                Stopwatch sw = Stopwatch.StartNew();
                switch (moveMethod)
                {
                    case RotatorMember.Move:
                        if (canReadPosition) // Get us to a starting point of 10 degrees
                        {
                            LogCallToDriver(memberName, $"About to get Position property");
                            rotatorStartPosition = rotator.Position;
                            LogDebug(memberName, $"The rotator is currently at position: {rotator.Position:0.000}, the test position is: {requiredPosition:0.000}");
                        }
                        else
                        {
                            rotatorStartPosition = 0.0F;
                            LogDebug(memberName, $"Can't read the current rotator position - assuming 0.0, the test position is: {requiredPosition:0.000}");
                        }

                        LogDebug(memberName, "Starting relative move");
                        LogCallToDriver(memberName, $"About to call Move method");
                        TimeMethod(memberName, () => rotator.Move(requiredPosition), TargetTime.Standard);
                        LogDebug(memberName, $"Returned from Move method in {sw.Elapsed.TotalSeconds:0.000} seconds.");
                        break;

                    case RotatorMember.MoveAbsolute:
                        if (canReadPosition) // Get us to a starting point of 10 degrees
                        {
                            LogCallToDriver(memberName, $"About to get Position property");
                            rotatorStartPosition = rotator.Position;
                            LogDebug(memberName, $"The rotator is currently at position: {rotator.Position:0.000}, the test position is: {requiredPosition:0.000}");
                        }
                        else
                        {
                            rotatorStartPosition = 0.0F;
                            LogDebug(memberName, $"Can't read the current rotator position - assuming 0.0, the test position is: {requiredPosition:0.000}");
                        }

                        LogDebug(memberName, "Starting absolute move");
                        LogCallToDriver(memberName, $"About to call MoveAbsolute method");
                        TimeMethod(memberName, () => rotator.MoveAbsolute(requiredPosition), TargetTime.Standard);
                        LogDebug(memberName, $"Returned from MoveAbsolute method in {sw.Elapsed.TotalSeconds:0.000} seconds.");
                        break;

                    case RotatorMember.MoveMechanical:
                        if (canReadMechanicalPosition) // Get us to a starting point of 10 degrees
                        {
                            LogCallToDriver(memberName, $"About to get MechanicalPosition property");
                            rotatorStartPosition = rotator.MechanicalPosition;
                            LogDebug(memberName, $"The rotator is currently at mechanical position: {rotator.MechanicalPosition:0.000}, the test position is: {requiredPosition:0.000}");
                        }
                        else
                        {
                            rotatorStartPosition = 0.0F;
                            LogDebug(memberName, $"Can't read the current rotator mechanical position - assuming 0.0, the test position is: {requiredPosition:0.000}");
                        }

                        LogDebug(memberName, "Starting mechanical move");
                        LogCallToDriver(memberName, $"About to call MoveMechanical method");
                        TimeMethod(memberName, () => rotator.MoveMechanical(requiredPosition), TargetTime.Standard);
                        LogDebug(memberName, $"Returned from MoveMechanical method in {sw.Elapsed.TotalSeconds:0.000} seconds.");
                        break;

                    default:
                        LogIssue(memberName, $"RotatorMoveTest: Unknown test type - {moveMethod}");
                        break;
                }

                // Report synchronous behaviour for Platform 7 and later interfaces
                if (DeviceCapabilities.IsPlatform7OrLater(DeviceTypes.Rotator, GetInterfaceVersion())) // This is a Platform 7 or later interface
                {
                    // Check whether IsMoving can be read
                    if (canReadIsMoving) // IsMoving can be read
                    {
                        // Check whether the rotator is not moving and the Move method exceeded the standard response time
                        if (!rotator.IsMoving & sw.Elapsed.TotalSeconds > standardTargetResponseTime) // Not moving and time > standard response time ==> synchronous behaviour so rais issue
                        {
                            LogIssue(memberName, $"The {memberName} method operated synchronously when it should operate asynchronously in IRotatorV4 and later devices.");
                            LogInfo(memberName, $"The {memberName} method took {sw.Elapsed.TotalSeconds:0.0} seconds to complete, which is longer than the standard response time of {standardTargetResponseTime:0.0} seconds.");
                            LogInfo(memberName, $"The {memberName} method should complete quickly and IsMoving should be set true until the rotator finishes movement.");
                        }
                    }
                }

                // Wait for movement to stop
                RotatorWait(moveMethod, memberName, requiredPosition, rotatorStartPosition);

                // Report outcome
                switch (moveMethod)
                {
                    case RotatorMember.Move:
                        if (lastMoveWasAsync) // Asynchronous move
                        {
                            if (canReadPosition)
                            {
                                LogCallToDriver(memberName, $"About to get Position property");
                                rotatorPositionAfterMove = rotator.Position;
                                LogDebug(memberName, $"Asynchronous move completed. Rotator position: {rotatorPositionAfterMove} degrees");
                            }
                            else
                            {
                                LogOk(memberName, "Asynchronous move completed");
                            }
                        }
                        else // Synchronous move
                        {
                            if (canReadPosition) // Can read position
                            {
                                LogCallToDriver(memberName, $"About to get Position property");
                                rotatorPositionAfterMove = rotator.Position;
                                LogDebug(memberName, $"Synchronous move completed. Rotator position: {rotatorPositionAfterMove} degrees");
                            }
                            else // Cannot read position
                            {
                                LogOk(memberName, "Synchronous move completed");
                            }
                        }
                        break;

                    case RotatorMember.MoveMechanical:
                        if (lastMoveWasAsync) // Asynchronous move
                        {
                            if (canReadMechanicalPosition)
                            {
                                LogCallToDriver(memberName, $"About to get MechanicalPosition property");
                                rotatorPositionAfterMove = rotator.MechanicalPosition;
                                LogDebug(memberName, $"Asynchronous mechanical move completed. Rotator mechanical position: {rotatorPositionAfterMove} degrees");
                            }
                            else
                            {
                                LogOk(memberName, "Asynchronous mechanical move completed.");
                            }
                        }
                        else // Synchronous move
                        {
                            if (canReadPosition) // Can read position
                            {
                                LogCallToDriver(memberName, $"About to get MechanicalPosition property");
                                rotatorPositionAfterMove = rotator.MechanicalPosition;
                                LogDebug(memberName, $"Synchronous mechanical move completed. Rotator mechanical position: {rotatorPositionAfterMove} degrees");
                            }
                            else // Cannot read position
                            {
                                LogOk(memberName, "Synchronous mechanical move completed");
                            }
                        }
                        break;

                    case RotatorMember.MoveAbsolute:
                        if (lastMoveWasAsync) // Asynchronous move
                        {
                            if (canReadPosition)
                            {
                                LogCallToDriver(memberName, $"About to get Position property");
                                rotatorPositionAfterMove = rotator.Position;
                                LogDebug(memberName, $"Asynchronous absolute move completed. Rotator position: {rotatorPositionAfterMove} degrees");
                            }
                            else
                            {
                                LogOk(memberName, "Asynchronous absolute move completed.");
                            }
                        }
                        else // Synchronous move
                        {
                            if (canReadPosition) // Can read position
                            {
                                LogCallToDriver(memberName, $"About to get Position property");
                                rotatorPositionAfterMove = rotator.Position;
                                LogDebug(memberName, $"Synchronous absolute move completed. Rotator position: {rotatorPositionAfterMove} degrees");
                            }
                            else // Cannot read position
                            {
                                LogOk(memberName, "Synchronous absolute move completed");
                            }
                        }
                        break;
                }

                // Now test whether we got to where we expected to go
                if (rotatorPosition != ROTATOR_POSITION_UNKNOWN) // We were able to read the rotator's position or mechanical position
                {
                    if (canReadStepSize)
                    {
                        okLimit = 1.1d * rotatorStepSize; // Set to 110% of step size to allow tolerance on reporting within 1 step of required location
                    }
                    else
                    {
                        okLimit = ROTATOR_OK_TOLERANCE;
                    }

                    if (GetInterfaceVersion() < 3) // Interface V1 and V2 behaviour
                    {
                        if (rotatorPositionAfterMove < 0.0d)
                            LogInfo(memberName, "Rotator supports angles < 0.0");
                        if (rotatorPositionAfterMove > 360.0d)
                            LogInfo(memberName, "Rotator supports angles > 360.0");
                    }
                    else if (rotatorPositionAfterMove < 0.0d | rotatorPositionAfterMove >= 360.0d) // Interface V3 behaviour (Position must be 0..359.99999...)
                        LogIssue(memberName, $"Rotator position {rotatorPositionAfterMove:0.000} is outside the valid range: 0.0 to 359.99999...");

                    LogDebug($"{memberName}", $"Confirming that the rotator moved where commanded - Required position: {requiredPosition:0.000}, Actual position: {rotatorPositionAfterMove:0.000}, Starting position: {rotatorStartPosition:0.000}, Tolerance: {okLimit:0.000}");

                    // Calculate the position offset from the required position

                    if (moveMethod == RotatorMember.Move)
                    {
                        positionOffset = Math.Abs((720.0d + rotatorPositionAfterMove - (requiredPosition + rotatorStartPosition)) % 360.0d); // Account for rotator positions that report < 0.0 or > 360.0 degrees
                    }
                    else
                    {
                        positionOffset = Math.Abs((720.0d + rotatorPositionAfterMove - requiredPosition) % 360.0d); // Account for rotator positions that report < 0.0 or > 360.0 degrees
                    }

                    if (positionOffset > 180.0d)
                        positionOffset = 360.0d - positionOffset; // Cope with positions that return just under the expected value
                    switch (Math.Abs(positionOffset))
                    {
                        case 0.0d:
                            LogOk(memberName, $"Rotator is at the expected position: {rotatorPositionAfterMove}");
                            break;

                        case var @case when 0.0d <= @case && @case <= okLimit:
                            LogOk(memberName, $"Rotator is within {okLimit:0.000)} {((positionOffset <= 1.0d) ? " degree" : " degrees")} of the expected position: {rotatorPositionAfterMove}");
                            break;

                        case var case1 when 0.0d <= case1 && case1 <= ROTATOR_INFO_TOLERANCE:
                            LogInfo(memberName, $"Rotator is {positionOffset:0.000} degrees from expected position: {rotatorPositionAfterMove}");
                            break;

                        default:
                            LogIssue(memberName, $"Rotator is {positionOffset:0.000} degrees from expected position {rotatorPositionAfterMove}, which is more than the conformance value of {ROTATOR_INFO_TOLERANCE:0.0} degrees");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(expectedErrorMsg)) // Test for normal behaviour
                {
                    if (GetInterfaceVersion() < 3)
                    {
                        HandleException(memberName, MemberType.Method, Required.Optional, ex, "");
                    }
                    else
                    {
                        HandleException(memberName, MemberType.Method, Required.Mandatory, ex, "");
                    }
                }
                else if (IsInvalidValueException(memberName, ex)) // This is expected to fail because a bad position was used
                                                                  // Test for an InvalidValueException and handle if found
                {
                    HandleInvalidValueExceptionAsOk(memberName, MemberType.Method, Required.Mandatory, ex, "", expectedErrorMsg);
                }
                else if (GetInterfaceVersion() < 3) // Some other type of exception occurred
                {
                    HandleException(memberName, MemberType.Method, Required.Optional, ex, "");
                }
                else
                {
                    HandleException(memberName, MemberType.Method, Required.Mandatory, ex, "");
                }
            }

            SetAction("");
            SetStatus("");
            SetTest("");
        }

        private void RotatorWait(RotatorMember moveMethod, string memberName, float positionValue, float rotatorStartPosition)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            if (canReadIsMoving) // Can read IsMoving so test for asynchronous and synchronous behaviour
            {
                LogCallToDriver(memberName, $"About to get IsMoving property");
                if (rotator.IsMoving)
                {
                    LogDebug(memberName, "RotatorWait - Rotator is moving, waiting for move to complete");
                    SetTest(memberName);
                    SetAction("Waiting for move to complete");

                    LogCallToDriver(memberName, $"About to get Position and IsMoving properties repeatedly");
                    switch (moveMethod)
                    {
                        case RotatorMember.Move:
                            WaitWhile($"Moving by {positionValue:000} degrees", () => rotator.IsMoving, 500, settings.RotatorTimeout, () => $"{Math.Abs(rotator.Position - rotatorStartPosition):000} / {Math.Abs(positionValue % 360.0):000} relative");
                            break;

                        case RotatorMember.MoveMechanical:
                        case RotatorMember.MoveAbsolute:
                            WaitWhile($"Moving to {positionValue:000} degrees", () => rotator.IsMoving, 500, settings.RotatorTimeout, () => $"{Math.Abs(rotator.Position - rotatorStartPosition):000} / {Math.Abs((positionValue - rotatorStartPosition) % 360.0):000} absolute");
                            break;

                        default:
                            WaitWhile("Waiting for move to complete", () => rotator.IsMoving, 500, settings.RotatorTimeout, null);
                            break;
                    }

                    SetAction("");
                    lastMoveWasAsync = true;
                }
                else
                {
                    lastMoveWasAsync = false;
                }
            }
            else // Can not read IsMoving so we can only test for synchronous movement
            {
                LogDebug(memberName, "Cannot Read IsMoving");
                lastMoveWasAsync = false;
            }

            LogDebug(memberName, $"Time spent waiting for move to complete: {stopwatch.Elapsed.TotalSeconds:0.000} seconds");
        }

        private void RelativeMoveTest(float relativeStepSize)
        {
            float target;
            if (canReadPosition)
            {
                LogCallToDriver("Move", $"About to get Position property");
                if (rotator.Position < relativeStepSize) // Set a value that should succeed OK
                {
                    target = relativeStepSize;
                }
                else
                {
                    target = -relativeStepSize;
                }
            }
            else
            {
                target = relativeStepSize;
            }

            RotatorMoveTest(RotatorMember.Move, "Move", target, "");
            if (cancellationToken.IsCancellationRequested)
                return;

            RotatorMoveTest(RotatorMember.Move, "Move", -target, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            // Should now be back where we started

        }

        #endregion

    }
}

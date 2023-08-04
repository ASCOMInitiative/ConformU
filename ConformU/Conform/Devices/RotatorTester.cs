using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.Tools;
using System;
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
        private bool mCanReadIsMoving, canReadPosition, mCanReadTargetPosition, mCanReadStepSize;
        private bool mCanReverse, mIsMoving;
        private float mRotatorStepSize, mRotatorPosition, mechanicalPosition;
        private bool mReverse;
        private bool mLastMoveWasAsync;
        private bool canReadMechanicalPosition;
        private float initialPosiiton = ROTATOR_POSITION_UNKNOWN;
        private float initialMechanicalPosiiton = ROTATOR_POSITION_UNKNOWN;
        private float initialSyncOffset = ROTATOR_POSITION_UNKNOWN;

        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        private IRotatorV4 mRotator;

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
                    mRotator?.Dispose();
                    mRotator = null;
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
                            GExNotImplemented = (int)0x80040400;
                            GExInvalidValue1 = (int)0x80040405;
                            GExInvalidValue2 = (int)0x80040405;
                            GExNotSet1 = (int)0x80040403;
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
                        mRotator = new AlpacaRotator(
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
                                mRotator = new RotatorFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                mRotator = new Rotator(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }

                        LogInfo("CreateDevice", "Successfully created driver");
                        SetDevice(mRotator, DeviceTypes.Rotator); // Assign the driver to the base class

                        SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                        WaitFor(1000, 100);
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }
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

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(mRotator, DeviceTypes.Rotator);
        }

        public override void ReadCanProperties()
        {
            try
            {
                LogCallToDriver("CanReverse", "About to get CanReverse property");
                mCanReverse = mRotator.CanReverse;
                LogOk("CanReverse", mCanReverse.ToString());
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
                mRotator.Halt();
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

                if (!mRotator.IsMoving)
                {
                    LogOk("Pre-run Check", "Rotator is stationary");

                    // Try to record the current position of the rotator so that it can be restored after testing. If this fails the initial position will be set to unknown value.
                    try
                    {
                        LogCallToDriver("PreRun Check", "About to get Position property");
                        initialPosiiton = mRotator.Position;
                        LogOk("Pre-run Check", $"Rotator initial position: {initialPosiiton}");

                        // Attempt to get the rotator's current mechanical position. If this fails the initial mechanical position will be set to unknown value.
                        try
                        {
                            LogCallToDriver("PreRun Check", "About to get MechanicalPosition property");
                            initialMechanicalPosiiton = mRotator.MechanicalPosition;
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
                    currentPosition = mRotator.Position;
                    LogOk("Post-run Check", $"Current position: {currentPosition}");

                    // Restore the original sync offset, if possible, for IRotatorV3 and later devices
                    if (GetInterfaceVersion() >= 3)
                    {

                        if (!Single.IsNaN(initialMechanicalPosiiton))
                        {
                            // Get the current mechanical position
                            LogCallToDriver("Post-run Check", $"About to get MechanicalPosition property");
                            currentMechanicalPosition = mRotator.MechanicalPosition;
                            LogOk("Post-run Check", $"Current mechanical position: {currentMechanicalPosition}");

                            syncPosition = (float)Utilities.Range((double)(currentMechanicalPosition + initialSyncOffset), 0.0, true, 360.0, false);
                            LogOk("Post-run Check", $"New sync position: {syncPosition}");

                            LogCallToDriver("Post-run Check", $"About to call Sync method. Position: {syncPosition}");
                            mRotator.Sync(syncPosition);
                            LogOk("Post-run Check", $"Completed Sync ({initialSyncOffset} degrees) from position: {currentPosition} to {syncPosition}");

                            // Re-get the current position because the sync will have changed it
                            LogCallToDriver("Post-run Check", $"About to get Position property");
                            currentPosition = mRotator.Position;
                            LogOk("Post-run Check", $"New current position: {currentPosition}");
                        }
                    }

                    // Calculate the smallest relative movement required to get to the initial position
                    relativeMovement = (float)Utilities.Range((double)(initialPosiiton - currentPosition), -180.0, true, 180.0, true);

                    // Move to the starting position
                    LogCallToDriver("Post-run Check", $"About to move by {relativeMovement} to {initialPosiiton}");
                    mRotator.Move(relativeMovement);

                    // Wait for the move to complete
                    RotatorWait(RotatorPropertyMethod.Move, "Restoring original position", relativeMovement, currentPosition);

                    if (!mRotator.IsMoving)
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
                mCanReadIsMoving = false;
                LogCallToDriver("IsMoving", "About to get IsMoving property");
                mIsMoving = mRotator.IsMoving;
                mCanReadIsMoving = true; // Can read OK, doesn't generate an exception
                if (mIsMoving)
                {
                    LogIssue("IsMoving", "IsMoving is True before any movement has been commanded!");
                    LogInfo("IsMoving", "Further tests have been skipped");
                }
                else
                {
                    LogOk("IsMoving", mIsMoving.ToString());
                }
            }
            catch (Exception ex)
            {
                if (GetInterfaceVersion() < 3)
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
            mRotatorPosition = RotatorPropertyTestSingle(RotatorPropertyMethod.Position, "Position", 0.0f, 360.0f, Required.Mandatory);
            if (cancellationToken.IsCancellationRequested)
                return;

            // TargetPosition - Optional (V1,V2), Mandatory (V3)
            RotatorPropertyTestSingle(RotatorPropertyMethod.TargetPosition, "TargetPosition", 0.0f, 360.0f, Required.Mandatory);
            if (cancellationToken.IsCancellationRequested)
                return;

            // StepSize - Optional (V1,V2 and V3)
            mRotatorStepSize = RotatorPropertyTestSingle(RotatorPropertyMethod.StepSize, "StepSize", 0.0f, 360.0f, Required.Optional);
            if (cancellationToken.IsCancellationRequested)
                return;

            // Reverse Read - Optional if CanReverse is False, Mandatory if CanReverse is True (V1,V2), Mandatory (V3)
            try
            {
                LogCallToDriver("Reverse", "About to get Reverse property");
                mReverse = mRotator.Reverse;
                if (mCanReverse)
                {
                    LogOk("Reverse Read", mReverse.ToString());
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
                    if (mCanReverse)
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
                if (mReverse) // Try and set reverse to the opposite state
                {
                    LogCallToDriver("Reverse", "About to set Reverse property");
                    mRotator.Reverse = false;
                }
                else
                {
                    LogCallToDriver("Reverse", "About to set Reverse property");
                    mRotator.Reverse = true;
                }

                LogCallToDriver("Reverse", "About to set Reverse property");
                mRotator.Reverse = mReverse; // Restore original value
                if (mCanReverse)
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
                    if (mCanReverse)
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
            if (GetInterfaceVersion() >= 3)
            {
                try
                {
                    canReadMechanicalPosition = false;
                    LogCallToDriver("MechanicalPosition", "About to set MechanicalPosition property");
                    mechanicalPosition = mRotator.MechanicalPosition;
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
                        LogInfo("MechanicalPosition", $"Rotator sync offset: {mRotatorPosition - mechanicalPosition}");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("MechanicalPosition", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
        }

        private float RotatorPropertyTestSingle(RotatorPropertyMethod pType, string pName, float pMin, float pMax, Required requiredIRotatorV3State)
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
                    case RotatorPropertyMethod.Position:
                        {
                            canReadPosition = false;
                            rotatorPropertyTestSingleRet = mRotator.Position;
                            canReadPosition = true; // Can read position OK, doesn't generate an exception
                            break;
                        }

                    case RotatorPropertyMethod.StepSize:
                        {
                            mCanReadStepSize = false;
                            rotatorPropertyTestSingleRet = mRotator.StepSize;
                            mCanReadStepSize = true;
                            break;
                        }

                    case RotatorPropertyMethod.TargetPosition:
                        {
                            mCanReadTargetPosition = false;
                            rotatorPropertyTestSingleRet = mRotator.TargetPosition;
                            mCanReadTargetPosition = true;
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

        public override void CheckMethods()
        {
            LogCallToDriver("AccessChecks", "About to get Connected property");
            LogDebug("CheckMethods", $"Rotator is connected: {mRotator.Connected}");

            // Halt - Optional (V1,V2 and V3)
            try
            {
                LogCallToDriver("Halt", $"About to call Halt method");
                mRotator.Halt();
                LogOk("Halt", "Halt command successful");
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
            if (GetInterfaceVersion() >= 3)
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

        private void RotatorSynctest(float syncAngle, float mechanicalAngle)
        {
            float syncAngleDifference;
            RotatorMoveTest(RotatorPropertyMethod.MoveMechanical, "Sync", mechanicalAngle, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            try
            {
                LogCallToDriver("Sync", $"About to call Sync method");
                mRotator.Sync(syncAngle);
                LogOk("Sync", "Synced OK");

                // Check that Position and MechanicalPosition are now the same
                syncAngleDifference = mRotator.Position - syncAngle;
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

        private void RotatorMoveTest(RotatorPropertyMethod pType, string pName, float pValue, string pExpectErrorMsg)
        {
            float lRotatorStartPosition = default, rotatorPosition;
            double lOkLimit, lPositionOffset;
            LogCallToDriver(pName, $"About to get Position property");
            LogDebug("RotatorMoveTest", $"Start value, position: {pValue:0.000} {mRotator.Position:0.000}");

            SetTest(pName);
            SetAction("Setting position");
            try
            {
                // Move to requested position
                switch (pType)
                {
                    case RotatorPropertyMethod.Move:
                        {
                            LogDebug("RotatorMoveTest", $"Reading rotator start position: {canReadPosition}");
                            if (canReadPosition) // Get us to a starting point of 10 degrees
                            {
                                LogCallToDriver(pName, $"About to get Position property");
                                lRotatorStartPosition = mRotator.Position;
                            }

                            LogDebug("RotatorMoveTest", "Starting relative move");
                            LogCallToDriver(pName, $"About to call Move method");
                            mRotator.Move(pValue);
                            LogDebug("RotatorMoveTest", "Starting relative move");
                            break;
                        }

                    case RotatorPropertyMethod.MoveAbsolute:
                        {
                            LogDebug("RotatorMoveTest", "Starting absolute move");
                            lRotatorStartPosition = 0.0f;
                            LogCallToDriver(pName, $"About to call MoveAbsolute method");
                            mRotator.MoveAbsolute(pValue);
                            LogDebug("RotatorMoveTest", "Completed absolute move");
                            break;
                        }

                    case RotatorPropertyMethod.MoveMechanical:
                        {
                            LogDebug("RotatorMoveTest", "Starting mechanical move");
                            lRotatorStartPosition = 0.0f;
                            LogCallToDriver(pName, $"About to call MoveMechanical method");
                            mRotator.MoveMechanical(pValue);
                            LogDebug("RotatorMoveTest", "Completed mechanical move");
                            break;
                        }

                    default:
                        {
                            LogIssue(pName, $"RotatorMoveTest: Unknown test type - {pType}");
                            break;
                        }
                }

                RotatorWait(pType, pName, pValue, lRotatorStartPosition);
                if (mLastMoveWasAsync) // Asynchronous move
                {
                    switch (pType)
                    {
                        case RotatorPropertyMethod.Move:
                            {
                                if (canReadPosition)
                                {
                                    LogCallToDriver(pName, $"About to get Position property");
                                    LogOk(pName, $"Asynchronous move successful - moved by {pValue} degrees to: {mRotator.Position} degrees");
                                }
                                else
                                {
                                    LogOk(pName, "Asynchronous move successful");
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
                                    LogCallToDriver(pName, $"About to get Position property");
                                    LogOk(pName, $"Asynchronous move successful to: {mRotator.Position} degrees");
                                }
                                else
                                {
                                    LogOk(pName, "Asynchronous move successful");
                                }

                                break;
                            }
                    }
                }
                else if (canReadPosition) // Synchronous move
                {
                    LogCallToDriver(pName, $"About to get Position property");
                    LogOk(pName, $"Synchronous move successful to: {mRotator.Position} degrees");
                }
                else
                {
                    LogOk(pName, "Synchronous move successful");
                }

                // Now test whether we got to where we expected to go
                if (canReadPosition)
                {
                    if (mCanReadStepSize)
                    {
                        lOkLimit = 1.1d * mRotatorStepSize; // Set to 110% of step size to allow tolerance on reporting within 1 step of required location
                    }
                    else
                    {
                        lOkLimit = ROTATOR_OK_TOLERANCE;
                    }

                    LogCallToDriver(pName, $"About to get Position property");
                    LogDebug($"{pName}1",
                        $"Position, value, start, tolerance: {mRotator.Position:0.000} {pValue:0.000} {lRotatorStartPosition:0.000} {lOkLimit:0.000}");
                    LogCallToDriver(pName, $"About to get Position property");
                    rotatorPosition = mRotator.Position;
                    if (GetInterfaceVersion() < 3) // Interface V1 and V2 behaviour
                    {
                        if (rotatorPosition < 0.0d)
                            LogInfo(pName, "Rotator supports angles < 0.0");
                        if (rotatorPosition > 360.0d)
                            LogInfo(pName, "Rotator supports angles > 360.0");
                    }
                    else if (rotatorPosition < 0.0d | rotatorPosition >= 360.0d) // Interface V3 behaviour (Position must be 0..359.99999...)
                        LogIssue(pName, $"Rotator position {rotatorPosition:0.000} is outside the valid range: 0.0 to 359.99999...");

                    // Get the relevant position value
                    if (pType == RotatorPropertyMethod.MoveMechanical) // Use the MechanicalPosition property
                    {
                        LogCallToDriver(pName, $"About to get MechanicalPosition property");
                        rotatorPosition = mRotator.MechanicalPosition;
                    }
                    else // Use the Position property for all other methods
                    {
                        LogCallToDriver(pName, $"About to get Position property");
                        rotatorPosition = mRotator.Position;
                    }
                    // Calculate the position offset from the required position
                    lPositionOffset = Math.Abs((720.0d + rotatorPosition - (pValue + lRotatorStartPosition)) % 360.0d); // Account for rotator positions that report < 0.0 or > 360.0 degrees
                    if (lPositionOffset > 180.0d)
                        lPositionOffset = 360.0d - lPositionOffset; // Cope with positions that return just under the expected value
                    switch (Math.Abs(lPositionOffset))
                    {
                        case 0.0d:
                            {
                                LogOk(pName, $"Rotator is at the expected position: {rotatorPosition}");
                                break;
                            }

                        case var @case when 0.0d <= @case && @case <= lOkLimit:
                            {
                                LogOk(pName, $"Rotator is within {lOkLimit:0.000)} {((lPositionOffset <= 1.0d) ? " degree" : " degrees")} of the expected position: {rotatorPosition}");
                                break;
                            }

                        case var case1 when 0.0d <= case1 && case1 <= ROTATOR_INFO_TOLERANCE:
                            {
                                LogInfo(pName, $"Rotator is {lPositionOffset:0.000} degrees from expected position: {rotatorPosition}");
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"Rotator is {lPositionOffset:0.000} degrees from expected position {rotatorPosition}, which is more than the conformance value of {ROTATOR_INFO_TOLERANCE:0.0} degrees");
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                if (string.IsNullOrEmpty(pExpectErrorMsg)) // Test for normal behaviour
                {
                    if (GetInterfaceVersion() < 3)
                    {
                        HandleException(pName, MemberType.Method, Required.Optional, ex, "");
                    }
                    else
                    {
                        HandleException(pName, MemberType.Method, Required.Mandatory, ex, "");
                    }
                }
                else if (IsInvalidValueException(pName, ex)) // This is expected to fail because a bad position was used
                                                              // Test for an InvalidValueException and handle if found
                {
                    HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.Mandatory, ex, "", pExpectErrorMsg);
                }
                else if (GetInterfaceVersion() < 3) // Some other type of exception occurred
                {
                    HandleException(pName, MemberType.Method, Required.Optional, ex, "");
                }
                else
                {
                    HandleException(pName, MemberType.Method, Required.Mandatory, ex, "");
                }
            }

            SetAction("");
            SetStatus("");
            SetTest("");
        }

        private void RotatorWait(RotatorPropertyMethod pType, string pName, float pValue, float pRotatorStartPosition)
        {
            LogDebug("RotatorWait", "Entered RotatorWait");
            if (mCanReadIsMoving) // Can read IsMoving so test for asynchronous and synchronous behaviour
            {
                LogDebug("RotatorWait", "Can Read IsMoving OK");
                LogCallToDriver(pName, $"About to get IsMoving property");
                if (mRotator.IsMoving)
                {
                    LogDebug("RotatorWait", "Rotator is moving, waiting for move to complete");
                    SetTest(pName);
                    SetAction("Waiting for move to complete");
                    LogCallToDriver(pName, $"About to get Position and IsMoving properties repeatedly");
                    switch (pType)
                    {
                        case RotatorPropertyMethod.Move:
                            {
                                WaitWhile($"Moving by {pValue:000} degrees", () => mRotator.IsMoving, 500, settings.RotatorTimeout, () => $"{Math.Abs(mRotator.Position - pRotatorStartPosition):000} / {Math.Abs(pValue % 360.0):000} relative");
                                break;
                            }

                        case RotatorPropertyMethod.MoveMechanical:
                        case RotatorPropertyMethod.MoveAbsolute:
                            {
                                WaitWhile($"Moving to {pValue:000} degrees", () => mRotator.IsMoving, 500, settings.RotatorTimeout, () => $"{Math.Abs(mRotator.Position - pRotatorStartPosition):000} / {Math.Abs((pValue - pRotatorStartPosition) % 360.0):000} absolute");
                                break;
                            }

                        default:
                            {
                                WaitWhile("Waiting for move to complete", () => mRotator.IsMoving, 500, settings.RotatorTimeout, null);
                                break;
                            }
                    }

                    LogDebug("RotatorWait", "Rotator has stopped moving");
                    SetAction("");
                    mLastMoveWasAsync = true;
                }
                else
                {
                    mLastMoveWasAsync = false;
                }
            }
            else // Can only test for synchronous move
            {
                LogDebug("RotatorWait", "Cannot Read IsMoving");
                mLastMoveWasAsync = false;
            }
        }

        private void RelativeMoveTest(float pRelativeStepSize)
        {
            float lTarget;
            if (canReadPosition)
            {
                LogCallToDriver("Move", $"About to get Position property");
                if (mRotator.Position < pRelativeStepSize) // Set a value that should succeed OK
                {
                    lTarget = pRelativeStepSize;
                }
                else
                {
                    lTarget = -pRelativeStepSize;
                }
            }
            else
            {
                lTarget = pRelativeStepSize;
            }

            RotatorMoveTest(RotatorPropertyMethod.Move, "Move", lTarget, "");
            if (cancellationToken.IsCancellationRequested)
                return;
            RotatorMoveTest(RotatorPropertyMethod.Move, "Move", -lTarget, "");
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
            if (mCanReadTargetPosition)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.TargetPosition, "TargetPosition");
            }
            else
            {
                LogInfo("TargetPosition", "Skipping test as property is not supported");
            }

            // StepSize
            if (mCanReadStepSize)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.StepSize, "StepSize");
            }
            else
            {
                LogInfo("StepSize", "Skipping test as property is not supported");
            }

            // IsMoving
            if (mCanReadIsMoving)
            {
                RotatorPerformanceTest(RotatorPropertyMethod.IsMoving, "IsMoving");
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

        private void RotatorPerformanceTest(RotatorPropertyMethod pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime;
            float lSingle;
            bool lBoolean;
            double lRate;
            SetAction(pName);
            try
            {
                lStartTime = DateTime.Now;
                lCount = 0.0d;
                lLastElapsedTime = 0.0d;
                do
                {
                    lCount += 1.0d;
                    switch (pType)
                    {
                        case RotatorPropertyMethod.Position:
                            {
                                lSingle = mRotator.Position;
                                break;
                            }

                        case RotatorPropertyMethod.TargetPosition:
                            {
                                lSingle = mRotator.TargetPosition;
                                break;
                            }

                        case RotatorPropertyMethod.StepSize:
                            {
                                lSingle = mRotator.StepSize;
                                break;
                            }

                        case RotatorPropertyMethod.IsMoving:
                            {
                                lBoolean = mRotator.IsMoving;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"RotatorPerformanceTest: Unknown test type {pType}");
                                break;
                            }
                    }

                    lElapsedTime = DateTime.Now.Subtract(lStartTime).TotalSeconds;
                    if (lElapsedTime > lLastElapsedTime + 1.0d)
                    {
                        SetStatus($"{lCount} transactions in {lElapsedTime:0} seconds");
                        lLastElapsedTime = lElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (lElapsedTime <= PERF_LOOP_TIME);
                lRate = lCount / lElapsedTime;
                switch (lRate)
                {
                    case var @case when @case > 10.0d:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case var case1 when 2.0d <= case1 && case1 <= 10.0d:
                        {
                            LogOk(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case var case2 when 1.0d <= case2 && case2 <= 2.0d:
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

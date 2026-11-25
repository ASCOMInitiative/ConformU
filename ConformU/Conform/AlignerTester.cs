using ASCOM.Alpaca.Clients;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Configuration;
using System.Reflection;
using System.Threading;

namespace ConformU
{
    internal class AlignerTester : DeviceTesterBaseClass
    {
        private IAlignerV1 aligner;
        private readonly Settings settings;
        private readonly ConformLogger logger;
        private bool disposedValue;
        private readonly CancellationToken cancellationToken;
        bool? canFindHome = null;
        bool? canMoveX = null;
        bool? canMoveY = null;
        bool? canMoveZ = null;
        bool? canMoveAbsolute = null;
        bool? canSync = null;
        IAlignerConfigurationV1 alignerConfiguration = null;
        AlignerType alignerType = AlignerType.Unset;

        AlignerUnits units = AlignerUnits.Unset;

        bool backlashCompensation = false;
        bool okToTestMethods = false; // Flag indicating whether position can be read and can methods are as expected.

        double? xMaxMove = null;
        double? xMaxPosition = null;
        double? xMinPosition = null;
        double? xResolution = null;
        double? yMaxMove = null;
        double? yMaxPosition = null;
        double? yMinPosition = null;
        double? yResolution = null;
        double? zMaxMove = null;
        double? zMaxPosition = null;
        double? zMinPosition = null;
        double? zResolution = null;

        bool isMoving = false;

        double? xPosition = null;
        double? yPosition = null;
        double? zPosition = null;

        private enum MovementType
        {
            Absolute,
            Relative
        }

        public AlignerTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken, CancellationTokenSource conformCancellationTokenSource, ConformResults conformResults)
            : base(true, true, true, false, false, true, false, true, conformConfiguration, logger, conformCancellationToken, conformCancellationTokenSource, conformResults)
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        protected override void Dispose(bool disposing)
        {
            LogDebug("Dispose", $"Disposing of device: {disposing} {disposedValue}");
            if (!disposedValue)
            {
                if (disposing)
                {
                    aligner?.Dispose();
                    aligner = null;
                }
            }

            base.Dispose(disposing);
            disposedValue = true;
        }

        public override void InitialiseTest()
        {
            unchecked
            {
                ExNotImplemented = (int)0x80040400;
                ExInvalidValue1 = (int)0x80040405;
                ExInvalidValue2 = (int)0x80040405;
                ExInvalidValue3 = (int)0x80040405;
                ExInvalidValue4 = (int)0x80040405;
                ExInvalidValue5 = (int)0x80040405;
                ExInvalidValue6 = (int)0x80040405;
                ExNotSet1 = (int)0x80040403;
            }

            base.InitialiseTest();
        }

        public override void CreateDevice()
        {
            try
            {
                if (settings.DeviceTechnology != DeviceTechnology.Alpaca)
                    throw new ASCOM.InvalidValueException("Aligner devices support Alpaca client access only; COM Aligner devices are not supported.");

                LogInfo("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                aligner = new AlpacaAligner(
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

                LogInfo("CreateDevice", "Alpaca device created OK");
                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(aligner, DeviceTypes.Aligner);

                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);

                ValidateInterfaceVersion();
            }
            catch (Exception ex)
            {
                LogDebug("CreateDevice", $"Exception thrown: {ex.Message}\r\n{ex}");
                throw;
            }
        }

        public override void ReadCanProperties()
        {
            // Test IAlignerV1.Configuration before all other IAlignerV1 properties.
            try
            {
                LogCallToDriver("Configuration", "About to get Configuration property");
                alignerConfiguration = aligner.Configuration;
                if (alignerConfiguration is null)
                {
                    LogIssue("Configuration", "The device did not return a valid IAlignerConfigurationV1 object.");
                    LogInfo("Configuration", "Cannot continue with further tests because the Configuration response was invalid.");
                    conformCancellationTokenSource.Cancel();
                    return;
                }

                LogDebug("Configuration", $"Configuration property correctly returned an IAlignerConfigurationV1 object.");
                if (cancellationToken.IsCancellationRequested) return;

                // Now validate the configuration properties
                canFindHome = CanTest("CanFindHome", () => alignerConfiguration.CanFindHome);
                canMoveX = CanTest("CanMoveX", () => alignerConfiguration.CanMoveX);
                canMoveY = CanTest("CanMoveY", () => alignerConfiguration.CanMoveY);
                canMoveZ = CanTest("CanMoveZ", () => alignerConfiguration.CanMoveZ);
                canMoveAbsolute = CanTest("CanMoveAbsolute", () => alignerConfiguration.CanMoveAbsolute);
                canSync = CanTest("CanSync", () => alignerConfiguration.CanSync);

            }
            catch (Exception ex)
            {
                LogIssue("CheckProperties", $"Exception thrown when accessing Configuration property: {ex.Message}");
                LogDebug("CheckProperties", $"Exception thrown: {ex.Message}\r\n{ex}");
            }
        }

        public override void CheckProperties()
        {
            // Test IAlignerV1.Configuration before all other IAlignerV1 properties.
            try
            {
                // If we get here, property reads were successful, which means that Configuration was read successfully so record it here
                LogOk("Configuration", $"Configuration property correctly returned an IAlignerConfigurationV1 object.");
                if (cancellationToken.IsCancellationRequested) return;

                // Now validate the configuration properties

                // Verify Configuration.AlignerType is accessible and cannot throw PropertyNotImplementedException.
                try
                {
                    alignerType = alignerConfiguration.AlignerType;
                    LogOk("AlignerType", alignerType.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("AlignerType", MemberType.Property, Required.Mandatory, ex, "");
                }

                // Verify Configuration.BacklashCompensation is accessible and cannot throw PropertyNotImplementedException.
                try
                {
                    backlashCompensation = alignerConfiguration.BacklashCompensation;
                    LogOk("BacklashCompensation", backlashCompensation.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("BacklashCompensation", MemberType.Property, Required.Mandatory, ex, "");
                }

                // Verify Configuration.Units is accessible and cannot throw PropertyNotImplementedException.
                try
                {
                    units = alignerConfiguration.Units;
                    LogOk("Units", units.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("Units", MemberType.Property, Required.Mandatory, ex, "");
                }

                xMaxMove = DoubleTest("XMaxMove", () => alignerConfiguration.XMaxMove);
                xMaxPosition = DoubleTest("XMaxPosition", () => alignerConfiguration.XMaxPosition);
                xMinPosition = DoubleTest("XMinPosition", () => alignerConfiguration.XMinPosition);
                xResolution = DoubleTest("XResolution", () => alignerConfiguration.XResolution);
                yMaxMove = DoubleTest("YMaxMove", () => alignerConfiguration.YMaxMove);
                yMaxPosition = DoubleTest("YMaxPosition", () => alignerConfiguration.YMaxPosition);
                yMinPosition = DoubleTest("YMinPosition", () => alignerConfiguration.YMinPosition);
                yResolution = DoubleTest("YResolution", () => alignerConfiguration.YResolution);
                zMaxMove = DoubleTest("ZMaxMove", () => alignerConfiguration.ZMaxMove);
                zMaxPosition = DoubleTest("ZMaxPosition", () => alignerConfiguration.ZMaxPosition);
                zMinPosition = DoubleTest("ZMinPosition", () => alignerConfiguration.ZMinPosition);
                zResolution = DoubleTest("ZResolution", () => alignerConfiguration.ZResolution);

                // Verify IsMoving is accessible and cannot throw PropertyNotImplementedException.
                try
                {
                    isMoving = aligner.IsMoving;
                    LogOk("IsMoving", isMoving.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("IsMoving", MemberType.Property, Required.Mandatory, ex, "");
                }

                // Verify XPosition is accessible and can throw PropertyNotImplementedException.
                try
                {
                    xPosition = aligner.XPosition;
                    LogOk("XPosition", xPosition.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("XPosition", MemberType.Property, Required.Optional, ex, "");
                }

                // Verify YPosition is accessible and can throw PropertyNotImplementedException.
                try
                {
                    yPosition = aligner.YPosition;
                    LogOk("YPosition", yPosition.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("YPosition", MemberType.Property, Required.Optional, ex, "");
                }

                // Verify ZPosition is accessible and can throw PropertyNotImplementedException.
                try
                {
                    zPosition = aligner.ZPosition;
                    LogOk("ZPosition", zPosition.ToString());
                }
                catch (Exception ex)
                {
                    HandleException("ZPosition", MemberType.Property, Required.Optional, ex, "");
                }

                LogNewLine();
                LogTestOnly("Configuration Rules");

                // Check that XMinPosition and XMaxPosition are both implemented or not implemented
                if ((xMinPosition is null & xMaxPosition is null) | (xMinPosition is not null & xMaxPosition is not null))
                    LogOk("XMinPosition/XMaxPosition", $"Both XMinPosition and XMaxPosition are {(xMinPosition is not null ? "implemented" : "not implemented")}.");
                else
                    LogIssue("XMinPosition/XMaxPosition", $"XMinPosition and XMaxPosition must either both be implemented or bot be not implemented. " +
                        $"XMinPosition is {(xMinPosition is not null ? "implemented" : "not implemented")} while XMaxPosition is {(xMaxPosition is not null ? "implemented" : "not implemented")}.");

                // Check that YMinPosition and YMaxPosition are both implemented or not implemented
                if ((yMinPosition is null & yMaxPosition is null) | (yMinPosition is not null & yMaxPosition is not null))
                    LogOk("YMinPosition/YMaxPosition", $"Both YMinPosition and YMaxPosition are {(yMinPosition is not null ? "implemented" : "not implemented")}.");
                else
                    LogIssue("YMinPosition/YMaxPosition", $"YMinPosition and YMaxPosition must either both be implemented or bot be not implemented. " +
                        $"YMinPosition is {(yMinPosition is not null ? "implemented" : "not implemented")} while YMaxPosition is {(yMaxPosition is not null ? "implemented" : "not implemented")}.");

                // Check that ZMinPosition and ZMaxPosition are both implemented or not implemented
                if ((zMinPosition is null & zMaxPosition is null) | (zMinPosition is not null & zMaxPosition is not null))
                    LogOk("ZMinPosition/ZMaxPosition", $"Both ZMinPosition and ZMaxPosition are {(zMinPosition is not null ? "implemented" : "not implemented")}.");
                else
                    LogIssue("ZMinPosition/ZMaxPosition", $"ZMinPosition and ZMaxPosition must either both be implemented or bot be not implemented. " +
                        $"ZMinPosition is {(zMinPosition is not null ? "implemented" : "not implemented")} while ZMaxPosition is {(zMaxPosition is not null ? "implemented" : "not implemented")}.");

                // Check that aligners can move appropriate axes
                switch (alignerType)
                {
                    // Aligner type not set
                    case AlignerType.Unset:
                        LogIssue("Axis Movement", $"The aligner type is \"Unset\" so axis movement rules cannot be checked.");
                        break;

                    // Polar aligner
                    case AlignerType.PolarAligner:
                        if (canMoveX.GetValueOrDefault() & canMoveY.GetValueOrDefault() & !canMoveZ.GetValueOrDefault())
                        {
                            LogOk("Axis Movement", "The polar aligner can move its X and Y axes and cannot move its Z axis.");

                            // Check whether it is OK to test the move methods. This is only possible if both X and Y positions can be read.
                            if (xPosition.HasValue & yPosition.HasValue) // Both axes are readable, so it is OK to test the move methods.
                                okToTestMethods = true;
                            else // There is a problem so don't test the move methods
                                LogIssue("Axis Movement", $"The polar aligner can move both X and Y axes, but at least one of the axes cannot be read. XPosition: {xPosition}, YPosition: {yPosition}");
                        }
                        else
                        {
                            LogIssue("Axis Movement", $"Polar aligners must be able to move both X and Y axes and not be able to move the Z axis. ");
                            LogInfo("Axis Movement", $"CanMoveX: {canMoveX.GetValueOrDefault()}, CanMoveY: {canMoveY.GetValueOrDefault()}, CanMoveZ: {canMoveZ.GetValueOrDefault()}");
                        }
                        break;

                    //Optical axis aligner and all other types
                    case AlignerType.OpticalAxisAligner:
                    default:
                        // Check that at least one axis is movable for optical axis aligners
                        if (canMoveX.GetValueOrDefault() | canMoveY.GetValueOrDefault() | canMoveZ.GetValueOrDefault())
                        {
                            LogOk("Axis Movement", $"At least one aligner axis can move.");

                            // Check whether it is OK to test the move methods. This is only possible if every movable axis position can be read.
                            if ((!canMoveX.GetValueOrDefault() || xPosition.HasValue) &
                                (!canMoveY.GetValueOrDefault() || yPosition.HasValue) &
                                (!canMoveZ.GetValueOrDefault() || zPosition.HasValue))
                                okToTestMethods = true;
                            else // There is a problem so don't test the move methods
                            {
                                LogIssue("Axis Movement", $"Method tests will be skipped because the aligner can move at least one axis, but at least one of the movable axes can not be read.");
                                LogInfo("Axis Movement", $"CanMoveX: {canMoveX.GetValueOrDefault()}, CanMoveY: {canMoveY.GetValueOrDefault()}, CanMoveZ: {canMoveZ.GetValueOrDefault()}, XPosition: {xPosition}, YPosition: {yPosition}, ZPosition: {zPosition}");
                            }
                        }
                        else
                            LogIssue("Axis Movement", $"None of the axes are configured as moveable - CanMoveX: {canMoveX.GetValueOrDefault()}, CanMoveY: {canMoveY.GetValueOrDefault()}, CanMoveZ: {canMoveZ.GetValueOrDefault()}");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogIssue("CheckProperties", $"Exception thrown when accessing Configuration property: {ex.Message}");
                LogDebug("CheckProperties", $"Exception thrown: {ex.Message}\r\n{ex}");
            }
        }

        public override void CheckMethods()
        {
            // Check whether there is sufficient information to test methods
            if (okToTestMethods) // OK to test methods
            {
                // Save the start position so we can return to it at the end of testing
                Vector3D startingPosition = new(xPosition.GetValueOrDefault(), yPosition.GetValueOrDefault(), zPosition.GetValueOrDefault());

                // Test relative movement when a safe non-zero vector can be calculated.
                Vector3D? relativeMovement = GetMovePosition(MovementType.Relative, canMoveX.GetValueOrDefault(), canMoveY.GetValueOrDefault(), canMoveZ.GetValueOrDefault());
                if (relativeMovement.HasValue)
                {
                    // Verify position accuracy only after the asynchronous move completes successfully.
                    if (MoveAligner("Move", () => aligner.Move(relativeMovement.Value.X, relativeMovement.Value.Y, relativeMovement.Value.Z)))
                    {
                        CheckPosition("Move", new Vector3D(
                            startingPosition.X + relativeMovement.Value.X,
                            startingPosition.Y + relativeMovement.Value.Y,
                            startingPosition.Z + relativeMovement.Value.Z));
                    }

                    RestorePosition(startingPosition);
                }
                else
                    LogInfo("Move", "Move test skipped because a safe relative movement could not be determined.");

                // Test absolute movement only when the configuration advertises support.
                if (canMoveAbsolute.GetValueOrDefault())
                {
                    Vector3D? absolutePosition = GetMovePosition(MovementType.Absolute, canMoveX.GetValueOrDefault(), canMoveY.GetValueOrDefault(), canMoveZ.GetValueOrDefault());
                    if (absolutePosition.HasValue)
                    {
                        if (MoveAligner("MoveAbsolute", () => aligner.MoveAbsolute(absolutePosition.Value.X, absolutePosition.Value.Y, absolutePosition.Value.Z)))
                            CheckPosition("MoveAbsolute", absolutePosition.Value);

                        RestorePosition(startingPosition);
                    }
                    else
                        LogInfo("MoveAbsolute", "MoveAbsolute test skipped because a safe absolute position could not be determined.");
                }
                else
                {
                    TestUnimplementedMethod("MoveAbsolute", () => aligner.MoveAbsolute(0.0, 0.0, 0.0), "CanMoveAbsolute is false");
                    RestorePosition(startingPosition);
                }

                // Test homing when supported; otherwise verify the required not-implemented response.
                if (canFindHome.GetValueOrDefault())
                {
                    MoveAligner("FindHome", () => aligner.FindHome());
                    RestorePosition(startingPosition);
                }
                else
                {
                    TestUnimplementedMethod("FindHome", () => aligner.FindHome(), "CanFindHome is false");
                    RestorePosition(startingPosition);
                }

                // TODO: Test IAlignerV1.Halt().
                // TODO: Test IAlignerV1.Sync(double X, double Y, double Z).
            }
            else
                LogInfo("CheckMethods", "Method tests were skipped because of issues in returned properties reported earlier..");
        }

        public override void CheckPerformance()
        {
            SetTest("Performance");

            PerformanceTestCommon(ApplicationCancellationToken);

            SetAction("");
            SetStatus("");
            SetTest("");
        }

        public override void CheckConfiguration()
        {
            try
            {
                if (!settings.TestProperties)
                    LogConfigurationAlert("Property tests were omitted due to Conform configuration.");

                if (!settings.TestMethods)
                    LogConfigurationAlert("Method tests were omitted due to Conform configuration.");

                if (settings.AllowConnectedTrueAfterDisconnect)
                    LogConfigurationAlert("Connected is allowed to be true after disconnection.");
            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n{ex}");
            }
        }

        #region Support Code

        /// <summary>
        /// Calculates a safe relative movement or absolute target for the movable axes.
        /// </summary>
        /// <returns>The calculated position, or null when a safe movement cannot be determined.</returns>
        private Vector3D? GetMovePosition(MovementType movementType, bool canMoveX, bool canMoveY, bool canMoveZ)
        {
            double? x = GetAxisMovePosition(movementType, canMoveX, xPosition, xMinPosition, xMaxPosition, xMaxMove, xResolution, "X");
            double? y = GetAxisMovePosition(movementType, canMoveY, yPosition, yMinPosition, yMaxPosition, yMaxMove, yResolution, "Y");
            double? z = GetAxisMovePosition(movementType, canMoveZ, zPosition, zMinPosition, zMaxPosition, zMaxMove, zResolution, "Z");

            if ((canMoveX & !x.HasValue) | (canMoveY & !y.HasValue) | (canMoveZ & !z.HasValue))
                return null;

            return new Vector3D(x.GetValueOrDefault(), y.GetValueOrDefault(), z.GetValueOrDefault());
        }

        /// <summary>
        /// Calculates movement for one axis in the direction with the most available range.
        /// </summary>
        /// <returns>A relative movement or absolute target, or null when the calculation is unsafe.</returns>
        private double? GetAxisMovePosition(
            MovementType movementType,
            bool canMoveAxis,
            double? currentPosition,
            double? minimumPosition,
            double? maximumPosition,
            double? maximumMove,
            double? resolution,
            string axisName)
        {
            // Non-movable axes must receive zero.
            if (!canMoveAxis)
                return 0.0;

            // Accurate movement verification requires a readable starting position.
            if (!currentPosition.HasValue)
            {
                LogIssue("Move Position", $"{axisName} can move but its current position is unavailable, so a safe test position cannot be calculated.");
                return null;
            }

            bool hasPositionRange = minimumPosition.HasValue & maximumPosition.HasValue;
            // A usable position range must either be fully implemented or fully unimplemented.
            if ((minimumPosition.HasValue ^ maximumPosition.HasValue) ||
                (hasPositionRange && minimumPosition.Value > maximumPosition.Value))
            {
                LogIssue("Move Position", $"{axisName} position limits are invalid: minimum {minimumPosition}, maximum {maximumPosition}.");
                return null;
            }

            // Absolute movement requires known bounds so the target cannot exceed the device range.
            if (movementType == MovementType.Absolute && !hasPositionRange)
            {
                LogInfo("Move Position", $"{axisName} position limits are not implemented, so a safe absolute test position cannot be calculated.");
                return null;
            }

            double direction = 1.0;
            double? availableMovement = null;
            double? movementRange = null;

            if (hasPositionRange)
            {
                double positiveMovement = maximumPosition.Value - currentPosition.Value;
                double negativeMovement = currentPosition.Value - minimumPosition.Value;
                movementRange = maximumPosition.Value - minimumPosition.Value;

                // Reject inconsistent devices rather than risk moving beyond a configured limit.
                if (positiveMovement < 0.0 || negativeMovement < 0.0)
                {
                    LogIssue("Move Position", $"{axisName} current position {currentPosition.Value} is outside its configured range {minimumPosition.Value} to {maximumPosition.Value}.");
                    return null;
                }

                // Move toward the larger side of the available position range.
                if (negativeMovement > positiveMovement)
                    direction = -1.0;

                availableMovement = Math.Max(positiveMovement, negativeMovement);
            }

            double movement = GetTestMovement(movementRange, maximumMove, resolution);
            if (availableMovement.HasValue)
            {
                // A resolvable move must fit in the larger side of the range.
                if (resolution.HasValue && resolution.Value > availableMovement.Value)
                {
                    LogIssue("Move Position", $"{axisName} has only {availableMovement.Value} {units} available in the larger direction, which is less than its resolution {resolution.Value}.");
                    return null;
                }

                movement = Math.Min(movement, availableMovement.Value);
            }

            if (!(movement > 0.0) || !double.IsFinite(movement) || (maximumMove.HasValue && maximumMove.Value >= 0.0 && movement > maximumMove.Value))
            {
                // Reject zero, non-finite, or command-limit-exceeding movements.
                LogIssue("Move Position", $"A valid non-zero {axisName} movement could not be calculated. Maximum move: {maximumMove}, range: {movementRange}, resolution: {resolution}.");
                return null;
            }

            movement *= direction;
            return movementType == MovementType.Absolute ? currentPosition.Value + movement : movement;
        }

        /// <summary>
        /// Chooses a small, executable test movement from the available range metadata.
        /// </summary>
        private double GetTestMovement(double? movementRange, double? maximumMove, double? resolution)
        {
            double movement = 0.0;

            if (movementRange.HasValue && movementRange.Value > 0.0)
            {
                movement = movementRange.Value / 10.0;

                if (maximumMove.HasValue && maximumMove.Value > 0.0)
                    movement = Math.Min(movement, Math.Abs(maximumMove.Value));
            }
            else if (maximumMove.HasValue && maximumMove.Value > 0.0)
            {
                movement = Math.Abs(maximumMove.Value) / 10.0;
            }

            if (resolution.HasValue && resolution.Value > 0.0)
                movement = movement > 0.0 ? Math.Max(movement, resolution.Value) : resolution.Value;

            if (!(movement > 0.0))
            {
                // Use a conservative unit-appropriate fallback when the device exposes no range metadata.
                movement = alignerType switch
                {
                    AlignerType.PolarAligner => 0.1,
                    AlignerType.OpticalAxisAligner => 0.1,
                    _ => units == AlignerUnits.Counts ? 1.0 : 0.1
                };
            }

            return movement;
        }

        /// <summary>
        /// Executes an asynchronous aligner operation and waits for completion.
        /// </summary>
        /// <returns>True when the operation completes and IsMoving becomes false.</returns>
        private bool MoveAligner(string testName, Action moveAction)
        {
            try
            {
                // Do not interfere with movement initiated before this test.
                LogCallToDriver(testName, "About to get IsMoving property");
                if (aligner.IsMoving)
                {
                    LogIssue(testName, "Aligner is already moving before the test, so the method test was skipped.");
                    return false;
                }

                SetAction(testName);
                LogCallToDriver(testName, $"About to call {testName} method");
                TimeMethod(testName, moveAction, TargetTime.Standard);

                LogCallToDriver(testName, "About to get IsMoving property repeatedly");
                WaitWhile($"Waiting for {testName} to complete", () => aligner.IsMoving, 500, settings.AlpacaConfiguration.LongResponseTimeout);

                if (cancellationToken.IsCancellationRequested)
                    return false;

                LogCallToDriver(testName, "About to get IsMoving property");
                if (aligner.IsMoving)
                {
                    LogIssue(testName, "IsMoving is still true after the move completed.");
                    return false;
                }

                LogOk(testName, $"{testName} completed and IsMoving is false.");
                return true;
            }
            catch (Exception ex)
            {
                HandleException(testName, MemberType.Method, Required.MustBeImplemented, ex, "");
                return false;
            }
        }

        /// <summary>
        /// Verifies that a capability-disabled method throws a not-implemented exception.
        /// </summary>
        private void TestUnimplementedMethod(string testName, Action testAction, string reason)
        {
            try
            {
                LogCallToDriver(testName, $"About to call {testName} method");
                testAction();
                LogIssue(testName, $"{reason}, but {testName} did not throw a MethodNotImplementedException.");

                LogCallToDriver(testName, "About to get IsMoving property");
                if (aligner.IsMoving)
                {
                    LogCallToDriver(testName, "About to call Halt method to stop unexpected movement");
                    aligner.Halt();
                }
            }
            catch (Exception ex)
            {
                HandleException(testName, MemberType.Method, Required.MustNotBeImplemented, ex, reason);
            }
        }

        /// <summary>
        /// Checks each movable axis against the expected final position.
        /// </summary>
        private void CheckPosition(string testName, Vector3D expectedPosition)
        {
            // Check only axes that the configuration says can move.
            if (canMoveX.GetValueOrDefault())
                CheckAxisPosition(testName, "X", expectedPosition.X, xResolution, () => aligner.XPosition);

            if (canMoveY.GetValueOrDefault())
                CheckAxisPosition(testName, "Y", expectedPosition.Y, yResolution, () => aligner.YPosition);

            if (canMoveZ.GetValueOrDefault())
                CheckAxisPosition(testName, "Z", expectedPosition.Z, zResolution, () => aligner.ZPosition);
        }

        /// <summary>
        /// Checks one axis position within the reported resolution or a numeric tolerance.
        /// </summary>
        private void CheckAxisPosition(string testName, string axisName, double expectedPosition, double? resolution, Func<double> getPosition)
        {
            try
            {
                LogCallToDriver(testName, $"About to get {axisName}Position property");
                double actualPosition = getPosition();
                double tolerance = resolution.HasValue && resolution.Value > 0.0
                    ? resolution.Value
                    : Math.Max(1.0e-9, Math.Abs(expectedPosition) * 1.0e-9);

                if (Math.Abs(actualPosition - expectedPosition) <= tolerance)
                    LogOk(testName, $"{axisName} position is {actualPosition}, matching expected position {expectedPosition} within tolerance {tolerance}.");
                else
                    LogIssue(testName, $"{axisName} position is {actualPosition}, which differs from expected position {expectedPosition} by {Math.Abs(actualPosition - expectedPosition)}; tolerance is {tolerance}.");
            }
            catch (Exception ex)
            {
                HandleException(testName, MemberType.Property, Required.MustBeImplemented, ex, $"{axisName}Position was readable earlier in the test run");
            }
        }

        /// <summary>
        /// Returns all movable axes to the position captured before method testing.
        /// </summary>
        private void RestorePosition(Vector3D startingPosition)
        {
            if (cancellationToken.IsCancellationRequested)
                return;

            SetAction("Restoring aligner position");

            if (canMoveAbsolute.GetValueOrDefault())
            {
                // Absolute-capable devices can restore all axes in one operation.
                MoveAligner(
                    "Restore Position",
                    () => aligner.MoveAbsolute(startingPosition.X, startingPosition.Y, startingPosition.Z));
                return;
            }

            for (int pass = 0; pass < 100 && !cancellationToken.IsCancellationRequested; pass++)
            {
                Vector3D? relativeMovement = GetRestoreMovement(startingPosition);
                if (!relativeMovement.HasValue)
                    return;

                if (MoveIsComplete(relativeMovement.Value))
                    return;

                if (!MoveAligner(
                    "Restore Position",
                    () => aligner.Move(relativeMovement.Value.X, relativeMovement.Value.Y, relativeMovement.Value.Z)))
                {
                    return;
                }
            }

            if (!cancellationToken.IsCancellationRequested)
                LogIssue("Restore Position", "The aligner did not return to its starting position after 100 relative moves.");
        }

        /// <summary>
        /// Calculates the relative movement required to restore all axes.
        /// </summary>
        /// <returns>The restoration movement, or null when a position cannot be read.</returns>
        private Vector3D? GetRestoreMovement(Vector3D startingPosition)
        {
            try
            {
                double x = GetRestoreAxisMovement(canMoveX.GetValueOrDefault(), () => aligner.XPosition, startingPosition.X, xMaxMove, xResolution);
                double y = GetRestoreAxisMovement(canMoveY.GetValueOrDefault(), () => aligner.YPosition, startingPosition.Y, yMaxMove, yResolution);
                double z = GetRestoreAxisMovement(canMoveZ.GetValueOrDefault(), () => aligner.ZPosition, startingPosition.Z, zMaxMove, zResolution);
                return new Vector3D(x, y, z);
            }
            catch (Exception ex)
            {
                HandleException("Restore Position", MemberType.Property, Required.MustBeImplemented, ex, "Unable to read the current position for restoration");
                return null;
            }
        }

        /// <summary>
        /// Calculates and limits one axis's movement toward its starting position.
        /// </summary>
        private double GetRestoreAxisMovement(bool canMoveAxis, Func<double> getPosition, double startingPosition, double? maximumMove, double? resolution)
        {
            if (!canMoveAxis)
                return 0.0;

            LogCallToDriver("Restore Position", "About to get current position property");
            double remainingMovement = startingPosition - getPosition();
            if (Math.Abs(remainingMovement) <= GetPositionTolerance(startingPosition, resolution))
                return 0.0;

            return maximumMove.HasValue && maximumMove.Value > 0.0
                ? Math.Clamp(remainingMovement, -maximumMove.Value, maximumMove.Value)
                : remainingMovement;
        }

        /// <summary>
        /// Determines whether all axis restoration movements are zero.
        /// </summary>
        private static bool MoveIsComplete(Vector3D movement)
        {
            return movement.X == 0.0 & movement.Y == 0.0 & movement.Z == 0.0;
        }

        /// <summary>
        /// Returns the position comparison tolerance for an axis.
        /// </summary>
        private static double GetPositionTolerance(double position, double? resolution)
        {
            return resolution.HasValue && resolution.Value > 0.0
                ? resolution.Value
                : Math.Max(1.0e-9, Math.Abs(position) * 1.0e-9);
        }

        /// <summary>
        /// Test a CanXXX function and return its value or null if an error occurs.
        /// </summary>
        /// <param name="testName">CanXXX function name</param>
        /// <param name="canFunction">The CanXXX function to test</param>
        /// <returns></returns>
        private bool? CanTest(string testName, Func<bool> canFunction)
        {
            // Initialise response to null, indicating failure
            bool? response = null;

            try
            {
                LogCallToDriver(testName, $"About to get {testName} property");

                TimeMethod(testName, () =>
                {
                    // Call the CanXXX function and record its value
                    response = canFunction();
                }, TargetTime.Fast);
                LogOk(testName, response.ToString());
            }
            catch (Exception ex)
            {
                HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
            }

            // Return the value
            return response;
        }

        /// <summary>
        /// Test a CanXXX function and return its value or null if an error occurs.
        /// </summary>
        /// <param name="testName">CanXXX function name</param>
        /// <param name="doubleFunction">The CanXXX function to test</param>
        /// <returns></returns>
        private double? DoubleTest(string testName, Func<double> doubleFunction)
        {
            // Initialise response to null, indicating failure
            double? response = null;

            try
            {
                LogCallToDriver(testName, $"About to get {testName} property");

                TimeMethod(testName, () =>
                {
                    // Call the double function and record its value
                    response = doubleFunction();
                }, TargetTime.Fast);
                LogOk(testName, response.ToString());
            }
            catch (Exception ex)
            {
                HandleException(testName, MemberType.Property, Required.Optional, ex, "");
            }

            // Return the value
            return response;
        }


        #endregion
    }
}

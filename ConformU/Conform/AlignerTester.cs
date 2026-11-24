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
            Abolute,
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

                            // Check whether it is OK to test the move methods. This is only possible if both X and Y positions can be read.
                            if ((canMoveX.GetValueOrDefault() & xPosition.HasValue) | (canMoveY.GetValueOrDefault() & yPosition.HasValue) | (canMoveZ.GetValueOrDefault() & zPosition.HasValue)) // All implemented axes are readable so it is OK to test the move methods.
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
                Vector3D? startingPosition = new Vector3D(xPosition.GetValueOrDefault(), yPosition.GetValueOrDefault(), zPosition.GetValueOrDefault());

                // TODO: Test IAlignerV1.FindHome().
                // TODO: Test IAlignerV1.Halt().
                // TODO: Test IAlignerV1.Move(double X, double Y, double Z).
                // TODO: Test IAlignerV1.MoveAbsolute(double X, double Y, double Z).
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

        private Vector3D? GetMovePosition(MovementType movementType, bool canMoveX, bool canMoveY, bool canMoveZ)
        {
            Vector3D? position = new Vector3D();

            return position;
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

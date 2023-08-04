using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConformU
{
    internal class FilterWheelTester : DeviceTesterBaseClass
    {
        private const int FWTEST_IS_MOVING = -1;

        private IFilterWheelV3 filterWheel;

        private enum FilterWheelProperties
        {
            FocusOffsets,
            Names,
            Position
        }

        // Helper variables
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose
        public FilterWheelTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, false, true, false, false, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
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
                    filterWheel?.Dispose();
                    filterWheel = null;
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
                            GExNotImplemented = (int)0x80040400;
                            GExInvalidValue1 = (int)0x80040404;
                            GExInvalidValue2 = (int)0x80040404;
                            GExNotSet1 = (int)0x80040403;
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
                        filterWheel = new AlpacaFilterWheel(
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
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComAccessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                filterWheel = new FilterWheelFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                filterWheel = new FilterWheel(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                BaseClassDevice = filterWheel; // Assign the driver to the base class

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

        public override void PreRunCheck()
        {
            DateTime startTime;

            // Get into a consistent state
            SetFullStatus("FilterWheel Pre-run Check", "Wait one second for initialisation", "");
            WaitFor(1000); // Wait for 1 second to allow any movement to start
            startTime = DateTime.Now;
            try
            {
                LogCallToDriver("Pre-run Check", "About to get Position property repeatedly");
                do
                {
                    SetFullStatus("FilterWheel Pre-run Check", "Waiting for movement to stop", DateTime.Now.Subtract(startTime).Seconds + " second(s)");
                    WaitFor(SLEEP_TIME);
                }
                while ((filterWheel.Position == FWTEST_IS_MOVING) & (DateTime.Now.Subtract(startTime).TotalSeconds <= settings.FilterWheelTimeout)); // Wait until movement has stopped or 30 seconds have passed
                if (filterWheel.Position != FWTEST_IS_MOVING)
                    LogOk("Pre-run Check", "Filter wheel is stationary, ready to start tests");
                else
                {
                    LogIssue("Pre-run Check", $"The filter wheel is still moving after {settings.FilterWheelTimeout} seconds, further tests abandoned because the device is not in the expected stationary state.");
                    ResetTestActionStatus();
                    return;
                }
            }
            catch (Exception ex)
            {
                LogIssue("Pre-run Check", $"Unable to determine that the Filter wheel is stationary: {ex.Message}");
                LogDebug("Pre-run Check", $"Exception detail: {ex}");
            }

            ResetTestActionStatus();
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(filterWheel, DeviceTypes.FilterWheel);
        }

        public override void CheckProperties()
        {
            int[] filterOffsets;
            int numberOfFilternames = 0, numberOfFilterOffsets = 0, filterNumber, startingFilterNumber;

            int maxFilterNumber = 0; // The highest populated filter slot number (1 less than the number of filters)
            int halfDistanceFilterNumber = 0; // A filter slot number close to the middle of the filter wheel

            string[] filterNames;

            // FocusOffsets - Required - Read only
            try
            {
                LogCallToDriver("FocusOffsets Get", "About to get FocusOffsets property");
                filterOffsets = filterWheel.FocusOffsets;
                numberOfFilterOffsets = filterOffsets.Length;
                if (numberOfFilterOffsets == 0)
                    LogIssue("FocusOffsets Get", "Found no offset values in the returned array");
                else
                {
                    LogOk("FocusOffsets Get", "Found " + numberOfFilterOffsets.ToString() + " filter offset values");
                    maxFilterNumber = Convert.ToInt16(numberOfFilterOffsets - 1);
                    halfDistanceFilterNumber = maxFilterNumber / 2;
                }

                filterNumber = 0;
                foreach (var offset in filterOffsets)
                {
                    LogInfo("FocusOffsets Get", "Filter " + filterNumber.ToString() + " Offset: " + offset.ToString());
                    filterNumber += 1;
                    if (cancellationToken.IsCancellationRequested)
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleException("FocusOffsets Get", MemberType.Property, Required.Mandatory, ex, "");
            }

            // Names - Required - Read only
            try
            {
                LogCallToDriver("Names Get", "About to get Names property");
                filterNames = filterWheel.Names;
                numberOfFilternames = filterNames.Length;
                if (numberOfFilternames == 0)
                    LogIssue("Names Get", "Did not find any names in the returned array");
                else
                    LogOk("Names Get", "Found " + numberOfFilternames.ToString() + " filter names");
                filterNumber = 0;
                foreach (var name in filterNames)
                {
                    if (name == null)
                        LogIssue("Names Get", "Filter " + filterNumber.ToString() + " has a value of nothing");
                    else if (name == "")
                        LogIssue("Names Get", "Filter " + filterNumber.ToString() + " has a value of \"\"");
                    else
                        LogInfo("Names Get", "Filter " + filterNumber.ToString() + " Name: " + name);
                    filterNumber += 1;
                }
            }
            catch (Exception ex)
            {
                HandleException("Names Get", MemberType.Property, Required.Mandatory, ex, "");
            }

            // Confirm number of array elements in filter names and filter offsets are the same
            if (numberOfFilternames == numberOfFilterOffsets)
                LogOk("Names Get", "Number of filter offsets and number of names are the same: " + numberOfFilternames.ToString());
            else
                LogIssue("Names Get", "Number of filter offsets and number of names are different: " + numberOfFilterOffsets.ToString() + " " + numberOfFilternames.ToString());

            // Position - Required - Read
            try
            {
                SetTest("Position Get");
                LogCallToDriver("Position Get", "About to get Position property");
                startingFilterNumber = filterWheel.Position;
                if ((startingFilterNumber < 0) | (startingFilterNumber >= numberOfFilterOffsets))
                    LogIssue("Position Get", $"Illegal filter position returned: {startingFilterNumber}");
                else
                    LogOk("Position Get", $"Currently at position: {startingFilterNumber}");
            }
            catch (Exception ex)
            {
                HandleException("Position", MemberType.Property, Required.Mandatory, ex, "");
            }

            // Position - Required - Write
            SetTest("Position Set");

            // Make sure some filter slots are available
            if (numberOfFilterOffsets <= 0) // No filter slots available so exist
            {
                LogIssue("Position", "Filter position tests skipped as number of filters appears to be 0: " + numberOfFilterOffsets.ToString());
                ResetTestActionStatus();
                return;
            }

            // Some filters are available so test movement
            try
            {
                // Move sequentially upwards to each position in turn
                LogNewLine();
                LogTestOnly("Testing ascending sequential movement");

                double upwardMovewentTime = 0.0;
                for (short i = 0; i <= Convert.ToInt16(numberOfFilterOffsets - 1); i++)
                {
                    upwardMovewentTime += MoveToPosition(i);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Move sequentially downwards to each position in turn
                LogNewLine();
                LogTestOnly("Testing descending sequential movement");

                double downwardMovewentTime = 0.0;

                for (short i = Convert.ToInt16(numberOfFilterOffsets - 1); i >= 0; i--)
                {
                    downwardMovewentTime += MoveToPosition(i);
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Test long distance moves
                LogNewLine();
                LogTestOnly("Testing potentially long distance moves");

                // Move to position 0
                MoveToPosition(0);
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Move to highest position
                MoveToPosition(maxFilterNumber);
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Move to one down from the highest position (uni-directional filter wheels will have to travel a long way)
                if (maxFilterNumber > 0)
                {
                    MoveToPosition(maxFilterNumber - 1);
                }
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Move back to position 0
                MoveToPosition(0);
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Move to the half way position
                if (halfDistanceFilterNumber > 0)
                {
                    MoveToPosition(halfDistanceFilterNumber);
                }
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Move back to position 0
                MoveToPosition(0);
                if (cancellationToken.IsCancellationRequested)
                    return;

                LogNewLine();
                LogTestOnly("Testing error conditions");

                // Confirm that an error is correctly generated for outside range values
                try
                {
                    LogCallToDriver("Position Set", "About to set Position property");
                    filterWheel.Position = -1; // Negative position, positions should never be negative
                    LogIssue("Position Set", "Failed to generate exception when selecting filter with negative filter number");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to - 1", "Correctly rejected bad position: -1");
                }
                if (cancellationToken.IsCancellationRequested)
                    return;

                // Confirm that an error is correctly generated for outside range values
                try
                {
                    LogCallToDriver("Position Set", "About to set Position property");
                    filterWheel.Position = (short)numberOfFilterOffsets; // This should be 1 above the highest array element returned
                    LogIssue("Position Set", "Failed to generate exception when selecting filter outside expected range");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to " + System.Convert.ToString(numberOfFilterOffsets), "Correctly rejected bad position: " + System.Convert.ToString(numberOfFilterOffsets));
                }

                // Report on the uni-directional and bi-directional behaviour.
                LogNewLine();
                LogTestOnly("Directionality assessment");

                if (numberOfFilternames > 1) // 3 or more filters so there will be a difference between uni-directional and bi-directional filter wheels
                {
                    if (Math.Abs(upwardMovewentTime - downwardMovewentTime) < Math.Min(upwardMovewentTime, downwardMovewentTime) / numberOfFilterOffsets)
                    {
                        LogInfo("Directionality", $"The filter wheel is bi-directional.");
                    }
                    else
                    {
                        LogInfo("Directionality", $"The filter wheel is uni-directional.");
                    }
                }
                else // 1 or less filters so cannot tell whether the filter wheel is bi-directional or uni-directional.
                {
                    LogInfo("Directionality", $"The filter wheel has 0 or 1 filters and it is not possible to differentiate between uni-directional and bi-directional behaviour.");
                }

                LogDebug("Directionality", $"Overall upward movement time: {upwardMovewentTime:0.0}, Overall downward movement time: {downwardMovewentTime:0.0}, " +
                   $"Difference: {Math.Abs(upwardMovewentTime - downwardMovewentTime):0.0}, Smallest average movement time: {Math.Min(upwardMovewentTime, downwardMovewentTime) / numberOfFilterOffsets:0.0}.");
            }
            catch (Exception ex)
            {
                HandleException("Position Get", MemberType.Property, Required.Mandatory, ex, "");
            }

            ResetTestActionStatus();
        }
        public override void CheckPerformance()
        {
            FilterWheelPerformanceTest(FilterWheelProperties.FocusOffsets, "FocusOffsets");
            FilterWheelPerformanceTest(FilterWheelProperties.Names, "Names");
            FilterWheelPerformanceTest(FilterWheelProperties.Position, "Position");
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


        private void FilterWheelPerformanceTest(FilterWheelProperties pType, string pName)
        {
            int[] lOffsets;
            string[] lNames;
            int lStartFilterNumber;
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime;
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
                        case FilterWheelProperties.FocusOffsets:
                            {
                                lOffsets = filterWheel.FocusOffsets;
                                break;
                            }

                        case FilterWheelProperties.Names:
                            {
                                lNames = filterWheel.Names;
                                break;
                            }

                        case FilterWheelProperties.Position:
                            {
                                lStartFilterNumber = filterWheel.Position;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, "FilterWheelPerformanceTest: Unknown test type " + pType.ToString());
                                break;
                            }
                    }

                    lElapsedTime = DateTime.Now.Subtract(lStartTime).TotalSeconds;
                    if (lElapsedTime > lLastElapsedTime + 1.0)
                    {
                        SetStatus(lCount + " transactions in " + lElapsedTime.ToString("0") + " seconds");
                        lLastElapsedTime = lElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (lElapsedTime <= PERF_LOOP_TIME);
                lRate = lCount / lElapsedTime;
                switch (lRate)
                {
                    case object _ when lRate > 10.0:
                        {
                            LogInfo(pName, "Transaction rate: " + lRate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= lRate && lRate <= 10.0:
                        {
                            LogOk(pName, "Transaction rate: " + lRate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= lRate && lRate <= 2.0:
                        {
                            LogInfo(pName, "Transaction rate: " + lRate.ToString("0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(pName, "Transaction rate: " + lRate.ToString("0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(pName, "Unable to complete test: " + ex.Message);
            }
        }

        /// <summary>
        /// Move the filter wheel to the specified position and report the outcome
        /// </summary>
        /// <param name="position">Filter wheel position</param>
        /// <returns>Operation duration (seconds).</returns>
        private double MoveToPosition(int position)
        {
            double duration = 0.0;

            // Create a test name that incorporates the filter wheel position for use in log messages
            string testName = $"Position Set {position}";

            try
            {
                // Set the required position
                LogCallToDriver("Position Set", $"About to set Position property {position}");
                SetAction($"Setting position {position}");
                filterWheel.Position = Convert.ToInt16(position);

                // Wait for the reported position to match the required position
                Stopwatch sw = Stopwatch.StartNew();
                LogCallToDriver(testName, "About to get Position property repeatedly");
                WaitWhile($"Moving to position {position}", () => { return filterWheel.Position != position; }, 100, settings.FilterWheelTimeout);

                // Record the duration of the wait
                duration = sw.Elapsed.TotalSeconds;

                // Exit if the STOP button has been pressed
                if (cancellationToken.IsCancellationRequested)
                    return 0.0;

                // Get the reported position
                short reportedPosition = filterWheel.Position;

                // Test whether the reported position matches the required position
                if (reportedPosition == position) // The filter wheel is at the required position
                    LogOk(testName, $"Reached position: {position} in: {duration:0.0} seconds");
                else // The filter wheel is not at the required position so must have timed out
                    LogIssue(testName, $"The filter wheel did not reach specified position: {position} within the {settings.FilterWheelTimeout} second timeout. The reported position after {duration:0.0} seconds is: {reportedPosition}.");

                // Add a 1 second wait so as not to push the filter wheel too hard
                sw.Restart();
                WaitWhile($"Waiting for wheel to stabilise at position {position}", () => { return sw.ElapsedMilliseconds < 1000; }, 500, 1);
            }
            catch (Exception ex)
            {
                HandleException(testName, MemberType.Property, Required.Mandatory, ex, "");
            }

            // Return the elapsed time for the move ignoring the stabilisation time added at the end
            return duration;
        }
    }
}

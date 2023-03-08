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
        const int FWTEST_IS_MOVING = -1;

        private IFilterWheelV2 filterWheel;
        enum FilterWheelProperties
        {
            FocusOffsets,
            Names,
            Position
        }

        // Helper variables
        private ITelescopeV3 telescopeDevice;
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
                    telescopeDevice?.Dispose();
                    telescopeDevice = null;
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
                baseClassDevice = filterWheel; // Assign the driver to the base class

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
                LogCallToDriver("Connected", "About to get Connected property");
                return filterWheel.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                filterWheel.Connected = value;
            }
        }
        public override void PreRunCheck()
        {
            DateTime StartTime;

            // Get into a consistent state
            SetFullStatus("FilterWheel Pre-run Check", "Wait one second for initialisation", "");
            WaitFor(1000); // Wait for 1 second to allow any movement to start
            StartTime = DateTime.Now;
            try
            {
                LogCallToDriver("Pre-run Check", "About to get Position property repeatedly");
                do
                {
                    SetFullStatus("FilterWheel Pre-run Check", "Waiting for movement to stop", DateTime.Now.Subtract(StartTime).Seconds + " second(s)");
                    WaitFor(SLEEP_TIME);
                }
                while ((filterWheel.Position == FWTEST_IS_MOVING) & (DateTime.Now.Subtract(StartTime).TotalSeconds <= settings.FilterWheelTimeout)); // Wait until movement has stopped or 30 seconds have passed
                if (filterWheel.Position != FWTEST_IS_MOVING)
                    LogOK("Pre-run Check", "Filter wheel is stationary, ready to start tests");
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
            string[] filterNames;

            DateTime startTime, endTime;

            // FocusOffsets - Required - Read only
            try
            {
                LogCallToDriver("FocusOffsets Get", "About to get FocusOffsets property");
                filterOffsets = filterWheel.FocusOffsets;
                numberOfFilterOffsets = filterOffsets.Length;
                if (numberOfFilterOffsets == 0)
                    LogIssue("FocusOffsets Get", "Found no offset values in the returned array");
                else
                    LogOK("FocusOffsets Get", "Found " + numberOfFilterOffsets.ToString() + " filter offset values");

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
                    LogOK("Names Get", "Found " + numberOfFilternames.ToString() + " filter names");
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
                LogOK("Names Get", "Number of filter offsets and number of names are the same: " + numberOfFilternames.ToString());
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
                    LogOK("Position Get", $"Currently at position: {startingFilterNumber}");
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

            try
            {
                // Move to each position in turn
                for (short i = 0; i <= Convert.ToInt16(numberOfFilterOffsets - 1); i++)
                {
                    try
                    {
                        LogCallToDriver("Position Set", "About to set Position property");
                        SetAction($"Setting position {i}");
                        filterWheel.Position = i;

                        startTime = DateTime.Now;
                        LogCallToDriver("Position Set", "About to get Position property repeatedly");
                        WaitWhile($"Moving to position {i}", () => { return filterWheel.Position != i; }, 500, settings.FilterWheelTimeout);

                        if (cancellationToken.IsCancellationRequested)
                            return;

                        endTime = DateTime.Now;
                        if (filterWheel.Position == i)
                            LogOK("Position Set", "Reached position: " + i.ToString() + " in: " + endTime.Subtract(startTime).TotalSeconds.ToString("0.0") + " seconds");
                        else
                            LogIssue("Position Set", "Filter wheel did not reach specified position: " + i.ToString() + " within timeout of: " + settings.FilterWheelTimeout.ToString());
                        //WaitFor(1000); // Pause to allow filter wheel to stabilise
                        Stopwatch sw = Stopwatch.StartNew();
                        WaitWhile($"Waiting for wheel to stabilise at position {i}", () => { return sw.ElapsedMilliseconds < 1000; }, 500, 1);

                    }
                    catch (Exception ex)
                    {
                        HandleException("Position Set", MemberType.Property, Required.Mandatory, ex, "");
                    }
                }

                // Confirm that an error is correctly generated for outside range values
                try
                {
                    LogCallToDriver("Position Set", "About to set Position property");
                    filterWheel.Position = -1; // Negative position, positions should never be negative
                    LogIssue("Position Set", "Failed to generate exception when selecting filter with negative filter number");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOK("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to - 1", "Correctly rejected bad position: -1");
                }

                // Confirm that an error is correctly generated for outside range values
                try
                {
                    LogCallToDriver("Position Set", "About to set Position property");
                    filterWheel.Position = (short)numberOfFilterOffsets; // This should be 1 above the highest array element returned
                    LogIssue("Position Set", "Failed to generate exception when selecting filter outside expected range");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOK("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to " + System.Convert.ToString(numberOfFilterOffsets), "Correctly rejected bad position: " + System.Convert.ToString(numberOfFilterOffsets));
                }

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
        private void FilterWheelPerformanceTest(FilterWheelProperties p_Type, string p_Name)
        {
            int[] l_Offsets;
            string[] l_Names;
            int l_StartFilterNumber;
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
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
                        case FilterWheelProperties.FocusOffsets:
                            {
                                l_Offsets = filterWheel.FocusOffsets;
                                break;
                            }

                        case FilterWheelProperties.Names:
                            {
                                l_Names = filterWheel.Names;
                                break;
                            }

                        case FilterWheelProperties.Position:
                            {
                                l_StartFilterNumber = filterWheel.Position;
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, "FilterWheelPerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
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

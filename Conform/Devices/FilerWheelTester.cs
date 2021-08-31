using System;
using ASCOM.Standard.Interfaces;
using System.Threading;
using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.COM.DriverAccess;

namespace ConformU
{
    internal class FilterWheelTester : DeviceTesterBaseClass
    {
        const int FILTER_WHEEL_TIME_OUT = 10; // Filter wheel command timeout (seconds)
        const int FWTEST_IS_MOVING = -1;
        const int FWTEST_TIMEOUT = 30;

        private IFilterWheelV2 m_FilterWheel;
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
                    if (telescopeDevice is not null) telescopeDevice.Dispose();
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
                        LogDebug("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_FilterWheel = new AlpacaFilterWheel(settings.AlpacaConfiguration.AccessServiceType.ToString(),
                            settings.AlpacaDevice.IpAddress,
                            settings.AlpacaDevice.IpPort,
                            settings.AlpacaDevice.AlpacaDeviceNumber,
                            settings.StrictCasing,
                            settings.TraceAlpacaCalls ? logger : null);
                        LogDebug("CreateDevice", $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogDebug("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_FilterWheel = new FilterWheelFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogDebug("CreateDevice", $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                m_FilterWheel = new FilterWheel(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogDebug("CreateDevice", "Successfully created driver");
                baseClassDevice = m_FilterWheel; // Assign the driver to the base class

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
                return m_FilterWheel.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                m_FilterWheel.Connected = value;
            }
        }
        public override void PreRunCheck()
        {
            DateTime StartTime;

            // Get into a consistent state
            SetStatus("FilterWheel Pre-run Check", "Wait one second for initialisation", "");
            WaitFor(1000); // Wait for 1 second to allow any movement to start
            StartTime = DateTime.Now;
            try
            {
                LogCallToDriver("Pre-run Check", "About to get Position property repeatedly");
                do
                {
                    SetStatus("FilterWheel Pre-run Check", "Waiting for movement to stop", DateTime.Now.Subtract(StartTime).Seconds + " second(s)");
                    WaitFor(SLEEP_TIME);
                }
                while ((m_FilterWheel.Position == FWTEST_IS_MOVING) & (DateTime.Now.Subtract(StartTime).TotalSeconds <= FWTEST_TIMEOUT)); // Wait until movement has stopped or 30 seconds have passed
                if (m_FilterWheel.Position != FWTEST_IS_MOVING)
                    LogOK("Pre-run Check", "Filter wheel is stationary, ready to start tests");
            }
            catch (Exception ex)
            {
                LogInfo("Pre-run Check", "Unable to determine that the Filter wheel is stationary");
                LogError("Pre-run Check", "Exception: " + ex.ToString());
            }
            ClearStatus();
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_FilterWheel, DeviceType.FilterWheel);
        }

        public override void CheckProperties()
        {
            int[] l_Offsets;
            int l_NNames = 0, l_NOffsets = 0, l_FilterNumber, l_StartFilterNumber;
            string[] l_Names;

            DateTime l_StartTime, l_EndTime;

            // FocusOffsets - Required - Read only
            try
            {
                LogCallToDriver("FocusOffsets Get", "About to get FocusOffsets property");
                l_Offsets = m_FilterWheel.FocusOffsets;
                l_NOffsets = l_Offsets.Length;
                if (l_NOffsets == 0)
                    LogError("FocusOffsets Get", "Found no offset values in the returned array");
                else
                    LogOK("FocusOffsets Get", "Found " + l_NOffsets.ToString() + " filter offset values");

                l_FilterNumber = 0;
                foreach (var offset in l_Offsets)
                {
                    LogInfo("FocusOffsets Get", "Filter " + l_FilterNumber.ToString() + " Offset: " + offset.ToString());
                    l_FilterNumber += 1;
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
                l_Names = m_FilterWheel.Names;
                l_NNames = l_Names.Length;
                if (l_NNames == 0)
                    LogError("Names Get", "Did not find any names in the returned array");
                else
                    LogOK("Names Get", "Found " + l_NNames.ToString() + " filter names");
                l_FilterNumber = 0;
                foreach (var name in l_Names)
                {
                    if (name == null)
                        LogIssue("Names Get", "Filter " + l_FilterNumber.ToString() + " has a value of nothing");
                    else if (name == "")
                        LogIssue("Names Get", "Filter " + l_FilterNumber.ToString() + " has a value of \"\"");
                    else
                        LogInfo("Names Get", "Filter " + l_FilterNumber.ToString() + " Name: " + name);
                    l_FilterNumber += 1;
                }
            }
            catch (Exception ex)
            {
                HandleException("Names Get", MemberType.Property, Required.Mandatory, ex, "");
            }

            // Confirm number of array elements in filter names and filter offsets are the same
            if (l_NNames == l_NOffsets)
                LogOK("Names Get", "Number of filter offsets and number of names are the same: " + l_NNames.ToString());
            else
                LogError("Names Get", "Number of filter offsets and number of names are different: " + l_NOffsets.ToString() + " " + l_NNames.ToString());

            // Position - Required - Read / Write
            switch (l_NOffsets)
            {
                case object _ when l_NOffsets <= 0:
                    {
                        LogIssue("Position", "Filter position tests skipped as number of filters appears to be 0: " + l_NOffsets.ToString());
                        break;
                    }

                default:
                    {
                        try
                        {
                            LogCallToDriver("Position Get", "About to get Position property");
                            l_StartFilterNumber = m_FilterWheel.Position;
                            if ((l_StartFilterNumber < 0) | (l_StartFilterNumber >= l_NOffsets))
                                LogError("Position Get", "Illegal filter position returned: " + l_StartFilterNumber.ToString());
                            else
                            {
                                LogOK("Position Get", "Currently at position: " + l_StartFilterNumber.ToString());
                                for (short i = 0; i <= Convert.ToInt16(l_NOffsets - 1); i++)
                                {
                                    try
                                    {
                                        LogCallToDriver("Position Set", "About to set Position property");
                                        m_FilterWheel.Position = i;
                                        l_StartTime = DateTime.Now;
                                        LogCallToDriver("Position Set", "About to get Position property repeatedly");
                                        do
                                        {
                                            Thread.Sleep(100);
                                        }
                                        while ((m_FilterWheel.Position != i) & (DateTime.Now.Subtract(l_StartTime).TotalSeconds <= FILTER_WHEEL_TIME_OUT) & !cancellationToken.IsCancellationRequested);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;

                                        l_EndTime = DateTime.Now;
                                        if (m_FilterWheel.Position == i)
                                            LogOK("Position Set", "Reached position: " + i.ToString() + " in: " + l_EndTime.Subtract(l_StartTime).TotalSeconds.ToString("0.0") + " seconds");
                                        else
                                            LogError("Position Set", "Filter wheel did not reach specified position: " + i.ToString() + " within timeout of: " + FILTER_WHEEL_TIME_OUT.ToString());
                                        WaitFor(1000); // Pause to allow filter wheel to stabilise
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException("Position Set", MemberType.Property, Required.Mandatory, ex, "");
                                    }
                                }
                                try // Confirm that an error is correctly generated for outside range values
                                {
                                    LogCallToDriver("Position Set", "About to set Position property");
                                    m_FilterWheel.Position = -1; // Negative position, positions should never be negative
                                    LogError("Position Set", "Failed to generate exception when selecting filter with negative filter number");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to - 1", "Correctly rejected bad position: -1");
                                }
                                try // Confirm that an error is correctly generated for outside range values
                                {
                                    LogCallToDriver("Position Set", "About to set Position property");
                                    m_FilterWheel.Position = (short)l_NOffsets; // This should be 1 above the highest array element returned
                                    LogError("Position Set", "Failed to generate exception when selecting filter outside expected range");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOK("Position Set", MemberType.Property, Required.MustBeImplemented, ex, "setting position to " + System.Convert.ToString(l_NOffsets), "Correctly rejected bad position: " + System.Convert.ToString(l_NOffsets));
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            HandleException("Position Get", MemberType.Property, Required.Mandatory, ex, "");
                        }

                        break;
                    }
            }
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
                        case FilterWheelProperties.FocusOffsets:
                            {
                                l_Offsets = m_FilterWheel.FocusOffsets;
                                break;
                            }

                        case FilterWheelProperties.Names:
                            {
                                l_Names = m_FilterWheel.Names;
                                break;
                            }

                        case FilterWheelProperties.Position:
                            {
                                l_StartFilterNumber = m_FilterWheel.Position;
                                break;
                            }

                        default:
                            {
                                LogError(p_Name, "FilterWheelPerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
                    {
                        Status(StatusType.staStatus, l_Count + " transactions in " + l_ElapsedTime.ToString( "0") + " seconds");
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

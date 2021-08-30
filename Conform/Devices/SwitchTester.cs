using System;
using ASCOM.Standard.Interfaces;
using ASCOM.Standard.AlpacaClients;
using System.Collections;
using System.Threading;

namespace ConformU
{
    internal class SwitchTester : DeviceTesterBaseClass
    {
        private int m_PerformanceGetSwitchName, m_PerformanceGetSwitch;
        private short m_MaxSwitch;
        private bool m_CanReadMaxSwitch;
        private readonly int ExtendedSwitchNumberTestRange; // Checks for usable switches either side of the expected range
        private readonly int SWITCH_WRITE_DELAY;
        private readonly int SWITCH_READ_DELAY;

        const int NUMBER_OF_SWITCH_TEST_STATES = 10;
        const double BAD_SWITCH_VALUE = double.NaN; // Do not change this value, the Double.IsNaN method is used in various tests in the code below

        enum SwitchPropertyMethod
        {
            MaxSwitch,
            GetSwitch,
            GetSwitchName,
            SwitchType
        }

        private enum SwitchMethod
        {
            CanWrite,
            GetSwitch,
            GetSwitchDescription,
            GetSwitchName,
            GetSwitchValue,
            MaxSwitchValue,
            MinSwitchValue,
            SetSwitch,
            SetSwitchName,
            SetSwitchValue,
            SwitchStep
        }
        // Helper variables
        private ISwitchV2 m_Switch;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose

        public SwitchTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;

            m_PerformanceGetSwitchName = int.MinValue; // Initialise to silly values
            m_PerformanceGetSwitch = int.MinValue;

            SWITCH_READ_DELAY = settings.SwitchReadDelay; // Get values for the two delay parameters as set by the user or the default values if not yet set
            SWITCH_WRITE_DELAY = settings.SwitchWriteDelay;
            ExtendedSwitchNumberTestRange = settings.SwitchExtendedNumberTestRange;
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
                    if (m_Switch is not null) m_Switch.Dispose();
                    m_Switch = null;
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
                            g_ExInvalidValue1 = (int)0x80040405;
                            g_ExInvalidValue2 = (int)0x80040405;
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
                        logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_Switch = new AlpacaSwitch(settings.AlpacaConfiguration.AccessServiceType.ToString(),
                            settings.AlpacaDevice.IpAddress,
                            settings.AlpacaDevice.IpPort,
                            settings.AlpacaDevice.AlpacaDeviceNumber,
                            settings.StrictCasing,
                            settings.TraceAlpacaCalls ? logger : null);
                        logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_Switch = new SwitchFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                logger.LogMessage("CreateDevice", MessageLevel.Debug, $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                m_Switch = new ASCOM.Standard.COM.DriverAccess.Switch(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogDebug("CreateDevice", "Successfully created driver");
                baseClassDevice = m_Switch; // Assign the driver to the base class

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
                return m_Switch.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to get Connected property");
                m_Switch.Connected = value;
            }
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_Switch, DeviceType.Switch);
        }

        public override void CheckProperties()
        {
            // MaxSwitch - Mandatory
            switch (g_InterfaceVersion)
            {
                case 1:
                case 2 // Original Platform 5 switch interface and ISwitchV2 have the same property
               :
                    {
                        m_MaxSwitch = SwitchPropertyTestShort(SwitchPropertyMethod.MaxSwitch, "MaxSwitch", 1, short.MaxValue); if (cancellationToken.IsCancellationRequested)
                            return;
                        break;
                    }

                default:
                    {
                        LogError("Switches", "Unknown switch interface version: " + g_InterfaceVersion);
                        break;
                    }
            }
        }

        public override void CheckMethods()
        {
            short i;
            bool l_GetSwitch = false, l_GetSwitchOriginal = false, l_NewSwitchState, l_GetSwitchOK, l_SetSwitchOK, l_GetSwitchValueOK, l_SetSwitchValueMinOK, l_SetSwitchValueMaxOK,  l_SwitchCanWrite;
            Exception l_GetSwitchException,  l_GetSwitchValueException;
            double l_GetSwitchValue, l_GetSwitchValueOriginal=0.0, l_SwitchMinimum, l_SwitchMaximum, l_SwitchStep, l_SwitchRange;
            string l_SwitchName, l_SwitchDescription;

            switch (g_InterfaceVersion)
            {
                case 1 // Platform 5 interface v1
               :
                    {
                        if (m_CanReadMaxSwitch)
                        {
                            Status(StatusType.staTest, "Read/Write Switches");
                            Status(StatusType.staAction, "Testing switch");
                            // Find valid GetSwitch values
                            for (i = 0; i <= System.Convert.ToInt16(m_MaxSwitch - 1); i++)
                            {
                                Status(StatusType.staStatus, i.ToString());
                                l_GetSwitchOK = false;
                                l_SetSwitchOK = false;
                                try // Read switch state
                                {
                                    LogCallToDriver("GetSwitch " + i, $"About to call GetSwitch({i}) method");
                                    l_GetSwitchOriginal = m_Switch.GetSwitch(i);
                                    LogOK("GetSwitch " + i, "Found switch, state: " + l_GetSwitchOriginal.ToString());
                                    l_GetSwitchOK = true;
                                    if (i > m_MaxSwitch)
                                        LogIssue("GetSwitch " + i, "Usable switch found above MaxSwitch!");
                                    else
                                        m_PerformanceGetSwitch = i;// Save last good switch number for performance test
                                }
                                catch (Exception ex)
                                {
                                    if (IsNotSetException(ex))
                                        LogInfo("GetSwitch " + i, "Switch read is not implemented");
                                    else
                                    {
                                        LogInfo("GetSwitch " + i, "Unable to read switch: " + ex.Message);
                                        LogDebug("GetSwitch " + i, "Exception: " + ex.ToString());
                                    }
                                }

                                try // Now try to write the value
                                {
                                    LogCallToDriver("SetSwitch " + i, $"About to call SetSwitch({i})");
                                    m_Switch.SetSwitch(i, !l_GetSwitchOriginal); // Swap the switch state
                                    l_SetSwitchOK = true;
                                    if (l_GetSwitchOK)
                                    {
                                        LogCallToDriver("SetSwitch", $"About to call GetSwitch({i}) method");
                                        l_NewSwitchState = m_Switch.GetSwitch(i); // Read the new switch state to confirm that value did change
                                        if (l_NewSwitchState == !l_GetSwitchOriginal)
                                        {
                                            LogOK("SetSwitch " + i, "Switch correctly changed state");
                                            LogCallToDriver("SetSwitch", "About to call SetSwitch method");
                                            m_Switch.SetSwitch(i, l_GetSwitchOriginal); // Now put switch back to original state
                                        }
                                        else
                                            LogIssue("SetSwitch " + i, "Switch did not change state, currently it is " + l_NewSwitchState.ToString());
                                    }
                                    else
                                        LogInfo("SetSwitch " + i, "You have a write only switch!");
                                    if (i > m_MaxSwitch)
                                        LogIssue("SetSwitch " + i, "Usable switch found above MaxSwitch!");
                                }
                                catch (Exception ex)
                                {
                                    if (IsNotSetException(ex))
                                        LogInfo("SetSwitch " + i, "Switch write is not implemented");
                                    else
                                    {
                                        LogInfo("SetSwitch " + i, "Unable to write to switch: " + ex.Message);
                                        LogDebug("SetSwitch " + i, "Exception: " + ex.ToString());
                                    }
                                }

                                try
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("GetSwitchName " + i, string.Format("About to get switch name {0}", i));
                                    l_SwitchName = m_Switch.GetSwitchName(i);
                                    if (l_GetSwitchOK | l_SetSwitchOK)
                                    {
                                        if (l_SwitchName == "")
                                            LogInfo("GetSwitchName " + i, "Switch name returns null string");
                                        else
                                        {
                                            m_PerformanceGetSwitchName = i; // Save last good name index for performance test
                                            LogOK("GetSwitchName " + i, "Found switch, name:  " + l_SwitchName);
                                        }
                                    }
                                    else if (l_SwitchName == "")
                                        LogError("GetSwitchName " + i, "Switch name returns null string but switch can neither read nor write!");
                                    else
                                        LogError("GetSwitchName " + i, "Found switch, name:  " + l_SwitchName + " which can neither read nor write!");
                                }
                                catch (Exception ex)
                                {
                                    LogDebug("GetSwitchName " + i, "Exception: " + ex.ToString());
                                }
                            }
                            Status(StatusType.staTest, "");
                            Status(StatusType.staAction, "");
                            Status(StatusType.staStatus, "");
                        }
                        else
                            LogError("SwitchCheckMethods", "Skipping further tests as there is no valid value for MaxSwitch");
                        break;
                    }

                case 2 // Platform 6 interface version 2
         :
                    {
                        if (m_CanReadMaxSwitch)
                        {
                            Status(StatusType.staTest, "Read/Write Switches");

                            CheckInaccessibleOutOfRange(SwitchMethod.CanWrite);
                            CheckInaccessibleOutOfRange(SwitchMethod.GetSwitch);
                            CheckInaccessibleOutOfRange(SwitchMethod.GetSwitchDescription);
                            CheckInaccessibleOutOfRange(SwitchMethod.GetSwitchName);
                            CheckInaccessibleOutOfRange(SwitchMethod.GetSwitchValue);
                            CheckInaccessibleOutOfRange(SwitchMethod.MaxSwitchValue);
                            CheckInaccessibleOutOfRange(SwitchMethod.MinSwitchValue);
                            CheckInaccessibleOutOfRange(SwitchMethod.SetSwitch);
                            CheckInaccessibleOutOfRange(SwitchMethod.SetSwitchValue);
                            CheckInaccessibleOutOfRange(SwitchMethod.SwitchStep);

                            // Find valid GetSwitch values
                            Status(StatusType.staAction, "Testing switch");
                            LogDebug("GetSwitchName ", string.Format("Extended switch number test range: {0} - {1}", -ExtendedSwitchNumberTestRange, m_MaxSwitch + ExtendedSwitchNumberTestRange - 1));
                            for (i = (short)-ExtendedSwitchNumberTestRange; i <= Convert.ToInt16(m_MaxSwitch + ExtendedSwitchNumberTestRange - 1); i++)
                            {
                                Status(StatusType.staStatus, i.ToString());

                                // Initialise status variables
                                l_SetSwitchValueMinOK = false;
                                l_SetSwitchValueMaxOK = false;
                                l_SwitchStep = BAD_SWITCH_VALUE;

                                try // Read switch name to determine whether this is a valid switch
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogTestAndMessage("GetSwitchName", string.Format("About to get switch {0} name", i));
                                    l_SwitchName = m_Switch.GetSwitchName(i);
                                    LogOK("GetSwitchName ", "Found switch " + i);

                                    // Test that the switch number is in the valid range of 0..MaxSwitch-1
                                    if (i > (m_MaxSwitch - 1))
                                        LogIssue("GetSwitchName ", "Usable switch found above MaxSwitch - 1!");
                                    else if (i < 0)
                                        LogIssue("GetSwitchName ", "Usable switch found below 0!");
                                    else
                                        m_PerformanceGetSwitch = i;// Save last good switch number for performance test

                                    LogOK("GetSwitchName ", "  Name: " + l_SwitchName);

                                    try // Read switch description
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage("GetSwitchDescription", string.Format("  About to get switch {0} description", i));
                                        l_SwitchDescription = m_Switch.GetSwitchDescription(i);
                                        LogOK("GetSwitchDescription ", "  Description: " + l_SwitchDescription);
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError("GetSwitchDescription ", "Mandatory parameter threw an exception: " + ex.Message);
                                    }

                                    try // Read switch minimum value
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage("MinSwitchValue", string.Format("  About to get switch {0} minimum value", i));
                                        l_SwitchMinimum = m_Switch.MinSwitchValue(i);
                                        LogOK("MinSwitchValue ", "  Minimum: " + l_SwitchMinimum.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError("MinSwitchValue ", "Mandatory parameter threw an exception: " + ex.Message);
                                        l_SwitchMinimum = BAD_SWITCH_VALUE;
                                    }

                                    try // Read switch maximum value
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage("MaxSwitchValue", string.Format("  About to get switch {0} maximum value", i));
                                        l_SwitchMaximum = m_Switch.MaxSwitchValue(i);

                                        if (IsGoodValue(l_SwitchMinimum))
                                        {
                                            // Maximum value must be > Minimum value
                                            if (l_SwitchMaximum > l_SwitchMinimum)
                                            {
                                                LogOK("MaxSwitchValue ", "  Maximum: " + l_SwitchMaximum.ToString());
                                                l_SwitchRange = l_SwitchMaximum - l_SwitchMinimum; // Calculate the range of values the switch might take
                                            }
                                            else
                                            {
                                                LogInfo("MaxSwitchValue ", "  Maximum: " + l_SwitchMaximum.ToString());
                                                LogIssue("MaxSwitchValue ", "MaxSwitchValue is less or equal to MinSwitchValue, it must be greater!");
                                                l_SwitchRange = BAD_SWITCH_VALUE; // Special value because the maximum or minimum values are bad
                                            }
                                        }
                                        else
                                        {
                                            LogInfo("MaxSwitchValue ", "  Maximum: " + l_SwitchMaximum.ToString());
                                            l_SwitchRange = BAD_SWITCH_VALUE;
                                            LogInfo("MaxSwitchValue ", "  Test that switch Maximum is greater than Minimum skipped because of an error reading the Minimum value.");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError("MaxSwitchValue ", "Mandatory parameter threw an exception: " + ex.Message);
                                        l_SwitchMaximum = BAD_SWITCH_VALUE;
                                        l_SwitchRange = BAD_SWITCH_VALUE;
                                    }

                                    try // Read switch step value
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage("SwitchStep", string.Format("  About to get switch {0} step size", i));
                                        l_SwitchStep = m_Switch.SwitchStep(i);
                                        LogOK("SwitchStep ", "  Step size: " + l_SwitchStep.ToString());

                                        // Step must be greater than 0
                                        if (l_SwitchStep > 0.0)
                                        {
                                            LogOK("SwitchStep ", "  Step size is greater than zero");

                                            // Step must be less than or equal to the range of possible values
                                            if (l_SwitchStep <= l_SwitchRange)
                                            {
                                                LogOK("SwitchStep ", "  Step size is less than the range of possible values");

                                                // Now check that the switch range is an integer multiple of the step size
                                                // Doubles are converted to the Decimal type (which has higher precision) in order to avoid unexpected outcomes from Mod due to rounding errors
                                                switch (Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)))
                                                {
                                                    case 0M:
                                                        {
                                                            LogOK("SwitchStep ", "  The switch range is an integer multiple of the step size.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) <= Convert.ToDecimal(l_SwitchStep / 100):
                                                        {
                                                            LogInfo("SwitchStep ", "  The switch range is within 1% of being an integer multiple of the step size.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) <= Convert.ToDecimal(l_SwitchStep / 10):
                                                        {
                                                            LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 10%.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) <= Convert.ToDecimal(l_SwitchStep / 5):
                                                        {
                                                            LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 20%.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep)) <= Convert.ToDecimal(l_SwitchStep / 2):
                                                        {
                                                            LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 50%.");
                                                            break;
                                                        }

                                                    default:
                                                        {
                                                            LogIssue("SwitchStep ", "The switch range must be an integer multiple of the step size. Remainder`: " + decimal.Subtract(Convert.ToDecimal(l_SwitchMaximum), Convert.ToDecimal(l_SwitchMinimum)) % Convert.ToDecimal(l_SwitchStep));
                                                            break;
                                                        }
                                                }
                                            }
                                            else
                                            {
                                                LogIssue("SwitchStep ", "Step size must be less than the range of possible values (MaxSwitchValue - MinSwitchValue");
                                                l_SwitchStep = BAD_SWITCH_VALUE;
                                            }
                                        }
                                        else
                                        {
                                            LogIssue("SwitchStep ", "Step size must be greater than zero");
                                            l_SwitchStep = BAD_SWITCH_VALUE;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError("SwitchStep ", "Mandatory parameter threw an exception: " + ex.Message);
                                    }

                                    LogDebug("SwitchMinimum ", l_SwitchMinimum.ToString());
                                    LogDebug("SwitchMaximum ", l_SwitchMaximum.ToString());
                                    LogDebug("SwitchStep ", l_SwitchStep.ToString());
                                    LogDebug("SwitchRange ", l_SwitchRange.ToString());

                                    try // Read CanWrite 
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage("CanWrite", string.Format("  About to get switch {0} CanWrite status", i));
                                        l_SwitchCanWrite = m_Switch.CanWrite(i);
                                        LogOK("CanWrite ", "  CanWrite: " + l_SwitchCanWrite);
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError("CanWrite ", "Mandatory parameter threw an exception: " + ex.Message);
                                        LogInfo("CanWrite ", "Assuming that CanWrite is false");
                                        l_SwitchCanWrite = false;
                                    }// Initialise to a default state

                                    // Access the Get Methods and record the outcomes
                                    try
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage("GetSwitch", string.Format("  About to call GetSwitch({0}) method", i));
                                        l_GetSwitchOriginal = m_Switch.GetSwitch(i);
                                        l_GetSwitchOK = true;
                                        LogOK("GetSwitch ", "  " + l_GetSwitchOriginal.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        l_GetSwitchException = ex;
                                        LogDebug("GetSwitch ", "Exception: " + ex.ToString());
                                        l_GetSwitchOK = false;
                                    }

                                    try
                                    {
                                        if (settings.DisplayMethodCalls)
                                            LogTestAndMessage("GetSwitchValue", string.Format("  About to call GetSwitchValue({0}) method", i));
                                        l_GetSwitchValueOriginal = m_Switch.GetSwitchValue(i);
                                        l_GetSwitchValueOK = true;
                                        LogOK("GetSwitchValue ", "  " + l_GetSwitchValueOriginal);
                                    }
                                    catch (Exception ex)
                                    {
                                        l_GetSwitchValueException = ex;
                                        LogDebug("GetSwitchValue ", "Exception: " + ex.ToString());
                                        l_GetSwitchValueOK = false;
                                    }

                                    // Now try to write to see which of these methods are available
                                    if (settings.SwitchEnableSet)
                                    {

                                        // Try to set the two boolean values through SetSwitch
                                        try
                                        {

                                            // Try SetSwitch(False)
                                            Status(StatusType.staStatus, "Setting SetSwitch - False");
                                            if (settings.DisplayMethodCalls)
                                                LogTestAndMessage("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, false));
                                            m_Switch.SetSwitch(i, false); // Set switch false
                                            WaitFor(SWITCH_READ_DELAY);

                                            // Check GetSwitch
                                            if (l_GetSwitchOK)
                                            {
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitch", string.Format("  About to call GetSwitch({0}) method", i));
                                                if (m_Switch.GetSwitch(i) == false)
                                                    LogOK("SetSwitch ", "  GetSwitch returned False after SetSwitch(False)");
                                                else
                                                    LogIssue("SetSwitch ", "  GetSwitch returned True after SetSwitch(False)");
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                            // Check GetSwitchValue returns the switch minimum value
                                            if (l_GetSwitchValueOK & IsGoodValue(l_SwitchMinimum))
                                            {
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitch", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                l_GetSwitchValue = m_Switch.GetSwitchValue(i);
                                                switch (l_GetSwitchValue)
                                                {
                                                    case object _ when l_GetSwitchValue == l_SwitchMinimum:
                                                        {
                                                            LogOK("SetSwitch ", "  GetSwitchValue returned MINIMUM_VALUE after SetSwitch(False)");
                                                            break;
                                                        }

                                                    case object _ when l_SwitchMinimum * 0.99 <= l_GetSwitchValue && l_GetSwitchValue <= l_SwitchMinimum * 1.01:
                                                        {
                                                            LogOK("SetSwitch ", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitch(False)");
                                                            break;
                                                        }

                                                    default:
                                                        {
                                                            LogIssue("SetSwitch ", "  GetSwitchValue did not return MINIMUM_VALUE after SetSwitch(False): " + l_GetSwitchValue);
                                                            break;
                                                        }
                                                }
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMinimum methods");
                                            WaitFor(SWITCH_WRITE_DELAY);

                                            // Try SetSwitch(True)
                                            Status(StatusType.staStatus, "Setting SetSwitch - True");
                                            if (settings.DisplayMethodCalls)
                                                LogTestAndMessage("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, true));
                                            m_Switch.SetSwitch(i, true); // Set switch true
                                            WaitFor(SWITCH_READ_DELAY);

                                            // Check GetSwitch
                                            if (l_GetSwitchOK)
                                            {
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitch", string.Format("  About to call GetSwitch({0}) method", i));
                                                if (m_Switch.GetSwitch(i) == true)
                                                    LogOK("SetSwitch ", "  GetSwitch read True after SetSwitch(True)");
                                                else
                                                    LogIssue("SetSwitch ", "  GetSwitch read False after SetSwitch(True)");
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                            // Check GetSwitchValue returns the switch maximum value
                                            if (l_GetSwitchValueOK & IsGoodValue(l_SwitchMaximum))
                                            {
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitch", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                l_GetSwitchValue = m_Switch.GetSwitchValue(i);
                                                switch (l_GetSwitchValue)
                                                {
                                                    case object _ when l_GetSwitchValue == l_SwitchMaximum:
                                                        {
                                                            LogOK("SetSwitch ", "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitch(True)");
                                                            break;
                                                        }

                                                    case object _ when l_SwitchMaximum * 0.99 <= l_SwitchMaximum && l_SwitchMaximum <= l_SwitchMaximum * 1.01:
                                                        {
                                                            LogOK("SetSwitch ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitch(True)");
                                                            break;
                                                        }

                                                    default:
                                                        {
                                                            LogIssue("SetSwitch ", "  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitch(True): " + l_GetSwitchValue);
                                                            break;
                                                        }
                                                }
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMaximum methods");
                                            WaitFor(SWITCH_WRITE_DELAY);

                                            // Return to original state if possible,otherwise set to false
                                            if (l_GetSwitchOK)
                                            {
                                                Status(StatusType.staStatus, "Returning boolean switch to its original value");
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, l_GetSwitch));
                                                m_Switch.SetSwitch(i, l_GetSwitch); // Return to the original state
                                                WaitFor(SWITCH_WRITE_DELAY);
                                            }
                                            else
                                            {
                                                Status(StatusType.staStatus, "Setting boolean switch to False");
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, false));
                                                m_Switch.SetSwitch(i, false); // Set to false
                                                WaitFor(SWITCH_WRITE_DELAY);
                                            }

                                            l_SetSwitchOK = true;
                                            LogDebug("SetSwitch ", "Set value OK");
                                        }
                                        catch (Exception ex)
                                        {
                                            if (l_SwitchCanWrite)
                                            {
                                                LogError("SetSwitch ", "Exception: " + ex.Message);
                                                LogDebug("SetSwitch ", "Exception: " + ex.ToString());
                                            }
                                            else if (IsMethodNotImplementedException(ex))
                                                LogOK("SetSwitch ", "  CanWrite is False and MethodNotImplementedException was thrown");
                                            else
                                            {
                                                LogError("SetSwitch ", "Exception: " + ex.Message);
                                                LogDebug("SetSwitch ", "Exception: " + ex.ToString());
                                            }
                                        }

                                        // Try to set the minimum and maximum values through SetSwitchValue
                                        try
                                        {
                                            if (IsGoodValue(l_SwitchMinimum))
                                            {
                                                Status(StatusType.staStatus, "Setting SetSwitchValue - MINIMUM_VALUE");
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the minimum permissible value", i, l_SwitchMinimum));
                                                m_Switch.SetSwitchValue(i, l_SwitchMinimum); // Set switch to minimum
                                                WaitFor(SWITCH_READ_DELAY);

                                                // Check GetSwitch
                                                if (l_GetSwitchOK)
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SetSwitchValue", string.Format("  About to call GetSwitch({0}) method", i));
                                                    if (m_Switch.GetSwitch(i) == false)
                                                        LogOK("SetSwitchValue", "  GetSwitch returned False after SetSwitchValue(MINIMUM_VALUE)");
                                                    else
                                                        LogIssue("SetSwitchValue", "  GetSwitch returned True after SetSwitchValue(MINIMUM_VALUE)");
                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                                // Check GetSwitchValue returns the switch minimum value
                                                if (l_GetSwitchValueOK)
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SetSwitchValue", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                    l_GetSwitchValue = m_Switch.GetSwitchValue(i);
                                                    switch (l_GetSwitchValue)
                                                    {
                                                        case object _ when l_GetSwitchValue == l_SwitchMinimum:
                                                            {
                                                                LogOK("SetSwitchValue", "  GetSwitchValue returned MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)");
                                                                l_SetSwitchValueMinOK = true;
                                                                break;
                                                            }

                                                        case object _ when l_SwitchMinimum * 0.99 <= l_GetSwitchValue && l_GetSwitchValue <= l_SwitchMinimum * 1.01:
                                                            {
                                                                LogOK("SetSwitchValue", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)");
                                                                l_SetSwitchValueMinOK = true;
                                                                break;
                                                            }

                                                        default:
                                                            {
                                                                LogIssue("SetSwitchValue", "  GetSwitchValue did not return MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE): " + l_GetSwitchValue);
                                                                break;
                                                            }
                                                    }
                                                    if (l_GetSwitchValue == l_SwitchMinimum)
                                                    {
                                                    }
                                                    else
                                                    {
                                                    }
                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");
                                                WaitFor(SWITCH_WRITE_DELAY);

                                                // Now try a value below minimum
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an invalid low value", i, l_SwitchMinimum - 1.0));
                                                    m_Switch.SetSwitchValue(i, l_SwitchMinimum - 1.0);
                                                    LogIssue("SetSwitchValue", "Switch did not throw an exception when a value below SwitchMinimum was set: " + (l_SwitchMinimum - 1.0).ToString());
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleInvalidValueExceptionAsOK("SetSwitchValue", MemberType.Method, Required.Mandatory, ex, "when setting a value below SwitchMinimum - " + (l_SwitchMinimum - 1.0).ToString(), "  Switch threw an InvalidOperationException when a value below SwitchMinimum was set: " + (l_SwitchMinimum - 1.0).ToString());
                                                }
                                                WaitFor(SWITCH_WRITE_DELAY);
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                            // Try SetSwitchValue(MAXIMUM_VALUE)
                                            if (IsGoodValue(l_SwitchMaximum))
                                            {
                                                Status(StatusType.staStatus, "Setting SetSwitchValue - MAXIMUM_VALUE");
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the maximum permissible value", i, l_SwitchMaximum));
                                                m_Switch.SetSwitchValue(i, l_SwitchMaximum); // Set switch to minimum
                                                WaitFor(SWITCH_READ_DELAY);

                                                // Check GetSwitch
                                                if (l_GetSwitchOK)
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SetSwitchValue", string.Format("  About to call GetSwitch({0}) method", i));
                                                    if (m_Switch.GetSwitch(i) == true)
                                                        LogOK("SetSwitchValue ", "  GetSwitch returned True after SetSwitchValue(MAXIMUM_VALUE)");
                                                    else
                                                        LogIssue("SetSwitchValue ", "  GetSwitch returned False after SetSwitchValue(MAXIMUM_VALUE)");
                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                                // Check GetSwitchValue returns the switch maximum value
                                                if (l_GetSwitchValueOK)
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SetSwitchValue", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                    l_GetSwitchValue = m_Switch.GetSwitchValue(i);
                                                    switch (l_GetSwitchValue)
                                                    {
                                                        case object _ when l_GetSwitchValue==l_SwitchMaximum:
                                                            {
                                                                LogOK("SetSwitchValue ", "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)");
                                                                l_SetSwitchValueMaxOK = true;
                                                                break;
                                                            }

                                                        case object _ when l_SwitchMaximum * 0.99 <= l_GetSwitchValue && l_GetSwitchValue <= l_SwitchMaximum * 1.01:
                                                            {
                                                                LogOK("SetSwitchValue ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)");
                                                                break;
                                                            }

                                                        default:
                                                            {
                                                                LogIssue("SetSwitchValue ", "  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE): " + l_GetSwitchValue);
                                                                break;
                                                            }
                                                    }
                                                    if (l_GetSwitchValue == l_SwitchMaximum)
                                                    {
                                                    }
                                                    else
                                                    {
                                                    }
                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");
                                                WaitFor(SWITCH_WRITE_DELAY);

                                                // Now try a value above maximum
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an invalid high value", i, l_SwitchMaximum + 1.0));
                                                    m_Switch.SetSwitchValue(i, l_SwitchMaximum + 1.0);
                                                    LogIssue("SetSwitchValue", "Switch did not throw an exception when a value above SwitchMaximum was set: " + l_SwitchMaximum + 1.0);
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleInvalidValueExceptionAsOK("SetSwitchValue", MemberType.Method, Required.Mandatory, ex, "when setting a value above SwitchMaximum was set: " + l_SwitchMaximum + 1.0, "  Switch threw an InvalidOperationException when a value above SwitchMaximum was set: " + l_SwitchMaximum + 1.0);
                                                }
                                                WaitFor(SWITCH_WRITE_DELAY);
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                            // Test some positions of the multi-state switch between the minimum and maximum values
                                            if (l_GetSwitchValueOK & l_SetSwitchValueMinOK & l_SetSwitchValueMaxOK & IsGoodValue(l_SwitchRange) & IsGoodValue(l_SwitchStep))
                                            {
                                                TestSetSwitchValue(i, 0.0, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep);
                                                TestSetSwitchValue(i, 0.25, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep);
                                                TestSetSwitchValue(i, 0.5, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep);
                                                TestSetSwitchValue(i, 0.75, l_SwitchMinimum, l_SwitchMaximum, l_SwitchRange, l_SwitchStep);
                                            }
                                            else
                                            {
                                                LogInfo("SetSwitchValue ", "Skipping multi state tests because of earlier errors");
                                                LogDebug("GetSwitchValueOK ", l_GetSwitchValueOK.ToString());
                                                LogDebug("SetSwitchValueMinOK ", l_SetSwitchValueMinOK.ToString());
                                                LogDebug("SetSwitchValueMaxOK ", l_SetSwitchValueMaxOK.ToString());
                                                LogDebug("SwitchRange ", l_SwitchRange.ToString());
                                                LogDebug("SwitchStep ", l_SwitchStep.ToString());
                                            }

                                            // Return to original state if possible,otherwise set to false
                                            if (l_GetSwitchValueOK)
                                            {
                                                Status(StatusType.staStatus, "Returning switch to its original value");
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to restore pre-test value", i, l_GetSwitchValueOriginal));
                                                m_Switch.SetSwitchValue(i, l_GetSwitchValueOriginal); // Return to the original state
                                                LogOK("SetSwitchValue ", "  Switch has been reset to its original state");
                                                WaitFor(SWITCH_WRITE_DELAY);
                                            }
                                            else if (IsGoodValue(l_SwitchMinimum) & IsGoodValue(l_SwitchMaximum))
                                            {
                                                Status(StatusType.staStatus, "Setting switch to half its minimum to maximum range");
                                                if (settings.DisplayMethodCalls)
                                                    LogTestAndMessage("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the value to its mid-point", i, (l_SwitchMaximum - l_SwitchMinimum) / 2.0));
                                                m_Switch.SetSwitchValue(i, (l_SwitchMaximum - l_SwitchMinimum) / 2.0); // Return to the half way state
                                                LogOK("SetSwitchValue ", "  Switch has been reset to half its range");
                                                WaitFor(SWITCH_WRITE_DELAY);
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "Switch can not be returned to its default state because of issues with GetSwitchValue, GetSwitchMinimum or GetSwitchMaximum");
                                        }
                                        catch (Exception ex)
                                        {
                                            if (l_SwitchCanWrite)
                                            {
                                                LogError("SetSwitchValue ", "Exception: " + ex.Message);
                                                LogDebug("SetSwitchValue ", "Exception: " + ex.ToString());
                                            }
                                            else if (IsMethodNotImplementedException(ex))
                                                LogOK("SetSwitchValue ", "  CanWrite is False and MethodNotImplementedException was thrown");
                                            else
                                            {
                                                LogError("SetSwitchValue ", "Exception: " + ex.Message);
                                                LogDebug("SetSwitchValue ", "Exception: " + ex.ToString());
                                            }
                                        }
                                    }
                                    else
                                        LogInfo("SetSwitch", "  All write tests have been skipped because the \"Set Switches\" checkbox is unchecked");

                                    LogNewLine();
                                }
                                catch (Exception ex)
                                {
                                    if ((i >= 0) & (i < m_MaxSwitch))
                                        LogError("GetSwitchName ", "Mandatory method GetSwitchName threw an exception: " + ex.ToString());
                                    else
                                        LogDebug("GetSwitchName ", "Exception: " + ex.ToString());
                                }

                                if (cancellationToken.IsCancellationRequested)
                                    return;
                            }
                            Status(StatusType.staTest, "");
                            Status(StatusType.staAction, "");
                            Status(StatusType.staStatus, "");
                        }
                        else
                            LogInfo("SwitchCheckMethods", "Skipping further tests as there is no valid value for MaxSwitch");
                        break;
                    }
            }
        }

        public override void CheckPerformance()
        {
            // MaxSwitch
            if (m_CanReadMaxSwitch)
                SwitchPerformanceTest(SwitchPropertyMethod.MaxSwitch, "MaxSwitch");
            else
                LogInfo("MaxSwitch", "Test skipped as unable to read value");
            // GetSwitch
            if (System.Convert.ToBoolean(m_PerformanceGetSwitch))
                SwitchPerformanceTest(SwitchPropertyMethod.GetSwitch, "GetSwitch");
            else
                LogInfo("GetSwitch", "Test skipped as unable to read value");
            // GetSwitchName
            if (System.Convert.ToBoolean(m_PerformanceGetSwitchName))
                SwitchPerformanceTest(SwitchPropertyMethod.GetSwitchName, "GetSwitchName");
            else
                LogInfo("GetSwitchName", "Test skipped as unable to read value");
            Status(StatusType.staTest, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
        }

        private short SwitchPropertyTestShort(SwitchPropertyMethod p_Type, string p_Name, short p_Min, short p_Max)
        {
            short returnValue = 0;

            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage(p_Name, string.Format("About to get property {0}", p_Name));
                returnValue = 0;
                switch (p_Type)
                {
                    case SwitchPropertyMethod.MaxSwitch:
                        {
                            m_CanReadMaxSwitch = false;
                            returnValue = m_Switch.MaxSwitch;
                            break;
                        }

                    default:
                        {
                            LogError(p_Name, "SwitchPropertyTestInteger: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogError(p_Name, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogError(p_Name, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            m_CanReadMaxSwitch = true; // A valid value has been found
                            LogOK(p_Name, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
            }
            return returnValue;
        }

        private void SwitchPerformanceTest(SwitchPropertyMethod p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
            string l_String;
            bool l_Boolean;
            double l_Rate;
            short l_Short;
            Status(StatusType.staTest, "Performance Testing");
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
                        case SwitchPropertyMethod.MaxSwitch:
                            {
                                l_Short = m_Switch.MaxSwitch;
                                break;
                            }

                        case SwitchPropertyMethod.GetSwitch:
                            {
                                l_Boolean = m_Switch.GetSwitch((short)m_PerformanceGetSwitch);
                                break;
                            }

                        case SwitchPropertyMethod.GetSwitchName:
                            {
                                l_String = m_Switch.GetSwitchName((short)m_PerformanceGetSwitchName);
                                break;
                            }

                        default:
                            {
                                LogError(p_Name, "SwitchPerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
                    {
                        Status(StatusType.staStatus, l_Count + " transactions in " + l_ElapsedTime.ToString("0") + " seconds");
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

        /// <summary>
        ///     ''' Tests whether a double has a good value or the NaN bad value indicator
        ///     ''' </summary>
        ///     ''' <param name="value">Variable to be tested</param>
        ///     ''' <returns>Returns True if the variable has a good value, otherwise returns False</returns>
        ///     ''' <remarks></remarks>
        private static bool IsGoodValue(double value)
        {
            return !double.IsNaN(value);
        }

        /// <summary>
        ///     ''' Test that GetSwitchValue returns the same value as was set with SetSwitchValue
        ///     ''' </summary>
        ///     ''' <param name="i">Switch number</param>
        ///     ''' <param name="Offset">Offset size as a percentage of switch step</param>
        ///     ''' <param name="SwitchMinimum">Switch minimum value</param>
        ///     ''' <param name="SwitchRange">Switch value range</param>
        ///     ''' <param name="SwitchStep">Size of each switch step</param>
        ///     ''' <remarks></remarks>
        private void TestSetSwitchValue(int i, double Offset, double SwitchMinimum, double SwitchMaximum, double SwitchRange, double SwitchStep)
        {
            double l_MultiStateStepSize, TestValue2, l_SwitchValue;
            MessageLevel msgLevel0, msgLevel1, msgLevel2;
            int l_MultiStateNumberOfSteps;

            // Test the switch at the calculated positions
            try
            {
                if ((((SwitchMaximum - SwitchMinimum) / SwitchStep) + 1) >= NUMBER_OF_SWITCH_TEST_STATES)
                {
                    l_MultiStateStepSize = SwitchRange / NUMBER_OF_SWITCH_TEST_STATES;
                    l_MultiStateNumberOfSteps = NUMBER_OF_SWITCH_TEST_STATES;
                }
                else
                {
                    l_MultiStateStepSize = SwitchStep; // Use the specified switch step size
                    l_MultiStateNumberOfSteps =(int) Math.Floor(SwitchRange / SwitchStep);
                }
                LogDebug("MultiStateStepSize", l_MultiStateStepSize.ToString());
                LogDebug("MultiStateNumberOfSteps", l_MultiStateNumberOfSteps.ToString());

                if (Offset == 0.0)
                {
                    msgLevel0 = MessageLevel.OK;
                    msgLevel1 = MessageLevel.Info;
                    msgLevel2 = MessageLevel.Issue;
                }
                else
                {
                    msgLevel0 = MessageLevel.Info;
                    msgLevel1 = MessageLevel.Info;
                    msgLevel2 = MessageLevel.Info;
                }

                LogInfo("SetSwitchValue", "  Testing with steps that are " + Offset.ToString("P0") + " offset from integer SwitchStep values");

                for (double TestValue = SwitchMinimum; TestValue <= SwitchMinimum + l_MultiStateStepSize * l_MultiStateNumberOfSteps; TestValue += l_MultiStateStepSize)
                {

                    // Round the test value to the nearest lowest switch step
                    if (TestValue == SwitchMinimum)
                        TestValue2 = SwitchMinimum + Offset * SwitchStep;
                    else
                        TestValue2 = (Math.Round((TestValue - SwitchMinimum) / SwitchStep) * SwitchStep) + SwitchMinimum + Offset * SwitchStep;

                    if (TestValue2 <= SwitchMaximum)
                    {
                        Status(StatusType.staStatus, "Setting multi-state switch - " + TestValue2);
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an intermediate value", i, TestValue2));
                        m_Switch.SetSwitchValue((short)i, TestValue2); // Set the required switch value
                        WaitFor(SWITCH_READ_DELAY);
                        if (settings.DisplayMethodCalls)
                            LogTestAndMessage("SetSwitchValue", string.Format("  About to call GetSwitchValue({0})", i));
                        l_SwitchValue = m_Switch.GetSwitchValue((short)i); // Read back the switch value 

                        switch (Math.Abs(l_SwitchValue - TestValue2))
                        {
                            case 0.0:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel0, "  Set and read match: " + TestValue2);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.01:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel1, "   Set/Read differ by up to 1% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.1:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "   Set/Read differ by 1-10% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.2:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 10-20% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.3:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 20-30% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.4:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 30-40% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.5:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 40-50% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.6:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 50-60% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.7:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 60-70% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.8:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 70-80% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 0.9:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 80-90% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(l_SwitchValue - TestValue2) && Math.Abs(l_SwitchValue - TestValue2) <= SwitchStep * 1.0:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by 90-100% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }

                            default:
                                {
                                    LogMsg("SetSwitchValue " + "Offset: " + Offset.ToString("P0").PadLeft(4), msgLevel2, "  Set/Read differ by >100% of SwitchStep. Set: " + TestValue2 + ", Read: " + l_SwitchValue);
                                    break;
                                }
                        }

                        WaitFor(SWITCH_WRITE_DELAY);
                    }
                    if (cancellationToken.IsCancellationRequested)
                        return;
                }
            }
            catch (Exception ex)
            {
                HandleException("SetSwitchValue", MemberType.Property, Required.Mandatory, ex, "");
            }
        }

        private void CheckInaccessibleOutOfRange(SwitchMethod method)
        {
            const short LOW_TEST_VALUE = -1;
            const short HIGH_TEST_VALUE = 1;
            bool BoolValue;
            string StringValue;
            double DoubleValue;

            if (m_MaxSwitch > 0)
            {

                // Try a value below 0
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage(method.ToString(), string.Format("About to call {0} with invalid low value for switch number: {1} for ", method.ToString(), LOW_TEST_VALUE));
                    switch (method)
                    {
                        case SwitchMethod.CanWrite:
                            {
                                BoolValue = m_Switch.CanWrite(LOW_TEST_VALUE);
                                break;
                            }

                        case SwitchMethod.GetSwitch:
                            {
                                BoolValue = m_Switch.GetSwitch(LOW_TEST_VALUE);
                                break;
                            }

                        case SwitchMethod.GetSwitchDescription:
                            {
                                StringValue = m_Switch.GetSwitchDescription(LOW_TEST_VALUE);
                                break;
                            }

                        case SwitchMethod.GetSwitchName:
                            {
                                StringValue = m_Switch.GetSwitchName(LOW_TEST_VALUE);
                                break;
                            }

                        case SwitchMethod.GetSwitchValue:
                            {
                                DoubleValue = m_Switch.GetSwitchValue(LOW_TEST_VALUE);
                                break;
                            }

                        case SwitchMethod.MaxSwitchValue:
                            {
                                DoubleValue = m_Switch.MaxSwitchValue(LOW_TEST_VALUE);
                                break;
                            }

                        case SwitchMethod.MinSwitchValue:
                            {
                                DoubleValue = m_Switch.MinSwitchValue(LOW_TEST_VALUE);
                                break;
                            }

                        case SwitchMethod.SetSwitch:
                            {
                                m_Switch.SetSwitch(LOW_TEST_VALUE, false);
                                break;
                            }

                        case SwitchMethod.SetSwitchName:
                            {
                                break;
                            }

                        case SwitchMethod.SetSwitchValue:
                            {
                                m_Switch.SetSwitchValue(LOW_TEST_VALUE, 0.0);
                                break;
                            }

                        case SwitchMethod.SwitchStep:
                            {
                                DoubleValue = m_Switch.SwitchStep(LOW_TEST_VALUE);
                                break;
                            }

                        default:
                            {
                                LogError("CheckInaccessibleOutOfRange",$"Unknown value of SwitchMethod Enum: {method}");
                                break;
                            }
                    }
                    LogIssue("SwitchNumber", "Switch did not throw an exception when a switch ID below 0 was used in method: " + method.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOK("SwitchNumber", MemberType.Property, Required.MustBeImplemented, ex, "when a switch ID below 0 was used in method: " + method.ToString(), "Switch device threw an InvalidOperationException when a switch ID below 0 was used in method: " + method.ToString());
                }

                // Try a value above MaxSwitch
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage(method.ToString(), string.Format("About to call {0} with invalid high value for switch number: {1} for ", method.ToString(), m_MaxSwitch + HIGH_TEST_VALUE));
                    switch (method)
                    {
                        case SwitchMethod.CanWrite:
                            {
                                BoolValue = m_Switch.CanWrite((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        case SwitchMethod.GetSwitch:
                            {
                                BoolValue = m_Switch.GetSwitch((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        case SwitchMethod.GetSwitchDescription:
                            {
                                StringValue = m_Switch.GetSwitchDescription((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        case SwitchMethod.GetSwitchName:
                            {
                                StringValue = m_Switch.GetSwitchName((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        case SwitchMethod.GetSwitchValue:
                            {
                                DoubleValue = m_Switch.GetSwitchValue((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        case SwitchMethod.MaxSwitchValue:
                            {
                                DoubleValue = m_Switch.MaxSwitchValue((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        case SwitchMethod.MinSwitchValue:
                            {
                                DoubleValue = m_Switch.MinSwitchValue((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        case SwitchMethod.SetSwitch:
                            {
                                m_Switch.SetSwitch((short)(m_MaxSwitch + HIGH_TEST_VALUE), false);
                                break;
                            }

                        case SwitchMethod.SetSwitchName:
                            {
                                break;
                            }

                        case SwitchMethod.SetSwitchValue:
                            {
                                m_Switch.SetSwitchValue((short)(m_MaxSwitch + HIGH_TEST_VALUE), 0.0);
                                break;
                            }

                        case SwitchMethod.SwitchStep:
                            {
                                DoubleValue = m_Switch.SwitchStep((short)(m_MaxSwitch + HIGH_TEST_VALUE));
                                break;
                            }

                        default:
                            {
                                LogError("CheckInaccessibleOutOfRange", $"Unknown value of SwitchMethod Enum: {method}");
                                break;
                            }
                    }
                    LogIssue("SwitchNumber", "Switch did not throw an exception when a switch ID above MaxSwitch was used in method: " + method.ToString());
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOK("SwitchNumber", MemberType.Property, Required.MustBeImplemented, ex, "when a switch ID above MaxSwitch was used in method: " + method.ToString(), "Switch device threw an InvalidOperationException when a switch ID above MaxSwitch was used in method: " + method.ToString());
                }
            }
            else
                LogInfo("SwitchNumber", "Skipping range tests because MaxSwitch cannot be read");
        }
    }
}

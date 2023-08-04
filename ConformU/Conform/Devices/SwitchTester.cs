using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ASCOM;

namespace ConformU
{
    internal class SwitchTester : DeviceTesterBaseClass
    {
        private int performanceGetSwitchName, performanceGetSwitch;
        private short maxSwitch;
        private bool canReadMaxSwitch;
        private readonly int extendedSwitchNumberTestRange; // Checks for usable switches either side of the expected range
        private readonly int switchWriteDelay;
        private readonly int switchReadDelay;

        private const int NUMBER_OF_SWITCH_TEST_STATES = 10;
        private const double BAD_SWITCH_VALUE = double.NaN; // Do not change this value, the Double.IsNaN method is used in various tests in the code below

        private enum SwitchPropertyMethod
        {
            MaxSwitch,
            GetSwitch,
            GetSwitchName,
            SwitchType
        }

        private enum SwitchMethod
        {
            CanAsync,
            CancelAsync,
            CanWrite,
            GetSwitch,
            GetSwitchDescription,
            GetSwitchName,
            GetSwitchValue,
            MaxSwitchValue,
            MinSwitchValue,
            SetAsync,
            SetSwitch,
            SetSwitchName,
            SetSwitchValue,
            SetAsyncValue,
            StateChangeComplete,
            SwitchStep
        }

        // Helper variables
        private ISwitchV3 switchDevice;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose

        public SwitchTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;

            performanceGetSwitchName = int.MinValue; // Initialise to silly values
            performanceGetSwitch = int.MinValue;

            switchReadDelay = settings.SwitchReadDelay; // Get values for the two delay parameters as set by the user or the default values if not yet set
            switchWriteDelay = settings.SwitchWriteDelay;
            extendedSwitchNumberTestRange = settings.SwitchExtendedNumberTestRange;
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
                    switchDevice?.Dispose();
                    switchDevice = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        public override void InitialiseTest()
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
                        switchDevice = new AlpacaSwitch(
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
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                switchDevice = new SwitchFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                switchDevice = new Switch(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(switchDevice, DeviceTypes.Switch); // Assign the driver to the base class

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

        public override void CheckProperties()
        {
            // MaxSwitch - Mandatory
            switch (GetInterfaceVersion())
            {
                case 1: // Original Platform 5 switch interface, ISwitchV2 and ISwitchV3 have the same property
                case 2:
                case 3:
                    maxSwitch = SwitchPropertyTestShort(SwitchPropertyMethod.MaxSwitch, "MaxSwitch", 1, short.MaxValue);
                    break;

                default:
                    LogIssue("Switches", $"Unknown switch interface version: {GetInterfaceVersion()}");
                    break;
            }

            if (cancellationToken.IsCancellationRequested)
                return;
        }

        public override void CheckMethods()
        {
            short i;
            bool getSwitch = false, getSwitchOriginal = false;
            bool getSwitchOk, setSwitchOk;
            double getSwitchValueOriginal = 0.0;
            string switchName;

            switch (GetInterfaceVersion())
            {
                case 1: // Platform 5 interface v1
                    if (canReadMaxSwitch)
                    {

                        // Find valid GetSwitch values
                        for (i = 0; i <= System.Convert.ToInt16(maxSwitch - 1); i++)
                        {
                            SetTest($"Read/Write Switch {i}");
                            getSwitchOk = false;
                            setSwitchOk = false;
                            try // Read switch state
                            {
                                LogCallToDriver($"GetSwitch {i}", $"About to call GetSwitch({i}) method");
                                SetAction($"GetSwitch");
                                getSwitchOriginal = switchDevice.GetSwitch(i);
                                WaitForReadDelay("GetSwitch");

                                LogOk($"GetSwitch {i}", $"Found switch, state: {getSwitchOriginal}");
                                getSwitchOk = true;
                                if (i > maxSwitch)
                                    LogIssue($"GetSwitch {i}", "Usable switch found above MaxSwitch!");
                                else
                                    performanceGetSwitch = i;// Save last good switch number for performance test
                            }
                            catch (Exception ex)
                            {
                                if (IsNotSetException(ex))
                                    LogInfo($"GetSwitch {i}", "Switch read is not implemented");
                                else
                                {
                                    LogInfo($"GetSwitch {i}", $"Unable to read switch: {ex.Message}");
                                    LogDebug($"GetSwitch {i}", $"Exception: {ex}");
                                }
                            }

                            try // Now try to write the value
                            {
                                LogCallToDriver($"SetSwitch {i}", $"About to call SetSwitch({i})");
                                SetAction($"SetSwitch {!getSwitchOriginal}");
                                switchDevice.SetSwitch(i, !getSwitchOriginal); // Swap the switch state
                                WaitForWriteDelay($"SetSwitch {!getSwitchOriginal}");

                                setSwitchOk = true;
                                if (getSwitchOk)
                                {
                                    LogCallToDriver("SetSwitch", $"About to call GetSwitch({i}) method");
                                    SetAction($"GetSwitch");
                                    bool newSwitchState = switchDevice.GetSwitch(i);
                                    WaitForReadDelay("GetSwitch");

                                    if (newSwitchState == !getSwitchOriginal)
                                    {
                                        LogOk($"SetSwitch {i}", "Switch correctly changed state");
                                        LogCallToDriver("SetSwitch", "About to call SetSwitch method");
                                        SetAction($"SetSwitch {getSwitchOriginal}");
                                        switchDevice.SetSwitch(i, getSwitchOriginal); // Now put switch back to original state
                                        WaitForWriteDelay($"SetSwitch {getSwitchOriginal}");
                                    }
                                    else
                                        LogIssue($"SetSwitch {i}",
                                            $"Switch did not change state, currently it is {newSwitchState}");
                                }
                                else
                                    LogInfo($"SetSwitch {i}", "You have a write only switch!");
                                if (i > maxSwitch)
                                    LogIssue($"SetSwitch {i}", "Usable switch found above MaxSwitch!");
                            }
                            catch (Exception ex)
                            {
                                if (IsNotSetException(ex))
                                    LogInfo($"SetSwitch {i}", "Switch write is not implemented");
                                else
                                {
                                    LogInfo($"SetSwitch {i}", $"Unable to write to switch: {ex.Message}");
                                    LogDebug($"SetSwitch {i}", $"Exception: {ex}");
                                }
                            }

                            try
                            {
                                LogCallToDriver($"GetSwitchName {i}", $"About to get switch name {i}");
                                switchName = switchDevice.GetSwitchName(i);
                                if (getSwitchOk | setSwitchOk)
                                {
                                    if (switchName == "")
                                        LogInfo($"GetSwitchName {i}", "Switch name returns null string");
                                    else
                                    {
                                        performanceGetSwitchName = i; // Save last good name index for performance test
                                        LogOk($"GetSwitchName {i}", $"Found switch, name:  {switchName}");
                                    }
                                }
                                else if (switchName == "")
                                    LogIssue($"GetSwitchName {i}", "Switch name returns null string but switch can neither read nor write!");
                                else
                                    LogIssue($"GetSwitchName {i}",
                                        $"Found switch, name:  {switchName} which can neither read nor write!");
                            }
                            catch (Exception ex)
                            {
                                LogDebug($"GetSwitchName {i}", $"Exception: {ex}");
                            }
                        }
                        SetTest("");
                        SetAction("");
                        SetStatus("");
                    }
                    else
                        LogIssue("SwitchCheckMethods", "Skipping further tests as there is no valid value for MaxSwitch");
                    break;

                case 2: // Platform 6 interface version 2
                case 3: // Platform 6 interface version 3
                    if (canReadMaxSwitch)
                    {
                        SetTest("Read/Write Switches");

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

                        // Check the async methods if present
                        if (DeviceCapabilities.HasAsyncSwitch(GetInterfaceVersion()))
                        {
                            CheckInaccessibleOutOfRange(SwitchMethod.CanAsync);
                            CheckInaccessibleOutOfRange(SwitchMethod.CancelAsync);
                            CheckInaccessibleOutOfRange(SwitchMethod.SetAsync);
                            CheckInaccessibleOutOfRange(SwitchMethod.SetAsyncValue);
                            CheckInaccessibleOutOfRange(SwitchMethod.StateChangeComplete);
                        }

                        // Find valid GetSwitch values
                        LogDebug("GetSwitchName ",
                            $"Extended switch number test range: {-extendedSwitchNumberTestRange} - {maxSwitch + extendedSwitchNumberTestRange - 1}");
                        for (i = (short)-extendedSwitchNumberTestRange; i <= Convert.ToInt16(maxSwitch + extendedSwitchNumberTestRange - 1); i++)
                        {
                            SetStatus(i.ToString());

                            // Initialise status variables
                            bool setSwitchValueMinOk = false;
                            bool setSwitchValueMaxOk = false;
                            double switchStep = BAD_SWITCH_VALUE;

                            try // Read switch name to determine whether this is a valid switch
                            {
                                LogCallToDriver("GetSwitchName", $"About to get switch {i} name");
                                switchName = switchDevice.GetSwitchName(i);
                                LogOk("GetSwitchName ", $"Found switch {i}");
                                SetTest($"Testing switch {i}");

                                // Test that the switch number is in the valid range of 0..MaxSwitch-1
                                if (i > (maxSwitch - 1))
                                    LogIssue("GetSwitchName ", "Usable switch found above MaxSwitch - 1!");
                                else if (i < 0)
                                    LogIssue("GetSwitchName ", "Usable switch found below 0!");
                                else
                                    performanceGetSwitch = i;// Save last good switch number for performance test

                                LogOk("GetSwitchName ", $"  Name: {switchName}");

                                // Read switch description
                                try
                                {
                                    SetAction("Getting switch description"); LogCallToDriver("GetSwitchDescription",
                                        $"  About to get switch {i} description");
                                    string switchDescription = switchDevice.GetSwitchDescription(i);
                                    LogOk("GetSwitchDescription ", $"  Description: {switchDescription}");
                                }
                                catch (Exception ex)
                                {
                                    LogIssue("GetSwitchDescription ",
                                        $"Mandatory parameter threw an exception: {ex.Message}");
                                }

                                // Read switch minimum value
                                double switchMinimum;
                                try
                                {
                                    SetAction("Getting switch minimum value"); LogCallToDriver("MinSwitchValue",
                                        $"  About to get switch {i} minimum value");
                                    switchMinimum = switchDevice.MinSwitchValue(i);
                                    LogOk("MinSwitchValue ", $"  Minimum: {switchMinimum}");
                                }
                                catch (Exception ex)
                                {
                                    LogIssue("MinSwitchValue ",
                                        $"Mandatory parameter threw an exception: {ex.Message}");
                                    switchMinimum = BAD_SWITCH_VALUE;
                                }

                                // Read switch maximum value
                                double switchMaximum;
                                double switchRange;
                                try
                                {
                                    SetAction("Getting switch maximum value"); LogCallToDriver("MaxSwitchValue",
                                        $"  About to get switch {i} maximum value");
                                    switchMaximum = switchDevice.MaxSwitchValue(i);

                                    if (IsGoodValue(switchMinimum))
                                    {
                                        // Maximum value must be > Minimum value
                                        if (switchMaximum > switchMinimum)
                                        {
                                            LogOk("MaxSwitchValue ", $"  Maximum: {switchMaximum}");
                                            switchRange = switchMaximum - switchMinimum; // Calculate the range of values the switch might take
                                        }
                                        else
                                        {
                                            LogInfo("MaxSwitchValue ", $"  Maximum: {switchMaximum}");
                                            LogIssue("MaxSwitchValue ", "MaxSwitchValue is less or equal to MinSwitchValue, it must be greater!");
                                            switchRange = BAD_SWITCH_VALUE; // Special value because the maximum or minimum values are bad
                                        }
                                    }
                                    else
                                    {
                                        LogInfo("MaxSwitchValue ", $"  Maximum: {switchMaximum}");
                                        switchRange = BAD_SWITCH_VALUE;
                                        LogInfo("MaxSwitchValue ", "  Test that switch Maximum is greater than Minimum skipped because of an error reading the Minimum value.");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogIssue("MaxSwitchValue ",
                                        $"Mandatory parameter threw an exception: {ex.Message}");
                                    switchMaximum = BAD_SWITCH_VALUE;
                                    switchRange = BAD_SWITCH_VALUE;
                                }

                                // Read switch step value
                                try
                                {
                                    SetAction("Getting switch step size"); LogCallToDriver("SwitchStep",
                                        $"  About to get switch {i} step size");
                                    switchStep = switchDevice.SwitchStep(i);
                                    LogOk("SwitchStep ", $"  Step size: {switchStep}");

                                    // Step must be greater than 0
                                    if (switchStep > 0.0)
                                    {
                                        LogOk("SwitchStep ", "  Step size is greater than zero");

                                        // Step must be less than or equal to the range of possible values
                                        if (switchStep <= switchRange)
                                        {
                                            LogOk("SwitchStep ", "  Step size is less than the range of possible values");

                                            // Now check that the switch range is an integer multiple of the step size
                                            // Doubles are converted to the Decimal type (which has higher precision) in order to avoid unexpected outcomes from Mod due to rounding errors
                                            switch (Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)))
                                            {
                                                case 0M:
                                                    LogOk("SwitchStep ", "  The switch range is an integer multiple of the step size.");
                                                    break;

                                                case object o when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) <= Convert.ToDecimal(switchStep / 100):
                                                    LogInfo("SwitchStep ", "  The switch range is within 1% of being an integer multiple of the step size.");
                                                    break;

                                                case object o when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) <= Convert.ToDecimal(switchStep / 10):
                                                    LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 10%.");
                                                    break;

                                                case object o when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) <= Convert.ToDecimal(switchStep / 5):
                                                    LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 20%.");
                                                    break;

                                                case object o when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)) <= Convert.ToDecimal(switchStep / 2):
                                                    LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 50%.");
                                                    break;

                                                default:
                                                    LogIssue("SwitchStep ",
                                                        $"The switch range must be an integer multiple of the step size. Remainder`: {decimal.Subtract(Convert.ToDecimal(switchMaximum), Convert.ToDecimal(switchMinimum)) % Convert.ToDecimal(switchStep)}");
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            LogIssue("SwitchStep ", "Step size must be less than the range of possible values (MaxSwitchValue - MinSwitchValue");
                                            switchStep = BAD_SWITCH_VALUE;
                                        }
                                    }
                                    else
                                    {
                                        LogIssue("SwitchStep ", "Step size must be greater than zero");
                                        switchStep = BAD_SWITCH_VALUE;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogIssue("SwitchStep ", $"Mandatory parameter threw an exception: {ex.Message}");
                                }

                                LogDebug("SwitchMinimum ", switchMinimum.ToString());
                                LogDebug("SwitchMaximum ", switchMaximum.ToString());
                                LogDebug("SwitchStep ", switchStep.ToString());
                                LogDebug("SwitchRange ", switchRange.ToString());

                                // Read CanWrite 
                                bool switchCanWrite;
                                try
                                {
                                    LogCallToDriver("CanWrite", $"  About to get switch {i} CanWrite status");
                                    switchCanWrite = switchDevice.CanWrite(i);
                                    LogOk("CanWrite ", $"  CanWrite: {switchCanWrite}");
                                }
                                catch (Exception ex)
                                {
                                    LogIssue("CanWrite ", $"Mandatory parameter threw an exception: {ex.Message}");
                                    LogInfo("CanWrite ", "Assuming that CanWrite is false");
                                    switchCanWrite = false; // Initialise to a default state
                                }

                                // Read and test CanAsync if present
                                bool switchCanAsync = false;
                                if (DeviceCapabilities.HasAsyncSwitch(GetInterfaceVersion()))
                                {
                                    // Read CanAsync 
                                    try
                                    {
                                        LogCallToDriver("CanAsync", $"  About to get switch {i} CanAsync status");
                                        switchCanAsync = switchDevice.CanAsync(i);
                                        LogDebug("CanAsync",$"CanAsync: {switchCanAsync}, CanWrite: {switchCanWrite}");

                                        if (switchCanWrite & switchCanAsync) // Switch can write and async - OK
                                        {
                                            LogOk("CanAsync ", $"  CanAsync: {switchCanAsync}");
                                        }
                                        else // The device is not capable of one or both of write or async so test the remaining 3 combinations
                                        {
                                            if (switchCanWrite) // Switch can write but not async - OK
                                                LogOk("CanAsync ", $"  CanAsync: {switchCanAsync}");
                                            else // Switch cannot write so check whether CanAsync is set
                                            {
                                                if (switchCanAsync) // Switch cannot write but can async - BAD
                                                    LogIssue("CanAsync ", $"  CanWrite is False but CanAsync is True. CanWrite must be true if CanAsync is true.");
                                                else // Switch cannot write or async - OK
                                                    LogOk("CanAsync ", $"  CanAsync: {switchCanAsync}");
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogIssue("CanAsync ", $"Mandatory parameter threw an exception: {ex.Message}");
                                        LogInfo("CanAsync ", "Assuming that CanAsync is false");
                                        switchCanAsync = false; // Initialise to a default state
                                    }
                                }

                                // Access GetSwitch and record the outcome
                                try
                                {
                                    SetAction($"GetSwitch"); LogCallToDriver("GetSwitch",
                                        $"  About to call GetSwitch({i}) method");
                                    getSwitchOriginal = switchDevice.GetSwitch(i);
                                    WaitForReadDelay("GetSwitch");

                                    getSwitchOk = true;
                                    LogOk("GetSwitch ", $"  {getSwitchOriginal}");
                                }
                                catch (Exception ex)
                                {
                                    LogDebug("GetSwitch ", $"Exception: {ex}");
                                    getSwitchOk = false;
                                }

                                // Access GetSwitchValue and record the outcome
                                bool getSwitchValueOk;
                                try
                                {
                                    SetAction($"GetSwitchValue"); LogCallToDriver("GetSwitchValue",
                                        $"  About to call GetSwitchValue({i}) method");
                                    getSwitchValueOriginal = switchDevice.GetSwitchValue(i);
                                    WaitForReadDelay("GetSwitchValue");
                                    getSwitchValueOk = true;
                                    LogOk("GetSwitchValue ", $"  {getSwitchValueOriginal}");
                                }
                                catch (Exception ex)
                                {
                                    LogDebug("GetSwitchValue ", $"Exception: {ex}");
                                    getSwitchValueOk = false;
                                }

                                // Now try to write to see which of these methods are available
                                if (settings.SwitchEnableSet)
                                {

                                    // Try to set the two boolean values through SetSwitch
                                    double getSwitchValue;
                                    try
                                    {
                                        // Try SetSwitch(False)
                                        SetAction($"SetSwitch {i} False"); LogCallToDriver("SetSwitch", $"  About to call SetSwitch({i}, {false}) method");
                                        SetAction("SetSwitch false");
                                        switchDevice.SetSwitch(i, false); // Set switch false
                                        WaitForWriteDelay($"SetSwitch False");

                                        // Check GetSwitch
                                        if (getSwitchOk)
                                        {
                                            LogCallToDriver("SetSwitch", $"  About to call GetSwitch({i}) method");

                                            SetAction($"GetSwitch");
                                            if (switchDevice.GetSwitch(i) == false)
                                                LogOk("SetSwitch ", "  GetSwitch returned False after SetSwitch(False)");
                                            else
                                                LogIssue("SetSwitch ", "  GetSwitch returned True after SetSwitch(False)");
                                            WaitForReadDelay("GetSwitch");
                                        }
                                        else
                                            LogInfo("SetSwitch ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                        // Check GetSwitchValue returns the switch minimum value
                                        if (getSwitchValueOk & IsGoodValue(switchMinimum))
                                        {
                                            LogCallToDriver("SetSwitch",
                                                $"  About to call GetSwitchValue({i}) method");
                                            SetAction($"GetSwitchValue");
                                            getSwitchValue = switchDevice.GetSwitchValue(i);
                                            WaitForReadDelay("GetSwitchValue");
                                            switch (getSwitchValue)
                                            {
                                                case object o when Math.Abs(getSwitchValue - switchMinimum) < getSwitchValue * 0.0001:
                                                    LogOk("SetSwitch ", "  GetSwitchValue returned MINIMUM_VALUE after SetSwitch(False)");
                                                    break;

                                                case object o when switchMinimum * 0.99 <= getSwitchValue && getSwitchValue <= switchMinimum * 1.01:
                                                    LogOk("SetSwitch ", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitch(False)");
                                                    break;

                                                default:
                                                    LogIssue("SetSwitch ",
                                                        $"  GetSwitchValue did not return MINIMUM_VALUE after SetSwitch(False): {getSwitchValue}");
                                                    break;
                                            }
                                        }
                                        else
                                            LogInfo("SetSwitch ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMinimum methods");
                                        if (cancellationToken.IsCancellationRequested) return;

                                        // Try SetSwitch(True)
                                        SetAction($"SetSwitch {i} True"); LogCallToDriver("SetSwitch",
                                            $"  About to call SetSwitch({i}, {true}) method");
                                        SetAction($"SetSwitch true");
                                        switchDevice.SetSwitch(i, true); // Set switch true
                                        WaitForWriteDelay("SetSwitch true");

                                        // Check GetSwitch
                                        if (getSwitchOk)
                                        {
                                            LogCallToDriver("SetSwitch", $"  About to call GetSwitch({i}) method");
                                            SetAction($"GetSwitch");
                                            if (switchDevice.GetSwitch(i) == true)
                                                LogOk("SetSwitch ", "  GetSwitch read True after SetSwitch(True)");
                                            else
                                                LogIssue("SetSwitch ", "  GetSwitch read False after SetSwitch(True)");
                                            WaitForReadDelay("GetSwitch");
                                        }
                                        else
                                            LogInfo("SetSwitch ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                        // Check GetSwitchValue returns the switch maximum value
                                        if (getSwitchValueOk & IsGoodValue(switchMaximum))
                                        {
                                            LogCallToDriver("SetSwitch",
                                                $"  About to call GetSwitchValue({i}) method");
                                            SetAction($"GetSwitchValue");
                                            getSwitchValue = switchDevice.GetSwitchValue(i);
                                            WaitForReadDelay("GetSwitchValue");
                                            switch (getSwitchValue)
                                            {
                                                case object o when Math.Abs(getSwitchValue - switchMaximum) < getSwitchValue * 0.0001:
                                                    LogOk("SetSwitch ", "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitch(True)");
                                                    break;

                                                case object o when switchMaximum * 0.99 <= switchMaximum && switchMaximum <= switchMaximum * 1.01:
                                                    LogOk("SetSwitch ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitch(True)");
                                                    break;

                                                default:
                                                    LogIssue("SetSwitch ",
                                                        $"  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitch(True): {getSwitchValue}");
                                                    break;
                                            }
                                        }
                                        else
                                            LogInfo("SetSwitch ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMaximum methods");
                                        if (cancellationToken.IsCancellationRequested) return;

                                        // Return to original state if possible,otherwise set to false
                                        if (getSwitchOk)
                                        {
                                            LogCallToDriver("SetSwitch",
                                                $"  About to call SetSwitch({i}, {getSwitch}) method");
                                            SetAction($"SetSwitch {getSwitch} to its original value");
                                            switchDevice.SetSwitch(i, getSwitch); // Return to the original state
                                            WaitForWriteDelay($"SetSwitch {getSwitch} to its original value");
                                        }
                                        else
                                        {
                                            LogCallToDriver("SetSwitch",
                                                $"  About to call SetSwitch({i}, {false}) method");
                                            SetAction("SeSwitch false");
                                            switchDevice.SetSwitch(i, false); // Set to false
                                            WaitForWriteDelay("SetSwitch false");
                                        }

                                        setSwitchOk = true;
                                        LogDebug("SetSwitch ", "Set value OK");
                                        if (cancellationToken.IsCancellationRequested) return;

                                    }
                                    catch (Exception ex)
                                    {
                                        LogDebug("SetSwitch ", $"Exception generated - Switch can write: {switchCanWrite}, IsMethodNotImplementedException: {IsMethodNotImplementedException(ex)}, IsNotImplementedException: {IsNotImplementedException(ex)}, COM Access mechanic: {settings.ComConfiguration.ComAccessMechanic}, Device Technology: {settings.DeviceTechnology}");
                                        if (switchCanWrite)
                                        {
                                            LogIssue("SetSwitch ", $"Exception: {ex.Message}");
                                            LogDebug("SetSwitch ", $"Exception: {ex}");
                                        }
                                        else if (IsMethodNotImplementedException(ex))
                                            LogOk("SetSwitch ", "  CanWrite is False and MethodNotImplementedException was thrown");
                                        // Determine whether we are testing a COM device using the cross platform DriverAccess module and, if so, test for the NotimplementedExceptionso that it returns.
                                        else if ((settings.ComConfiguration.ComAccessMechanic == ComAccessMechanic.DriverAccess) & (settings.DeviceTechnology == DeviceTechnology.COM))
                                        {
                                            if (IsNotImplementedException(ex)) // Got a NotImplementedException, which is OK
                                            {
                                                LogOk("SetSwitch ", "  CanWrite is False and NotImplementedException was thrown by DriverAccess");
                                            }
                                            else // Got some other exception, which is an issue.
                                            {
                                                LogIssue("SetSwitch ", $"Exception: {ex.Message}");
                                                LogDebug("SetSwitch ", $"Exception: {ex}");
                                            }
                                        }
                                        else // Received an unexpected exception, which is an issue.
                                        {
                                            LogIssue("SetSwitch ", $"Exception: {ex.Message}");
                                            LogDebug("SetSwitch ", $"Exception: {ex}");
                                        }
                                    }

                                    // Try to set the minimum and maximum values through SetSwitchValue
                                    try
                                    {
                                        if (IsGoodValue(switchMinimum))
                                        {
                                            LogCallToDriver("SetSwitchValue",
                                                $"  About to call SetSwitchValue({i}, {switchMinimum}), attempting to set the minimum permissible value");
                                            SetAction($"SetSwitchValue {switchMinimum}");
                                            switchDevice.SetSwitchValue(i, switchMinimum); // Set switch to minimum
                                            WaitForWriteDelay($"SetSwitchValue {switchMinimum}");

                                            // Check GetSwitch
                                            if (getSwitchOk)
                                            {
                                                LogCallToDriver("SetSwitchValue",
                                                    $"  About to call GetSwitch({i}) method");
                                                SetAction("GetSwitch");
                                                if (switchDevice.GetSwitch(i) == false)
                                                    LogOk("SetSwitchValue", "  GetSwitch returned False after SetSwitchValue(MINIMUM_VALUE)");
                                                else
                                                    LogIssue("SetSwitchValue", "  GetSwitch returned True after SetSwitchValue(MINIMUM_VALUE)");
                                                WaitForReadDelay("GetSwitch");
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                            // Check GetSwitchValue returns the switch minimum value
                                            if (getSwitchValueOk)
                                            {
                                                LogCallToDriver("SetSwitchValue",
                                                    $"  About to call GetSwitchValue({i}) method");
                                                SetAction("GetSwitchValue");
                                                getSwitchValue = switchDevice.GetSwitchValue(i);
                                                WaitForReadDelay("GetSwitchValue");

                                                switch (getSwitchValue)
                                                {
                                                    case object o when Math.Abs(getSwitchValue - switchMinimum) < getSwitchValue * 0.0001:
                                                        LogOk("SetSwitchValue", "  GetSwitchValue returned MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)");
                                                        setSwitchValueMinOk = true;
                                                        break;

                                                    case object o when switchMinimum * 0.99 <= getSwitchValue && getSwitchValue <= switchMinimum * 1.01:
                                                        LogOk("SetSwitchValue", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)");
                                                        setSwitchValueMinOk = true;
                                                        break;

                                                    default:
                                                        LogIssue("SetSwitchValue",
                                                            $"  GetSwitchValue did not return MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE): {getSwitchValue}");
                                                        break;
                                                }
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");
                                            if (cancellationToken.IsCancellationRequested) return;

                                            // Now try a value below minimum
                                            try
                                            {
                                                LogCallToDriver("SetSwitchValue",
                                                    $"  About to call SetSwitchValue({i}, {switchMinimum - 1.0}), attempting to set an invalid low value");
                                                SetAction($"SetSwitchValue {switchMinimum - 1.0}");
                                                switchDevice.SetSwitchValue(i, switchMinimum - 1.0);
                                                WaitForWriteDelay($"SetSwitchValue {switchMinimum - 1.0}");

                                                LogIssue("SetSwitchValue",
                                                    $"Switch did not throw an exception when a value below SwitchMinimum was set: {(switchMinimum - 1.0)}");
                                            }
                                            catch (Exception ex)
                                            {
                                                HandleInvalidValueExceptionAsOk("SetSwitchValue", MemberType.Method, Required.Mandatory, ex,
                                                    $"when setting a value below SwitchMinimum - {(switchMinimum - 1.0)}",
                                                    $"  Switch threw an InvalidOperationException when a value below SwitchMinimum was set: {(switchMinimum - 1.0)}");
                                            }
                                        }
                                        else
                                            LogInfo("SetSwitchValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                        // Try SetSwitchValue(MAXIMUM_VALUE)
                                        if (IsGoodValue(switchMaximum))
                                        {
                                            LogCallToDriver("SetSwitchValue",
                                                $"  About to call SetSwitchValue({i}, {switchMaximum}), attempting to set the maximum permissible value");
                                            SetAction($"SetSwitchValue maximum {switchMaximum}");
                                            switchDevice.SetSwitchValue(i, switchMaximum); // Set switch to maximum
                                            WaitForWriteDelay($"SetSwitchValue maximum {switchMaximum}");

                                            // Check GetSwitch
                                            if (getSwitchOk)
                                            {
                                                LogCallToDriver("SetSwitchValue",
                                                    $"  About to call GetSwitch({i}) method");
                                                SetAction("GetSwitch");
                                                if (switchDevice.GetSwitch(i) == true)
                                                    LogOk("SetSwitchValue ", "  GetSwitch returned True after SetSwitchValue(MAXIMUM_VALUE)");
                                                else
                                                    LogIssue("SetSwitchValue ", "  GetSwitch returned False after SetSwitchValue(MAXIMUM_VALUE)");
                                                WaitForReadDelay("GetSwitch");

                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                            // Check GetSwitchValue returns the switch maximum value
                                            if (getSwitchValueOk)
                                            {
                                                LogCallToDriver("SetSwitchValue",
                                                    $"  About to call GetSwitchValue({i}) method");
                                                SetAction("GetSwitchValue");
                                                getSwitchValue = switchDevice.GetSwitchValue(i);
                                                WaitForReadDelay("GetSwitchValue");

                                                switch (getSwitchValue)
                                                {
                                                    case object o when Math.Abs(getSwitchValue - switchMaximum) < getSwitchValue * 0.0001:
                                                        LogOk("SetSwitchValue ", "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)");
                                                        setSwitchValueMaxOk = true;
                                                        break;

                                                    case object o when switchMaximum * 0.99 <= getSwitchValue && getSwitchValue <= switchMaximum * 1.01:
                                                        LogOk("SetSwitchValue ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)");
                                                        break;

                                                    default:
                                                        LogIssue("SetSwitchValue ",
                                                            $"  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE): {getSwitchValue}");
                                                        break;
                                                }
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");

                                            // Now try a value above maximum
                                            try
                                            {
                                                LogCallToDriver("SetSwitchValue",
                                                    $"  About to call SetSwitchValue({i}, {switchMaximum + 1.0}), attempting to set an invalid high value");
                                                SetAction($"SetSwitchValue {switchMaximum + 1.0}");
                                                switchDevice.SetSwitchValue(i, switchMaximum + 1.0);
                                                WaitForWriteDelay($"SetSwitchValue {switchMaximum + 1.0}");

                                                LogIssue("SetSwitchValue",
                                                    $"Switch did not throw an exception when a value above SwitchMaximum was set: {switchMaximum}{1.0}");
                                            }
                                            catch (Exception ex)
                                            {
                                                HandleInvalidValueExceptionAsOk("SetSwitchValue", MemberType.Method, Required.Mandatory, ex,
                                                    $"when setting a value above SwitchMaximum was set: {switchMaximum}{1.0}",
                                                    $"  Switch threw an InvalidOperationException when a value above SwitchMaximum was set: {switchMaximum}{1.0}");
                                            }
                                            if (cancellationToken.IsCancellationRequested) return;
                                        }
                                        else
                                            LogInfo("SetSwitchValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                        // Test some positions of the multi-state switch between the minimum and maximum values
                                        if (getSwitchValueOk & setSwitchValueMinOk & setSwitchValueMaxOk & IsGoodValue(switchRange) & IsGoodValue(switchStep))
                                        {
                                            TestSetSwitchValue(i, 0.0, switchMinimum, switchMaximum, switchRange, switchStep); if (cancellationToken.IsCancellationRequested) return;
                                            TestSetSwitchValue(i, 0.25, switchMinimum, switchMaximum, switchRange, switchStep); if (cancellationToken.IsCancellationRequested) return;
                                            TestSetSwitchValue(i, 0.5, switchMinimum, switchMaximum, switchRange, switchStep); if (cancellationToken.IsCancellationRequested) return;
                                            TestSetSwitchValue(i, 0.75, switchMinimum, switchMaximum, switchRange, switchStep); if (cancellationToken.IsCancellationRequested) return;
                                        }
                                        else
                                        {
                                            LogInfo("SetSwitchValue ", "Skipping multi state tests because of earlier errors");
                                            LogDebug("GetSwitchValueOK ", getSwitchValueOk.ToString());
                                            LogDebug("SetSwitchValueMinOK ", setSwitchValueMinOk.ToString());
                                            LogDebug("SetSwitchValueMaxOK ", setSwitchValueMaxOk.ToString());
                                            LogDebug("SwitchRange ", switchRange.ToString());
                                            LogDebug("SwitchStep ", switchStep.ToString());
                                        }

                                        // Return to original state if possible,otherwise set to false
                                        if (getSwitchValueOk)
                                        {
                                            LogCallToDriver("SetSwitchValue",
                                                $"  About to call SetSwitchValue({i}, {getSwitchValueOriginal}), attempting to restore pre-test value");
                                            SetAction($"SetSwitchValue to initial value {getSwitchValueOriginal}");
                                            switchDevice.SetSwitchValue(i, getSwitchValueOriginal); // Return to the original state
                                            LogOk("SetSwitchValue ", "  Switch has been reset to its original state");
                                            WaitForWriteDelay($"SetSwitchValue to initial value {getSwitchValueOriginal}");
                                        }
                                        else if (IsGoodValue(switchMinimum) & IsGoodValue(switchMaximum))
                                        {
                                            LogCallToDriver("SetSwitchValue",
                                                $"  About to call SetSwitchValue({i}, {(switchMaximum - switchMinimum) / 2.0}), attempting to set the value to its mid-point");
                                            SetAction($"SetSwitchValue to midpoint {(switchMaximum - switchMinimum) / 2.0}");
                                            switchDevice.SetSwitchValue(i, (switchMaximum - switchMinimum) / 2.0); // Return to the half way state
                                            LogOk("SetSwitchValue ", "  Switch has been reset to half its range");
                                            WaitForWriteDelay($"SetSwitchValue to midpoint {(switchMaximum - switchMinimum) / 2.0}");
                                        }
                                        else
                                            LogInfo("SetSwitchValue ", "Switch can not be returned to its default state because of issues with GetSwitchValue, GetSwitchMinimum or GetSwitchMaximum");
                                        if (cancellationToken.IsCancellationRequested) return;
                                    }
                                    catch (Exception ex)
                                    {
                                        LogDebug("SetSwitchValue ", $"EXCEPTION GENERATED - Switch can write: {switchCanWrite}, IsMethodNotImplementedException: {IsMethodNotImplementedException(ex)}, IsNotImplementedException: {IsNotImplementedException(ex)}, COM Access mechanic: {settings.ComConfiguration.ComAccessMechanic}, Device Technology: {settings.DeviceTechnology}");
                                        if (switchCanWrite)
                                        {
                                            LogIssue("SetSwitchValue ", $"Exception: {ex.Message}");
                                            LogDebug("SetSwitchValue ", $"Exception: {ex}");
                                        }
                                        else if (IsMethodNotImplementedException(ex))
                                            LogOk("SetSwitchValue ", "  CanWrite is False and MethodNotImplementedException was thrown");
                                        // Determine whether we are testing a COM device using the cross platform DriverAccess module and, if so, test for the NotimplementedExceptionso that it returns.
                                        else if ((settings.ComConfiguration.ComAccessMechanic == ComAccessMechanic.DriverAccess) & (settings.DeviceTechnology == DeviceTechnology.COM))
                                        {
                                            if (IsNotImplementedException(ex)) // Got a NotImplementedException, which is OK
                                            {
                                                LogOk("SetSwitchValue ", "  CanWrite is False and NotImplementedException was thrown by DriverAccess");
                                            }
                                            else // Got some other exception, which is an issue.
                                            {
                                                LogIssue("SetSwitchValue ", $"Exception: {ex.Message}");
                                                LogDebug("SetSwitchValue ", $"Exception: {ex}");
                                            }
                                        }
                                        else
                                        {
                                            LogIssue("SetSwitchValue ", $"Exception: {ex.Message}");
                                            LogDebug("SetSwitchValue ", $"Exception: {ex}");
                                        }
                                    }

                                    // Test the async methods if present
                                    if (DeviceCapabilities.HasAsyncSwitch(GetInterfaceVersion()))
                                    {
                                        // Only test if the switch canb write and supports async operation
                                        LogDebug("CanAsync", $"CanAsync: {switchCanAsync}, CanWrite: {switchCanWrite}");
                                        if (switchCanWrite & switchCanAsync)
                                        {
                                            // Test SetAsync
                                            try
                                            {
                                                // Try SetAsync(False)
                                                SetAction($"SetAsync {i} False"); LogCallToDriver("SetAsync", $"  About to call SetAsync({i}, {false}) method");
                                                SetAction("SetAsync false");
                                                switchDevice.SetAsync(i, false); // Set switch false

                                                // Wait for the operation to complete
                                                short i1 = i;
                                                WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);

                                                WaitForWriteDelay($"SetAsync False");

                                                // Check GetSwitch
                                                if (getSwitchOk)
                                                {
                                                    LogCallToDriver("SetAsync", $"  About to call GetSwitch({i}) method");

                                                    SetAction($"GetSwitch");
                                                    if (switchDevice.GetSwitch(i) == false)
                                                        LogOk("SetAsync ", "  GetSwitch returned False after SetAsync(False)");
                                                    else
                                                        LogIssue("SetAsync ", "  GetSwitch returned True after SetAsync(False)");
                                                    WaitForReadDelay("GetSwitch");
                                                }
                                                else
                                                    LogInfo("SetAsync ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                                // Check GetSwitchValue returns the switch minimum value
                                                if (getSwitchValueOk & IsGoodValue(switchMinimum))
                                                {
                                                    LogCallToDriver("SetAsync",
                                                        $"  About to call GetSwitchValue({i}) method");
                                                    SetAction($"GetSwitchValue");
                                                    getSwitchValue = switchDevice.GetSwitchValue(i);
                                                    WaitForReadDelay("GetSwitchValue");
                                                    switch (getSwitchValue)
                                                    {
                                                        case object o when Math.Abs(getSwitchValue - switchMinimum) < getSwitchValue * 0.0001:
                                                            LogOk("SetAsync ", "  GetSwitchValue returned MINIMUM_VALUE after SetAsync(False)");
                                                            break;

                                                        case object o when switchMinimum * 0.99 <= getSwitchValue && getSwitchValue <= switchMinimum * 1.01:
                                                            LogOk("SetAsync ", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetAsync(False)");
                                                            break;

                                                        default:
                                                            LogIssue("SetAsync ",
                                                                $"  GetSwitchValue did not return MINIMUM_VALUE after SetAsync(False): {getSwitchValue}");
                                                            break;
                                                    }
                                                }
                                                else
                                                    LogInfo("SetAsync ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMinimum methods");
                                                if (cancellationToken.IsCancellationRequested) return;

                                                // Try SetAsync(True)
                                                SetAction($"SetAsync {i} True"); LogCallToDriver("SetAsync",
                                                    $"  About to call SetAsync({i}, {true}) method");
                                                SetAction($"SetAsync true");
                                                switchDevice.SetAsync(i, true); // Set switch true
                                                WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);

                                                // Check GetSwitch
                                                if (getSwitchOk)
                                                {
                                                    LogCallToDriver("SetAsync", $"  About to call GetSwitch({i}) method");
                                                    SetAction($"GetSwitch");
                                                    if (switchDevice.GetSwitch(i) == true)
                                                        LogOk("SetAsync ", "  GetSwitch read True after SetAsync(True)");
                                                    else
                                                        LogIssue("SetAsync ", "  GetSwitch read False after SetAsync(True)");
                                                    WaitForReadDelay("GetSwitch");
                                                }
                                                else
                                                    LogInfo("SetAsync ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                                // Check GetSwitchValue returns the switch maximum value
                                                if (getSwitchValueOk & IsGoodValue(switchMaximum))
                                                {
                                                    LogCallToDriver("SetAsync",
                                                        $"  About to call GetSwitchValue({i}) method");
                                                    SetAction($"GetSwitchValue");
                                                    getSwitchValue = switchDevice.GetSwitchValue(i);
                                                    WaitForReadDelay("GetSwitchValue");
                                                    switch (getSwitchValue)
                                                    {
                                                        case object o when Math.Abs(getSwitchValue - switchMaximum) < getSwitchValue * 0.0001:
                                                            LogOk("SetAsync ", "  GetSwitchValue returned MAXIMUM_VALUE after SetAsync(True)");
                                                            break;

                                                        case object o when switchMaximum * 0.99 <= switchMaximum && switchMaximum <= switchMaximum * 1.01:
                                                            LogOk("SetAsync ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetAsync(True)");
                                                            break;

                                                        default:
                                                            LogIssue("SetAsync ",
                                                                $"  GetSwitchValue did not return MAXIMUM_VALUE after SetAsync(True): {getSwitchValue}");
                                                            break;
                                                    }
                                                }
                                                else
                                                    LogInfo("SetAsync ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMaximum methods");
                                                if (cancellationToken.IsCancellationRequested) return;

                                                // Return to original state if possible,otherwise set to false
                                                if (getSwitchOk)
                                                {
                                                    LogCallToDriver("SetAsync",
                                                        $"  About to call SetAsync({i}, {getSwitch}) method");
                                                    SetAction($"SetAsync {getSwitch} to its original value");
                                                    switchDevice.SetAsync(i, getSwitch); // Return to the original state
                                                    WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);
                                                }
                                                else
                                                {
                                                    LogCallToDriver("SetAsync",
                                                        $"  About to call SetAsync({i}, {false}) method");
                                                    SetAction("SeSwitch false");
                                                    switchDevice.SetAsync(i, false); // Set to false
                                                    WaitForWriteDelay("SetAsync false");
                                                }

                                                setSwitchOk = true;
                                                LogDebug("SetAsync ", "Set value OK");
                                                if (cancellationToken.IsCancellationRequested) return;

                                            }
                                            catch (Exception ex)
                                            {
                                                LogDebug("SetAsync ", $"Exception generated - Switch can write: {switchCanWrite}, IsMethodNotImplementedException: {IsMethodNotImplementedException(ex)}, IsNotImplementedException: {IsNotImplementedException(ex)}, COM Access mechanic: {settings.ComConfiguration.ComAccessMechanic}, Device Technology: {settings.DeviceTechnology}");
                                                if (switchCanWrite)
                                                {
                                                    LogIssue("SetAsync ", $"Exception: {ex.Message}");
                                                    LogDebug("SetAsync ", $"Exception: {ex}");
                                                }
                                                else if (IsMethodNotImplementedException(ex))
                                                    LogOk("SetAsync ", "  CanWrite is False and MethodNotImplementedException was thrown");
                                                // Determine whether we are testing a COM device using the cross platform DriverAccess module and, if so, test for the NotimplementedExceptionso that it returns.
                                                else if ((settings.ComConfiguration.ComAccessMechanic == ComAccessMechanic.DriverAccess) & (settings.DeviceTechnology == DeviceTechnology.COM))
                                                {
                                                    if (IsNotImplementedException(ex)) // Got a NotImplementedException, which is OK
                                                    {
                                                        LogOk("SetAsync ", "  CanWrite is False and NotImplementedException was thrown by DriverAccess");
                                                    }
                                                    else // Got some other exception, which is an issue.
                                                    {
                                                        LogIssue("SetAsync ", $"Exception: {ex.Message}");
                                                        LogDebug("SetAsync ", $"Exception: {ex}");
                                                    }
                                                }
                                                else // Received an unexpected exception, which is an issue.
                                                {
                                                    LogIssue("SetAsync ", $"Exception: {ex.Message}");
                                                    LogDebug("SetAsync ", $"Exception: {ex}");
                                                }
                                            }

                                            // Try to set the minimum and maximum values through SetSwitchValue
                                            try
                                            {
                                                if (IsGoodValue(switchMinimum))
                                                {
                                                    LogCallToDriver("SetAsyncValue",
                                                        $"  About to call SetAsyncValue({i}, {switchMinimum}), attempting to set the minimum permissible value");
                                                    SetAction($"SetAsyncValue {switchMinimum}");
                                                    switchDevice.SetAsyncValue(i, switchMinimum); // Set switch to minimum

                                                    // Wait for the operation to complete
                                                    short i1 = i;
                                                    WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);

                                                    // Check GetSwitchValue returns the switch minimum value
                                                    if (getSwitchValueOk)
                                                    {
                                                        LogCallToDriver("SetAsyncValue",
                                                            $"  About to call GetSwitchValue({i}) method");
                                                        SetAction("GetSwitchValue");
                                                        getSwitchValue = switchDevice.GetSwitchValue(i);
                                                        WaitForReadDelay("GetSwitchValue");

                                                        switch (getSwitchValue)
                                                        {
                                                            case object o when Math.Abs(getSwitchValue - switchMinimum) < getSwitchValue * 0.0001:
                                                                LogOk("SetAsyncValue", "  GetSwitchValue returned MINIMUM_VALUE after SetAsyncValue(MINIMUM_VALUE)");
                                                                setSwitchValueMinOk = true;
                                                                break;

                                                            case object o when switchMinimum * 0.99 <= getSwitchValue && getSwitchValue <= switchMinimum * 1.01:
                                                                LogOk("SetAsyncValue", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetAsyncValue(MINIMUM_VALUE)");
                                                                setSwitchValueMinOk = true;
                                                                break;

                                                            default:
                                                                LogIssue("SetAsyncValue",
                                                                    $"  GetSwitchValue did not return MINIMUM_VALUE after SetAsyncValue(MINIMUM_VALUE): {getSwitchValue}");
                                                                break;
                                                        }
                                                    }
                                                    else
                                                        LogInfo("SetAsyncValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");
                                                    if (cancellationToken.IsCancellationRequested) return;

                                                    // Now try a value below minimum
                                                    try
                                                    {
                                                        LogCallToDriver("SetAsyncValue",
                                                            $"  About to call SetAsyncValue({i}, {switchMinimum - 1.0}), attempting to set an invalid low value");
                                                        SetAction($"SetAsyncValue {switchMinimum - 1.0}");
                                                        switchDevice.SetAsyncValue(i, switchMinimum - 1.0);

                                                        // Wait for the operation to complete
                                                        WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);

                                                        LogIssue("SetAsyncValue",
                                                            $"Switch did not throw an exception when a value below SwitchMinimum was set: {(switchMinimum - 1.0)}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        HandleInvalidValueExceptionAsOk("SetAsyncValue", MemberType.Method, Required.Mandatory, ex,
                                                            $"when setting a value below SwitchMinimum - {(switchMinimum - 1.0)}",
                                                            $"  Switch threw an InvalidOperationException when a value below SwitchMinimum was set: {(switchMinimum - 1.0)}");
                                                    }
                                                }
                                                else
                                                    LogInfo("SetAsyncValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                                // Try SetAsyncValue(MAXIMUM_VALUE)
                                                if (IsGoodValue(switchMaximum))
                                                {
                                                    LogCallToDriver("SetAsyncValue",
                                                        $"  About to call SetAsyncValue({i}, {switchMaximum}), attempting to set the maximum permissible value");
                                                    SetAction($"SetAsyncValue maximum {switchMaximum}");
                                                    switchDevice.SetAsyncValue(i, switchMaximum); // Set switch to maximum

                                                    // Wait for the operation to complete
                                                    short i1 = i;
                                                    WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);

                                                    // Check GetSwitchValue returns the switch maximum value
                                                    if (getSwitchValueOk)
                                                    {
                                                        LogCallToDriver("SetAsyncValue",
                                                            $"  About to call GetSwitchValue({i}) method");
                                                        SetAction("GetSwitchValue");
                                                        getSwitchValue = switchDevice.GetSwitchValue(i);

                                                        // Wait for the operation to complete
                                                        WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);

                                                        switch (getSwitchValue)
                                                        {
                                                            case object o when Math.Abs(getSwitchValue - switchMaximum) < getSwitchValue * 0.0001:
                                                                LogOk("SetAsyncValue ", "  GetSwitchValue returned MAXIMUM_VALUE after SetAsyncValue(MAXIMUM_VALUE)");
                                                                setSwitchValueMaxOk = true;
                                                                break;

                                                            case object o when switchMaximum * 0.99 <= getSwitchValue && getSwitchValue <= switchMaximum * 1.01:
                                                                LogOk("SetAsyncValue ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetAsyncValue(MAXIMUM_VALUE)");
                                                                break;

                                                            default:
                                                                LogIssue("SetAsyncValue ",
                                                                    $"  GetSwitchValue did not return MAXIMUM_VALUE after SetAsyncValue(MAXIMUM_VALUE): {getSwitchValue}");
                                                                break;
                                                        }
                                                    }
                                                    else
                                                        LogInfo("SetAsyncValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");

                                                    // Now try a value above maximum
                                                    try
                                                    {
                                                        LogCallToDriver("SetAsyncValue",
                                                            $"  About to call SetAsyncValue({i}, {switchMaximum + 1.0}), attempting to set an invalid high value");
                                                        SetAction($"SetAsyncValue {switchMaximum + 1.0}");
                                                        switchDevice.SetAsyncValue(i, switchMaximum + 1.0);

                                                        // Wait for the operation to complete
                                                        WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);

                                                        LogIssue("SetAsyncValue",
                                                            $"Switch did not throw an exception when a value above SwitchMaximum was set: {switchMaximum}{1.0}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        HandleInvalidValueExceptionAsOk("SetAsyncValue", MemberType.Method, Required.Mandatory, ex,
                                                            $"when setting a value above SwitchMaximum was set: {switchMaximum}{1.0}",
                                                            $"  Switch threw an InvalidOperationException when a value above SwitchMaximum was set: {switchMaximum}{1.0}");
                                                    }
                                                    if (cancellationToken.IsCancellationRequested) return;
                                                }
                                                else
                                                    LogInfo("SetAsyncValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                                // Return to original state if possible,otherwise set to false
                                                if (getSwitchValueOk)
                                                {
                                                    LogCallToDriver("SetAsyncValue",
                                                        $"  About to call SetAsyncValue({i}, {getSwitchValueOriginal}), attempting to restore pre-test value");
                                                    SetAction($"SetAsyncValue to initial value {getSwitchValueOriginal}");
                                                    switchDevice.SetAsyncValue(i, getSwitchValueOriginal); // Return to the original state

                                                    // Wait for the operation to complete
                                                    short i1 = i;
                                                    WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);
                                                    LogOk("SetAsyncValue ", "  Switch has been reset to its original state");
                                                }
                                                else if (IsGoodValue(switchMinimum) & IsGoodValue(switchMaximum))
                                                {
                                                    LogCallToDriver("SetAsyncValue",
                                                        $"  About to call SetAsyncValue({i}, {(switchMaximum - switchMinimum) / 2.0}), attempting to set the value to its mid-point");
                                                    SetAction($"SetAsyncValue to midpoint {(switchMaximum - switchMinimum) / 2.0}");
                                                    switchDevice.SetAsyncValue(i, (switchMaximum - switchMinimum) / 2.0); // Return to the half way state
                                                    LogOk("SetAsyncValue ", "  Switch has been reset to half its range");

                                                    // Wait for the operation to complete
                                                    short i1 = i;
                                                    WaitWhile("Waiting for async operation to complete...", () => !switchDevice.StateChangeComplete(i1), 500, 10);
                                                }
                                                else
                                                    LogInfo("SetAsyncValue ", "Switch can not be returned to its default state because of issues with GetSwitchValue, GetSwitchMinimum or GetSwitchMaximum");
                                                if (cancellationToken.IsCancellationRequested) return;
                                            }
                                            catch (Exception ex)
                                            {
                                                LogDebug("SetAsyncValue ", $"EXCEPTION GENERATED - Switch can write: {switchCanWrite}, IsMethodNotImplementedException: {IsMethodNotImplementedException(ex)}, IsNotImplementedException: {IsNotImplementedException(ex)}, COM Access mechanic: {settings.ComConfiguration.ComAccessMechanic}, Device Technology: {settings.DeviceTechnology}");
                                                if (switchCanWrite)
                                                {
                                                    LogIssue("SetAsyncValue ", $"Exception: {ex.Message}");
                                                    LogDebug("SetAsyncValue ", $"Exception: {ex}");
                                                }
                                                else if (IsMethodNotImplementedException(ex))
                                                    LogOk("SetAsyncValue ", "  CanWrite is False and MethodNotImplementedException was thrown");
                                                // Determine whether we are testing a COM device using the cross platform DriverAccess module and, if so, test for the NotimplementedExceptionso that it returns.
                                                else if ((settings.ComConfiguration.ComAccessMechanic == ComAccessMechanic.DriverAccess) & (settings.DeviceTechnology == DeviceTechnology.COM))
                                                {
                                                    if (IsNotImplementedException(ex)) // Got a NotImplementedException, which is OK
                                                    {
                                                        LogOk("SetAsyncValue ", "  CanWrite is False and NotImplementedException was thrown by DriverAccess");
                                                    }
                                                    else // Got some other exception, which is an issue.
                                                    {
                                                        LogIssue("SetAsyncValue ", $"Exception: {ex.Message}");
                                                        LogDebug("SetAsyncValue ", $"Exception: {ex}");
                                                    }
                                                }
                                                else
                                                {
                                                    LogIssue("SetAsyncValue ", $"Exception: {ex.Message}");
                                                    LogDebug("SetAsyncValue ", $"Exception: {ex}");
                                                }
                                            }

                                        }
                                        else
                                            LogInfo("Async", "Skipping asynchropnous method tests because the switch cannot write or doesn't support asynchronous operation.");
                                    }

                                }
                                else
                                    LogInfo("SetSwitch", "  All write tests have been skipped because the \"Set Switches\" checkbox is unchecked");

                                LogNewLine();
                            }
                            catch (Exception ex)
                            {
                                if ((i >= 0) & (i < maxSwitch))
                                    LogIssue("GetSwitchName ", $"Mandatory method GetSwitchName threw an exception: {ex}");
                                else
                                    LogDebug("GetSwitchName ", $"Exception: {ex}");
                            }

                            if (cancellationToken.IsCancellationRequested)
                                return;
                        }
                        SetTest("");
                        SetAction("");
                        SetStatus("");
                    }
                    else
                        LogInfo("SwitchCheckMethods", "Skipping further tests as there is no valid value for MaxSwitch");
                    break;

                default:
                    throw new InvalidValueException($"Unsupported interface version: {GetInterfaceVersion()}");
            }
        }

        public override void CheckPerformance()
        {
            // MaxSwitch
            if (canReadMaxSwitch)
                SwitchPerformanceTest(SwitchPropertyMethod.MaxSwitch, "MaxSwitch");
            else
                LogInfo("MaxSwitch", "Test skipped as unable to read value");
            // GetSwitch
            if (System.Convert.ToBoolean(performanceGetSwitch))
                SwitchPerformanceTest(SwitchPropertyMethod.GetSwitch, "GetSwitch");
            else
                LogInfo("GetSwitch", "Test skipped as unable to read value");
            // GetSwitchName
            if (System.Convert.ToBoolean(performanceGetSwitchName))
                SwitchPerformanceTest(SwitchPropertyMethod.GetSwitchName, "GetSwitchName");
            else
                LogInfo("GetSwitchName", "Test skipped as unable to read value");
            SetTest("");
            SetAction("");
            SetStatus("");
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

                // Miscellaneous configuration
                if (!settings.SwitchEnableSet)
                    LogConfigurationAlert("Set switch tests were omitted due to Conform configuration.");

            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n:{ex}");
            }
        }

        #region Support code

        private short SwitchPropertyTestShort(SwitchPropertyMethod pType, string pName, short pMin, short pMax)
        {
            short returnValue = 0;

            try
            {
                LogCallToDriver(pName, $"About to get property {pName}");
                returnValue = 0;
                switch (pType)
                {
                    case SwitchPropertyMethod.MaxSwitch:
                        {
                            canReadMaxSwitch = false;
                            returnValue = switchDevice.MaxSwitch;
                            break;
                        }

                    default:
                        {
                            LogIssue(pName, $"SwitchPropertyTestInteger: Unknown test type - {pType}");
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < pMin // Lower than minimum value
                   :
                        {
                            LogIssue(pName, $"Invalid value: {returnValue}");
                            break;
                        }

                    case object _ when returnValue > pMax // Higher than maximum value
             :
                        {
                            LogIssue(pName, $"Invalid value: {returnValue}");
                            break;
                        }

                    default:
                        {
                            canReadMaxSwitch = true; // A valid value has been found
                            LogOk(pName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.Mandatory, ex, "");
            }
            return returnValue;
        }

        private void SwitchPerformanceTest(SwitchPropertyMethod pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime;
            string lString;
            bool lBoolean;
            double lRate;
            short lShort;
            SetTest("Performance Testing");
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
                        case SwitchPropertyMethod.MaxSwitch:
                            {
                                lShort = switchDevice.MaxSwitch;
                                break;
                            }

                        case SwitchPropertyMethod.GetSwitch:
                            {
                                lBoolean = switchDevice.GetSwitch((short)performanceGetSwitch);
                                break;
                            }

                        case SwitchPropertyMethod.GetSwitchName:
                            {
                                lString = switchDevice.GetSwitchName((short)performanceGetSwitchName);
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"SwitchPerformanceTest: Unknown test type {pType}");
                                break;
                            }
                    }

                    lElapsedTime = DateTime.Now.Subtract(lStartTime).TotalSeconds;
                    if (lElapsedTime > lLastElapsedTime + 1.0)
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
                    case object _ when lRate > 10.0:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case object _ when 2.0 <= lRate && lRate <= 10.0:
                        {
                            LogOk(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case object _ when 1.0 <= lRate && lRate <= 2.0:
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
        ///     ''' <param name="offset">Offset size as a percentage of switch step</param>
        ///     ''' <param name="switchMinimum">Switch minimum value</param>
        ///     ''' <param name="switchRange">Switch value range</param>
        ///     ''' <param name="switchStep">Size of each switch step</param>
        ///     ''' <remarks></remarks>
        private void TestSetSwitchValue(int i, double offset, double switchMinimum, double switchMaximum, double switchRange, double switchStep)
        {
            double lMultiStateStepSize, testValue2, lSwitchValue;
            MessageLevel msgLevel0, msgLevel1, msgLevel2;
            int lMultiStateNumberOfSteps;

            // Test the switch at the calculated positions
            try
            {
                if ((((switchMaximum - switchMinimum) / switchStep) + 1) >= NUMBER_OF_SWITCH_TEST_STATES)
                {
                    lMultiStateStepSize = switchRange / NUMBER_OF_SWITCH_TEST_STATES;
                    lMultiStateNumberOfSteps = NUMBER_OF_SWITCH_TEST_STATES;
                }
                else
                {
                    lMultiStateStepSize = switchStep; // Use the specified switch step size
                    lMultiStateNumberOfSteps = (int)Math.Floor(switchRange / switchStep);
                }
                LogDebug("MultiStateStepSize", lMultiStateStepSize.ToString());
                LogDebug("MultiStateNumberOfSteps", lMultiStateNumberOfSteps.ToString());

                if (offset == 0.0)
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

                LogInfo("SetSwitchValue",
                    $"  Testing with steps that are {offset:P0} offset from integer SwitchStep values");

                for (double testValue = switchMinimum; testValue <= switchMinimum + lMultiStateStepSize * lMultiStateNumberOfSteps; testValue += lMultiStateStepSize)
                {

                    // Round the test value to the nearest lowest switch step
                    if (testValue == switchMinimum)
                        testValue2 = switchMinimum + offset * switchStep;
                    else
                        testValue2 = (Math.Round((testValue - switchMinimum) / switchStep) * switchStep) + switchMinimum + offset * switchStep;

                    if (testValue2 <= switchMaximum)
                    {
                        LogCallToDriver("SetSwitchValue",
                            $"  About to call SetSwitchValue({i}, {testValue2}), attempting to set an intermediate value");
                        SetAction($"SetSwitchValue {testValue2}");
                        switchDevice.SetSwitchValue((short)i, testValue2); // Set the required switch value
                        WaitForWriteDelay($"SetSwitchValue {testValue2}"); LogCallToDriver("SetSwitchValue",
                            $"  About to call GetSwitchValue({i})");
                        SetAction("GetSwitchValue");
                        lSwitchValue = switchDevice.GetSwitchValue((short)i); // Read back the switch value 
                        WaitForReadDelay("GetSwitchValue");

                        switch (Math.Abs(lSwitchValue - testValue2))
                        {
                            case 0.0:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel0,
                                        $"  Set and read match: {testValue2}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.01:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel1,
                                        $"   Set/Read differ by up to 1% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.1:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"   Set/Read differ by 1-10% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.2:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 10-20% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.3:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 20-30% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.4:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 30-40% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.5:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 40-50% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.6:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 50-60% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.7:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 60-70% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.8:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 70-80% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.9:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 80-90% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 1.0:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by 90-100% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            default:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset,4:P0}", msgLevel2,
                                        $"  Set/Read differ by >100% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }
                        }

                        // Exit if the test has been cancelled
                        if (cancellationToken.IsCancellationRequested) return;

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
            const short lowTestValue = -1;
            const short highTestValue = 1;
            bool boolValue;
            string stringValue;
            double doubleValue;

            if (maxSwitch > 0)
            {

                // Try a value below 0
                try
                {
                    LogCallToDriver(method.ToString(),
                        $"About to call {method} with invalid low value for switch number: {lowTestValue} for ");
                    switch (method)
                    {
                        case SwitchMethod.CanWrite:
                            boolValue = switchDevice.CanWrite(lowTestValue);
                            break;

                        case SwitchMethod.GetSwitch:
                            boolValue = switchDevice.GetSwitch(lowTestValue);
                            break;

                        case SwitchMethod.GetSwitchDescription:
                            stringValue = switchDevice.GetSwitchDescription(lowTestValue);
                            break;

                        case SwitchMethod.GetSwitchName:
                            stringValue = switchDevice.GetSwitchName(lowTestValue);
                            break;

                        case SwitchMethod.GetSwitchValue:
                            doubleValue = switchDevice.GetSwitchValue(lowTestValue);
                            break;

                        case SwitchMethod.MaxSwitchValue:
                            doubleValue = switchDevice.MaxSwitchValue(lowTestValue);
                            break;

                        case SwitchMethod.MinSwitchValue:
                            doubleValue = switchDevice.MinSwitchValue(lowTestValue);
                            break;

                        case SwitchMethod.SetSwitch:
                            switchDevice.SetSwitch(lowTestValue, false);
                            break;

                        case SwitchMethod.SetSwitchName:
                            break;

                        case SwitchMethod.SetSwitchValue:
                            switchDevice.SetSwitchValue(lowTestValue, 0.0);
                            break;

                        case SwitchMethod.SwitchStep:
                            doubleValue = switchDevice.SwitchStep(lowTestValue);
                            break;

                        case SwitchMethod.CanAsync:
                            boolValue = switchDevice.CanAsync(lowTestValue);
                            break;

                        case SwitchMethod.CancelAsync:
                            switchDevice.CancelAsync(lowTestValue);
                            break;

                        case SwitchMethod.SetAsync:
                            switchDevice.SetAsync(lowTestValue, false);
                            break;

                        case SwitchMethod.SetAsyncValue:
                            switchDevice.SetAsyncValue(lowTestValue, 0.0);
                            break;

                        case SwitchMethod.StateChangeComplete:
                            boolValue = switchDevice.StateChangeComplete(lowTestValue);
                            break;

                        default:
                            {
                                LogIssue("CheckInaccessibleOutOfRange", $"Unknown value of SwitchMethod Enum: {method}");
                                break;
                            }
                    }
                    LogIssue("SwitchNumber",
                        $"Switch did not throw an exception when a switch ID below 0 was used in method: {method}");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("SwitchNumber", MemberType.Property, Required.MustBeImplemented, ex,
                        $"when a switch ID below 0 was used in method: {method}",
                        $"Switch device threw an InvalidOperationException when a switch ID below 0 was used in method: {method}");
                }

                // Try a value above MaxSwitch
                try
                {
                    LogCallToDriver(method.ToString(),
                        $"About to call {method} with invalid high value for switch number: {maxSwitch + highTestValue} for ");
                    switch (method)
                    {
                        case SwitchMethod.CanWrite:
                            boolValue = switchDevice.CanWrite((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.GetSwitch:
                            boolValue = switchDevice.GetSwitch((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.GetSwitchDescription:
                            stringValue = switchDevice.GetSwitchDescription((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.GetSwitchName:
                            stringValue = switchDevice.GetSwitchName((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.GetSwitchValue:
                            doubleValue = switchDevice.GetSwitchValue((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.MaxSwitchValue:
                            doubleValue = switchDevice.MaxSwitchValue((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.MinSwitchValue:
                            doubleValue = switchDevice.MinSwitchValue((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.SetSwitch:
                            switchDevice.SetSwitch((short)(maxSwitch + highTestValue), false);
                            break;

                        case SwitchMethod.SetSwitchName:
                            break;

                        case SwitchMethod.SetSwitchValue:
                            switchDevice.SetSwitchValue((short)(maxSwitch + highTestValue), 0.0);
                            break;

                        case SwitchMethod.SwitchStep:
                            doubleValue = switchDevice.SwitchStep((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.CanAsync:
                            boolValue = switchDevice.CanAsync((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.CancelAsync:
                            switchDevice.CancelAsync((short)(maxSwitch + highTestValue));
                            break;

                        case SwitchMethod.SetAsync:
                            switchDevice.SetAsync((short)(maxSwitch + highTestValue), false);
                            break;

                        case SwitchMethod.SetAsyncValue:
                            switchDevice.SetAsyncValue((short)(maxSwitch + highTestValue), 0.0);
                            break;

                        case SwitchMethod.StateChangeComplete:
                            boolValue = switchDevice.StateChangeComplete((short)(maxSwitch + highTestValue));
                            break;

                        default:
                            {
                                LogIssue("CheckInaccessibleOutOfRange", $"Unknown value of SwitchMethod Enum: {method}");
                                break;
                            }
                    }
                    LogIssue("SwitchNumber",
                        $"Switch did not throw an exception when a switch ID above MaxSwitch was used in method: {method}");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("SwitchNumber", MemberType.Property, Required.MustBeImplemented, ex,
                        $"when a switch ID above MaxSwitch was used in method: {method}",
                        $"Switch device threw an InvalidOperationException when a switch ID above MaxSwitch was used in method: {method}");
                }
            }
            else
                LogInfo("SwitchNumber", "Skipping range tests because MaxSwitch cannot be read");
        }
        private void WaitForReadDelay(string message)
        {
            if (switchReadDelay > 0)
            {
                SetAction($"{message} post read delay");
                WaitFor(switchReadDelay);
            }
            SetAction("");
        }

        private void WaitForWriteDelay(string message)
        {
            if (switchWriteDelay > 0)
            {
                SetAction($"{message} post write delay");
                WaitFor(switchWriteDelay);
            }
            SetAction("");
        }

        #endregion

    }
}

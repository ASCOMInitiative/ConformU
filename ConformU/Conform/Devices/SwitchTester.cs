using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConformU
{
    internal class SwitchTester : DeviceTesterBaseClass
    {
        private int mPerformanceGetSwitchName, mPerformanceGetSwitch;
        private short mMaxSwitch;
        private bool mCanReadMaxSwitch;
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
        private ISwitchV3 mSwitch;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose

        public SwitchTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;

            mPerformanceGetSwitchName = int.MinValue; // Initialise to silly values
            mPerformanceGetSwitch = int.MinValue;

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
                    if (mSwitch is not null) mSwitch.Dispose();
                    mSwitch = null;
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
                            GExInvalidValue1 = (int)0x80040405;
                            GExInvalidValue2 = (int)0x80040405;
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
                        mSwitch = new AlpacaSwitch(
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
                                mSwitch = new SwitchFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                mSwitch = new Switch(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                BaseClassDevice = mSwitch; // Assign the driver to the base class

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

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(mSwitch, DeviceTypes.Switch);
        }

        public override void CheckProperties()
        {
            // MaxSwitch - Mandatory
            switch (GetInterfaceVersion())
            {
                case 1: // Original Platform 5 switch interface, ISwitchV2 and ISwitchV3 have the same property
                case 2:
                case 3:
                    mMaxSwitch = SwitchPropertyTestShort(SwitchPropertyMethod.MaxSwitch, "MaxSwitch", 1, short.MaxValue);
                    break;

                default:
                    LogIssue("Switches", $"Unknown switch interface version: {GetInterfaceVersion()}");
                    break;
            }

            if (cancellationToken.IsCancellationRequested)
                return;
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

        public override void CheckMethods()
        {
            short i;
            bool lGetSwitch = false, lGetSwitchOriginal = false, lNewSwitchState, lGetSwitchOk, lSetSwitchOk, lGetSwitchValueOk, lSetSwitchValueMinOk, lSetSwitchValueMaxOk, lSwitchCanWrite;
            Exception lGetSwitchException, lGetSwitchValueException;
            double lGetSwitchValue, lGetSwitchValueOriginal = 0.0, lSwitchMinimum, lSwitchMaximum, lSwitchStep, lSwitchRange;
            string lSwitchName, lSwitchDescription;

            switch (GetInterfaceVersion())
            {
                case 1 // Platform 5 interface v1
               :
                    {
                        if (mCanReadMaxSwitch)
                        {

                            // Find valid GetSwitch values
                            for (i = 0; i <= System.Convert.ToInt16(mMaxSwitch - 1); i++)
                            {
                                SetTest($"Read/Write Switch {i}");
                                lGetSwitchOk = false;
                                lSetSwitchOk = false;
                                try // Read switch state
                                {
                                    LogCallToDriver($"GetSwitch {i}", $"About to call GetSwitch({i}) method");
                                    SetAction($"GetSwitch");
                                    lGetSwitchOriginal = mSwitch.GetSwitch(i);
                                    WaitForReadDelay("GetSwitch");

                                    LogOk($"GetSwitch {i}", $"Found switch, state: {lGetSwitchOriginal}");
                                    lGetSwitchOk = true;
                                    if (i > mMaxSwitch)
                                        LogIssue($"GetSwitch {i}", "Usable switch found above MaxSwitch!");
                                    else
                                        mPerformanceGetSwitch = i;// Save last good switch number for performance test
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
                                    SetAction($"SetSwitch {!lGetSwitchOriginal}");
                                    mSwitch.SetSwitch(i, !lGetSwitchOriginal); // Swap the switch state
                                    WaitForWriteDelay($"SetSwitch {!lGetSwitchOriginal}");

                                    lSetSwitchOk = true;
                                    if (lGetSwitchOk)
                                    {
                                        LogCallToDriver("SetSwitch", $"About to call GetSwitch({i}) method");
                                        SetAction($"GetSwitch");
                                        lNewSwitchState = mSwitch.GetSwitch(i); // Read the new switch state to confirm that value did change
                                        WaitForReadDelay("GetSwitch");

                                        if (lNewSwitchState == !lGetSwitchOriginal)
                                        {
                                            LogOk($"SetSwitch {i}", "Switch correctly changed state");
                                            LogCallToDriver("SetSwitch", "About to call SetSwitch method");
                                            SetAction($"SetSwitch {lGetSwitchOriginal}");
                                            mSwitch.SetSwitch(i, lGetSwitchOriginal); // Now put switch back to original state
                                            WaitForWriteDelay($"SetSwitch {lGetSwitchOriginal}");
                                        }
                                        else
                                            LogIssue($"SetSwitch {i}",
                                                $"Switch did not change state, currently it is {lNewSwitchState}");
                                    }
                                    else
                                        LogInfo($"SetSwitch {i}", "You have a write only switch!");
                                    if (i > mMaxSwitch)
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
                                    LogCallToDriver($"GetSwitchName {i}", string.Format("About to get switch name {0}", i));
                                    lSwitchName = mSwitch.GetSwitchName(i);
                                    if (lGetSwitchOk | lSetSwitchOk)
                                    {
                                        if (lSwitchName == "")
                                            LogInfo($"GetSwitchName {i}", "Switch name returns null string");
                                        else
                                        {
                                            mPerformanceGetSwitchName = i; // Save last good name index for performance test
                                            LogOk($"GetSwitchName {i}", $"Found switch, name:  {lSwitchName}");
                                        }
                                    }
                                    else if (lSwitchName == "")
                                        LogIssue($"GetSwitchName {i}", "Switch name returns null string but switch can neither read nor write!");
                                    else
                                        LogIssue($"GetSwitchName {i}",
                                            $"Found switch, name:  {lSwitchName} which can neither read nor write!");
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
                    }

                case 2 // Platform 6 interface version 2
         :
                    {
                        if (mCanReadMaxSwitch)
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

                            // Find valid GetSwitch values
                            LogDebug("GetSwitchName ", string.Format("Extended switch number test range: {0} - {1}", -extendedSwitchNumberTestRange, mMaxSwitch + extendedSwitchNumberTestRange - 1));
                            for (i = (short)-extendedSwitchNumberTestRange; i <= Convert.ToInt16(mMaxSwitch + extendedSwitchNumberTestRange - 1); i++)
                            {
                                SetStatus(i.ToString());

                                // Initialise status variables
                                lSetSwitchValueMinOk = false;
                                lSetSwitchValueMaxOk = false;
                                lSwitchStep = BAD_SWITCH_VALUE;

                                try // Read switch name to determine whether this is a valid switch
                                {
                                    LogCallToDriver("GetSwitchName", string.Format("About to get switch {0} name", i));
                                    lSwitchName = mSwitch.GetSwitchName(i);
                                    LogOk("GetSwitchName ", $"Found switch {i}");
                                    SetTest($"Testing switch {i}");

                                    // Test that the switch number is in the valid range of 0..MaxSwitch-1
                                    if (i > (mMaxSwitch - 1))
                                        LogIssue("GetSwitchName ", "Usable switch found above MaxSwitch - 1!");
                                    else if (i < 0)
                                        LogIssue("GetSwitchName ", "Usable switch found below 0!");
                                    else
                                        mPerformanceGetSwitch = i;// Save last good switch number for performance test

                                    LogOk("GetSwitchName ", $"  Name: {lSwitchName}");

                                    try // Read switch description
                                    {
                                        SetAction("Getting switch description"); LogCallToDriver("GetSwitchDescription", string.Format("  About to get switch {0} description", i));
                                        lSwitchDescription = mSwitch.GetSwitchDescription(i);
                                        LogOk("GetSwitchDescription ", $"  Description: {lSwitchDescription}");
                                    }
                                    catch (Exception ex)
                                    {
                                        LogIssue("GetSwitchDescription ",
                                            $"Mandatory parameter threw an exception: {ex.Message}");
                                    }

                                    try // Read switch minimum value
                                    {
                                        SetAction("Getting switch minimum value"); LogCallToDriver("MinSwitchValue", string.Format("  About to get switch {0} minimum value", i));
                                        lSwitchMinimum = mSwitch.MinSwitchValue(i);
                                        LogOk("MinSwitchValue ", $"  Minimum: {lSwitchMinimum}");
                                    }
                                    catch (Exception ex)
                                    {
                                        LogIssue("MinSwitchValue ",
                                            $"Mandatory parameter threw an exception: {ex.Message}");
                                        lSwitchMinimum = BAD_SWITCH_VALUE;
                                    }

                                    try // Read switch maximum value
                                    {
                                        SetAction("Getting switch maximum value"); LogCallToDriver("MaxSwitchValue", string.Format("  About to get switch {0} maximum value", i));
                                        lSwitchMaximum = mSwitch.MaxSwitchValue(i);

                                        if (IsGoodValue(lSwitchMinimum))
                                        {
                                            // Maximum value must be > Minimum value
                                            if (lSwitchMaximum > lSwitchMinimum)
                                            {
                                                LogOk("MaxSwitchValue ", $"  Maximum: {lSwitchMaximum}");
                                                lSwitchRange = lSwitchMaximum - lSwitchMinimum; // Calculate the range of values the switch might take
                                            }
                                            else
                                            {
                                                LogInfo("MaxSwitchValue ", $"  Maximum: {lSwitchMaximum}");
                                                LogIssue("MaxSwitchValue ", "MaxSwitchValue is less or equal to MinSwitchValue, it must be greater!");
                                                lSwitchRange = BAD_SWITCH_VALUE; // Special value because the maximum or minimum values are bad
                                            }
                                        }
                                        else
                                        {
                                            LogInfo("MaxSwitchValue ", $"  Maximum: {lSwitchMaximum}");
                                            lSwitchRange = BAD_SWITCH_VALUE;
                                            LogInfo("MaxSwitchValue ", "  Test that switch Maximum is greater than Minimum skipped because of an error reading the Minimum value.");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogIssue("MaxSwitchValue ",
                                            $"Mandatory parameter threw an exception: {ex.Message}");
                                        lSwitchMaximum = BAD_SWITCH_VALUE;
                                        lSwitchRange = BAD_SWITCH_VALUE;
                                    }

                                    try // Read switch step value
                                    {
                                        SetAction("Getting switch step size"); LogCallToDriver("SwitchStep", string.Format("  About to get switch {0} step size", i));
                                        lSwitchStep = mSwitch.SwitchStep(i);
                                        LogOk("SwitchStep ", $"  Step size: {lSwitchStep}");

                                        // Step must be greater than 0
                                        if (lSwitchStep > 0.0)
                                        {
                                            LogOk("SwitchStep ", "  Step size is greater than zero");

                                            // Step must be less than or equal to the range of possible values
                                            if (lSwitchStep <= lSwitchRange)
                                            {
                                                LogOk("SwitchStep ", "  Step size is less than the range of possible values");

                                                // Now check that the switch range is an integer multiple of the step size
                                                // Doubles are converted to the Decimal type (which has higher precision) in order to avoid unexpected outcomes from Mod due to rounding errors
                                                switch (Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)))
                                                {
                                                    case 0M:
                                                        {
                                                            LogOk("SwitchStep ", "  The switch range is an integer multiple of the step size.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) <= Convert.ToDecimal(lSwitchStep / 100):
                                                        {
                                                            LogInfo("SwitchStep ", "  The switch range is within 1% of being an integer multiple of the step size.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) <= Convert.ToDecimal(lSwitchStep / 10):
                                                        {
                                                            LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 10%.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) <= Convert.ToDecimal(lSwitchStep / 5):
                                                        {
                                                            LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 20%.");
                                                            break;
                                                        }

                                                    case object _ when 0M <= Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) && Math.Abs(decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)) <= Convert.ToDecimal(lSwitchStep / 2):
                                                        {
                                                            LogIssue("SwitchStep ", "  The switch range is not an integer multiple of the step size, but is within 50%.");
                                                            break;
                                                        }

                                                    default:
                                                        {
                                                            LogIssue("SwitchStep ",
                                                                $"The switch range must be an integer multiple of the step size. Remainder`: {decimal.Subtract(Convert.ToDecimal(lSwitchMaximum), Convert.ToDecimal(lSwitchMinimum)) % Convert.ToDecimal(lSwitchStep)}");
                                                            break;
                                                        }
                                                }
                                            }
                                            else
                                            {
                                                LogIssue("SwitchStep ", "Step size must be less than the range of possible values (MaxSwitchValue - MinSwitchValue");
                                                lSwitchStep = BAD_SWITCH_VALUE;
                                            }
                                        }
                                        else
                                        {
                                            LogIssue("SwitchStep ", "Step size must be greater than zero");
                                            lSwitchStep = BAD_SWITCH_VALUE;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogIssue("SwitchStep ", $"Mandatory parameter threw an exception: {ex.Message}");
                                    }

                                    LogDebug("SwitchMinimum ", lSwitchMinimum.ToString());
                                    LogDebug("SwitchMaximum ", lSwitchMaximum.ToString());
                                    LogDebug("SwitchStep ", lSwitchStep.ToString());
                                    LogDebug("SwitchRange ", lSwitchRange.ToString());

                                    try // Read CanWrite 
                                    {
                                        LogCallToDriver("CanWrite", string.Format("  About to get switch {0} CanWrite status", i));
                                        lSwitchCanWrite = mSwitch.CanWrite(i);
                                        LogOk("CanWrite ", $"  CanWrite: {lSwitchCanWrite}");
                                    }
                                    catch (Exception ex)
                                    {
                                        LogIssue("CanWrite ", $"Mandatory parameter threw an exception: {ex.Message}");
                                        LogInfo("CanWrite ", "Assuming that CanWrite is false");
                                        lSwitchCanWrite = false;
                                    }// Initialise to a default state

                                    // Access the Get Methods and record the outcomes
                                    try
                                    {
                                        SetAction($"GetSwitch"); LogCallToDriver("GetSwitch", string.Format("  About to call GetSwitch({0}) method", i));
                                        lGetSwitchOriginal = mSwitch.GetSwitch(i);
                                        WaitForReadDelay("GetSwitch");

                                        lGetSwitchOk = true;
                                        LogOk("GetSwitch ", $"  {lGetSwitchOriginal}");
                                    }
                                    catch (Exception ex)
                                    {
                                        lGetSwitchException = ex;
                                        LogDebug("GetSwitch ", $"Exception: {ex}");
                                        lGetSwitchOk = false;
                                    }

                                    try
                                    {
                                        SetAction($"GetSwitchValue"); LogCallToDriver("GetSwitchValue", string.Format("  About to call GetSwitchValue({0}) method", i));
                                        lGetSwitchValueOriginal = mSwitch.GetSwitchValue(i);
                                        WaitForReadDelay("GetSwitchValue");
                                        lGetSwitchValueOk = true;
                                        LogOk("GetSwitchValue ", $"  {lGetSwitchValueOriginal}");
                                    }
                                    catch (Exception ex)
                                    {
                                        lGetSwitchValueException = ex;
                                        LogDebug("GetSwitchValue ", $"Exception: {ex}");
                                        lGetSwitchValueOk = false;
                                    }

                                    // Now try to write to see which of these methods are available
                                    if (settings.SwitchEnableSet)
                                    {

                                        // Try to set the two boolean values through SetSwitch
                                        try
                                        {
                                            // Try SetSwitch(False)
                                            SetAction($"SetSwitch {i} False"); LogCallToDriver("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, false));
                                            SetAction("SetSwitch false");
                                            mSwitch.SetSwitch(i, false); // Set switch false
                                            WaitForWriteDelay($"SetSwitch False");

                                            // Check GetSwitch
                                            if (lGetSwitchOk)
                                            {
                                                LogCallToDriver("SetSwitch", string.Format("  About to call GetSwitch({0}) method", i));

                                                SetAction($"GetSwitch");
                                                if (mSwitch.GetSwitch(i) == false)
                                                    LogOk("SetSwitch ", "  GetSwitch returned False after SetSwitch(False)");
                                                else
                                                    LogIssue("SetSwitch ", "  GetSwitch returned True after SetSwitch(False)");
                                                WaitForReadDelay("GetSwitch");
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                            // Check GetSwitchValue returns the switch minimum value
                                            if (lGetSwitchValueOk & IsGoodValue(lSwitchMinimum))
                                            {
                                                LogCallToDriver("SetSwitch", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                SetAction($"GetSwitchValue");
                                                lGetSwitchValue = mSwitch.GetSwitchValue(i);
                                                WaitForReadDelay("GetSwitchValue");
                                                switch (lGetSwitchValue)
                                                {
                                                    case object _ when lGetSwitchValue == lSwitchMinimum:
                                                        {
                                                            LogOk("SetSwitch ", "  GetSwitchValue returned MINIMUM_VALUE after SetSwitch(False)");
                                                            break;
                                                        }

                                                    case object _ when lSwitchMinimum * 0.99 <= lGetSwitchValue && lGetSwitchValue <= lSwitchMinimum * 1.01:
                                                        {
                                                            LogOk("SetSwitch ", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitch(False)");
                                                            break;
                                                        }

                                                    default:
                                                        {
                                                            LogIssue("SetSwitch ",
                                                                $"  GetSwitchValue did not return MINIMUM_VALUE after SetSwitch(False): {lGetSwitchValue}");
                                                            break;
                                                        }
                                                }
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMinimum methods");
                                            if (cancellationToken.IsCancellationRequested) return;

                                            // Try SetSwitch(True)
                                            SetAction($"SetSwitch {i} True"); LogCallToDriver("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, true));
                                            SetAction($"SetSwitch true");
                                            mSwitch.SetSwitch(i, true); // Set switch true
                                            WaitForWriteDelay("SetSwitch true");

                                            // Check GetSwitch
                                            if (lGetSwitchOk)
                                            {
                                                LogCallToDriver("SetSwitch", string.Format("  About to call GetSwitch({0}) method", i));
                                                SetAction($"GetSwitch");
                                                if (mSwitch.GetSwitch(i) == true)
                                                    LogOk("SetSwitch ", "  GetSwitch read True after SetSwitch(True)");
                                                else
                                                    LogIssue("SetSwitch ", "  GetSwitch read False after SetSwitch(True)");
                                                WaitForReadDelay("GetSwitch");
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                            // Check GetSwitchValue returns the switch maximum value
                                            if (lGetSwitchValueOk & IsGoodValue(lSwitchMaximum))
                                            {
                                                LogCallToDriver("SetSwitch", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                SetAction($"GetSwitchValue");
                                                lGetSwitchValue = mSwitch.GetSwitchValue(i);
                                                WaitForReadDelay("GetSwitchValue");
                                                switch (lGetSwitchValue)
                                                {
                                                    case object _ when lGetSwitchValue == lSwitchMaximum:
                                                        {
                                                            LogOk("SetSwitch ", "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitch(True)");
                                                            break;
                                                        }

                                                    case object _ when lSwitchMaximum * 0.99 <= lSwitchMaximum && lSwitchMaximum <= lSwitchMaximum * 1.01:
                                                        {
                                                            LogOk("SetSwitch ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitch(True)");
                                                            break;
                                                        }

                                                    default:
                                                        {
                                                            LogIssue("SetSwitch ",
                                                                $"  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitch(True): {lGetSwitchValue}");
                                                            break;
                                                        }
                                                }
                                            }
                                            else
                                                LogInfo("SetSwitch ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue or GetSwitchMaximum methods");
                                            if (cancellationToken.IsCancellationRequested) return;

                                            // Return to original state if possible,otherwise set to false
                                            if (lGetSwitchOk)
                                            {
                                                LogCallToDriver("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, lGetSwitch));
                                                SetAction($"SetSwitch {lGetSwitch} to its original value");
                                                mSwitch.SetSwitch(i, lGetSwitch); // Return to the original state
                                                WaitForWriteDelay($"SetSwitch {lGetSwitch} to its original value");
                                            }
                                            else
                                            {
                                                LogCallToDriver("SetSwitch", string.Format("  About to call SetSwitch({0}, {1}) method", i, false));
                                                SetAction("SeSwitch false");
                                                mSwitch.SetSwitch(i, false); // Set to false
                                                WaitForWriteDelay("SetSwitch false");
                                            }

                                            lSetSwitchOk = true;
                                            LogDebug("SetSwitch ", "Set value OK");
                                            if (cancellationToken.IsCancellationRequested) return;

                                        }
                                        catch (Exception ex)
                                        {
                                            LogDebug("SetSwitch ", $"Exception generated - Switch can write: {lSwitchCanWrite}, IsMethodNotImplementedException: {IsMethodNotImplementedException(ex)}, IsNotImplementedException: {IsNotImplementedException(ex)}, COM Access mechanic: {settings.ComConfiguration.ComAccessMechanic}, Device Technology: {settings.DeviceTechnology}");
                                            if (lSwitchCanWrite)
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
                                            if (IsGoodValue(lSwitchMinimum))
                                            {
                                                LogCallToDriver("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the minimum permissible value", i, lSwitchMinimum));
                                                SetAction($"SetSwitchValue {lSwitchMinimum}");
                                                mSwitch.SetSwitchValue(i, lSwitchMinimum); // Set switch to minimum
                                                WaitForWriteDelay($"SetSwitchValue {lSwitchMinimum}");

                                                // Check GetSwitch
                                                if (lGetSwitchOk)
                                                {
                                                    LogCallToDriver("SetSwitchValue", string.Format("  About to call GetSwitch({0}) method", i));
                                                    SetAction("GetSwitch");
                                                    if (mSwitch.GetSwitch(i) == false)
                                                        LogOk("SetSwitchValue", "  GetSwitch returned False after SetSwitchValue(MINIMUM_VALUE)");
                                                    else
                                                        LogIssue("SetSwitchValue", "  GetSwitch returned True after SetSwitchValue(MINIMUM_VALUE)");
                                                    WaitForReadDelay("GetSwitch");
                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                                // Check GetSwitchValue returns the switch minimum value
                                                if (lGetSwitchValueOk)
                                                {
                                                    LogCallToDriver("SetSwitchValue", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                    SetAction("GetSwitchValue");
                                                    lGetSwitchValue = mSwitch.GetSwitchValue(i);
                                                    WaitForReadDelay("GetSwitchValue");

                                                    switch (lGetSwitchValue)
                                                    {
                                                        case object _ when lGetSwitchValue == lSwitchMinimum:
                                                            {
                                                                LogOk("SetSwitchValue", "  GetSwitchValue returned MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)");
                                                                lSetSwitchValueMinOk = true;
                                                                break;
                                                            }

                                                        case object _ when lSwitchMinimum * 0.99 <= lGetSwitchValue && lGetSwitchValue <= lSwitchMinimum * 1.01:
                                                            {
                                                                LogOk("SetSwitchValue", "  GetSwitchValue returned a value within 1% of MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE)");
                                                                lSetSwitchValueMinOk = true;
                                                                break;
                                                            }

                                                        default:
                                                            {
                                                                LogIssue("SetSwitchValue",
                                                                    $"  GetSwitchValue did not return MINIMUM_VALUE after SetSwitchValue(MINIMUM_VALUE): {lGetSwitchValue}");
                                                                break;
                                                            }
                                                    }
                                                    if (lGetSwitchValue == lSwitchMinimum)
                                                    {
                                                    }
                                                    else
                                                    {
                                                    }
                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");
                                                if (cancellationToken.IsCancellationRequested) return;

                                                // Now try a value below minimum
                                                try
                                                {
                                                    LogCallToDriver("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an invalid low value", i, lSwitchMinimum - 1.0));
                                                    SetAction($"SetSwitchValue {lSwitchMinimum - 1.0}");
                                                    mSwitch.SetSwitchValue(i, lSwitchMinimum - 1.0);
                                                    WaitForWriteDelay($"SetSwitchValue {lSwitchMinimum - 1.0}");

                                                    LogIssue("SetSwitchValue",
                                                        $"Switch did not throw an exception when a value below SwitchMinimum was set: {(lSwitchMinimum - 1.0)}");
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleInvalidValueExceptionAsOk("SetSwitchValue", MemberType.Method, Required.Mandatory, ex,
                                                        $"when setting a value below SwitchMinimum - {(lSwitchMinimum - 1.0)}",
                                                        $"  Switch threw an InvalidOperationException when a value below SwitchMinimum was set: {(lSwitchMinimum - 1.0)}");
                                                }
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                            // Try SetSwitchValue(MAXIMUM_VALUE)
                                            if (IsGoodValue(lSwitchMaximum))
                                            {
                                                LogCallToDriver("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the maximum permissible value", i, lSwitchMaximum));
                                                SetAction($"SetSwitchValue maximum {lSwitchMaximum}");
                                                mSwitch.SetSwitchValue(i, lSwitchMaximum); // Set switch to maximum
                                                WaitForWriteDelay($"SetSwitchValue maximum {lSwitchMaximum}");

                                                // Check GetSwitch
                                                if (lGetSwitchOk)
                                                {
                                                    LogCallToDriver("SetSwitchValue", string.Format("  About to call GetSwitch({0}) method", i));
                                                    SetAction("GetSwitch");
                                                    if (mSwitch.GetSwitch(i) == true)
                                                        LogOk("SetSwitchValue ", "  GetSwitch returned True after SetSwitchValue(MAXIMUM_VALUE)");
                                                    else
                                                        LogIssue("SetSwitchValue ", "  GetSwitch returned False after SetSwitchValue(MAXIMUM_VALUE)");
                                                    WaitForReadDelay("GetSwitch");

                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitch confirmation because of an issue with the GetSwitch method");

                                                // Check GetSwitchValue returns the switch maximum value
                                                if (lGetSwitchValueOk)
                                                {
                                                    LogCallToDriver("SetSwitchValue", string.Format("  About to call GetSwitchValue({0}) method", i));
                                                    SetAction("GetSwitchValue");
                                                    lGetSwitchValue = mSwitch.GetSwitchValue(i);
                                                    WaitForReadDelay("GetSwitchValue");

                                                    switch (lGetSwitchValue)
                                                    {
                                                        case object _ when lGetSwitchValue == lSwitchMaximum:
                                                            {
                                                                LogOk("SetSwitchValue ", "  GetSwitchValue returned MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)");
                                                                lSetSwitchValueMaxOk = true;
                                                                break;
                                                            }

                                                        case object _ when lSwitchMaximum * 0.99 <= lGetSwitchValue && lGetSwitchValue <= lSwitchMaximum * 1.01:
                                                            {
                                                                LogOk("SetSwitchValue ", "  GetSwitchValue returned a value within 1% of MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE)");
                                                                break;
                                                            }

                                                        default:
                                                            {
                                                                LogIssue("SetSwitchValue ",
                                                                    $"  GetSwitchValue did not return MAXIMUM_VALUE after SetSwitchValue(MAXIMUM_VALUE): {lGetSwitchValue}");
                                                                break;
                                                            }
                                                    }
                                                    if (lGetSwitchValue == lSwitchMaximum)
                                                    {
                                                    }
                                                    else
                                                    {
                                                    }
                                                }
                                                else
                                                    LogInfo("SetSwitchValue ", "  Skipping GetSwitchValue confirmation because of an issue with the GetSwitchValue method");

                                                // Now try a value above maximum
                                                try
                                                {
                                                    LogCallToDriver("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an invalid high value", i, lSwitchMaximum + 1.0));
                                                    SetAction($"SetSwitchValue {lSwitchMaximum + 1.0}");
                                                    mSwitch.SetSwitchValue(i, lSwitchMaximum + 1.0);
                                                    WaitForWriteDelay($"SetSwitchValue {lSwitchMaximum + 1.0}");

                                                    LogIssue("SetSwitchValue",
                                                        $"Switch did not throw an exception when a value above SwitchMaximum was set: {lSwitchMaximum}{1.0}");
                                                }
                                                catch (Exception ex)
                                                {
                                                    HandleInvalidValueExceptionAsOk("SetSwitchValue", MemberType.Method, Required.Mandatory, ex,
                                                        $"when setting a value above SwitchMaximum was set: {lSwitchMaximum}{1.0}",
                                                        $"  Switch threw an InvalidOperationException when a value above SwitchMaximum was set: {lSwitchMaximum}{1.0}");
                                                }
                                                if (cancellationToken.IsCancellationRequested) return;
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "  Skipping test because of an issue with retrieving the switch minimum value through GetSwitchMinimim");

                                            // Test some positions of the multi-state switch between the minimum and maximum values
                                            if (lGetSwitchValueOk & lSetSwitchValueMinOk & lSetSwitchValueMaxOk & IsGoodValue(lSwitchRange) & IsGoodValue(lSwitchStep))
                                            {
                                                TestSetSwitchValue(i, 0.0, lSwitchMinimum, lSwitchMaximum, lSwitchRange, lSwitchStep); if (cancellationToken.IsCancellationRequested) return;
                                                TestSetSwitchValue(i, 0.25, lSwitchMinimum, lSwitchMaximum, lSwitchRange, lSwitchStep); if (cancellationToken.IsCancellationRequested) return;
                                                TestSetSwitchValue(i, 0.5, lSwitchMinimum, lSwitchMaximum, lSwitchRange, lSwitchStep); if (cancellationToken.IsCancellationRequested) return;
                                                TestSetSwitchValue(i, 0.75, lSwitchMinimum, lSwitchMaximum, lSwitchRange, lSwitchStep); if (cancellationToken.IsCancellationRequested) return;
                                            }
                                            else
                                            {
                                                LogInfo("SetSwitchValue ", "Skipping multi state tests because of earlier errors");
                                                LogDebug("GetSwitchValueOK ", lGetSwitchValueOk.ToString());
                                                LogDebug("SetSwitchValueMinOK ", lSetSwitchValueMinOk.ToString());
                                                LogDebug("SetSwitchValueMaxOK ", lSetSwitchValueMaxOk.ToString());
                                                LogDebug("SwitchRange ", lSwitchRange.ToString());
                                                LogDebug("SwitchStep ", lSwitchStep.ToString());
                                            }

                                            // Return to original state if possible,otherwise set to false
                                            if (lGetSwitchValueOk)
                                            {
                                                LogCallToDriver("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to restore pre-test value", i, lGetSwitchValueOriginal));
                                                SetAction($"SetSwitchValue to initial value {lGetSwitchValueOriginal}");
                                                mSwitch.SetSwitchValue(i, lGetSwitchValueOriginal); // Return to the original state
                                                LogOk("SetSwitchValue ", "  Switch has been reset to its original state");
                                                WaitForWriteDelay($"SetSwitchValue to initial value {lGetSwitchValueOriginal}");
                                            }
                                            else if (IsGoodValue(lSwitchMinimum) & IsGoodValue(lSwitchMaximum))
                                            {
                                                LogCallToDriver("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set the value to its mid-point", i, (lSwitchMaximum - lSwitchMinimum) / 2.0));
                                                SetAction($"SetSwitchValue to midpoint {(lSwitchMaximum - lSwitchMinimum) / 2.0}");
                                                mSwitch.SetSwitchValue(i, (lSwitchMaximum - lSwitchMinimum) / 2.0); // Return to the half way state
                                                LogOk("SetSwitchValue ", "  Switch has been reset to half its range");
                                                WaitForWriteDelay($"SetSwitchValue to midpoint {(lSwitchMaximum - lSwitchMinimum) / 2.0}");
                                            }
                                            else
                                                LogInfo("SetSwitchValue ", "Switch can not be returned to its default state because of issues with GetSwitchValue, GetSwitchMinimum or GetSwitchMaximum");
                                            if (cancellationToken.IsCancellationRequested) return;
                                        }
                                        catch (Exception ex)
                                        {
                                            LogDebug("SetSwitchValue ", $"EXCEPTION GENERATED - Switch can write: {lSwitchCanWrite}, IsMethodNotImplementedException: {IsMethodNotImplementedException(ex)}, IsNotImplementedException: {IsNotImplementedException(ex)}, COM Access mechanic: {settings.ComConfiguration.ComAccessMechanic}, Device Technology: {settings.DeviceTechnology}");
                                            if (lSwitchCanWrite)
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
                                    }
                                    else
                                        LogInfo("SetSwitch", "  All write tests have been skipped because the \"Set Switches\" checkbox is unchecked");

                                    LogNewLine();
                                }
                                catch (Exception ex)
                                {
                                    if ((i >= 0) & (i < mMaxSwitch))
                                        LogIssue("GetSwitchName ",
                                            $"Mandatory method GetSwitchName threw an exception: {ex}");
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
                    }
            }
        }

        public override void CheckPerformance()
        {
            // MaxSwitch
            if (mCanReadMaxSwitch)
                SwitchPerformanceTest(SwitchPropertyMethod.MaxSwitch, "MaxSwitch");
            else
                LogInfo("MaxSwitch", "Test skipped as unable to read value");
            // GetSwitch
            if (System.Convert.ToBoolean(mPerformanceGetSwitch))
                SwitchPerformanceTest(SwitchPropertyMethod.GetSwitch, "GetSwitch");
            else
                LogInfo("GetSwitch", "Test skipped as unable to read value");
            // GetSwitchName
            if (System.Convert.ToBoolean(mPerformanceGetSwitchName))
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

        private short SwitchPropertyTestShort(SwitchPropertyMethod pType, string pName, short pMin, short pMax)
        {
            short returnValue = 0;

            try
            {
                LogCallToDriver(pName, string.Format("About to get property {0}", pName));
                returnValue = 0;
                switch (pType)
                {
                    case SwitchPropertyMethod.MaxSwitch:
                        {
                            mCanReadMaxSwitch = false;
                            returnValue = mSwitch.MaxSwitch;
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
                            mCanReadMaxSwitch = true; // A valid value has been found
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
                                lShort = mSwitch.MaxSwitch;
                                break;
                            }

                        case SwitchPropertyMethod.GetSwitch:
                            {
                                lBoolean = mSwitch.GetSwitch((short)mPerformanceGetSwitch);
                                break;
                            }

                        case SwitchPropertyMethod.GetSwitchName:
                            {
                                lString = mSwitch.GetSwitchName((short)mPerformanceGetSwitchName);
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
                        LogCallToDriver("SetSwitchValue", string.Format("  About to call SetSwitchValue({0}, {1}), attempting to set an intermediate value", i, testValue2));
                        SetAction($"SetSwitchValue {testValue2}");
                        mSwitch.SetSwitchValue((short)i, testValue2); // Set the required switch value
                        WaitForWriteDelay($"SetSwitchValue {testValue2}"); LogCallToDriver("SetSwitchValue", string.Format("  About to call GetSwitchValue({0})", i));
                        SetAction("GetSwitchValue");
                        lSwitchValue = mSwitch.GetSwitchValue((short)i); // Read back the switch value 
                        WaitForReadDelay("GetSwitchValue");

                        switch (Math.Abs(lSwitchValue - testValue2))
                        {
                            case 0.0:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel0,
                                        $"  Set and read match: {testValue2}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.01:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel1,
                                        $"   Set/Read differ by up to 1% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.1:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"   Set/Read differ by 1-10% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.2:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 10-20% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.3:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 20-30% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.4:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 30-40% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.5:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 40-50% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.6:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 50-60% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.7:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 60-70% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.8:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 70-80% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 0.9:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 80-90% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            case object _ when 0.0 <= Math.Abs(lSwitchValue - testValue2) && Math.Abs(lSwitchValue - testValue2) <= switchStep * 1.0:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
                                        $"  Set/Read differ by 90-100% of SwitchStep. Set: {testValue2}, Read: {lSwitchValue}");
                                    break;
                                }

                            default:
                                {
                                    LogMsg($"SetSwitchValue Offset: {offset.ToString("P0").PadLeft(4)}", msgLevel2,
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

            if (mMaxSwitch > 0)
            {

                // Try a value below 0
                try
                {
                    LogCallToDriver(method.ToString(), string.Format("About to call {0} with invalid low value for switch number: {1} for ", method.ToString(), lowTestValue));
                    switch (method)
                    {
                        case SwitchMethod.CanWrite:
                            {
                                boolValue = mSwitch.CanWrite(lowTestValue);
                                break;
                            }

                        case SwitchMethod.GetSwitch:
                            {
                                boolValue = mSwitch.GetSwitch(lowTestValue);
                                break;
                            }

                        case SwitchMethod.GetSwitchDescription:
                            {
                                stringValue = mSwitch.GetSwitchDescription(lowTestValue);
                                break;
                            }

                        case SwitchMethod.GetSwitchName:
                            {
                                stringValue = mSwitch.GetSwitchName(lowTestValue);
                                break;
                            }

                        case SwitchMethod.GetSwitchValue:
                            {
                                doubleValue = mSwitch.GetSwitchValue(lowTestValue);
                                break;
                            }

                        case SwitchMethod.MaxSwitchValue:
                            {
                                doubleValue = mSwitch.MaxSwitchValue(lowTestValue);
                                break;
                            }

                        case SwitchMethod.MinSwitchValue:
                            {
                                doubleValue = mSwitch.MinSwitchValue(lowTestValue);
                                break;
                            }

                        case SwitchMethod.SetSwitch:
                            {
                                mSwitch.SetSwitch(lowTestValue, false);
                                break;
                            }

                        case SwitchMethod.SetSwitchName:
                            {
                                break;
                            }

                        case SwitchMethod.SetSwitchValue:
                            {
                                mSwitch.SetSwitchValue(lowTestValue, 0.0);
                                break;
                            }

                        case SwitchMethod.SwitchStep:
                            {
                                doubleValue = mSwitch.SwitchStep(lowTestValue);
                                break;
                            }

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
                    LogCallToDriver(method.ToString(), string.Format("About to call {0} with invalid high value for switch number: {1} for ", method.ToString(), mMaxSwitch + highTestValue));
                    switch (method)
                    {
                        case SwitchMethod.CanWrite:
                            {
                                boolValue = mSwitch.CanWrite((short)(mMaxSwitch + highTestValue));
                                break;
                            }

                        case SwitchMethod.GetSwitch:
                            {
                                boolValue = mSwitch.GetSwitch((short)(mMaxSwitch + highTestValue));
                                break;
                            }

                        case SwitchMethod.GetSwitchDescription:
                            {
                                stringValue = mSwitch.GetSwitchDescription((short)(mMaxSwitch + highTestValue));
                                break;
                            }

                        case SwitchMethod.GetSwitchName:
                            {
                                stringValue = mSwitch.GetSwitchName((short)(mMaxSwitch + highTestValue));
                                break;
                            }

                        case SwitchMethod.GetSwitchValue:
                            {
                                doubleValue = mSwitch.GetSwitchValue((short)(mMaxSwitch + highTestValue));
                                break;
                            }

                        case SwitchMethod.MaxSwitchValue:
                            {
                                doubleValue = mSwitch.MaxSwitchValue((short)(mMaxSwitch + highTestValue));
                                break;
                            }

                        case SwitchMethod.MinSwitchValue:
                            {
                                doubleValue = mSwitch.MinSwitchValue((short)(mMaxSwitch + highTestValue));
                                break;
                            }

                        case SwitchMethod.SetSwitch:
                            {
                                mSwitch.SetSwitch((short)(mMaxSwitch + highTestValue), false);
                                break;
                            }

                        case SwitchMethod.SetSwitchName:
                            {
                                break;
                            }

                        case SwitchMethod.SetSwitchValue:
                            {
                                mSwitch.SetSwitchValue((short)(mMaxSwitch + highTestValue), 0.0);
                                break;
                            }

                        case SwitchMethod.SwitchStep:
                            {
                                doubleValue = mSwitch.SwitchStep((short)(mMaxSwitch + highTestValue));
                                break;
                            }

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
    }
}

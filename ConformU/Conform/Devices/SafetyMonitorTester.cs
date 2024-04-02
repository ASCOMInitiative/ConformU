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

    internal class SafetyMonitorTester : DeviceTesterBaseClass
    {
        private bool mIsSafe;
        private ISafetyMonitorV3 mSafetyMonitor;

        // Helper variables
        internal static Utilities GUtil;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        private enum RequiredProperty
        {
            PropIsSafe
        }
        private enum PerformanceProperty
        {
            PropIsSafe
        }

        #region New and Dispose

        public SafetyMonitorTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, false, false, true, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            GUtil = new();
            //g_settings.MessageLevel = MessageLevel.Debug;

            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        protected override void Dispose(bool disposing)
        {
            LogDebug("Dispose", $"Disposing of test device: {disposing} {disposedValue}");
            if (!disposedValue)
            {
                if (disposing)
                {
                    mSafetyMonitor?.Dispose();
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
                switch (settings.ComDevice.ProgId ?? "")
                {
                    default:
                        {
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040405;
                            ExInvalidValue2 = (int)0x80040405;
                            ExInvalidValue3 = (int)0x80040405;
                            ExInvalidValue4 = (int)0x80040405;
                            ExInvalidValue5 = (int)0x80040405;
                            ExInvalidValue6 = (int)0x80040405;
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
                        mSafetyMonitor = new AlpacaSafetyMonitor(
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
                                mSafetyMonitor = new SafetyMonitorFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                mSafetyMonitor = new SafetyMonitor(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(mSafetyMonitor, DeviceTypes.SafetyMonitor); // Assign the driver to the base class

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

        public override void PreConnectChecks()
        {
            // Confirm that key properties are false when not connected
            try
            {
                LogCallToDriver("IsSafe", "About to get IsSafe property");
                mIsSafe = mSafetyMonitor.IsSafe;
                if (!mIsSafe)
                    LogOk("IsSafe", "Reports false before connection");
                else
                    LogIssue("IsSafe", "Reports true before connection rather than false");
            }
            catch (Exception ex)
            {
                LogIssue("IsSafe",
                    $"Cannot confirm that IsSafe is false before connection because it threw an exception: {ex.Message}");
            }
        }

        public override void CheckProperties()
        {
            RequiredPropertiesTest(RequiredProperty.PropIsSafe, "IsSafe");
        }
        public override void CheckPerformance()
        {
            SetTest("Performance");
            PerformanceTest(PerformanceProperty.PropIsSafe, "IsSafe");
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

            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n:{ex}");
            }
        }

        private void RequiredPropertiesTest(RequiredProperty pType, string pName)
        {
            try
            {
                switch (pType)
                {
                    case RequiredProperty.PropIsSafe:
                        {
                            LogCallToDriver("IsSafe", "About to get IsSafe property");
                            mIsSafe = TimeFuncNoParams("IsSafe", () => mSafetyMonitor.IsSafe);
                            LogOk(pName, mIsSafe.ToString());
                            break;
                        }

                    default:
                        {
                            LogIssue(pName, $"RequiredPropertiesTest: Unknown test type {pType}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.Mandatory, ex, "");
            }
        }
        private void PerformanceTest(PerformanceProperty pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime, lRate;
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
                        case PerformanceProperty.PropIsSafe:
                            {
                                mIsSafe = mSafetyMonitor.IsSafe;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"PerformanceTest: Unknown test type {pType}");
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
                LogInfo(pName, $"Unable to complete test: {ex}");
            }
        }
    }

}

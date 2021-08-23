using System;
using ASCOM.Standard.Interfaces;
using System.Threading;
using ASCOM.Standard.Utilities;
using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.COM.DriverAccess;

namespace ConformU
{

    internal class SafetyMonitorTester : DeviceTesterBaseClass
    {
        private bool m_IsSafe;
        private ISafetyMonitor m_SafetyMonitor;

        // Helper variables
        internal static Utilities g_Util;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        private enum RequiredProperty
        {
            propIsSafe
        }
        private enum PerformanceProperty
        {
            propIsSafe
        }

        #region New and Dispose

        public SafetyMonitorTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, false, false, true, true, false, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            g_Util = new();
            //g_settings.MessageLevel = MessageLevel.Debug;

            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        protected override void Dispose(bool disposing)
        {
            LogMsg("Dispose", MessageLevel.msgDebug, "Disposing of test device: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (m_SafetyMonitor is not null) m_SafetyMonitor.Dispose();
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
                switch (settings.ComDevice.ProgId ?? "")
                {
                    default:
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040405;
                            g_ExInvalidValue2 = (int)0x80040405;
                            g_ExInvalidValue3 = (int)0x80040405;
                            g_ExInvalidValue4 = (int)0x80040405;
                            g_ExInvalidValue5 = (int)0x80040405;
                            g_ExInvalidValue6 = (int)0x80040405;
                            g_ExNotSet1 = (int)0x80040403;
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
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_SafetyMonitor = new AlpacaSafetyMonitor(settings.AlpacaConfiguration.AccessServiceType.ToString(),
                            settings.AlpacaDevice.IpAddress,
                            settings.AlpacaDevice.IpPort,
                            settings.AlpacaDevice.AlpacaDeviceNumber,
                            settings.StrictCasing,
                            settings.DisplayMethodCalls ? logger : null);
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_SafetyMonitor = new SafetyMonitorFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                m_SafetyMonitor = new SafetyMonitor(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");
                baseClassDevice = m_SafetyMonitor; // Assign the driver to the base class

                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
                g_Stop = false;
            }
            catch (Exception ex)
            {
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Exception thrown: " + ex.Message);
                throw; // Re throw exception 
            }

            if (g_Stop) WaitFor(200);

        }
        public override void PreConnectChecks()
        {
            // Confirm that key properties are false when not connected
            try
            {
                LogCallToDriver("IsSafe", "About to get IsSafe property");
                m_IsSafe = m_SafetyMonitor.IsSafe;
                if (!m_IsSafe)
                    LogMsg("IsSafe", MessageLevel.msgOK, "Reports false before connection");
                else
                    LogMsg("IsSafe", MessageLevel.msgIssue, "Reports true before connection rather than false");
            }
            catch (Exception ex)
            {
                LogMsg("IsSafe", MessageLevel.msgError, "Cannot confirm that IsSafe is false before connection because it threw an exception: " + ex.Message);
            }
        }
        public override bool Connected
        {
            get
            {
                LogCallToDriver("Connected", "About to get Connected property");
                return m_SafetyMonitor.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                m_SafetyMonitor.Connected = value;
            }
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_SafetyMonitor, DeviceType.SafetyMonitor);
        }

        public override void CheckProperties()
        {
            RequiredPropertiesTest(RequiredProperty.propIsSafe, "IsSafe");
        }
        public override void CheckPerformance()
        {
            Status(StatusType.staTest, "Performance");
            PerformanceTest(PerformanceProperty.propIsSafe, "IsSafe");
            Status(StatusType.staTest, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
        }

        private void RequiredPropertiesTest(RequiredProperty p_Type, string p_Name)
        {
            try
            {
                switch (p_Type)
                {
                    case RequiredProperty.propIsSafe:
                        {
                            m_IsSafe = m_SafetyMonitor.IsSafe;
                            LogCallToDriver("IsSafe", "About to get IsSafe property");
                            LogMsg(p_Name, MessageLevel.msgOK, m_IsSafe.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "RequiredPropertiesTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
            }
        }
        private void PerformanceTest(PerformanceProperty p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate;
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
                        case PerformanceProperty.propIsSafe:
                            {
                                m_IsSafe = m_SafetyMonitor.IsSafe;
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "PerformanceTest: Unknown test type " + p_Type.ToString());
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
                while (!(l_ElapsedTime > PERF_LOOP_TIME));
                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case object _ when l_Rate > 10.0:
                        {
                            LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= l_Rate && l_Rate <= 10.0:
                        {
                            LogMsg(p_Name, MessageLevel.msgOK, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= l_Rate && l_Rate <= 2.0:
                        {
                            LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgInfo, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " + ex.ToString());
            }
        }
    }

}

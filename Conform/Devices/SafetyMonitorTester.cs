using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
//using Microsoft.VisualBasic;
using ASCOM.Standard.Interfaces;
using static ConformU.ConformConstants;
using System.Threading;
using ASCOM.Standard.Utilities;

namespace ConformU
{

    internal class SafetyMonitorTester : DeviceTesterBaseClass
    {
        private bool m_CanIsGood, m_CanEmergencyShutdown;
        private bool m_IsSafe, m_IsGood;
        private string m_Description, m_DriverInfo, m_DriverVersion;

        /* TODO ERROR: Skipped IfDirectiveTrivia */
        /* TODO ERROR: Skipped DisabledTextTrivia */
        /* TODO ERROR: Skipped ElseDirectiveTrivia */
        private ISafetyMonitor m_SafetyMonitor;
        /* TODO ERROR: Skipped EndIfDirectiveTrivia */

        // Helper variables
        internal static Utilities g_Util;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ILogger logger;

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

        public SafetyMonitorTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, true, true, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            g_Util = new();
            //g_settings.MessageLevel = MessageLevel.Debug;

            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        protected override void Dispose(bool disposing)
        {
            LogMsg("Dispose", MessageLevel.Debug, "Disposing of Telescope driver: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (true) // Should be True but make False to stop Conform from cleanly dropping the telescope object (useful for retaining driver in memory to change flags)
                    {
                        if (telescopeDevice is not null) telescopeDevice.Dispose();
                        telescopeDevice = null;
                        GC.Collect();
                    }
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion


        public override void CheckInitialise()
        {
            // Set the error type numbers according to the standards adopted by individual authors.
            // Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
            // messages to driver authors!
            unchecked
            {
                switch (g_SafetyMonitorProgID)
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
            base.CheckInitialise(g_SafetyMonitorProgID);
        }

        public override void CreateDevice()
        {
            /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
            if (g_Settings.UseDriverAccess)
            {
                LogMsg("Conform", MessageLevel.Always, "is using ASCOM.DriverAccess.SafetyMonitor to get a SafetyMonitor object");
                m_SafetyMonitor = new ASCOM.DriverAccess.SafetyMonitor(g_SafetyMonitorProgID);
                LogMsg("CreateDevice", MessageLevel.Debug, "Successfully created driver");
            }
            else
            {
                m_SafetyMonitor = CreateObject(g_SafetyMonitorProgID);
                LogMsg("CreateDevice", MessageLevel.Debug, "Successfully created driver");
            }
            /* TODO ERROR: Skipped EndIfDirectiveTrivia */
            g_Stop = false; // connected OK so clear stop flag to allow other tests to run
        }
        public override void PreConnectChecks()
        {
            // Confirm that key properties are false when not connected
            try
            {
                LogCallToDriver("IsSafe", "About to get IsSafe property");
                m_IsSafe = m_SafetyMonitor.IsSafe;
                if (!m_IsSafe)
                    LogMsg("IsSafe", MessageLevel.OK, "Reports false before connection");
                else
                    LogMsg("IsSafe", MessageLevel.Issue, "Reports true before connection rather than false");
            }
            catch (Exception ex)
            {
                LogMsg("IsSafe", MessageLevel.Error, "Cannot confirm that IsSafe is false before connection because it threw an exception: " + ex.Message);
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
                            LogMsg(p_Name, MessageLevel.OK, m_IsSafe.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.Error, "RequiredPropertiesTest: Unknown test type " + p_Type.ToString());
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
                                LogMsg(p_Name, MessageLevel.Error, "PerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
                    {
                        Status(StatusType.staStatus, l_Count + " transactions in " + l_ElapsedTime.ToString( "0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        if (TestStop())
                            return;
                    }
                }
                while (!(l_ElapsedTime > PERF_LOOP_TIME));
                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case object _ when l_Rate > 10.0:
                        {
                            LogMsg(p_Name, MessageLevel.Info, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= l_Rate && l_Rate <= 10.0:
                        {
                            LogMsg(p_Name, MessageLevel.OK, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= l_Rate && l_Rate <= 2.0:
                        {
                            LogMsg(p_Name, MessageLevel.Info, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.Info, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogMsg(p_Name, MessageLevel.Info, "Unable to complete test: " + ex.ToString());
            }
        }
    }

}

using System;
//using Microsoft.VisualBasic;
using ASCOM.Standard.Interfaces;
using System.Threading;
using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.COM.DriverAccess;

namespace ConformU
{

    internal class DomeTester : DeviceTesterBaseClass
    {
        const double DOME_SYNC_OFFSET = 45.0; // Amount to offset the azimuth when testing ability to sync
        const double DOME_ILLEGAL_ALTITUDE_LOW = -10.0; // Illegal value to test dome driver exception generation
        const double DOME_ILLEGAL_ALTITUDE_HIGH = 100.0; // Illegal value to test dome driver exception generation
        const double DOME_ILLEGAL_AZIMUTH_LOW = -10.0; // Illegal value to test dome driver exception generation
        const double DOME_ILLEGAL_AZIMUTH_HIGH = 370.0; // Illegal value to test dome driver exception generation

        // Dome variables
        private bool m_CanSetAltitude, m_CanSetAzimuth, m_CanSetShutter, m_CanSlave, m_CanSyncAzimuth, m_Slaved;
        private ShutterState m_ShutterStatus;
        private bool m_CanReadAltitude, m_CanReadAtPark, m_CanReadAtHome, m_CanReadSlewing, m_CanReadSlaved, m_CanReadShutterStatus, m_CanReadAzimuth, m_CanSlewToAzimuth;
        private bool m_AsyncSlewAzimuth, m_AsyncSlewAltitude;

        // General variables
        private bool m_Slewing, m_AtHome, m_AtPark, m_CanFindHome, m_CanFindPark, m_CanPark, m_CanSetPark, m_Connected;
        private string m_Description, m_DriverINfo, m_Name;
        private short m_InterfaceVersion;
        private double m_Altitude, m_Azimuth;


        private enum DomePropertyMethod
        {
            // Properties
            Altitude,
            AtHome,
            AtPark,
            Azimuth,
            CanFindHome,
            CanPark,
            CanSetAltitude,
            CanSetAzimuth,
            CanSetPark,
            CanSetShutter,
            CanSlave,
            CanSyncAzimuth,
            Connected,
            Description,
            DriverInfo,
            InterfaceVersion,
            Name,
            ShutterStatus,
            SlavedRead,
            SlavedWrite,
            Slewing,

            // Methods
            AbortSlew,
            CloseShutter,
            CommandBlind,
            CommandBool,
            CommandString,
            FindHome,
            OpenShutter,
            Park,
            SetPark,
            SlewToAltitude,
            SlewToAzimuth,
            SyncToAzimuth
        }

        private IDomeV2 domeDevice;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose
        public DomeTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, false, true, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        protected override void Dispose(bool disposing)
        {
            LogMsg("Dispose", MessageLevel.msgDebug, "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (domeDevice is not null) domeDevice.Dispose();
                    domeDevice = null;
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
                switch (settings.ComDevice.ProgId.ToUpper())
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
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        domeDevice = new AlpacaDome(settings.AlpacaConfiguration.AccessServiceType.ToString(),
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
                                domeDevice = new DomeFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                domeDevice = new Dome(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");
                baseClassDevice = domeDevice; // Assign the driver to the base class

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

        public override bool Connected
        {
            get
            {
                LogCallToDriver("Connected", "About to get Connected property");
                return domeDevice.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                domeDevice.Connected = value;
                g_Stop = false;
            }
        }
        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(domeDevice, DeviceType.Dome);
        }

        public override void PreRunCheck()
        {
            int l_VStringPtr, l_V1, l_V2, l_V3;
            string l_VString;

            // Add a test for a back level version of the Dome simulator - just abandon this process if any errors occur
            if (settings.ComDevice.ProgId.ToUpper() == "DOMESIM.DOME")
            {
                l_VString = "";
                try
                {
                    LogCallToDriver("PreRunCheck", "About to get DriverInfo property");
                    l_VStringPtr = domeDevice.DriverInfo.ToUpper().IndexOf("ASCOM DOME SIMULATOR "); // Point at the start of the version string
                    if (l_VStringPtr > 0)
                    {
                        LogCallToDriver("PreRunCheck", "About to get DriverInfo property");
                        l_VString = domeDevice.DriverInfo.ToUpper().Substring(l_VStringPtr + 21); // Get the version string
                        l_VStringPtr = l_VString.IndexOf(".");
                        if (l_VStringPtr > 1)
                        {
                            l_V1 = System.Convert.ToInt32(l_VString.Substring(1, l_VStringPtr - 1)); // Extract the number
                            l_VString = l_VString.Substring(l_VStringPtr + 1); // Get the second version number part
                            l_VStringPtr = l_VString.IndexOf(".");
                            if (l_VStringPtr > 1)
                            {
                                l_V2 = int.Parse(l_VString.Substring(1, l_VStringPtr - 1)); // Extract the number
                                l_VString = l_VString.Substring(l_VStringPtr + 1); // Get the third version number part
                                                                                   // Find the next non numeric character
                                l_VStringPtr = 0;
                                do
                                    l_VStringPtr += 1;
                                while (int.TryParse(l_VString.Substring(l_VStringPtr, 1), out _));

                                if (l_VStringPtr > 1)
                                {
                                    l_V3 = System.Convert.ToInt32(l_VString.Substring(1, l_VStringPtr - 1)); // Extract the number
                                                                                                             // Turn the version parts into a whole number
                                    l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3;
                                    if (l_V1 < 5000007)
                                    {
                                        LogMsg("Version Check", MessageLevel.msgIssue, "*** This version of the dome simulator has known conformance issues, ***");
                                        LogMsg("Version Check", MessageLevel.msgIssue, "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***");
                                        LogMsg("", MessageLevel.msgAlways, "");
                                    }
                                    else
                                        LogMsg("Version Check", MessageLevel.msgDebug, "Version check OK");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogMsg("ConformanceCheck", MessageLevel.msgError, ex.ToString());
                }
            }
            if (!cancellationToken.IsCancellationRequested)
            {
                // Get into a consistent state
                try
                {
                    LogCallToDriver("PreRunCheck", "About to get Slewing property");
                    m_Slewing = domeDevice.Slewing; // Try to read the Slewing property
                    if (m_Slewing)
                        LogMsg("DomeSafety", MessageLevel.msgInfo, $"The Slewing property is true at device start-up. This could be by design or possibly Slewing logic is inverted?");// Display a message if slewing is True
                    DomeWaitForSlew(settings.DomeAzimuthTimeout); // Wait for slewing to finish
                }
                catch (Exception ex)
                {
                    LogMsg("DomeSafety", MessageLevel.msgWarning, $"The Slewing property threw an exception and should not have: {ex.Message}"); // Display a warning message because Slewing should not throw an exception!
                    LogMsg("DomeSafety", MessageLevel.msgDebug, $"{ex}");
                }// Log the full message in debug mode
                if (settings.DomeOpenShutter)
                {
                    LogMsg("DomeSafety", MessageLevel.msgComment, "Attempting to open shutter as some tests may fail if it is closed...");
                    try
                    {
                        LogCallToDriver("PreRunCheck", "About to call OpenShutter");
                        domeDevice.OpenShutter();
                        try
                        {
                            DomeShutterWait(ShutterState.Open);
                        }
                        catch
                        {
                        }
                        if (cancellationToken.IsCancellationRequested)
                        {
                            LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                            LogMsg("DomeSafety", MessageLevel.msgComment, "Stop button pressed, further testing abandoned, shutter status: " + domeDevice.ShutterStatus.ToString());
                        }
                        else
                        {
                            LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                            if (domeDevice.ShutterStatus == ShutterState.Open)
                            {
                                LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                                LogMsg("DomeSafety", MessageLevel.msgOK, "Shutter status: " + domeDevice.ShutterStatus.ToString());
                            }
                            else
                            {
                                LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                                LogMsg("DomeSafety", MessageLevel.msgWarning, "Shutter status: " + domeDevice.ShutterStatus.ToString());
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMsg("DomeSafety", MessageLevel.msgComment, "Unable to open shutter, some tests may fail: " + ex.Message);
                    }
                    Status(StatusType.staTest, "");
                }
                else
                    LogMsg("DomeSafety", MessageLevel.msgComment, "Open shutter check box is unchecked so shutter not opened");
            }
        }
        public override void ReadCanProperties()
        {
            DomeMandatoryTest(DomePropertyMethod.CanFindHome, "CanFindHome"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.CanPark, "CanPark"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.CanSetAltitude, "CanSetAltitude"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.CanSetAzimuth, "CanSetAzimuth"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.CanSetPark, "CanSetPark"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.CanSetShutter, "CanSetShutter"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.CanSlave, "CanSlave"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.CanSyncAzimuth, "CanSyncAzimuth"); if (cancellationToken.IsCancellationRequested)
                return;
        }
        public override void CheckProperties()
        {
            if (!settings.DomeOpenShutter)
                LogMsgInfo("Altitude", "You have configured Conform not to open the shutter so the following test may fail.");
            DomeOptionalTest(DomePropertyMethod.Altitude, MemberType.Property, "Altitude"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.AtHome, MemberType.Property, "AtHome"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.AtPark, MemberType.Property, "AtPark"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.Azimuth, MemberType.Property, "Azimuth"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.ShutterStatus, MemberType.Property, "ShutterStatus"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.SlavedRead, "Slaved Read"); if (cancellationToken.IsCancellationRequested)
                return;
            if (m_Slaved & (!m_CanSlave))
                LogMsg("Slaved Read", MessageLevel.msgIssue, "Dome is slaved but CanSlave is false");
            DomeOptionalTest(DomePropertyMethod.SlavedWrite, MemberType.Property, "Slaved Write"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.Slewing, "Slewing"); if (cancellationToken.IsCancellationRequested)
                return;
        }
        public override void CheckMethods()
        {
            DomeMandatoryTest(DomePropertyMethod.AbortSlew, "AbortSlew"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.SlewToAltitude, MemberType.Method, "SlewToAltitude"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.SlewToAzimuth, MemberType.Method, "SlewToAzimuth"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.SyncToAzimuth, MemberType.Method, "SyncToAzimuth"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.CloseShutter, MemberType.Method, "CloseShutter"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.OpenShutter, MemberType.Method, "OpenShutter"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.FindHome, MemberType.Method, "FindHome"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.Park, MemberType.Method, "Park"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeOptionalTest(DomePropertyMethod.SetPark, MemberType.Method, "SetPark"); if (cancellationToken.IsCancellationRequested)
                return; // SetPark must follow Park
        }
        public override void CheckPerformance()
        {
            if (m_CanReadAltitude)
            {
                DomePerformanceTest(DomePropertyMethod.Altitude, "Altitude"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (m_CanReadAzimuth)
            {
                DomePerformanceTest(DomePropertyMethod.Azimuth, "Azimuth"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (m_CanReadShutterStatus)
            {
                DomePerformanceTest(DomePropertyMethod.ShutterStatus, "ShutterStatus"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (m_CanReadSlaved)
            {
                DomePerformanceTest(DomePropertyMethod.SlavedRead, "Slaved"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (m_CanReadSlewing)
            {
                DomePerformanceTest(DomePropertyMethod.Slewing, "Slewing"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
        }
        public override void PostRunCheck()
        {
            if (settings.DomeOpenShutter)
            {
                if (m_CanSetShutter)
                {
                    LogMsg("DomeSafety", MessageLevel.msgInfo, "Attempting to close shutter...");
                    try // Close shutter
                    {
                        LogCallToDriver("DomeSafety", "About to call CloseShutter");
                        domeDevice.CloseShutter();
                        DomeShutterWait(ShutterState.Closed);
                        LogMsg("DomeSafety", MessageLevel.msgOK, "Shutter successfully closed");
                    }
                    catch (Exception ex)
                    {
                        LogMsg("DomeSafety", MessageLevel.msgComment, "Exception closing shutter: " + ex.Message);
                        LogMsg("DomeSafety", MessageLevel.msgComment, "Please close shutter manually");
                    }
                }
                else
                    LogMsg("DomeSafety", MessageLevel.msgInfo, "CanSetShutter is false, please close the shutter manually");
            }
            else
                LogMsg("DomeSafety", MessageLevel.msgInfo, "Open shutter check box is unchecked so close shutter bypassed");
            // 3.0.0.17 - Added check for CanPark
            if (m_CanPark)
            {
                LogMsg("DomeSafety", MessageLevel.msgInfo, "Attempting to park dome...");
                try // Park
                {
                    LogCallToDriver("DomeSafety", "About to call Park");
                    domeDevice.Park();
                    DomeWaitForSlew(settings.DomeAzimuthTimeout);
                    LogMsg("DomeSafety", MessageLevel.msgOK, "Dome successfully parked");
                }
                catch (Exception)
                {
                    LogMsg("DomeSafety", MessageLevel.msgError, "Exception generated, unable to park dome");
                }
            }
            else
                LogMsg("DomeSafety", MessageLevel.msgInfo, "CanPark is false - skipping dome parking");
        }

        private void DomeSlewToAltitude(string p_Name, double p_Altitude)
        {
            DateTime l_StartTime;

            if (!settings.DomeOpenShutter)
                LogMsgInfo("SlewToAltitude", "You have configured Conform not to open the shutter so the following slew may fail.");

            Status(StatusType.staAction, "Slew to " + p_Altitude + " degrees");
            LogCallToDriver(p_Name, "About to call SlewToAltitude");
            domeDevice.SlewToAltitude(p_Altitude);
            if (m_CanReadSlewing)
            {
                l_StartTime = DateTime.Now;
                LogCallToDriver(p_Name, "About to get Slewing property");
                if (domeDevice.Slewing)
                {
                    DomeWaitForSlew(settings.DomeAltitudeTimeout); if (cancellationToken.IsCancellationRequested)
                        return;
                    m_AsyncSlewAltitude = true;
                    LogMsg(p_Name + " " + p_Altitude, MessageLevel.msgOK, "Asynchronous slew OK");
                }
                else
                {
                    m_AsyncSlewAltitude = false;
                    LogMsg(p_Name + " " + p_Altitude, MessageLevel.msgOK, "Synchronous slew OK");
                }
            }
            else
                LogMsg(p_Name + " " + p_Altitude, MessageLevel.msgOK, "Can't read Slewing so assume synchronous slew OK");
            DomeStabliisationWait();
        }
        private void DomeSlewToAzimuth(string p_Name, double p_Azimuth)
        {
            Status(StatusType.staAction, "Slew to " + p_Azimuth + " degrees");
            if (p_Azimuth >= 0.0 & p_Azimuth <= 359.9999999)
            {
                m_CanSlewToAzimuth = false;
                LogCallToDriver(p_Name, "About to call SlewToAzimuth");
                domeDevice.SlewToAzimuth(p_Azimuth);
                m_CanSlewToAzimuth = true; // Command is supported and didn't generate an exception
            }
            else
            {
                LogCallToDriver(p_Name, "About to call SlewToAzimuth");
                domeDevice.SlewToAzimuth(p_Azimuth);
            }
            if (m_CanReadSlewing)
            {
                LogCallToDriver(p_Name, "About to get Slewing property");
                if (domeDevice.Slewing)
                {
                    DomeWaitForSlew(settings.DomeAzimuthTimeout); if (cancellationToken.IsCancellationRequested)
                        return;
                    m_AsyncSlewAzimuth = true;
                    LogMsg(p_Name + " " + p_Azimuth, MessageLevel.msgOK, "Asynchronous slew OK");
                }
                else
                {
                    m_AsyncSlewAzimuth = false;
                    LogMsg(p_Name + " " + p_Azimuth, MessageLevel.msgOK, "Synchronous slew OK");
                }
            }
            else
                LogMsg(p_Name + " " + p_Azimuth, MessageLevel.msgOK, "Can't read Slewing so assume synchronous slew OK");
            DomeStabliisationWait();
        }
        private void DomeWaitForSlew(double p_TimeOut)
        {
            DateTime l_StartTime;
            l_StartTime = DateTime.Now;
            do
            {
                WaitFor(SLEEP_TIME);
                Status(StatusType.staStatus, "Slewing Status: " + domeDevice.Slewing + ", Timeout: " + DateTime.Now.Subtract(l_StartTime).TotalSeconds.ToString("#0") + "/" + p_TimeOut + ", press stop to abandon wait");
            }
            while (domeDevice.Slewing & !cancellationToken.IsCancellationRequested & (DateTime.Now.Subtract(l_StartTime).TotalSeconds <= p_TimeOut));

            Status(StatusType.staStatus, "");
            if ((DateTime.Now.Subtract(l_StartTime).TotalSeconds > p_TimeOut))
            {
                LogMsg("DomeWaitForSlew", MessageLevel.msgError, "Timed out waiting for Dome slew, consider increasing time-outs in Options/Conform Options.");
                LogMsg("DomeWaitForSlew", MessageLevel.msgInfo, "Another cause of time-outs is if your Slewing Property logic is inverted or is not operating correctly.");
            }
        }
        private void DomeMandatoryTest(DomePropertyMethod p_Type, string p_Name)
        {
            try
            {
                switch (p_Type)
                {
                    case DomePropertyMethod.CanFindHome:
                        {
                            LogCallToDriver(p_Name, "About to get CanFindHome property");
                            m_CanFindHome = domeDevice.CanFindHome;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanFindHome.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanPark:
                        {
                            LogCallToDriver(p_Name, "About to get CanPark property");
                            m_CanPark = domeDevice.CanPark;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanPark.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetAltitude:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetAltitude property");
                            m_CanSetAltitude = domeDevice.CanSetAltitude;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanSetAltitude.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetAzimuth:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetAzimuth property");
                            m_CanSetAzimuth = domeDevice.CanSetAzimuth;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanSetAzimuth.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetPark:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetPark property");
                            m_CanSetPark = domeDevice.CanSetPark;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanSetPark.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetShutter:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetShutter property");
                            m_CanSetShutter = domeDevice.CanSetShutter;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanSetShutter.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSlave:
                        {
                            LogCallToDriver(p_Name, "About to get CanSlave property");
                            m_CanSlave = domeDevice.CanSlave;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanSlave.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSyncAzimuth:
                        {
                            LogCallToDriver(p_Name, "About to get CanSyncAzimuth property");
                            m_CanSyncAzimuth = domeDevice.CanSyncAzimuth;
                            LogMsg(p_Name, MessageLevel.msgOK, m_CanSyncAzimuth.ToString());
                            break;
                        }

                    case DomePropertyMethod.Connected:
                        {
                            LogCallToDriver(p_Name, "About to get Connected property");
                            m_Connected = domeDevice.Connected;
                            LogMsg(p_Name, MessageLevel.msgOK, m_Connected.ToString());
                            break;
                        }

                    case DomePropertyMethod.Description:
                        {
                            LogCallToDriver(p_Name, "About to get Description property");
                            m_Description = domeDevice.Description;
                            LogMsg(p_Name, MessageLevel.msgOK, m_Description.ToString());
                            break;
                        }

                    case DomePropertyMethod.DriverInfo:
                        {
                            LogCallToDriver(p_Name, "About to get DriverInfo property");
                            m_DriverINfo = domeDevice.DriverInfo;
                            LogMsg(p_Name, MessageLevel.msgOK, m_DriverINfo.ToString());
                            break;
                        }

                    case DomePropertyMethod.InterfaceVersion:
                        {
                            LogCallToDriver(p_Name, "About to get InterfaceVersion property");
                            m_InterfaceVersion = domeDevice.InterfaceVersion;
                            LogMsg(p_Name, MessageLevel.msgOK, m_InterfaceVersion.ToString());
                            break;
                        }

                    case DomePropertyMethod.Name:
                        {
                            LogCallToDriver(p_Name, "About to get Name property");
                            m_Name = domeDevice.Name;
                            LogMsg(p_Name, MessageLevel.msgOK, m_Name.ToString());
                            break;
                        }

                    case DomePropertyMethod.SlavedRead:
                        {
                            m_CanReadSlaved = false;
                            LogCallToDriver(p_Name, "About to get Slaved property");
                            m_Slaved = domeDevice.Slaved;
                            m_CanReadSlaved = true;
                            LogMsg(p_Name, MessageLevel.msgOK, m_Slaved.ToString());
                            break;
                        }

                    case DomePropertyMethod.Slewing:
                        {
                            m_CanReadSlewing = false;
                            LogCallToDriver(p_Name, "About to get Slewing property");
                            m_Slewing = domeDevice.Slewing;
                            m_CanReadSlewing = true;
                            LogMsg(p_Name, MessageLevel.msgOK, m_Slewing.ToString());
                            break;
                        }

                    case DomePropertyMethod.AbortSlew:
                        {
                            LogCallToDriver(p_Name, "About to call AbortSlew method");
                            domeDevice.AbortSlew();
                            // Confirm that slaved is false
                            if (m_CanReadSlaved)
                            {
                                LogCallToDriver(p_Name, "About to get Slaved property");
                                if (domeDevice.Slaved)
                                    LogMsg("AbortSlew", MessageLevel.msgError, "Slaved property Is true after AbortSlew");
                                else
                                    LogMsg("AbortSlew", MessageLevel.msgOK, "AbortSlew command issued successfully");
                            }
                            else
                                LogMsg("AbortSlew", MessageLevel.msgOK, "Can't read Slaved property AbortSlew command was successful");
                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "DomeMandatoryTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Property, Required.Mandatory, ex, "");
            }
        }
        private void DomeOptionalTest(DomePropertyMethod p_Type, MemberType p_MemberType, string p_Name)
        {
            double l_SlewAngle, l_OriginalAzimuth, l_NewAzimuth;
            try
            {
                switch (p_Type)
                {
                    case DomePropertyMethod.Altitude:
                        {
                            m_CanReadAltitude = false;
                            LogCallToDriver(p_Name, "About to get Altitude property");
                            m_Altitude = domeDevice.Altitude;
                            m_CanReadAltitude = true;
                            LogMsg(p_Name, MessageLevel.msgOK, m_Altitude.ToString());
                            break;
                        }

                    case DomePropertyMethod.AtHome:
                        {
                            m_CanReadAtHome = false;
                            LogCallToDriver(p_Name, "About to get AtHome property");
                            m_AtHome = domeDevice.AtHome;
                            m_CanReadAtHome = true;
                            LogMsg(p_Name, MessageLevel.msgOK, m_AtHome.ToString());
                            break;
                        }

                    case DomePropertyMethod.AtPark:
                        {
                            m_CanReadAtPark = false;
                            LogCallToDriver(p_Name, "About to get AtPark property");
                            m_AtPark = domeDevice.AtPark;
                            m_CanReadAtPark = true;
                            LogMsg(p_Name, MessageLevel.msgOK, m_AtPark.ToString());
                            break;
                        }

                    case DomePropertyMethod.Azimuth:
                        {
                            m_CanReadAzimuth = false;
                            LogCallToDriver(p_Name, "About to get Azimuth property");
                            m_Azimuth = domeDevice.Azimuth;
                            m_CanReadAzimuth = true;
                            LogMsg(p_Name, MessageLevel.msgOK, m_Azimuth.ToString());
                            break;
                        }

                    case DomePropertyMethod.ShutterStatus:
                        {
                            m_CanReadShutterStatus = false;
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            m_ShutterStatus = domeDevice.ShutterStatus;
                            m_CanReadShutterStatus = true;
                            LogMsg(p_Name, MessageLevel.msgOK, m_ShutterStatus.ToString());
                            break;
                        }

                    case DomePropertyMethod.SlavedWrite:
                        {
                            if (m_CanSlave)
                            {
                                if (m_CanReadSlaved)
                                {
                                    if (m_Slaved)
                                    {
                                        LogCallToDriver(p_Name, "About to set Slaved property");
                                        domeDevice.Slaved = false;
                                    }
                                    else
                                    {
                                        LogCallToDriver(p_Name, "About to set Slaved property");
                                        domeDevice.Slaved = true;
                                    }
                                    LogCallToDriver(p_Name, "About to set Slaved property");
                                    domeDevice.Slaved = m_Slaved; // Restore original value
                                    LogMsg("Slaved Write", MessageLevel.msgOK, "Slave state changed successfully");
                                }
                                else
                                    LogMsg("Slaved Write", MessageLevel.msgInfo, "Test skipped since Slaved property can't be read");
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to set Slaved property");
                                domeDevice.Slaved = true;
                                LogMsg(p_Name, MessageLevel.msgError, "CanSlave is false but setting Slaved true did not raise an exception");
                                LogCallToDriver(p_Name, "About to set Slaved property");
                                domeDevice.Slaved = false; // Unslave to continue tests
                            }

                            break;
                        }

                    case DomePropertyMethod.CloseShutter:
                        {
                            if (m_CanSetShutter)
                            {
                                try
                                {
                                    DomeShutterTest(ShutterState.Closed, p_Name);
                                    DomeStabliisationWait();
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetShutter is True");
                                }
                            }
                            else
                            {
                                domeDevice.CloseShutter();
                                LogMsg(p_Name, MessageLevel.msgError, "CanSetShutter is false but CloseShutter did not raise an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.CommandBlind:
                        {
                            LogCallToDriver(p_Name, "About to call CommandBlind method");
                            domeDevice.CommandBlind(""); // m_Dome.CommandBlind("", True)
                            LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent");
                            break;
                        }

                    case DomePropertyMethod.CommandBool:
                        {
                            LogCallToDriver(p_Name, "About to call CommandBool method");
                            domeDevice.CommandBool(""); // m_Dome.CommandBool("", True)
                            LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent");
                            break;
                        }

                    case DomePropertyMethod.CommandString:
                        {
                            LogCallToDriver(p_Name, "About to call CommandString method");
                            domeDevice.CommandString(""); // m_Dome.CommandString("", True)
                            LogMsg(p_Name, MessageLevel.msgOK, "Null string successfully sent");
                            break;
                        }

                    case DomePropertyMethod.FindHome:
                        {
                            if (m_CanFindHome)
                            {
                                Status(StatusType.staTest, p_Name);
                                Status(StatusType.staAction, "Waiting for movement to stop");
                                try
                                {
                                    LogCallToDriver(p_Name, "About to call FindHome method");
                                    domeDevice.FindHome();
                                    if (m_CanReadSlaved)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slaved Property");
                                        if (domeDevice.Slaved)
                                            LogMsg(p_Name, MessageLevel.msgError, "Slaved is true but Home did not raise an exception");
                                    }
                                    if (m_CanReadSlewing)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slewing property repeatedly");
                                        do
                                        {
                                            WaitFor(SLEEP_TIME);
                                            Status(StatusType.staStatus, "Slewing Status: " + domeDevice.Slewing);
                                        }
                                        while (domeDevice.Slewing & !cancellationToken.IsCancellationRequested);
                                    }
                                    if (!cancellationToken.IsCancellationRequested)
                                    {
                                        if (m_CanReadAtHome)
                                        {
                                            LogCallToDriver(p_Name, "About to get AtHome property");
                                            if (domeDevice.AtHome)
                                                LogMsg(p_Name, MessageLevel.msgOK, "Dome homed successfully");
                                            else
                                                LogMsg(p_Name, MessageLevel.msgError, "Home command completed but AtHome is false");
                                        }
                                        else
                                            LogMsg(p_Name, MessageLevel.msgOK, "Can't read AtHome so assume that dome has homed successfully");
                                        DomeStabliisationWait();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanFindHome is True");
                                    DomeStabliisationWait();
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call FindHome method");
                                domeDevice.FindHome();
                                LogMsg(p_Name, MessageLevel.msgError, "CanFindHome is false but FindHome did not throw an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.OpenShutter:
                        {
                            if (m_CanSetShutter)
                            {
                                try
                                {
                                    DomeShutterTest(ShutterState.Open, p_Name);
                                    DomeStabliisationWait();
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetShutter is True");
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call OpenShutter method");
                                domeDevice.OpenShutter();
                                LogMsg(p_Name, MessageLevel.msgError, "CanSetShutter is false but OpenShutter did not raise an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.Park:
                        {
                            if (m_CanPark)
                            {
                                Status(StatusType.staTest, p_Name);
                                Status(StatusType.staAction, "Waiting for movement to stop");
                                try
                                {
                                    LogCallToDriver(p_Name, "About to call Park method");
                                    domeDevice.Park();
                                    if (m_CanReadSlaved)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slaved property");
                                        if (domeDevice.Slaved)
                                            LogMsg(p_Name, MessageLevel.msgError, "Slaved is true but Park did not raise an exception");
                                    }
                                    if (m_CanReadSlewing)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slewing property repeatedly");
                                        do
                                        {
                                            WaitFor(SLEEP_TIME);
                                            Status(StatusType.staStatus, "Slewing Status: " + domeDevice.Slewing);
                                        }
                                        while (domeDevice.Slewing & !cancellationToken.IsCancellationRequested);
                                    }
                                    if (!cancellationToken.IsCancellationRequested)
                                    {
                                        if (m_CanReadAtPark)
                                        {
                                            LogCallToDriver(p_Name, "About to get AtPark property");
                                            if (domeDevice.AtPark)
                                                LogMsg(p_Name, MessageLevel.msgOK, "Dome parked successfully");
                                            else
                                                LogMsg(p_Name, MessageLevel.msgError, "Park command completed but AtPark is false");
                                        }
                                        else
                                            LogMsg(p_Name, MessageLevel.msgOK, "Can't read AtPark so assume that dome has parked successfully");
                                    }
                                    DomeStabliisationWait();
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanPark is True");
                                    DomeStabliisationWait();
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call Park method");
                                domeDevice.Park();
                                LogMsg(p_Name, MessageLevel.msgError, "CanPark is false but Park did not raise an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.SetPark:
                        {
                            if (m_CanSetPark)
                            {
                                try
                                {
                                    LogCallToDriver(p_Name, "About to call SetPark method");
                                    domeDevice.SetPark();
                                    LogMsg(p_Name, MessageLevel.msgOK, "SetPark issued OK");
                                }
                                catch (Exception ex)
                                {
                                    HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetPark is True");
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call SetPark method");
                                domeDevice.SetPark();
                                LogMsg(p_Name, MessageLevel.msgError, "CanSetPath is false but SetPath did not throw an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.SlewToAltitude:
                        {
                            if (m_CanSetAltitude)
                            {
                                Status(StatusType.staTest, p_Name);
                                /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
                                for (l_SlewAngle = 0; l_SlewAngle <= 90; l_SlewAngle += 15)
                                {
                                    /* TODO ERROR: Skipped EndIfDirectiveTrivia */
                                    try
                                    {
                                        DomeSlewToAltitude(p_Name, l_SlewAngle);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetAltitude is True");
                                    }
                                }
                                // Test out of range values -10 and 100 degrees
                                if (m_CanSetAltitude)
                                {
                                    try
                                    {
                                        DomeSlewToAltitude(p_Name, DOME_ILLEGAL_ALTITUDE_LOW);
                                        LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal altitude " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees", "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees");
                                    }
                                    try
                                    {
                                        DomeSlewToAltitude(p_Name, DOME_ILLEGAL_ALTITUDE_HIGH);
                                        LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal altitude " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees", "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees");
                                    }
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call SlewToAltitude method");
                                domeDevice.SlewToAltitude(45.0);
                                LogMsg(p_Name, MessageLevel.msgError, "CanSetAltitude is false but SlewToAltitude did not raise an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.SlewToAzimuth:
                        {
                            if (m_CanSetAzimuth)
                            {
                                Status(StatusType.staTest, p_Name);
                                /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
                                for (l_SlewAngle = 0; l_SlewAngle <= 315; l_SlewAngle += 45)
                                {
                                    /* TODO ERROR: Skipped EndIfDirectiveTrivia */
                                    try
                                    {
                                        DomeSlewToAzimuth(p_Name, l_SlewAngle);
                                        if (cancellationToken.IsCancellationRequested)
                                            return;
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleException(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "CanSetAzimuth is True");
                                    }
                                }

                                if (m_CanSetAzimuth)
                                {
                                    // Test out of range values -10 and 370 degrees
                                    try
                                    {
                                        DomeSlewToAzimuth(p_Name, DOME_ILLEGAL_AZIMUTH_LOW);
                                        LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees", "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    try
                                    {
                                        DomeSlewToAzimuth(p_Name, DOME_ILLEGAL_AZIMUTH_HIGH);
                                        LogMsg(p_Name, MessageLevel.msgError, "No exception generated when slewing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees", "COM invalid value exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call SlewToAzimuth method");
                                domeDevice.SlewToAzimuth(45.0);
                                LogMsg(p_Name, MessageLevel.msgError, "CanSetAzimuth is false but SlewToAzimuth did not throw an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.SyncToAzimuth:
                        {
                            if (m_CanSyncAzimuth)
                            {
                                if (m_CanSlewToAzimuth)
                                {
                                    if (m_CanReadAzimuth)
                                    {
                                        LogCallToDriver(p_Name, "About to get Azimuth property");
                                        l_OriginalAzimuth = domeDevice.Azimuth;
                                        if (l_OriginalAzimuth > 300.0)
                                            l_NewAzimuth = l_OriginalAzimuth - DOME_SYNC_OFFSET;
                                        else
                                            l_NewAzimuth = l_OriginalAzimuth + DOME_SYNC_OFFSET;
                                        domeDevice.SyncToAzimuth(l_NewAzimuth); // Sync to new azimuth
                                                                                // OK Dome hasn't moved but should now show azimuth as a new value
                                        switch (Math.Abs(l_NewAzimuth - domeDevice.Azimuth))
                                        {
                                            case object _ when Math.Abs(l_NewAzimuth - domeDevice.Azimuth) < 1.0 // very close so give it an OK
                                           :
                                                {
                                                    LogMsg(p_Name, MessageLevel.msgOK, "Dome synced OK to within +- 1 degree");
                                                    break;
                                                }

                                            case object _ when Math.Abs(l_NewAzimuth - domeDevice.Azimuth) < 2.0 // close so give it an INFO
                                     :
                                                {
                                                    LogMsg(p_Name, MessageLevel.msgInfo, "Dome synced to within +- 2 degrees");
                                                    break;
                                                }

                                            case object _ when Math.Abs(l_NewAzimuth - domeDevice.Azimuth) < 5.0 // Closish so give an issue
                                     :
                                                {
                                                    LogMsg(p_Name, MessageLevel.msgIssue, "Dome only synced to within +- 5 degrees");
                                                    break;
                                                }

                                            case object _ when (DOME_SYNC_OFFSET - 2.0) <= Math.Abs(l_NewAzimuth - domeDevice.Azimuth) && Math.Abs(l_NewAzimuth - domeDevice.Azimuth) <= (DOME_SYNC_OFFSET + 2) // Hasn't really moved
                                     :
                                                {
                                                    LogMsg(p_Name, MessageLevel.msgError, "Dome did not sync, Azimuth didn't change value after sync command");
                                                    break;
                                                }

                                            default:
                                                {
                                                    LogMsg(p_Name, MessageLevel.msgIssue, "Dome azimuth was " + Math.Abs(l_NewAzimuth - domeDevice.Azimuth) + " degrees away from expected value");
                                                    break;
                                                }
                                        }
                                        // Now try and restore original value
                                        LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                        domeDevice.SyncToAzimuth(l_OriginalAzimuth);
                                    }
                                    else
                                    {
                                        LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                        domeDevice.SyncToAzimuth(45.0); // Sync to an arbitrary direction
                                        LogMsg(p_Name, MessageLevel.msgOK, "Dome successfully synced to 45 degrees but unable to read azimuth to confirm this");
                                    }

                                    // Now test sync to illegal values
                                    try
                                    {
                                        LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                        domeDevice.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_LOW);
                                        LogMsg(p_Name, MessageLevel.msgError, "No exception generated when syncing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "sync to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees", "COM invalid value exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    try
                                    {
                                        LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                        domeDevice.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_HIGH);
                                        LogMsg(p_Name, MessageLevel.msgError, "No exception generated when syncing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "sync to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees", "COM invalid value exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                }
                                else
                                    LogMsg(p_Name, MessageLevel.msgInfo, "SyncToAzimuth test skipped since SlewToAzimuth throws an exception");
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                domeDevice.SyncToAzimuth(45.0);
                                LogMsg(p_Name, MessageLevel.msgError, "CanSyncAzimuth is false but SyncToAzimuth did not raise an exception");
                            }

                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "DomeOptionalTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, p_MemberType, Required.Optional, ex, "");
            }
            Status(StatusType.staTest, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
        }
        private void DomeShutterTest(ShutterState p_RequiredShutterState, string p_Name)
        {
            ShutterState l_ShutterState;

            if (settings.DomeOpenShutter)
            {
                Status(StatusType.staTest, p_Name);
                if (m_CanReadShutterStatus)
                {
                    LogCallToDriver(p_Name, "About to get ShutterStatus property");
                    l_ShutterState = (ShutterState)domeDevice.ShutterStatus;

                    // Make sure we are in the required state to start the test
                    switch (l_ShutterState)
                    {
                        case ShutterState.Closed:
                            {
                                if (p_RequiredShutterState == ShutterState.Closed)
                                {
                                    // Wrong state, get to the required state
                                    Status(StatusType.staAction, "Opening shutter ready for close test");
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Opening shutter ready for close test");
                                    LogCallToDriver(p_Name, "About to call OpenShutter method");
                                    domeDevice.OpenShutter();
                                    if (!DomeShutterWait(ShutterState.Open))
                                        return; // Wait for shutter to open
                                    DomeStabliisationWait();
                                }
                                else
                                {
                                }

                                break;
                            }

                        case ShutterState.Closing:
                            {
                                if (p_RequiredShutterState == ShutterState.Closed)
                                {
                                    Status(StatusType.staAction, "Waiting for shutter to close before opening ready for close test");
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close before opening ready for close test");
                                    if (!DomeShutterWait(ShutterState.Closed))
                                        return; // Wait for shutter to close
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Opening shutter ready for close test");
                                    Status(StatusType.staAction, "Opening shutter ready for close test");
                                    LogCallToDriver(p_Name, "About to call OpenShutter method");
                                    domeDevice.OpenShutter(); // Then open it
                                    if (!DomeShutterWait(ShutterState.Open))
                                        return;
                                    DomeStabliisationWait();
                                }
                                else
                                {
                                    Status(StatusType.staAction, "Waiting for shutter to close ready for open test");
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close ready for open test");
                                    if (!DomeShutterWait(ShutterState.Closed))
                                        return; // Wait for shutter to close
                                    DomeStabliisationWait();
                                }

                                break;
                            }

                        case ShutterState.Opening:
                            {
                                if (p_RequiredShutterState == ShutterState.Closed)
                                {
                                    Status(StatusType.staAction, "Waiting for shutter to open ready for close test");
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open ready for close test");
                                    if (!DomeShutterWait(ShutterState.Open))
                                        return; // Wait for shutter to open
                                    DomeStabliisationWait();
                                }
                                else
                                {
                                    Status(StatusType.staAction, "Waiting for shutter to open before closing ready for open test");
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open before closing ready for open test");
                                    if (!DomeShutterWait(ShutterState.Open))
                                        return; // Wait for shutter to open
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Closing shutter ready for open test");
                                    Status(StatusType.staAction, "Closing shutter ready for open test");
                                    LogCallToDriver(p_Name, "About to call CloseShutter method");
                                    domeDevice.CloseShutter(); // Then close it
                                    if (!DomeShutterWait(ShutterState.Closed))
                                        return;
                                    DomeStabliisationWait();
                                }

                                break;
                            }

                        case ShutterState.Error:
                            {
                                LogMsg("DomeShutterTest", MessageLevel.msgError, $"Shutter state is Error: {l_ShutterState}");
                                break;
                            }

                        case ShutterState.Open:
                            {
                                if (p_RequiredShutterState == ShutterState.Closed)
                                {
                                }
                                else
                                {
                                    // Wrong state, get to the required state
                                    Status(StatusType.staAction, "Closing shutter ready for open  test");
                                    LogMsg(p_Name, MessageLevel.msgDebug, "Closing shutter ready for open test");
                                    LogCallToDriver(p_Name, "About to call CloseShutter method");
                                    domeDevice.CloseShutter();
                                    if (!DomeShutterWait(ShutterState.Closed))
                                        return; // Wait for shutter to open
                                    DomeStabliisationWait();
                                }

                                break;
                            }

                        default:
                            {
                                LogMsg("DomeShutterTest", MessageLevel.msgError, "Unexpected shutter status: " + l_ShutterState.ToString());
                                break;
                            }
                    }

                    // Now test that we can get to the required state
                    if (p_RequiredShutterState == ShutterState.Closed)
                    {
                        // Shutter is now open so close it
                        Status(StatusType.staAction, "Closing shutter");
                        LogCallToDriver(p_Name, "About to call CloseShutter method");
                        domeDevice.CloseShutter();
                        Status(StatusType.staAction, "Waiting for shutter to close");
                        LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to close");
                        if (!DomeShutterWait(ShutterState.Closed))
                        {
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            l_ShutterState = domeDevice.ShutterStatus;
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            LogMsg(p_Name, MessageLevel.msgError, "Unable to close shutter - ShutterStatus: " + domeDevice.ShutterStatus.ToString());
                            return;
                        }
                        else
                            LogMsg(p_Name, MessageLevel.msgOK, "Shutter closed successfully");
                        DomeStabliisationWait();
                    }
                    else
                    {
                        Status(StatusType.staAction, "Opening shutter");
                        domeDevice.OpenShutter();
                        Status(StatusType.staAction, "Waiting for shutter to open");
                        LogMsg(p_Name, MessageLevel.msgDebug, "Waiting for shutter to open");
                        if (!DomeShutterWait(ShutterState.Open))
                        {
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            l_ShutterState = domeDevice.ShutterStatus;
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            LogMsg(p_Name, MessageLevel.msgError, "Unable to open shutter - ShutterStatus: " + domeDevice.ShutterStatus.ToString());
                            return;
                        }
                        else
                            LogMsg(p_Name, MessageLevel.msgOK, "Shutter opened successfully");
                        DomeStabliisationWait();
                    }
                }
                else
                {
                    LogMsg(p_Name, MessageLevel.msgDebug, "Can't read shutter status!");
                    if (p_RequiredShutterState == ShutterState.Closed)
                    {
                        // Just issue command to see if it doesn't generate an error
                        LogCallToDriver(p_Name, "About to call CloseShutter method");
                        domeDevice.CloseShutter();
                        DomeStabliisationWait();
                    }
                    else
                    {
                        // Just issue command to see if it doesn't generate an error
                        LogCallToDriver(p_Name, "About to call OpenShutter method");
                        domeDevice.OpenShutter();
                        DomeStabliisationWait();
                    }
                    LogMsg(p_Name, MessageLevel.msgOK, "Command issued successfully but can't read ShutterStatus to confirm shutter is closed");
                }
                Status(StatusType.staTest, "");
                Status(StatusType.staAction, "");
                Status(StatusType.staStatus, "");
            }
            else
                LogMsg("DomeSafety", MessageLevel.msgComment, "Open shutter check box is unchecked so shutter test bypassed");
        }

        private bool DomeShutterWait(ShutterState p_RequiredStatus)
        {
            DateTime l_StartTime;
            // Wait for shutter to reach required stats or user presses stop or timeout occurs
            // Returns true if required state is reached
            ShutterState l_ShutterState;
            bool domeShutterWait = false;
            l_StartTime = DateTime.Now;
            try
            {
                LogCallToDriver("DomeShutterWait", "About to get ShutterStatus property repeatedly");
                do
                {
                    WaitFor(SLEEP_TIME);
                    l_ShutterState = domeDevice.ShutterStatus;
                    Status(StatusType.staStatus, "Shutter State: " + l_ShutterState.ToString() + " Timeout: " + DateTime.Now.Subtract(l_StartTime).Seconds + "/" + settings.DomeShutterTimeout);
                }
                while (!(l_ShutterState == p_RequiredStatus) & !cancellationToken.IsCancellationRequested & (DateTime.Now.Subtract(l_StartTime).TotalSeconds <= settings.DomeShutterTimeout));
                LogCallToDriver("DomeShutterWait", "About to get ShutterStatus property");
                if ((domeDevice.ShutterStatus == p_RequiredStatus))
                    domeShutterWait = true; // All worked so return True
                if ((DateTime.Now.Subtract(l_StartTime).TotalSeconds > settings.DomeShutterTimeout))
                    LogMsg("DomeShutterWait", MessageLevel.msgError, "Timed out waiting for shutter to reach state: " + p_RequiredStatus.ToString() + ", consider increasing the timeout setting in Options / Conformance Options");
            }
            catch (Exception ex)
            {
                LogMsg("DomeShutterWait", MessageLevel.msgError, "Unexpected exception: " + ex.ToString());
            }

            return domeShutterWait;
        }
        private void DomePerformanceTest(DomePropertyMethod p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
            double l_Rate;
            bool l_Boolean;
            double l_Double;
            ShutterState l_ShutterState;
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
                        case DomePropertyMethod.Altitude:
                            {
                                l_Double = domeDevice.Altitude;
                                break;
                            }

                        case DomePropertyMethod.Azimuth:
                            {
                                l_Double = domeDevice.Azimuth;
                                break;
                            }

                        case DomePropertyMethod.ShutterStatus:
                            {
                                l_ShutterState = domeDevice.ShutterStatus;
                                break;
                            }

                        case DomePropertyMethod.SlavedRead:
                            {
                                l_Boolean = domeDevice.Slaved;
                                break;
                            }

                        case DomePropertyMethod.Slewing:
                            {
                                l_Boolean = domeDevice.Slewing;
                                break;
                            }

                        default:
                            {
                                LogMsg(p_Name, MessageLevel.msgError, "DomePerformanceTest: Unknown test type " + p_Type.ToString());
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
                LogMsg(p_Name, MessageLevel.msgInfo, "Unable to complete test: " + ex.Message);
            }
        }

        public void DomeStabliisationWait()
        {
            Status(StatusType.staStatus, ""); // Clear status field
            for (double i = 1.0; i <= settings.DomeStabilisationWaitTime; i++)
            {
                Status(StatusType.staAction, "Waiting for Dome to stabilise - " + System.Convert.ToString(i) + "/" + settings.DomeStabilisationWaitTime + " seconds");
                WaitFor(1000); // Wait for 1 second
            }
        }
    }
}

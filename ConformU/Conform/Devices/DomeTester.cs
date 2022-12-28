using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;

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

        // General variables
        private bool m_Slewing, m_AtHome, m_AtPark, m_CanFindHome, m_CanPark, m_CanSetPark, m_Connected;
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
        public DomeTester(ConformConfigurationService conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, false, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
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
                    domeDevice?.Dispose();
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
                        LogInfo("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        domeDevice = new AlpacaDome(settings.AlpacaConfiguration.AccessServiceType,
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
                                                    Assembly.GetExecutingAssembly().GetName().Version.ToString(4));


                        LogInfo("CreateDevice", $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComAccessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                domeDevice = new DomeFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                domeDevice = new Dome(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                baseClassDevice = domeDevice; // Assign the driver to the base class

                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);

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
                return domeDevice.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                domeDevice.Connected = value;

            }
        }
        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(domeDevice, DeviceTypes.Dome);
        }

        public override void PreRunCheck()
        {
            int l_VStringPtr, l_V1, l_V2, l_V3;
            string l_VString;

            // Add a test for a back level version of the Dome simulator - just abandon this process if any errors occur
            if (settings.ComDevice.ProgId.ToUpper() == "DOMESIM.DOME")
            {
                try
                {
                    LogCallToDriver("PreRunCheck", "About to get DriverInfo property");
                    l_VStringPtr = domeDevice.DriverInfo.ToUpper().IndexOf("ASCOM DOME SIMULATOR "); // Point at the start of the version string
                    if (l_VStringPtr > 0)
                    {
                        LogCallToDriver("PreRunCheck", "About to get DriverInfo property");
                        l_VString = domeDevice.DriverInfo.ToUpper()[(l_VStringPtr + 21)..]; // Get the version string
                        l_VStringPtr = l_VString.IndexOf(".");
                        if (l_VStringPtr > 1)
                        {
                            l_V1 = System.Convert.ToInt32(l_VString[1..l_VStringPtr]); // Extract the number
                            l_VString = l_VString[(l_VStringPtr + 1)..]; // Get the second version number part
                            l_VStringPtr = l_VString.IndexOf(".");
                            if (l_VStringPtr > 1)
                            {
                                l_V2 = int.Parse(l_VString[1..l_VStringPtr]); // Extract the number
                                l_VString = l_VString[(l_VStringPtr + 1)..]; // Get the third version number part
                                                                             // Find the next non numeric character
                                l_VStringPtr = 0;
                                do
                                    l_VStringPtr += 1;
                                while (int.TryParse(l_VString.AsSpan(l_VStringPtr, 1), out _));

                                if (l_VStringPtr > 1)
                                {
                                    l_V3 = System.Convert.ToInt32(l_VString[1..l_VStringPtr]); // Extract the number
                                                                                               // Turn the version parts into a whole number
                                    l_V1 = l_V1 * 1000000 + l_V2 * 1000 + l_V3;
                                    if (l_V1 < 5000007)
                                    {
                                        LogIssue("Version Check", "*** This version of the dome simulator has known conformance issues, ***");
                                        LogIssue("Version Check", "*** please update it from the ASCOM site https://ascom-standards.org/Downloads/Index.htm ***");
                                        LogNewLine();
                                    }
                                    else
                                        LogDebug("Version Check", "Version check OK");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogIssue("ConformanceCheck", ex.ToString());
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
                        LogInfo("DomeSafety", $"The Slewing property is true at device start-up. This could be by design or possibly Slewing logic is inverted?");// Display a message if slewing is True
                    DomeWaitForSlew(settings.DomeAzimuthMovementTimeout, null); // Wait for slewing to finish
                }
                catch (Exception ex)
                {
                    LogIssue("DomeSafety", $"The Slewing property threw an exception and should not have: {ex.Message}"); // Display a warning message because Slewing should not throw an exception!
                    LogDebug("DomeSafety", $"{ex}");
                }// Log the full message in debug mode
                if (settings.DomeOpenShutter)
                {
                    LogTestAndMessage("DomeSafety", "Attempting to open shutter as some tests may fail if it is closed...");
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
                            LogTestAndMessage("DomeSafety", "Stop button pressed, further testing abandoned, shutter status: " + domeDevice.ShutterStatus.ToString());
                        }
                        else
                        {
                            LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                            if (domeDevice.ShutterStatus == ShutterState.Open)
                            {
                                LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                                LogOK("DomeSafety", "Shutter status: " + domeDevice.ShutterStatus.ToString());
                            }
                            else
                            {
                                LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                                LogIssue("DomeSafety", "Shutter status: " + domeDevice.ShutterStatus.ToString());
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogTestAndMessage("DomeSafety", "Unable to open shutter, some tests may fail: " + ex.Message);
                    }
                    SetTest("");
                }
                else
                    LogTestAndMessage("DomeSafety", "Open shutter check box is unchecked so shutter not opened");
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
                LogInfo("Altitude", "You have configured Conform not to open the shutter so the following test may fail.");
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
                LogIssue("Slaved Read", "Dome is slaved but CanSlave is false");
            DomeOptionalTest(DomePropertyMethod.SlavedWrite, MemberType.Property, "Slaved Write"); if (cancellationToken.IsCancellationRequested)
                return;
            DomeMandatoryTest(DomePropertyMethod.Slewing, "Slewing"); if (cancellationToken.IsCancellationRequested)
                return;
        }
        public override void CheckMethods()
        {
            DomeMandatoryTest(DomePropertyMethod.AbortSlew, "AbortSlew"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.SlewToAltitude, MemberType.Method, "SlewToAltitude"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.SlewToAzimuth, MemberType.Method, "SlewToAzimuth"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.SyncToAzimuth, MemberType.Method, "SyncToAzimuth"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.CloseShutter, MemberType.Method, "CloseShutter"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.OpenShutter, MemberType.Method, "OpenShutter"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.FindHome, MemberType.Method, "FindHome"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.Park, MemberType.Method, "Park"); if (cancellationToken.IsCancellationRequested) return;
            DomeOptionalTest(DomePropertyMethod.SetPark, MemberType.Method, "SetPark"); if (cancellationToken.IsCancellationRequested) return; // SetPark must follow Park
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
                    LogInfo("DomeSafety", "Attempting to close shutter...");
                    try // Close shutter
                    {
                        LogCallToDriver("DomeSafety", "About to call CloseShutter");
                        domeDevice.CloseShutter();
                        DomeShutterWait(ShutterState.Closed);
                        LogOK("DomeSafety", "Shutter successfully closed");
                    }
                    catch (Exception ex)
                    {
                        LogTestAndMessage("DomeSafety", "Exception closing shutter: " + ex.Message);
                        LogTestAndMessage("DomeSafety", "Please close shutter manually");
                    }
                }
                else
                    LogInfo("DomeSafety", "CanSetShutter is false, please close the shutter manually");
            }
            else
                LogInfo("DomeSafety", "Open shutter check box is unchecked so close shutter bypassed");
            // 3.0.0.17 - Added check for CanPark
            if (m_CanPark)
            {
                LogInfo("DomeSafety", "Attempting to park dome...");
                try // Park
                {
                    LogCallToDriver("DomeSafety", "About to call Park");
                    domeDevice.Park();
                    DomeWaitForSlew(settings.DomeAzimuthMovementTimeout, null);
                    LogOK("DomeSafety", "Dome successfully parked");
                }
                catch (Exception)
                {
                    LogIssue("DomeSafety", "Exception generated, unable to park dome");
                }
            }
            else
                LogInfo("DomeSafety", "CanPark is false - skipping dome parking");
        }

        private void DomeSlewToAltitude(string p_Name, double p_Altitude)
        {
            if (!settings.DomeOpenShutter) LogInfo("SlewToAltitude", "You have configured Conform not to open the shutter so the following slew may fail.");

            SetTest("SlewToAltitude");
            SetAction($"Slewing to altitude {p_Altitude} degrees");
            LogCallToDriver(p_Name, "About to call SlewToAltitude");
            domeDevice.SlewToAltitude(p_Altitude);
            if (m_CanReadSlewing)
            {
                LogCallToDriver(p_Name, "About to get Slewing property");
                if (domeDevice.Slewing)
                {
                    DomeWaitForSlew(settings.DomeAltitudeMovementTimeout, () => { return $"{domeDevice.Altitude:00} / {p_Altitude:00} degrees"; }); if (cancellationToken.IsCancellationRequested) return;
                    LogOK(p_Name + " " + p_Altitude, "Asynchronous slew OK");
                }
                else
                {
                    LogOK(p_Name + " " + p_Altitude, "Synchronous slew OK");
                }
            }
            else
            {
                LogOK(p_Name + " " + p_Altitude, "Can't read Slewing so assume synchronous slew OK");
            }
            DomeStabliisationWait();

            // Check whether the reported altitude matches the requested altitude
            if (m_CanReadAltitude)
            {
                LogCallToDriver(p_Name, "About to get Altitude property");
                double altitude = domeDevice.Altitude;

                if (Math.Abs(altitude - p_Altitude) <= settings.DomeSlewTolerance)
                {
                    LogOK(p_Name + " " + p_Altitude, $"Reached the required altitude: {p_Altitude:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported altitude: {altitude:0.0} degrees");
                }
                else
                {
                    LogIssue(p_Name + " " + p_Altitude, $"Failed to reach the required altitude: {p_Altitude:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported altitude: {altitude:0.0} degrees");
                }
            }
        }
        private void DomeSlewToAzimuth(string p_Name, double p_Azimuth)
        {
            SetAction($"Slewing to azimuth {p_Azimuth} degrees");
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
                    DomeWaitForSlew(settings.DomeAzimuthMovementTimeout, () => { return $"{domeDevice.Azimuth:000} / {p_Azimuth:000} degrees"; }); if (cancellationToken.IsCancellationRequested) return;
                    LogOK(p_Name + " " + p_Azimuth, "Asynchronous slew OK");
                }
                else
                {
                    LogOK(p_Name + " " + p_Azimuth, "Synchronous slew OK");
                }
            }
            else
            {
                LogOK(p_Name + " " + p_Azimuth, "Can't read Slewing so assume synchronous slew OK");
            }
            DomeStabliisationWait();

            // Check whether the reported azimuth matches the requested azimuth
            if (m_CanReadAzimuth)
            {
                LogCallToDriver(p_Name, "About to get Azimuth property");
                double azimuth = domeDevice.Azimuth;

                if (Math.Abs(azimuth - p_Azimuth) <= settings.DomeSlewTolerance)
                {
                    LogOK(p_Name + " " + p_Azimuth, $"Reached the required azimuth: {p_Azimuth:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported azimuth: {azimuth:0.0}");
                }
                else
                {
                    LogIssue(p_Name + " " + p_Azimuth, $"Failed to reach the required azimuth: {p_Azimuth:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported azimuth: {azimuth:0.0}");
                }
            }

        }
        private void DomeWaitForSlew(double p_TimeOut, Func<string> reportingFunction)
        {
            DateTime l_StartTime;
            l_StartTime = DateTime.Now;

            WaitWhile("", () => { return domeDevice.Slewing; }, 500, Convert.ToInt32(p_TimeOut), reportingFunction);

            SetStatus("");
            if ((DateTime.Now.Subtract(l_StartTime).TotalSeconds > p_TimeOut))
            {
                LogIssue("DomeWaitForSlew", "Timed out waiting for Dome slew, consider increasing time-outs in Options/Conform Options.");
                LogInfo("DomeWaitForSlew", "Another cause of time-outs is if your Slewing Property logic is inverted or is not operating correctly.");
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
                            LogOK(p_Name, m_CanFindHome.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanPark:
                        {
                            LogCallToDriver(p_Name, "About to get CanPark property");
                            m_CanPark = domeDevice.CanPark;
                            LogOK(p_Name, m_CanPark.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetAltitude:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetAltitude property");
                            m_CanSetAltitude = domeDevice.CanSetAltitude;
                            LogOK(p_Name, m_CanSetAltitude.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetAzimuth:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetAzimuth property");
                            m_CanSetAzimuth = domeDevice.CanSetAzimuth;
                            LogOK(p_Name, m_CanSetAzimuth.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetPark:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetPark property");
                            m_CanSetPark = domeDevice.CanSetPark;
                            LogOK(p_Name, m_CanSetPark.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSetShutter:
                        {
                            LogCallToDriver(p_Name, "About to get CanSetShutter property");
                            m_CanSetShutter = domeDevice.CanSetShutter;
                            LogOK(p_Name, m_CanSetShutter.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSlave:
                        {
                            LogCallToDriver(p_Name, "About to get CanSlave property");
                            m_CanSlave = domeDevice.CanSlave;
                            LogOK(p_Name, m_CanSlave.ToString());
                            break;
                        }

                    case DomePropertyMethod.CanSyncAzimuth:
                        {
                            LogCallToDriver(p_Name, "About to get CanSyncAzimuth property");
                            m_CanSyncAzimuth = domeDevice.CanSyncAzimuth;
                            LogOK(p_Name, m_CanSyncAzimuth.ToString());
                            break;
                        }

                    case DomePropertyMethod.Connected:
                        {
                            LogCallToDriver(p_Name, "About to get Connected property");
                            m_Connected = domeDevice.Connected;
                            LogOK(p_Name, m_Connected.ToString());
                            break;
                        }

                    case DomePropertyMethod.Description:
                        {
                            LogCallToDriver(p_Name, "About to get Description property");
                            m_Description = domeDevice.Description;
                            LogOK(p_Name, m_Description.ToString());
                            break;
                        }

                    case DomePropertyMethod.DriverInfo:
                        {
                            LogCallToDriver(p_Name, "About to get DriverInfo property");
                            m_DriverINfo = domeDevice.DriverInfo;
                            LogOK(p_Name, m_DriverINfo.ToString());
                            break;
                        }

                    case DomePropertyMethod.InterfaceVersion:
                        {
                            LogCallToDriver(p_Name, "About to get InterfaceVersion property");
                            m_InterfaceVersion = domeDevice.InterfaceVersion;
                            LogOK(p_Name, m_InterfaceVersion.ToString());
                            break;
                        }

                    case DomePropertyMethod.Name:
                        {
                            LogCallToDriver(p_Name, "About to get Name property");
                            m_Name = domeDevice.Name;
                            LogOK(p_Name, m_Name.ToString());
                            break;
                        }

                    case DomePropertyMethod.SlavedRead:
                        {
                            m_CanReadSlaved = false;
                            LogCallToDriver(p_Name, "About to get Slaved property");
                            m_Slaved = domeDevice.Slaved;
                            m_CanReadSlaved = true;
                            LogOK(p_Name, m_Slaved.ToString());
                            break;
                        }

                    case DomePropertyMethod.Slewing:
                        {
                            m_CanReadSlewing = false;
                            LogCallToDriver(p_Name, "About to get Slewing property");
                            m_Slewing = domeDevice.Slewing;
                            m_CanReadSlewing = true;
                            LogOK(p_Name, m_Slewing.ToString());
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
                                    LogIssue("AbortSlew", "Slaved property Is true after AbortSlew");
                                else
                                    LogOK("AbortSlew", "AbortSlew command issued successfully");
                            }
                            else
                                LogOK("AbortSlew", "Can't read Slaved property AbortSlew command was successful");
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "DomeMandatoryTest: Unknown test type " + p_Type.ToString());
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
                            LogOK(p_Name, m_Altitude.ToString());
                            break;
                        }

                    case DomePropertyMethod.AtHome:
                        {
                            m_CanReadAtHome = false;
                            LogCallToDriver(p_Name, "About to get AtHome property");
                            m_AtHome = domeDevice.AtHome;
                            m_CanReadAtHome = true;
                            LogOK(p_Name, m_AtHome.ToString());
                            break;
                        }

                    case DomePropertyMethod.AtPark:
                        {
                            m_CanReadAtPark = false;
                            LogCallToDriver(p_Name, "About to get AtPark property");
                            m_AtPark = domeDevice.AtPark;
                            m_CanReadAtPark = true;
                            LogOK(p_Name, m_AtPark.ToString());
                            break;
                        }

                    case DomePropertyMethod.Azimuth:
                        {
                            m_CanReadAzimuth = false;
                            LogCallToDriver(p_Name, "About to get Azimuth property");
                            m_Azimuth = domeDevice.Azimuth;
                            m_CanReadAzimuth = true;
                            LogOK(p_Name, m_Azimuth.ToString());
                            break;
                        }

                    case DomePropertyMethod.ShutterStatus:
                        {
                            m_CanReadShutterStatus = false;
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            m_ShutterStatus = domeDevice.ShutterStatus;
                            m_CanReadShutterStatus = true;
                            LogOK(p_Name, m_ShutterStatus.ToString());
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
                                    LogOK("Slaved Write", "Slave state changed successfully");
                                }
                                else
                                    LogInfo("Slaved Write", "Test skipped since Slaved property can't be read");
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to set Slaved property");
                                domeDevice.Slaved = true;
                                LogIssue(p_Name, "CanSlave is false but setting Slaved true did not raise an exception");
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
                                LogIssue(p_Name, "CanSetShutter is false but CloseShutter did not raise an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.FindHome:
                        {
                            if (m_CanFindHome)
                            {
                                SetTest(p_Name);
                                SetAction("Finding home");
                                SetStatus("Waiting for movement to stop");
                                try
                                {
                                    LogCallToDriver(p_Name, "About to call FindHome method");
                                    domeDevice.FindHome();
                                    if (m_CanReadSlaved)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slaved Property");
                                        if (domeDevice.Slaved)
                                            LogIssue(p_Name, "Slaved is true but Home did not raise an exception");
                                    }
                                    if (m_CanReadSlewing)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slewing property repeatedly");
                                        //do
                                        //{
                                        //    WaitFor(SLEEP_TIME);
                                        //    SetStatus("Slewing Status: " + domeDevice.Slewing);
                                        //}
                                        //while (domeDevice.Slewing & !cancellationToken.IsCancellationRequested);
                                        WaitWhile("Finding home", () => { return domeDevice.Slewing; }, 500, settings.DomeAzimuthMovementTimeout);

                                    }
                                    if (!cancellationToken.IsCancellationRequested)
                                    {
                                        if (m_CanReadAtHome)
                                        {
                                            LogCallToDriver(p_Name, "About to get AtHome property");
                                            if (domeDevice.AtHome)
                                                LogOK(p_Name, "Dome homed successfully");
                                            else
                                                LogIssue(p_Name, "Home command completed but AtHome is false");
                                        }
                                        else
                                            LogOK(p_Name, "Can't read AtHome so assume that dome has homed successfully");
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
                                LogIssue(p_Name, "CanFindHome is false but FindHome did not throw an exception");
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
                                LogIssue(p_Name, "CanSetShutter is false but OpenShutter did not raise an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.Park:
                        {
                            if (m_CanPark)
                            {
                                SetTest(p_Name);
                                SetAction("Parking");
                                SetStatus("Waiting for movement to stop");
                                try
                                {
                                    LogCallToDriver(p_Name, "About to call Park method");
                                    domeDevice.Park();
                                    if (m_CanReadSlaved)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slaved property");
                                        if (domeDevice.Slaved)
                                            LogIssue(p_Name, "Slaved is true but Park did not raise an exception");
                                    }
                                    if (m_CanReadSlewing)
                                    {
                                        LogCallToDriver(p_Name, "About to get Slewing property repeatedly");
                                        //do
                                        //{
                                        //    WaitFor(SLEEP_TIME);
                                        //    SetStatus("Slewing Status: " + domeDevice.Slewing);
                                        //}
                                        //while (domeDevice.Slewing & !cancellationToken.IsCancellationRequested);
                                        WaitWhile("Parking", () => { return domeDevice.Slewing; }, 500, settings.DomeAzimuthMovementTimeout);

                                    }
                                    if (!cancellationToken.IsCancellationRequested)
                                    {
                                        if (m_CanReadAtPark)
                                        {
                                            LogCallToDriver(p_Name, "About to get AtPark property");
                                            if (domeDevice.AtPark)
                                                LogOK(p_Name, "Dome parked successfully");
                                            else
                                                LogIssue(p_Name, "Park command completed but AtPark is false");
                                        }
                                        else
                                            LogOK(p_Name, "Can't read AtPark so assume that dome has parked successfully");
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
                                LogIssue(p_Name, "CanPark is false but Park did not raise an exception");
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
                                    LogOK(p_Name, "SetPark issued OK");
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
                                LogIssue(p_Name, "CanSetPath is false but SetPath did not throw an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.SlewToAltitude:
                        {
                            if (m_CanSetAltitude)
                            {
                                SetTest(p_Name);
                                for (l_SlewAngle = 0; l_SlewAngle <= 90; l_SlewAngle += 15)
                                {
                                    try
                                    {
                                        DomeSlewToAltitude(p_Name, l_SlewAngle);
                                        if (cancellationToken.IsCancellationRequested) return;
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
                                        LogIssue(p_Name, "No exception generated when slewing to illegal altitude " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees", "Invalid value exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_LOW + " degrees");
                                    }
                                    try
                                    {
                                        DomeSlewToAltitude(p_Name, DOME_ILLEGAL_ALTITUDE_HIGH);
                                        LogIssue(p_Name, "No exception generated when slewing to illegal altitude " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees", "Invalid value exception correctly raised for slew to " + DOME_ILLEGAL_ALTITUDE_HIGH + " degrees");
                                    }
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call SlewToAltitude method");
                                domeDevice.SlewToAltitude(45.0);
                                LogIssue(p_Name, "CanSetAltitude is false but SlewToAltitude did not raise an exception");
                            }

                            break;
                        }

                    case DomePropertyMethod.SlewToAzimuth:
                        {
                            if (m_CanSetAzimuth)
                            {
                                SetTest(p_Name);
                                for (l_SlewAngle = 0; l_SlewAngle <= 315; l_SlewAngle += 45)
                                {
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
                                        LogIssue(p_Name, "No exception generated when slewing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees", "Invalid value exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    try
                                    {
                                        DomeSlewToAzimuth(p_Name, DOME_ILLEGAL_AZIMUTH_HIGH);
                                        LogIssue(p_Name, "No exception generated when slewing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "slew to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees", "Invalid value exception correctly raised for slew to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                }
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call SlewToAzimuth method");
                                domeDevice.SlewToAzimuth(45.0);
                                LogIssue(p_Name, "CanSetAzimuth is false but SlewToAzimuth did not throw an exception");
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
                                                    LogOK(p_Name, "Dome synced OK to within +- 1 degree");
                                                    break;
                                                }

                                            case object _ when Math.Abs(l_NewAzimuth - domeDevice.Azimuth) < 2.0 // close so give it an INFO
                                     :
                                                {
                                                    LogInfo(p_Name, "Dome synced to within +- 2 degrees");
                                                    break;
                                                }

                                            case object _ when Math.Abs(l_NewAzimuth - domeDevice.Azimuth) < 5.0 // Closish so give an issue
                                     :
                                                {
                                                    LogIssue(p_Name, "Dome only synced to within +- 5 degrees");
                                                    break;
                                                }

                                            case object _ when (DOME_SYNC_OFFSET - 2.0) <= Math.Abs(l_NewAzimuth - domeDevice.Azimuth) && Math.Abs(l_NewAzimuth - domeDevice.Azimuth) <= (DOME_SYNC_OFFSET + 2) // Hasn't really moved
                                     :
                                                {
                                                    LogIssue(p_Name, "Dome did not sync, Azimuth didn't change value after sync command");
                                                    break;
                                                }

                                            default:
                                                {
                                                    LogIssue(p_Name, "Dome azimuth was " + Math.Abs(l_NewAzimuth - domeDevice.Azimuth) + " degrees away from expected value");
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
                                        LogOK(p_Name, "Dome successfully synced to 45 degrees but unable to read azimuth to confirm this");
                                    }

                                    // Now test sync to illegal values
                                    try
                                    {
                                        LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                        domeDevice.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_LOW);
                                        LogIssue(p_Name, "No exception generated when syncing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "sync to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees", "Invalid value exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_LOW + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                    try
                                    {
                                        LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                        domeDevice.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_HIGH);
                                        LogIssue(p_Name, "No exception generated when syncing to illegal azimuth " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    catch (Exception ex)
                                    {
                                        HandleInvalidValueExceptionAsOK(p_Name, MemberType.Method, Required.MustBeImplemented, ex, "sync to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees", "Invalid value exception correctly raised for sync to " + DOME_ILLEGAL_AZIMUTH_HIGH + " degrees");
                                    }
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                }
                                else
                                    LogInfo(p_Name, "SyncToAzimuth test skipped since SlewToAzimuth throws an exception");
                            }
                            else
                            {
                                LogCallToDriver(p_Name, "About to call SyncToAzimuth method");
                                domeDevice.SyncToAzimuth(45.0);
                                LogIssue(p_Name, "CanSyncAzimuth is false but SyncToAzimuth did not raise an exception");
                            }

                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "DomeOptionalTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, p_MemberType, Required.Optional, ex, "");
            }
            ClearStatus();
        }
        private void DomeShutterTest(ShutterState p_RequiredShutterState, string p_Name)
        {
            ShutterState l_ShutterState;

            if (settings.DomeOpenShutter)
            {
                SetTest(p_Name);
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
                                    SetAction("Opening shutter ready for close test");
                                    LogDebug(p_Name, "Opening shutter ready for close test");
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
                                    SetAction("Waiting for shutter to close before opening ready for close test");
                                    LogDebug(p_Name, "Waiting for shutter to close before opening ready for close test");
                                    if (!DomeShutterWait(ShutterState.Closed))
                                        return; // Wait for shutter to close
                                    LogDebug(p_Name, "Opening shutter ready for close test");
                                    SetAction("Opening shutter ready for close test");
                                    LogCallToDriver(p_Name, "About to call OpenShutter method");
                                    domeDevice.OpenShutter(); // Then open it
                                    if (!DomeShutterWait(ShutterState.Open))
                                        return;
                                    DomeStabliisationWait();
                                }
                                else
                                {
                                    SetAction("Waiting for shutter to close ready for open test");
                                    LogDebug(p_Name, "Waiting for shutter to close ready for open test");
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
                                    SetAction("Waiting for shutter to open ready for close test");
                                    LogDebug(p_Name, "Waiting for shutter to open ready for close test");
                                    if (!DomeShutterWait(ShutterState.Open))
                                        return; // Wait for shutter to open
                                    DomeStabliisationWait();
                                }
                                else
                                {
                                    SetAction("Waiting for shutter to open before closing ready for open test");
                                    LogDebug(p_Name, "Waiting for shutter to open before closing ready for open test");
                                    if (!DomeShutterWait(ShutterState.Open))
                                        return; // Wait for shutter to open
                                    LogDebug(p_Name, "Closing shutter ready for open test");
                                    SetAction("Closing shutter ready for open test");
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
                                LogIssue("DomeShutterTest", $"Shutter state is Error: {l_ShutterState}");
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
                                    SetAction("Closing shutter ready for open  test");
                                    LogDebug(p_Name, "Closing shutter ready for open test");
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
                                LogIssue("DomeShutterTest", "Unexpected shutter status: " + l_ShutterState.ToString());
                                break;
                            }
                    }

                    // Now test that we can get to the required state
                    if (p_RequiredShutterState == ShutterState.Closed)
                    {
                        // Shutter is now open so close it
                        SetAction("Closing shutter");
                        LogCallToDriver(p_Name, "About to call CloseShutter method");
                        domeDevice.CloseShutter();
                        SetAction("Waiting for shutter to close");
                        LogDebug(p_Name, "Waiting for shutter to close");
                        if (!DomeShutterWait(ShutterState.Closed))
                        {
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            l_ShutterState = domeDevice.ShutterStatus;
                            LogIssue(p_Name, "Unable to close shutter - ShutterStatus: " + l_ShutterState.ToString());
                            return;
                        }
                        else
                            LogOK(p_Name, "Shutter closed successfully");
                        DomeStabliisationWait();
                    }
                    else
                    {
                        SetAction("Opening shutter");
                        domeDevice.OpenShutter();
                        SetAction("Waiting for shutter to open");
                        LogDebug(p_Name, "Waiting for shutter to open");
                        if (!DomeShutterWait(ShutterState.Open))
                        {
                            LogCallToDriver(p_Name, "About to get ShutterStatus property");
                            l_ShutterState = domeDevice.ShutterStatus;
                            LogIssue(p_Name, "Unable to open shutter - ShutterStatus: " + l_ShutterState.ToString());
                            return;
                        }
                        else
                            LogOK(p_Name, "Shutter opened successfully");
                        DomeStabliisationWait();
                    }
                }
                else
                {
                    LogDebug(p_Name, "Can't read shutter status!");
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
                    LogOK(p_Name, "Command issued successfully but can't read ShutterStatus to confirm shutter is closed");
                }

                ClearStatus();
            }
            else
                LogTestAndMessage("DomeSafety", "Open shutter check box is unchecked so shutter test bypassed");
        }

        private bool DomeShutterWait(ShutterState p_RequiredStatus)
        {
            DateTime l_StartTime;
            // Wait for shutter to reach required stats or user presses stop or timeout occurs
            // Returns true if required state is reached
            bool returnValue = false;
            l_StartTime = DateTime.Now;
            try
            {
                LogCallToDriver("DomeShutterWait", "About to get ShutterStatus property multiple times");
                WaitWhile($"Waiting for shutter state {p_RequiredStatus}", () => { return (domeDevice.ShutterStatus != p_RequiredStatus); }, 500, settings.DomeShutterMovementTimeout);

                if ((domeDevice.ShutterStatus == p_RequiredStatus)) returnValue = true; // All worked so return True

                if ((DateTime.Now.Subtract(l_StartTime).TotalSeconds > settings.DomeShutterMovementTimeout))
                    LogIssue("DomeShutterWait", "Timed out waiting for shutter to reach state: " + p_RequiredStatus.ToString() + ", consider increasing the timeout setting in Options / Conformance Options");
            }
            catch (Exception ex)
            {
                LogIssue("DomeShutterWait", "Unexpected exception: " + ex.ToString());
            }

            return returnValue;
        }

        private void DomePerformanceTest(DomePropertyMethod p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime;
            double l_Rate;
            bool l_Boolean;
            double l_Double;
            ShutterState l_ShutterState;
            SetTest("Performance Testing");
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
                                LogIssue(p_Name, "DomePerformanceTest: Unknown test type " + p_Type.ToString());
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

        public void DomeStabliisationWait()
        {
            // Only wait if a non-zero wait time has been configured
            if (settings.DomeStabilisationWaitTime > 0)
            {
                Stopwatch sw = Stopwatch.StartNew();
                WaitWhile("Waiting for dome to stabilise", () => { return sw.Elapsed.TotalSeconds < settings.DomeStabilisationWaitTime; }, 500, settings.DomeStabilisationWaitTime);
            }
        }
    }
}

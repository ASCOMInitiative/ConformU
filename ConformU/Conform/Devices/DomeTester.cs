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

    internal class DomeTester : DeviceTesterBaseClass
    {
        private const double DOME_SYNC_OFFSET = 45.0; // Amount to offset the azimuth when testing ability to sync
        private const double DOME_ILLEGAL_ALTITUDE_LOW = -10.0; // Illegal value to test dome driver exception generation
        private const double DOME_ILLEGAL_ALTITUDE_HIGH = 100.0; // Illegal value to test dome driver exception generation
        private const double DOME_ILLEGAL_AZIMUTH_LOW = -10.0; // Illegal value to test dome driver exception generation
        private const double DOME_ILLEGAL_AZIMUTH_HIGH = 370.0; // Illegal value to test dome driver exception generation

        // Dome variables
        private bool mCanSetAltitude, mCanSetAzimuth, mCanSetShutter, mCanSlave, mCanSyncAzimuth, mSlaved;
        private ShutterState mShutterStatus;
        private bool mCanReadAltitude, mCanReadAtPark, mCanReadAtHome, mCanReadSlewing, mCanReadSlaved, mCanReadShutterStatus, mCanReadAzimuth, mCanSlewToAzimuth;

        // General variables
        private bool mSlewing, mAtHome, mAtPark, mCanFindHome, mCanPark, mCanSetPark, mConnected;
        private string mDescription, mDriverINfo, mName;
        private short mInterfaceVersion;
        private double mAltitude, mAzimuth;


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

        private IDomeV3 domeDevice;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose
        public DomeTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, true, false, false, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
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
                    domeDevice?.Dispose();
                    domeDevice = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        #region Conform Process

        public override void InitialiseTest()
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
                                                    Assembly.GetExecutingAssembly().GetName().Version.ToString(4),
                                                    settings.AlpacaConfiguration.TrustUserGeneratedSslCertificates);

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
                SetDevice(domeDevice, DeviceTypes.Dome); // Assign the driver to the base class

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

        public override void ReadCanProperties()
        {
            DomeMandatoryTest(DomePropertyMethod.CanFindHome, "CanFindHome");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.CanPark, "CanPark");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.CanSetAltitude, "CanSetAltitude");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.CanSetAzimuth, "CanSetAzimuth");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.CanSetPark, "CanSetPark");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.CanSetShutter, "CanSetShutter");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.CanSlave, "CanSlave");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.CanSyncAzimuth, "CanSyncAzimuth");
            if (cancellationToken.IsCancellationRequested)
                return;
        }

        public override void PreRunCheck()
        {
            int lVStringPtr, lV1, lV2, lV3;
            string lVString;

            // Add a test for a back level version of the Dome simulator - just abandon this process if any errors occur
            if (settings.ComDevice.ProgId.ToUpper() == "DOMESIM.DOME")
            {
                try
                {
                    LogCallToDriver("PreRunCheck", "About to get DriverInfo property");
                    lVStringPtr = domeDevice.DriverInfo.ToUpper().IndexOf("ASCOM DOME SIMULATOR "); // Point at the start of the version string
                    if (lVStringPtr > 0)
                    {
                        LogCallToDriver("PreRunCheck", "About to get DriverInfo property");
                        lVString = domeDevice.DriverInfo.ToUpper()[(lVStringPtr + 21)..]; // Get the version string
                        lVStringPtr = lVString.IndexOf(".");
                        if (lVStringPtr > 1)
                        {
                            lV1 = System.Convert.ToInt32(lVString[1..lVStringPtr]); // Extract the number
                            lVString = lVString[(lVStringPtr + 1)..]; // Get the second version number part
                            lVStringPtr = lVString.IndexOf(".");
                            if (lVStringPtr > 1)
                            {
                                lV2 = int.Parse(lVString[1..lVStringPtr]); // Extract the number
                                lVString = lVString[(lVStringPtr + 1)..]; // Get the third version number part
                                                                          // Find the next non numeric character
                                lVStringPtr = 0;
                                do
                                    lVStringPtr += 1;
                                while (int.TryParse(lVString.AsSpan(lVStringPtr, 1), out _));

                                if (lVStringPtr > 1)
                                {
                                    lV3 = System.Convert.ToInt32(lVString[1..lVStringPtr]); // Extract the number
                                                                                            // Turn the version parts into a whole number
                                    lV1 = lV1 * 1000000 + lV2 * 1000 + lV3;
                                    if (lV1 < 5000007)
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
                    mSlewing = domeDevice.Slewing; // Try to read the Slewing property
                    if (mSlewing)
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
                            LogTestAndMessage("DomeSafety",
                                $"Stop button pressed, further testing abandoned, shutter status: {domeDevice.ShutterStatus}");
                        }
                        else
                        {
                            LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                            if (domeDevice.ShutterStatus == ShutterState.Open)
                            {
                                LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                                LogOk("DomeSafety", $"Shutter status: {domeDevice.ShutterStatus}");
                            }
                            else
                            {
                                LogCallToDriver("PreRunCheck", "About to get ShutterStatus property");
                                LogIssue("DomeSafety", $"Shutter status: {domeDevice.ShutterStatus}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogTestAndMessage("DomeSafety", $"Unable to open shutter, some tests may fail: {ex.Message}");
                    }
                    SetTest("");
                }
                else
                    LogTestAndMessage("DomeSafety", "Open shutter check box is unchecked so shutter not opened");
            }
        }

        public override void CheckProperties()
        {
            if (!settings.DomeOpenShutter)
                LogInfo("Altitude", "You have configured Conform not to open the shutter so the following test may fail.");

            DomeOptionalTest(DomePropertyMethod.Altitude, MemberType.Property, "Altitude");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeOptionalTest(DomePropertyMethod.AtHome, MemberType.Property, "AtHome");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeOptionalTest(DomePropertyMethod.AtPark, MemberType.Property, "AtPark");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeOptionalTest(DomePropertyMethod.Azimuth, MemberType.Property, "Azimuth");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeOptionalTest(DomePropertyMethod.ShutterStatus, MemberType.Property, "ShutterStatus");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.SlavedRead, "Slaved Read");
            if (cancellationToken.IsCancellationRequested)
                return;

            if (mSlaved & (!mCanSlave))
                LogIssue("Slaved Read", "Dome is slaved but CanSlave is false");

            DomeOptionalTest(DomePropertyMethod.SlavedWrite, MemberType.Property, "Slaved Write");
            if (cancellationToken.IsCancellationRequested)
                return;

            DomeMandatoryTest(DomePropertyMethod.Slewing, "Slewing");
            if (cancellationToken.IsCancellationRequested)
                return;
        }

        public override void CheckMethods()
        {
            DomeMandatoryTest(DomePropertyMethod.AbortSlew, "AbortSlew");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.SlewToAltitude, MemberType.Method, "SlewToAltitude");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.SlewToAzimuth, MemberType.Method, "SlewToAzimuth");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.SyncToAzimuth, MemberType.Method, "SyncToAzimuth");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.CloseShutter, MemberType.Method, "CloseShutter");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.OpenShutter, MemberType.Method, "OpenShutter");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.FindHome, MemberType.Method, "FindHome");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.Park, MemberType.Method, "Park");
            if (cancellationToken.IsCancellationRequested) return;

            DomeOptionalTest(DomePropertyMethod.SetPark, MemberType.Method, "SetPark");
            if (cancellationToken.IsCancellationRequested) return; // SetPark must follow Park
        }

        public override void CheckPerformance()
        {
            if (mCanReadAltitude)
            {
                DomePerformanceTest(DomePropertyMethod.Altitude, "Altitude"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (mCanReadAzimuth)
            {
                DomePerformanceTest(DomePropertyMethod.Azimuth, "Azimuth"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (mCanReadShutterStatus)
            {
                DomePerformanceTest(DomePropertyMethod.ShutterStatus, "ShutterStatus"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (mCanReadSlaved)
            {
                DomePerformanceTest(DomePropertyMethod.SlavedRead, "Slaved"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
            if (mCanReadSlewing)
            {
                DomePerformanceTest(DomePropertyMethod.Slewing, "Slewing"); if (cancellationToken.IsCancellationRequested)
                    return;
            }
        }

        public override void PostRunCheck()
        {
            if (settings.DomeOpenShutter)
            {
                if (mCanSetShutter)
                {
                    LogInfo("DomeSafety", "Attempting to close shutter...");
                    try // Close shutter
                    {
                        LogCallToDriver("DomeSafety", "About to call CloseShutter");
                        domeDevice.CloseShutter();
                        DomeShutterWait(ShutterState.Closed);
                        LogOk("DomeSafety", "Shutter successfully closed");
                    }
                    catch (Exception ex)
                    {
                        LogTestAndMessage("DomeSafety", $"Exception closing shutter: {ex.Message}");
                        LogTestAndMessage("DomeSafety", "Please close shutter manually");
                    }
                }
                else
                    LogInfo("DomeSafety", "CanSetShutter is false, please close the shutter manually");
            }
            else
                LogInfo("DomeSafety", "Open shutter check box is unchecked so close shutter bypassed");
            // 3.0.0.17 - Added check for CanPark
            if (mCanPark)
            {
                LogInfo("DomeSafety", "Attempting to park dome...");
                try // Park
                {
                    LogCallToDriver("DomeSafety", "About to call Park");
                    domeDevice.Park();
                    DomeWaitForSlew(settings.DomeAzimuthMovementTimeout, null);
                    LogOk("DomeSafety", "Dome successfully parked");
                }
                catch (Exception)
                {
                    LogIssue("DomeSafety", "Exception generated, unable to park dome");
                }
            }
            else
                LogInfo("DomeSafety", "CanPark is false - skipping dome parking");
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
                if (!settings.DomeOpenShutter)
                    LogConfigurationAlert("Shutter related tests were omitted due to Conform configuration.");

            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n:{ex}");
            }
        }

        #endregion

        #region Support Code

        private void DomeSlewToAltitude(string pName, double pAltitude)
        {
            Stopwatch sw = Stopwatch.StartNew();

            if (!settings.DomeOpenShutter) LogInfo("SlewToAltitude", "You have configured Conform not to open the shutter so the following slew may fail.");

            SetTest("SlewToAltitude");
            SetAction($"Slewing to altitude {pAltitude} degrees");
            LogCallToDriver(pName, "About to call SlewToAltitude");
            TimeMethod(pName, () => domeDevice.SlewToAltitude(pAltitude), TargetTime.Standard);

            // Check whether Slewing can be read
            if (mCanReadSlewing) // Slewing can be read OK
            {
                // Check whether Slewing was set on return
                LogCallToDriver(pName, "About to get Slewing property");
                if (domeDevice.Slewing) // Asynchronous operation so wait for the slew to complete
                {
                    DomeWaitForSlew(settings.DomeAltitudeMovementTimeout, () => $"{domeDevice.Altitude:00} / {pAltitude:00} degrees"); if (cancellationToken.IsCancellationRequested) return;
                    LogOk($"{pName} {pAltitude}", "Asynchronous slew OK");
                }
                else // Synchronous operation
                {
                    // Check whether this is a Platform 7 or later device and message accordingly
                    if (IsPlatform7OrLater) // Platform 7 or later interface
                    {
                        if (sw.Elapsed.TotalSeconds < standardTargetResponseTime)

                            LogOk($"{pName} {pAltitude}", "Synchronous slew OK");
                        else
                        {
                            LogIssue($"{pName} {pAltitude}", $"Synchronous slew took {sw.Elapsed.TotalSeconds} seconds, which is longer than the standard response target: {standardTargetResponseTime} seconds.,");
                        }
                    }
                    else // Platform 6 interface
                        LogOk($"{pName} {pAltitude}", "Synchronous slew OK");

                }
            }
            else // Slewing can't be read
            {
                LogOk($"{pName} {pAltitude}", "Can't read Slewing so assume synchronous slew OK");
            }
            DomeStabliisationWait();

            // Check whether the reported altitude matches the requested altitude
            if (mCanReadAltitude)
            {
                LogCallToDriver(pName, "About to get Altitude property");
                double altitude = domeDevice.Altitude;

                if (Math.Abs(altitude - pAltitude) <= settings.DomeSlewTolerance)
                {
                    LogOk($"{pName} {pAltitude}", $"Reached the required altitude: {pAltitude:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported altitude: {altitude:0.0} degrees");
                }
                else
                {
                    LogIssue($"{pName} {pAltitude}", $"Failed to reach the required altitude: {pAltitude:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported altitude: {altitude:0.0} degrees");
                }
            }
        }

        private void DomeSlewToAzimuth(string pName, double pAzimuth)
        {
            SetAction($"Slewing to azimuth {pAzimuth} degrees");
            if (pAzimuth >= 0.0 & pAzimuth <= 359.9999999)
            {
                mCanSlewToAzimuth = false;
                LogCallToDriver(pName, "About to call SlewToAzimuth");
                TimeMethod(pName, () => domeDevice.SlewToAzimuth(pAzimuth), TargetTime.Standard);
                mCanSlewToAzimuth = true; // Command is supported and didn't generate an exception
            }
            else
            {
                LogCallToDriver(pName, "About to call SlewToAzimuth");
                TimeMethod(pName, () => domeDevice.SlewToAzimuth(pAzimuth), TargetTime.Standard);
            }

            if (mCanReadSlewing)
            {
                LogCallToDriver(pName, "About to get Slewing property");
                if (domeDevice.Slewing)
                {
                    DomeWaitForSlew(settings.DomeAzimuthMovementTimeout, () => $"{domeDevice.Azimuth:000} / {pAzimuth:000} degrees"); if (cancellationToken.IsCancellationRequested) return;
                    LogOk($"{pName} {pAzimuth}", "Asynchronous slew OK");
                }
                else
                {
                    LogOk($"{pName} {pAzimuth}", "Synchronous slew OK");
                }
            }
            else
            {
                LogOk($"{pName} {pAzimuth}", "Can't read Slewing so assume synchronous slew OK");
            }
            DomeStabliisationWait();

            // Check whether the reported azimuth matches the requested azimuth
            if (mCanReadAzimuth)
            {
                LogCallToDriver(pName, "About to get Azimuth property");
                double azimuth = domeDevice.Azimuth;

                if (Math.Abs(azimuth - pAzimuth) <= settings.DomeSlewTolerance)
                {
                    LogOk($"{pName} {pAzimuth}", $"Reached the required azimuth: {pAzimuth:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported azimuth: {azimuth:0.0}");
                }
                else
                {
                    LogIssue($"{pName} {pAzimuth}", $"Failed to reach the required azimuth: {pAzimuth:0.0} within tolerance ±{settings.DomeSlewTolerance} degrees. Reported azimuth: {azimuth:0.0}");
                }
            }

        }

        private void DomeWaitForSlew(double pTimeOut, Func<string> reportingFunction)
        {
            DateTime lStartTime;
            lStartTime = DateTime.Now;

            WaitWhile(GetAction(), () => domeDevice.Slewing, 500, Convert.ToInt32(pTimeOut), reportingFunction);

            SetStatus("");
            if ((DateTime.Now.Subtract(lStartTime).TotalSeconds > pTimeOut))
            {
                LogIssue("DomeWaitForSlew", "Timed out waiting for Dome slew, consider increasing time-outs in Options/Conform Options.");
                LogInfo("DomeWaitForSlew", "Another cause of time-outs is if your Slewing Property logic is inverted or is not operating correctly.");
            }
        }

        private void DomeMandatoryTest(DomePropertyMethod pType, string pName)
        {
            try
            {
                TimeMethod(pName, () =>
                {
                    switch (pType)
                    {
                        case DomePropertyMethod.CanFindHome:
                            {
                                LogCallToDriver(pName, "About to get CanFindHome property");
                                mCanFindHome = TimeFunc(pName, () => domeDevice.CanFindHome, TargetTime.Fast);
                                LogOk(pName, mCanFindHome.ToString());
                                break;
                            }

                        case DomePropertyMethod.CanPark:
                            {
                                LogCallToDriver(pName, "About to get CanPark property");
                                mCanPark = domeDevice.CanPark;
                                LogOk(pName, mCanPark.ToString());
                                break;
                            }

                        case DomePropertyMethod.CanSetAltitude:
                            LogCallToDriver(pName, "About to get CanSetAltitude property");
                            mCanSetAltitude = domeDevice.CanSetAltitude;
                            LogOk(pName, mCanSetAltitude.ToString());
                            break;

                        case DomePropertyMethod.CanSetAzimuth:
                            {
                                LogCallToDriver(pName, "About to get CanSetAzimuth property");
                                mCanSetAzimuth = domeDevice.CanSetAzimuth;
                                LogOk(pName, mCanSetAzimuth.ToString());
                                break;
                            }

                        case DomePropertyMethod.CanSetPark:
                            {
                                LogCallToDriver(pName, "About to get CanSetPark property");
                                mCanSetPark = domeDevice.CanSetPark;
                                LogOk(pName, mCanSetPark.ToString());
                                break;
                            }

                        case DomePropertyMethod.CanSetShutter:
                            {
                                LogCallToDriver(pName, "About to get CanSetShutter property");
                                mCanSetShutter = domeDevice.CanSetShutter;
                                LogOk(pName, mCanSetShutter.ToString());
                                break;
                            }

                        case DomePropertyMethod.CanSlave:
                            {
                                LogCallToDriver(pName, "About to get CanSlave property");
                                mCanSlave = domeDevice.CanSlave;
                                LogOk(pName, mCanSlave.ToString());
                                break;
                            }

                        case DomePropertyMethod.CanSyncAzimuth:
                            {
                                LogCallToDriver(pName, "About to get CanSyncAzimuth property");
                                mCanSyncAzimuth = domeDevice.CanSyncAzimuth;
                                LogOk(pName, mCanSyncAzimuth.ToString());
                                break;
                            }

                        case DomePropertyMethod.Connected:
                            {
                                LogCallToDriver(pName, "About to get Connected property");
                                mConnected = domeDevice.Connected;
                                LogOk(pName, mConnected.ToString());
                                break;
                            }

                        case DomePropertyMethod.Description:
                            {
                                LogCallToDriver(pName, "About to get Description property");
                                mDescription = domeDevice.Description;
                                LogOk(pName, mDescription.ToString());
                                break;
                            }

                        case DomePropertyMethod.DriverInfo:
                            {
                                LogCallToDriver(pName, "About to get DriverInfo property");
                                mDriverINfo = domeDevice.DriverInfo;
                                LogOk(pName, mDriverINfo.ToString());
                                break;
                            }

                        case DomePropertyMethod.InterfaceVersion:
                            {
                                LogCallToDriver(pName, "About to get InterfaceVersion property");
                                mInterfaceVersion = domeDevice.InterfaceVersion;
                                LogOk(pName, mInterfaceVersion.ToString());
                                break;
                            }

                        case DomePropertyMethod.Name:
                            {
                                LogCallToDriver(pName, "About to get Name property");
                                mName = domeDevice.Name;
                                LogOk(pName, mName.ToString());
                                break;
                            }

                        case DomePropertyMethod.SlavedRead:
                            {
                                mCanReadSlaved = false;
                                LogCallToDriver(pName, "About to get Slaved property");
                                mSlaved = domeDevice.Slaved;
                                mCanReadSlaved = true;
                                LogOk(pName, mSlaved.ToString());
                                break;
                            }

                        case DomePropertyMethod.Slewing:
                            {
                                mCanReadSlewing = false;
                                LogCallToDriver(pName, "About to get Slewing property");
                                mSlewing = domeDevice.Slewing;
                                mCanReadSlewing = true;
                                LogOk(pName, mSlewing.ToString());
                                break;
                            }

                        case DomePropertyMethod.AbortSlew:
                            {
                                LogCallToDriver(pName, "About to call AbortSlew method");
                                domeDevice.AbortSlew();
                                // Confirm that slaved is false
                                if (mCanReadSlaved)
                                {
                                    LogCallToDriver(pName, "About to get Slaved property");
                                    if (domeDevice.Slaved)
                                        LogIssue("AbortSlew", "Slaved property Is true after AbortSlew");
                                    else
                                        LogOk("AbortSlew", "AbortSlew command issued successfully");
                                }
                                else
                                    LogOk("AbortSlew", "Can't read Slaved property AbortSlew command was successful");
                                break;
                            }

                        default:
                            LogIssue(pName, $"DomeMandatoryTest: Unknown test type {pType}");
                            break;
                    }
                }, TargetTime.Fast);
            }
            catch (Exception ex)
            {
                HandleException(pName, MemberType.Property, Required.Mandatory, ex, "");
            }
        }

        private void DomeOptionalTest(DomePropertyMethod pType, MemberType pMemberType, string pName)
        {
            double lSlewAngle, lOriginalAzimuth, lNewAzimuth;
            try
            {
                switch (pType)
                {
                    case DomePropertyMethod.Altitude:
                        mCanReadAltitude = false;
                        LogCallToDriver(pName, "About to get Altitude property");
                        TimeMethod(pName, () => mAltitude = domeDevice.Altitude, TargetTime.Fast);
                        mCanReadAltitude = true;
                        LogOk(pName, mAltitude.ToString());
                        break;

                    case DomePropertyMethod.AtHome:
                        mCanReadAtHome = false;
                        LogCallToDriver(pName, "About to get AtHome property");
                        TimeMethod(pName, () => mAtHome = domeDevice.AtHome, TargetTime.Fast);
                        mCanReadAtHome = true;
                        LogOk(pName, mAtHome.ToString());
                        break;

                    case DomePropertyMethod.AtPark:
                        mCanReadAtPark = false;
                        LogCallToDriver(pName, "About to get AtPark property");
                        TimeMethod(pName, () => mAtPark = domeDevice.AtPark, TargetTime.Fast);
                        mCanReadAtPark = true;
                        LogOk(pName, mAtPark.ToString());
                        break;

                    case DomePropertyMethod.Azimuth:
                        mCanReadAzimuth = false;
                        LogCallToDriver(pName, "About to get Azimuth property");
                        TimeMethod(pName, () => mAzimuth = domeDevice.Azimuth, TargetTime.Fast);
                        mCanReadAzimuth = true;
                        LogOk(pName, mAzimuth.ToString());
                        break;

                    case DomePropertyMethod.ShutterStatus:
                        mCanReadShutterStatus = false;
                        LogCallToDriver(pName, "About to get ShutterStatus property");
                        TimeMethod(pName, () => mShutterStatus = domeDevice.ShutterStatus, TargetTime.Fast);
                        mCanReadShutterStatus = true;
                        LogOk(pName, mShutterStatus.ToString());
                        break;

                    case DomePropertyMethod.SlavedWrite:
                        if (mCanSlave)
                        {
                            if (mCanReadSlaved)
                            {
                                if (mSlaved)
                                {
                                    LogCallToDriver(pName, "About to set Slaved property");
                                    TimeMethod(pName, () => domeDevice.Slaved = false, TargetTime.Standard);
                                }
                                else
                                {
                                    LogCallToDriver(pName, "About to set Slaved property");
                                    TimeMethod(pName, () => domeDevice.Slaved = true, TargetTime.Standard);
                                }
                                LogCallToDriver(pName, "About to set Slaved property");
                                domeDevice.Slaved = mSlaved; // Restore original value
                                LogOk("Slaved Write", "Slave state changed successfully");
                            }
                            else
                                LogInfo("Slaved Write", "Test skipped since Slaved property can't be read");
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to set Slaved property");
                            domeDevice.Slaved = true;
                            LogIssue(pName, "CanSlave is false but setting Slaved true did not raise an exception");
                            LogCallToDriver(pName, "About to set Slaved property");
                            domeDevice.Slaved = false; // Un-slave to continue tests
                        }

                        break;

                    case DomePropertyMethod.CloseShutter:
                        if (mCanSetShutter)
                        {
                            try
                            {
                                DomeShutterTest(ShutterState.Closed, pName);
                                DomeStabliisationWait();
                            }
                            catch (Exception ex)
                            {
                                HandleException(pName, MemberType.Method, Required.MustBeImplemented, ex, "CanSetShutter is True");
                            }
                        }
                        else
                        {
                            domeDevice.CloseShutter();
                            LogIssue(pName, "CanSetShutter is false but CloseShutter did not raise an exception");
                        }

                        break;

                    case DomePropertyMethod.FindHome:
                        if (mCanFindHome)
                        {
                            SetTest(pName);
                            SetAction("Finding home");
                            SetStatus("Waiting for movement to stop");
                            try
                            {
                                LogCallToDriver(pName, "About to call FindHome method");
                                TimeMethod(pName, () => domeDevice.FindHome(), TargetTime.Standard);
                                if (mCanReadSlaved)
                                {
                                    LogCallToDriver(pName, "About to get Slaved Property");
                                    if (domeDevice.Slaved)
                                        LogIssue(pName, "Slaved is true but Home did not raise an exception");
                                }
                                if (mCanReadSlewing)
                                {
                                    LogCallToDriver(pName, "About to get Slewing property repeatedly");
                                    WaitWhile("Finding home", () => domeDevice.Slewing, 500, settings.DomeAzimuthMovementTimeout);
                                }
                                if (!cancellationToken.IsCancellationRequested)
                                {
                                    if (mCanReadAtHome)
                                    {
                                        LogCallToDriver(pName, "About to get AtHome property");
                                        if (domeDevice.AtHome)
                                            LogOk(pName, "Dome homed successfully");
                                        else
                                            LogIssue(pName, "Home command completed but AtHome is false");
                                    }
                                    else
                                        LogOk(pName, "Can't read AtHome so assume that dome has homed successfully");
                                    DomeStabliisationWait();
                                }
                            }
                            catch (Exception ex)
                            {
                                HandleException(pName, MemberType.Method, Required.MustBeImplemented, ex, "CanFindHome is True");
                                DomeStabliisationWait();
                            }
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to call FindHome method");
                            domeDevice.FindHome();
                            LogIssue(pName, "CanFindHome is false but FindHome did not throw an exception");
                        }

                        break;

                    case DomePropertyMethod.OpenShutter:
                        if (mCanSetShutter)
                        {
                            try
                            {
                                DomeShutterTest(ShutterState.Open, pName);
                                DomeStabliisationWait();
                            }
                            catch (Exception ex)
                            {
                                HandleException(pName, MemberType.Method, Required.MustBeImplemented, ex, "CanSetShutter is True");
                            }
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to call OpenShutter method");
                            domeDevice.OpenShutter();
                            LogIssue(pName, "CanSetShutter is false but OpenShutter did not raise an exception");
                        }

                        break;

                    case DomePropertyMethod.Park:
                        if (mCanPark)
                        {
                            SetTest(pName);
                            SetAction("Parking");
                            SetStatus("Waiting for movement to stop");
                            try
                            {
                                LogCallToDriver(pName, "About to call Park method");
                                TimeMethod(pName, () => domeDevice.Park(), TargetTime.Standard);
                                if (mCanReadSlaved)
                                {
                                    LogCallToDriver(pName, "About to get Slaved property");
                                    if (domeDevice.Slaved)
                                        LogIssue(pName, "Slaved is true but Park did not raise an exception");
                                }
                                if (mCanReadSlewing)
                                {
                                    LogCallToDriver(pName, "About to get Slewing property repeatedly");
                                    WaitWhile("Parking", () => domeDevice.Slewing, 500, settings.DomeAzimuthMovementTimeout);
                                }
                                if (!cancellationToken.IsCancellationRequested)
                                {
                                    if (mCanReadAtPark)
                                    {
                                        LogCallToDriver(pName, "About to get AtPark property");
                                        if (domeDevice.AtPark)
                                            LogOk(pName, "Dome parked successfully");
                                        else
                                            LogIssue(pName, "Park command completed but AtPark is false");
                                    }
                                    else
                                        LogOk(pName, "Can't read AtPark so assume that dome has parked successfully");
                                }
                                DomeStabliisationWait();
                            }
                            catch (Exception ex)
                            {
                                HandleException(pName, MemberType.Method, Required.MustBeImplemented, ex, "CanPark is True");
                                DomeStabliisationWait();
                            }
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to call Park method");
                            domeDevice.Park();
                            LogIssue(pName, "CanPark is false but Park did not raise an exception");
                        }

                        break;

                    case DomePropertyMethod.SetPark:
                        if (mCanSetPark)
                        {
                            try
                            {
                                LogCallToDriver(pName, "About to call SetPark method");
                                TimeMethod(pName, () => domeDevice.SetPark(), TargetTime.Standard);
                                LogOk(pName, "SetPark issued OK");
                            }
                            catch (Exception ex)
                            {
                                HandleException(pName, MemberType.Method, Required.MustBeImplemented, ex, "CanSetPark is True");
                            }
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to call SetPark method");
                            domeDevice.SetPark();
                            LogIssue(pName, "CanSetPath is false but SetPath did not throw an exception");
                        }

                        break;

                    case DomePropertyMethod.SlewToAltitude:
                        if (mCanSetAltitude)
                        {
                            SetTest(pName);
                            for (lSlewAngle = 0; lSlewAngle <= 90; lSlewAngle += 15)
                            {
                                try
                                {
                                    DomeSlewToAltitude(pName, lSlewAngle);
                                    if (cancellationToken.IsCancellationRequested) return;
                                }
                                catch (Exception ex)
                                {
                                    HandleException(pName, MemberType.Method, Required.MustBeImplemented, ex, "CanSetAltitude is True");
                                }
                            }

                            // Test out of range values -10 and 100 degrees
                            if (mCanSetAltitude)
                            {
                                try
                                {
                                    DomeSlewToAltitude(pName, DOME_ILLEGAL_ALTITUDE_LOW);
                                    LogIssue(pName,
                                        $"No exception generated when slewing to illegal altitude {DOME_ILLEGAL_ALTITUDE_LOW} degrees");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.MustBeImplemented, ex,
                                        $"slew to {DOME_ILLEGAL_ALTITUDE_LOW} degrees",
                                        $"Invalid value exception correctly raised for slew to {DOME_ILLEGAL_ALTITUDE_LOW} degrees");
                                }
                                try
                                {
                                    DomeSlewToAltitude(pName, DOME_ILLEGAL_ALTITUDE_HIGH);
                                    LogIssue(pName,
                                        $"No exception generated when slewing to illegal altitude {DOME_ILLEGAL_ALTITUDE_HIGH} degrees");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.MustBeImplemented, ex,
                                        $"slew to {DOME_ILLEGAL_ALTITUDE_HIGH} degrees",
                                        $"Invalid value exception correctly raised for slew to {DOME_ILLEGAL_ALTITUDE_HIGH} degrees");
                                }
                            }
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to call SlewToAltitude method");
                            domeDevice.SlewToAltitude(45.0);
                            LogIssue(pName, "CanSetAltitude is false but SlewToAltitude did not raise an exception");
                        }

                        break;

                    case DomePropertyMethod.SlewToAzimuth:
                        if (mCanSetAzimuth)
                        {
                            SetTest(pName);
                            for (lSlewAngle = 0; lSlewAngle <= 315; lSlewAngle += 45)
                            {
                                try
                                {
                                    DomeSlewToAzimuth(pName, lSlewAngle);
                                    if (cancellationToken.IsCancellationRequested)
                                        return;
                                }
                                catch (Exception ex)
                                {
                                    HandleException(pName, MemberType.Method, Required.MustBeImplemented, ex, "CanSetAzimuth is True");
                                }
                            }

                            if (mCanSetAzimuth)
                            {
                                // Test out of range values -10 and 370 degrees
                                try
                                {
                                    DomeSlewToAzimuth(pName, DOME_ILLEGAL_AZIMUTH_LOW);
                                    LogIssue(pName,
                                        $"No exception generated when slewing to illegal azimuth {DOME_ILLEGAL_AZIMUTH_LOW} degrees");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.MustBeImplemented, ex,
                                        $"slew to {DOME_ILLEGAL_AZIMUTH_LOW} degrees",
                                        $"Invalid value exception correctly raised for slew to {DOME_ILLEGAL_AZIMUTH_LOW} degrees");
                                }
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                try
                                {
                                    DomeSlewToAzimuth(pName, DOME_ILLEGAL_AZIMUTH_HIGH);
                                    LogIssue(pName,
                                        $"No exception generated when slewing to illegal azimuth {DOME_ILLEGAL_AZIMUTH_HIGH} degrees");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.MustBeImplemented, ex,
                                        $"slew to {DOME_ILLEGAL_AZIMUTH_HIGH} degrees",
                                        $"Invalid value exception correctly raised for slew to {DOME_ILLEGAL_AZIMUTH_HIGH} degrees");
                                }
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                            }
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to call SlewToAzimuth method");
                            domeDevice.SlewToAzimuth(45.0);
                            LogIssue(pName, "CanSetAzimuth is false but SlewToAzimuth did not throw an exception");
                        }

                        break;

                    case DomePropertyMethod.SyncToAzimuth:
                        if (mCanSyncAzimuth)
                        {
                            if (mCanSlewToAzimuth)
                            {
                                if (mCanReadAzimuth)
                                {
                                    LogCallToDriver(pName, "About to get Azimuth property");
                                    lOriginalAzimuth = domeDevice.Azimuth;
                                    if (lOriginalAzimuth > 300.0)
                                        lNewAzimuth = lOriginalAzimuth - DOME_SYNC_OFFSET;
                                    else
                                        lNewAzimuth = lOriginalAzimuth + DOME_SYNC_OFFSET;

                                    // Sync to new azimuth
                                    TimeMethod(pName, () => domeDevice.SyncToAzimuth(lNewAzimuth), TargetTime.Standard);

                                    // OK Dome hasn't moved but should now show azimuth as a new value
                                    switch (Math.Abs(lNewAzimuth - domeDevice.Azimuth))
                                    {
                                        case object _ when Math.Abs(lNewAzimuth - domeDevice.Azimuth) < 1.0: // very close so give it an OK
                                            LogOk(pName, "Dome synced OK to within +- 1 degree");
                                            break;

                                        case object _ when Math.Abs(lNewAzimuth - domeDevice.Azimuth) < 2.0: // close so give it an INFO
                                            LogInfo(pName, "Dome synced to within +- 2 degrees");
                                            break;

                                        case object _ when Math.Abs(lNewAzimuth - domeDevice.Azimuth) < 5.0: // Closish so give an issue
                                            LogIssue(pName, "Dome only synced to within +- 5 degrees");
                                            break;

                                        case object _ when (DOME_SYNC_OFFSET - 2.0) <= Math.Abs(lNewAzimuth - domeDevice.Azimuth) && Math.Abs(lNewAzimuth - domeDevice.Azimuth) <= (DOME_SYNC_OFFSET + 2): // Hasn't really moved
                                            LogIssue(pName, "Dome did not sync, Azimuth didn't change value after sync command");
                                            break;

                                        default:
                                            LogIssue(pName, $"Dome azimuth was {Math.Abs(lNewAzimuth - domeDevice.Azimuth)} degrees away from expected value");
                                            break;
                                    }

                                    // Now try and restore original value
                                    LogCallToDriver(pName, "About to call SyncToAzimuth method");
                                    domeDevice.SyncToAzimuth(lOriginalAzimuth);
                                }
                                else
                                {
                                    LogCallToDriver(pName, "About to call SyncToAzimuth method");
                                    TimeMethod(pName, () => domeDevice.SyncToAzimuth(45.0), TargetTime.Standard); // Sync to an arbitrary direction
                                    LogOk(pName, "Dome successfully synced to 45 degrees but unable to read azimuth to confirm this");
                                }

                                // Now test sync to illegal values
                                try
                                {
                                    LogCallToDriver(pName, "About to call SyncToAzimuth method");
                                    domeDevice.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_LOW);
                                    LogIssue(pName, $"No exception generated when syncing to illegal azimuth {DOME_ILLEGAL_AZIMUTH_LOW} degrees");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.MustBeImplemented, ex, $"sync to {DOME_ILLEGAL_AZIMUTH_LOW} degrees",
                                        $"Invalid value exception correctly raised for sync to {DOME_ILLEGAL_AZIMUTH_LOW} degrees");
                                }
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                                try
                                {
                                    LogCallToDriver(pName, "About to call SyncToAzimuth method");
                                    domeDevice.SyncToAzimuth(DOME_ILLEGAL_AZIMUTH_HIGH);
                                    LogIssue(pName, $"No exception generated when syncing to illegal azimuth {DOME_ILLEGAL_AZIMUTH_HIGH} degrees");
                                }
                                catch (Exception ex)
                                {
                                    HandleInvalidValueExceptionAsOk(pName, MemberType.Method, Required.MustBeImplemented, ex, $"sync to {DOME_ILLEGAL_AZIMUTH_HIGH} degrees",
                                        $"Invalid value exception correctly raised for sync to {DOME_ILLEGAL_AZIMUTH_HIGH} degrees");
                                }
                                if (cancellationToken.IsCancellationRequested)
                                    return;
                            }
                            else
                                LogInfo(pName, "SyncToAzimuth test skipped since SlewToAzimuth throws an exception");
                        }
                        else
                        {
                            LogCallToDriver(pName, "About to call SyncToAzimuth method");
                            domeDevice.SyncToAzimuth(45.0);
                            LogIssue(pName, "CanSyncAzimuth is false but SyncToAzimuth did not raise an exception");
                        }
                        break;

                    default:
                        LogIssue(pName, $"DomeOptionalTest: Unknown test type {pType}");
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleException(pName, pMemberType, Required.Optional, ex, "");
            }
            ClearStatus();
        }

        private void DomeShutterTest(ShutterState pRequiredFinalShutterState, string pName)
        {
            ShutterState lShutterState;

            if (settings.DomeOpenShutter)
            {
                SetTest(pName);
                if (mCanReadShutterStatus)
                {
                    LogCallToDriver(pName, "About to get ShutterStatus property");
                    lShutterState = (ShutterState)domeDevice.ShutterStatus;

                    // Make sure we are in the required Open state to start the close test
                    switch (lShutterState)
                    {
                        case ShutterState.Closed:
                            if (pRequiredFinalShutterState == ShutterState.Closed)
                            {
                                // Wrong state, get to the required state
                                SetAction("Opening shutter ready for close test");
                                LogDebug(pName, "Opening shutter ready for close test");
                                LogCallToDriver(pName, "About to call OpenShutter method");
                                domeDevice.OpenShutter();

                                // Wait for shutter to open
                                if (!DomeShutterWait(ShutterState.Open))
                                    return;

                                DomeStabliisationWait();
                            }
                            else 
                            {
                                // Already in Open state, no action required
                            }

                            break;

                        case ShutterState.Closing:
                            if (pRequiredFinalShutterState == ShutterState.Closed)
                            {
                                SetAction("Waiting for shutter to close before opening ready for close test");
                                LogDebug(pName, "Waiting for shutter to close before opening ready for close test");

                                // Wait for shutter to close
                                if (!DomeShutterWait(ShutterState.Closed))
                                    return;

                                LogDebug(pName, "Opening shutter ready for close test");
                                SetAction("Opening shutter ready for close test");
                                LogCallToDriver(pName, "About to call OpenShutter method");

                                // Then open it
                                domeDevice.OpenShutter();

                                if (!DomeShutterWait(ShutterState.Open))
                                    return;

                                DomeStabliisationWait();
                            }
                            else
                            {
                                SetAction("Waiting for shutter to close ready for open test");
                                LogDebug(pName, "Waiting for shutter to close ready for open test");

                                // Wait for shutter to close
                                if (!DomeShutterWait(ShutterState.Closed))
                                    return; 

                                DomeStabliisationWait();
                            }
                            break;

                        case ShutterState.Opening:
                            if (pRequiredFinalShutterState == ShutterState.Closed)
                            {
                                SetAction("Waiting for shutter to open ready for close test");
                                LogDebug(pName, "Waiting for shutter to open ready for close test");

                                // Wait for shutter to open
                                if (!DomeShutterWait(ShutterState.Open))
                                    return;

                                DomeStabliisationWait();
                            }
                            else
                            {
                                SetAction("Waiting for shutter to open before closing ready for open test");
                                LogDebug(pName, "Waiting for shutter to open before closing ready for open test");

                                // Wait for shutter to open
                                if (!DomeShutterWait(ShutterState.Open))
                                    return;

                                LogDebug(pName, "Closing shutter ready for open test");
                                SetAction("Closing shutter ready for open test");
                                LogCallToDriver(pName, "About to call CloseShutter method");

                                // Then close it
                                domeDevice.CloseShutter();
                                if (!DomeShutterWait(ShutterState.Closed))
                                    return;

                                DomeStabliisationWait();
                            }
                            break;

                        case ShutterState.Error:
                            LogIssue("DomeShutterTest", $"Shutter state is Error: {lShutterState}");
                            break;

                        case ShutterState.Open:
                            if (pRequiredFinalShutterState == ShutterState.Closed)
                            {
                                // Already in Closed state, no action required
                            }
                            else
                            {
                                // Wrong state, get to the required state
                                SetAction("Closing shutter ready for open  test");
                                LogDebug(pName, "Closing shutter ready for open test");
                                LogCallToDriver(pName, "About to call CloseShutter method");
                                domeDevice.CloseShutter();

                                // Wait for shutter to open
                                if (!DomeShutterWait(ShutterState.Closed))
                                    return;
                                DomeStabliisationWait();
                            }
                            break;

                        default:
                            LogIssue("DomeShutterTest", $"Unexpected shutter status: {lShutterState}");
                            break;
                    }

                    // Now test that we can get to the required state
                    if (pRequiredFinalShutterState == ShutterState.Closed) // Test that we can close the shutter
                    {
                        // Shutter is now open so close it
                        SetAction("Closing shutter");
                        LogCallToDriver(pName, "About to call CloseShutter method");
                        TimeMethod(pName, () => domeDevice.CloseShutter(), TargetTime.Standard);

                        SetAction("Waiting for shutter to close");
                        LogDebug(pName, "Waiting for shutter to close");
                        if (!DomeShutterWait(ShutterState.Closed))
                        {
                            LogCallToDriver(pName, "About to get ShutterStatus property");
                            lShutterState = domeDevice.ShutterStatus;
                            LogIssue(pName, $"Unable to close shutter - ShutterStatus: {lShutterState}");
                            return;
                        }
                        else
                            LogOk(pName, "Shutter closed successfully");
                        DomeStabliisationWait();
                    }
                    else // Test that we can open the shutter
                    {
                        // Shutter is now closed so open it
                        SetAction("Opening shutter");
                        LogCallToDriver(pName, "About to call OpenShutter method");
                        TimeMethod(pName, () => domeDevice.OpenShutter(), TargetTime.Standard);

                        SetAction("Waiting for shutter to open");
                        LogDebug(pName, "Waiting for shutter to open");
                        if (!DomeShutterWait(ShutterState.Open))
                        {
                            LogCallToDriver(pName, "About to get ShutterStatus property");
                            lShutterState = domeDevice.ShutterStatus;
                            LogIssue(pName, $"Unable to open shutter - ShutterStatus: {lShutterState}");
                            return;
                        }
                        else
                            LogOk(pName, "Shutter opened successfully");
                        DomeStabliisationWait();
                    }
                }
                else
                {
                    LogDebug(pName, "Can't read shutter status!");
                    if (pRequiredFinalShutterState == ShutterState.Closed)
                    {
                        // Just issue command to see if it doesn't generate an error
                        LogCallToDriver(pName, "About to call CloseShutter method");
                        domeDevice.CloseShutter();
                        DomeStabliisationWait();
                    }
                    else
                    {
                        // Just issue command to see if it doesn't generate an error
                        LogCallToDriver(pName, "About to call OpenShutter method");
                        domeDevice.OpenShutter();
                        DomeStabliisationWait();
                    }
                    LogOk(pName, "Command issued successfully but can't read ShutterStatus to confirm shutter is closed");
                }

                ClearStatus();
            }
            else
                LogTestAndMessage("DomeSafety", "Open shutter check box is unchecked so shutter test bypassed");
        }

        private bool DomeShutterWait(ShutterState pRequiredStatus)
        {
            DateTime lStartTime;
            // Wait for shutter to reach required stats or user presses stop or timeout occurs
            // Returns true if required state is reached
            bool returnValue = false;
            lStartTime = DateTime.Now;
            try
            {
                LogCallToDriver("DomeShutterWait", "About to get ShutterStatus property multiple times");
                WaitWhile($"Waiting for shutter state {pRequiredStatus}", () => (domeDevice.ShutterStatus != pRequiredStatus), 500, settings.DomeShutterMovementTimeout);

                if ((domeDevice.ShutterStatus == pRequiredStatus)) returnValue = true; // All worked so return True

                if ((DateTime.Now.Subtract(lStartTime).TotalSeconds > settings.DomeShutterMovementTimeout))
                    LogIssue("DomeShutterWait",
                        $"Timed out waiting for shutter to reach state: {pRequiredStatus}, consider increasing the timeout setting in Options / Conformance Options");
            }
            catch (Exception ex)
            {
                LogIssue("DomeShutterWait", $"Unexpected exception: {ex}");
            }

            return returnValue;
        }

        private void DomePerformanceTest(DomePropertyMethod pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime;
            double lRate;
            bool lBoolean;
            double lDouble;
            ShutterState lShutterState;
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
                        case DomePropertyMethod.Altitude:
                            {
                                lDouble = domeDevice.Altitude;
                                break;
                            }

                        case DomePropertyMethod.Azimuth:
                            {
                                lDouble = domeDevice.Azimuth;
                                break;
                            }

                        case DomePropertyMethod.ShutterStatus:
                            {
                                lShutterState = domeDevice.ShutterStatus;
                                break;
                            }

                        case DomePropertyMethod.SlavedRead:
                            {
                                lBoolean = domeDevice.Slaved;
                                break;
                            }

                        case DomePropertyMethod.Slewing:
                            {
                                lBoolean = domeDevice.Slewing;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"DomePerformanceTest: Unknown test type {pType}");
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

        public void DomeStabliisationWait()
        {
            // Only wait if a non-zero wait time has been configured
            if (settings.DomeStabilisationWaitTime > 0)
            {
                Stopwatch sw = Stopwatch.StartNew();
                WaitWhile("Waiting for dome to stabilise", () => sw.Elapsed.TotalSeconds < settings.DomeStabilisationWaitTime, 500, settings.DomeStabilisationWaitTime);
            }
        }

        #endregion

    }
}

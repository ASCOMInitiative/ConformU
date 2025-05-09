﻿// Base class from which particular device testers are derived
// Put all common elements in here
using ASCOM;
using ASCOM.Com;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.Tools;
using Microsoft.CSharp.RuntimeBinder;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.Linq;
using static ConformU.Globals;
using InvalidOperationException = System.InvalidOperationException;
using NotImplementedException = ASCOM.NotImplementedException;

namespace ConformU
{

    /// <summary>
    /// Base class for device tester classes. Contains common code and placeholders for the 
    /// methods that must be implemented in the device tester class
    /// </summary>
    internal class DeviceTesterBaseClass : IDisposable
    {
        // Implements IDisposable

        #region Constants

        internal const double PERF_LOOP_TIME = 5.0; // Performance loop run time in seconds
        internal const int SLEEP_TIME = 500; // Loop time for testing whether slewing has completed
        internal const int WAITFOR_UPDATE_INTERVAL = 500; // Time in milliseconds between updates in the WaitFor method

        // Class not registered COM exception error number
        internal const int REGDB_E_CLASSNOTREG = unchecked((int)0x80040154);

        // Member not found COM exceptions error number
        internal const int DISP_E_UNKNOWNNAME = unchecked((int)0x80020006);

        #endregion

        #region Variables

        private int? baseInterfaceVersion; // Variable to held interface version of the current device

        private bool hasProperties, hasCanProperties, hasMethods, hasPreRunCheck, hasPostRunCheck, hasPerformanceCheck;
        private bool hasPreConnectCheck;

        private IAscomDeviceV2 baseClassDevice;
        private DeviceTypes baseClassDeviceType;

        private string testName, testAction, testStatus;

        private readonly ConformLogger tl;
        internal readonly CancellationToken ApplicationCancellationToken;

        private readonly Settings settings;

        internal int ExNotImplemented, ExNotSet1, ExNotSet2;
        internal int ExInvalidValue1, ExInvalidValue2, ExInvalidValue3, ExInvalidValue4, ExInvalidValue5, ExInvalidValue6;

        #endregion

        #region Enums
        private enum CommandType
        {
            TstCommandString = 1,
            TstCommandBool = 2,
            TstCommandBlind = 3,
            TstCommandStringRaw = 4,
            TstCommandBoolRaw = 5,
            TstCommandBlindRaw = 6
        }

        protected enum MemberType
        {
            Property,
            Method
        }

        protected enum Required
        {
            Optional,
            Mandatory,
            MustBeImplemented,
            MustNotBeImplemented
        }

        internal enum TargetTime
        {
            Fast,
            Standard,
            Extended
        }

        #region Enums

        internal enum StatusType
        {
            StaTest = 1,
            StaAction = 2,
            StaStatus = 3
        }

        public enum SpecialTest
        {
            TelescopeSideOfPier,
            TelescopeDestinationSideOfPier,
            TelescopeSideOfPierAnalysis,
            TelescopeCommands
        }

        internal enum MandatoryMethod
        {
            Connected = 0,
            Description = 1,
            DriverInfo = 2,
            DriverVersion = 3,
            Name = 4,
            CommandXxx = 5
        }
        #endregion



        #endregion

        #region New and Dispose
        public DeviceTesterBaseClass() : base()
        {
            hasPreConnectCheck = false;
            hasPreRunCheck = false;
            hasCanProperties = false;
            hasProperties = true;
            hasMethods = true;
            hasPostRunCheck = false;
            hasPerformanceCheck = true;

            ClearStatus();
        }

        /// <summary>
        /// Initialise device tester base class
        /// </summary>
        /// <param name="hasCanProperties">Device has Can properties</param>
        /// <param name="hasProperties">Device has properties</param>
        /// <param name="hasMethods">Device has methods</param>
        /// <param name="hasPreRunCheck">Device requires a pre run safety check </param>
        /// <param name="hasPreConnectCheck">Device requires a pre connection check</param>
        /// <param name="hasPerformanceCheck">Device has a performance test</param>
        /// <param name="hasPostRunCheck">Device requires a post run safety check</param>
        /// <remarks></remarks>
        public DeviceTesterBaseClass(bool hasCanProperties, bool hasProperties, bool hasMethods, bool hasPreRunCheck, bool hasPreConnectCheck, bool hasPerformanceCheck, bool hasPostRunCheck, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken cancellationToken) : this()
        {
            this.hasPreConnectCheck = hasPreConnectCheck;
            this.hasPreRunCheck = hasPreRunCheck;
            this.hasCanProperties = hasCanProperties;
            this.hasProperties = hasProperties;
            this.hasMethods = hasMethods;
            this.hasPostRunCheck = hasPostRunCheck;
            this.hasPerformanceCheck = hasPerformanceCheck;
            tl = logger;
            this.ApplicationCancellationToken = cancellationToken;
            settings = conformConfiguration.Settings;
        }
        private bool disposedValue = false;        // To detect redundant calls
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    tl.LogMessage("DeviceTesterbaseClass", MessageLevel.Debug, "Dispose called");
                    try
                    {
                        tl.LogMessage("DeviceTesterbaseClass", MessageLevel.Debug, "Calling base class Dispose method...");
                        baseClassDevice.Dispose();
                        tl.LogMessage("DeviceTesterbaseClass", MessageLevel.Debug, "returned from base class Dispose method");
                    }
                    catch (Exception ex)
                    {
                        tl.LogMessage("DeviceTesterbaseClass", MessageLevel.Debug, $"Exception\r\n{ex}");
                    }

                    baseClassDevice = null;
                    tl.LogMessage("DeviceTesterbaseClass", MessageLevel.Debug, "Dispose finished");
                }
            }

            disposedValue = true;
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put clean-up code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        #endregion

        #region Code

        public void SetupDialog()
        {
            if (settings.DisplayMethodCalls) LogTestAndMessage("SetupDialog", "About to call SetupDialog");
            ((dynamic)baseClassDevice).SetupDialog();
        }

        public virtual void CheckCommonMethods()
        {
            LogTestOnly("Common Driver Methods");

            // InterfaceVersion - Required
            try
            {
                switch (TimeFunc("InterfaceVersion", () => GetInterfaceVersion(), TargetTime.Fast))
                {
                    case < 1:
                        LogIssue("InterfaceVersion",
                            $"InterfaceVersion must be 1 or greater but driver returned: {GetInterfaceVersion()}");
                        break;

                    default:
                        LogOk("InterfaceVersion", GetInterfaceVersion().ToString());
                        break;
                }
            }
            catch (Exception ex)
            {
                HandleException("InterfaceVersion", MemberType.Property, Required.Mandatory, ex, "");
            }
            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            // Connected - Required
            if (IncludeMethod(MandatoryMethod.Connected, baseClassDeviceType, GetInterfaceVersion()))
            {
                try
                {
                    LogCallToDriver("Connected", "About to get property Connected");
                    bool connected = TimeFunc("Connected", () => baseClassDevice.Connected, TargetTime.Fast);
                    LogOk("Connected", connected.ToString());
                }
                catch (Exception ex)
                {
                    LogIssue("Connected", ex.Message);
                }

                if (ApplicationCancellationToken.IsCancellationRequested)
                    return;
            }

            // Description - Required
            try
            {
                if (DeviceCapabilities.InterfaceHasMember(MemberNames.Description, baseClassDeviceType, GetInterfaceVersion()))
                {
                    LogCallToDriver("Description", "About to get property Description");
                    string description = TimeFunc("Description", () => baseClassDevice.Description, TargetTime.Fast);
                    switch (description ?? "")
                    {
                        case "":
                            LogInfo("Description", "No description string");
                            break;

                        default:
                            if (description.Length > 68 & baseClassDeviceType == DeviceTypes.Camera)
                            {
                                LogIssue("Description",
                                    $"Maximum number of characters is 68 for compatibility with FITS headers, found: {description.Length} characters: {description}");
                            }
                            else
                            {
                                LogOk("Description", description.ToString());
                            }
                            break;
                    }
                }
                else
                {
                    LogInfo("Description", "Description test omitted - Description is not a member of this interface version.");
                }
            }
            catch (Exception ex)
            {
                HandleException("Description", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            // DriverInfo - Required
            try
            {
                if (DeviceCapabilities.InterfaceHasMember(MemberNames.DriverInfo, baseClassDeviceType, GetInterfaceVersion()))
                {
                    LogCallToDriver("DriverInfo", "About to get property DriverInfo");
                    string driverInfo = TimeFunc("DriverInfo", () => baseClassDevice.DriverInfo, TargetTime.Fast);
                    switch (driverInfo ?? "")
                    {
                        case "":
                            LogInfo("DriverInfo", "No DriverInfo string");
                            break;

                        default:
                            LogOk("DriverInfo", driverInfo.ToString());
                            break;
                    }
                }
                else
                {
                    LogInfo("DriverInfo", "DriverInfo test omitted - DriverInfo is not a member of this interface version.");
                }
            }
            catch (Exception ex)
            {
                HandleException("DriverInfo", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            // DriverVersion - Required
            try
            {
                if (DeviceCapabilities.InterfaceHasMember(MemberNames.DriverVersion, baseClassDeviceType, GetInterfaceVersion()))
                {
                    LogCallToDriver("DriverVersion", "About to get property DriverVersion");
                    string driverVersion = TimeFunc("DriverVersion", () => baseClassDevice.DriverVersion, TargetTime.Fast);
                    switch (driverVersion ?? "")
                    {
                        case "":
                            LogInfo("DriverVersion", "No DriverVersion string");
                            break;

                        default:
                            LogOk("DriverVersion", driverVersion.ToString());
                            break;
                    }
                }
                else
                {
                    LogInfo("DriverVersion", "DriverVersion test omitted - DriverVersion is not a member of this interface version.");
                }
            }
            catch (Exception ex)
            {
                LogIssue("DriverVersion", ex.Message);
            }

            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            // Name - Required
            try
            {
                if (DeviceCapabilities.InterfaceHasMember(MemberNames.Name, baseClassDeviceType, GetInterfaceVersion()))
                {
                    LogCallToDriver("Name", "About to get property Name");
                    string name = TimeFunc("Name", () => baseClassDevice.Name, TargetTime.Fast);
                    switch (name ?? "")
                    {
                        case "":
                            LogInfo("Name", "Name is empty");
                            break;

                        default:
                            LogOk("Name", name);
                            break;
                    }
                }
                else
                {
                    LogInfo("Name", "Skipping test because Name is not a member of this interface version.");
                }
            }
            catch (Exception ex)
            {
                HandleException("Name", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (ApplicationCancellationToken.IsCancellationRequested)
                return;
            LogNewLine();

            // Action - optional but cannot be tested
            LogInfo("Action", "Conform cannot test the Action method");

            // Supported actions - Optional but Required through DriverAccess
            try
            {
                if (DeviceCapabilities.InterfaceHasMember(MemberNames.SupportedActions, baseClassDeviceType, GetInterfaceVersion()))
                {
                    LogCallToDriver("SupportedActions", "About to call method SupportedActions");
                    IList supportedActions = TimeFunc("SupportedActions", () => (IList)baseClassDevice.SupportedActions, TargetTime.Fast);
                    if (supportedActions.Count == 0)
                    {
                        LogOk("SupportedActions", "Driver returned an empty action list");
                    }
                    else
                    {
                        int i = 0;
                        foreach (object action in supportedActions)
                        {
                            i += 1;
                            if (action.GetType().Name == "String")
                            {
                                string actionString = action.ToString();
                                const string testParameters = "Conform test parameters";
                                switch (actionString ?? "")
                                {
                                    case "":
                                        LogIssue("SupportedActions", $"Supported action {i} Is an empty string"); // List the action that was found
                                        break;

                                    default:
                                        {
                                            LogOk("SupportedActions", $"Found action: {actionString}");

                                            // Carry out the following Action tests only when we are testing the Observing Conditions Hub and it is configured to use the Switch and OC simulators
                                            if (baseClassDeviceType == DeviceTypes.ObservingConditions & settings.DeviceTechnology == DeviceTechnology.COM & settings.ComDevice.ProgId.Equals("ASCOM.OCH.OBSERVINGCONDITIONS", StringComparison.OrdinalIgnoreCase))
                                            {
                                                string result;
                                                if (actionString.StartsWith("//OCSIMULATOR:", StringComparison.InvariantCultureIgnoreCase))
                                                {
                                                    try
                                                    {
                                                        LogCallToDriver("SupportedActions", "About to call method Action");
                                                        result = baseClassDevice.Action(actionString, testParameters);
                                                        LogOk("SupportedActions",
                                                            $"OC simulator action {actionString} gave result: {result}");
                                                    }
                                                    catch (Exception ex1)
                                                    {
                                                        LogIssue("SupportedActions",
                                                            $"Exception calling OCH simulator action {actionString}: {ex1.Message}");
                                                    }
                                                }
                                                else if (actionString.StartsWith("//ASCOM.SIMULATOR.OBSERVINGCONDITIONS:", StringComparison.InvariantCultureIgnoreCase))
                                                {
                                                    try
                                                    {
                                                        LogCallToDriver("SupportedActions", "About to call method Action");
                                                        result = baseClassDevice.Action(actionString, testParameters);
                                                        LogOk("SupportedActions",
                                                            $"OC simulator action {actionString} gave result: {result}");
                                                    }
                                                    catch (Exception ex1)
                                                    {
                                                        LogIssue("SupportedActions",
                                                            $"Exception calling OCH simulator action {actionString}: {ex1.Message}");
                                                    }
                                                }
                                                else if (actionString.StartsWith("//SWITCHSIMULATOR:", StringComparison.InvariantCultureIgnoreCase))
                                                {
                                                    try
                                                    {
                                                        LogCallToDriver("SupportedActions", "About to call method Action");
                                                        result = baseClassDevice.Action(actionString, testParameters);
                                                        LogOk("SupportedActions",
                                                            $"Switch simulator action {actionString} gave result: {result}");
                                                    }
                                                    catch (Exception ex1)
                                                    {
                                                        LogIssue("SupportedActions",
                                                            $"Exception calling switch simulator action {actionString}: {ex1.Message}");
                                                    }
                                                }
                                                else if (actionString.StartsWith("//ASCOM.SIMULATOR.SWITCH:", StringComparison.InvariantCultureIgnoreCase))
                                                {
                                                    try
                                                    {
                                                        LogCallToDriver("SupportedActions", "About to call method Action");
                                                        result = baseClassDevice.Action(actionString, testParameters);
                                                        LogOk("SupportedActions",
                                                            $"Switch simulator action {actionString} gave result: {result}");
                                                    }
                                                    catch (Exception ex1)
                                                    {
                                                        LogIssue("SupportedActions",
                                                            $"Exception calling switch simulator action {actionString}: {ex1.Message}");
                                                    }
                                                }
                                            }

                                            break;
                                        }
                                }
                            }
                            else
                            {
                                LogIssue("SupportedActions",
                                    $"Actions must be strings. The type of action {i} {action} is: {action.GetType().Name}");
                            }
                        }
                    }
                }
                else
                {
                    LogInfo("SupportedActions", "Skipping test because SupportedActions is not a member of this interface version.");
                }
            }
            catch (Exception ex)
            {
                if (baseClassDeviceType == DeviceTypes.Switch & ReferenceEquals(ex.GetType(), typeof(MissingMemberException)))
                {
                    LogOk("SupportedActions", "Switch V1 Driver does not have SupportedActions");
                }
                else
                {
                    HandleException("SupportedActions", MemberType.Property, Required.Optional, ex, "");
                    LogIssue("SupportedActions", ex.Message);
                }
            }
            LogNewLine();

            // DeviceState - Mandatory for Platform 7 and above, otherwise not present
            if (DeviceCapabilities.HasConnectAndDeviceState(baseClassDeviceType, GetInterfaceVersion()))
            {
                try
                {
                    LogCallToDriver("DeviceState", "About to get property DeviceState");
                    List<StateValue> deviceState = TimeFunc("DeviceState", () => baseClassDevice.DeviceState, TargetTime.Fast);

                    int numberOfItems = 0;
                    foreach (var item in deviceState)
                    {
                        numberOfItems++;
                    }
                    LogOk("DeviceState", $"Received {numberOfItems} operational state properties.");

                    // Create a case sensitive dictionary to hold property names
                    Dictionary<string, bool> casedPropertyNames = [];

                    // Define expected operational properties for each device type
                    switch (settings.DeviceType)
                    {
                        case DeviceTypes.Camera:
                            casedPropertyNames.Add(nameof(ICameraV4.CameraState), false);
                            casedPropertyNames.Add(nameof(ICameraV4.CCDTemperature), false);
                            casedPropertyNames.Add(nameof(ICameraV4.CoolerPower), false);
                            casedPropertyNames.Add(nameof(ICameraV4.HeatSinkTemperature), false);
                            casedPropertyNames.Add(nameof(ICameraV4.ImageReady), false);
                            casedPropertyNames.Add(nameof(ICameraV4.IsPulseGuiding), false);
                            casedPropertyNames.Add(nameof(ICameraV4.PercentCompleted), false);
                            break;

                        case DeviceTypes.CoverCalibrator:
                            casedPropertyNames.Add(nameof(ICoverCalibratorV2.Brightness), false);
                            casedPropertyNames.Add(nameof(ICoverCalibratorV2.CalibratorState), false);
                            casedPropertyNames.Add(nameof(ICoverCalibratorV2.CoverState), false);
                            casedPropertyNames.Add(nameof(ICoverCalibratorV2.CalibratorChanging), false);
                            casedPropertyNames.Add(nameof(ICoverCalibratorV2.CoverMoving), false);
                            break;

                        case DeviceTypes.Dome:
                            casedPropertyNames.Add(nameof(IDomeV3.Altitude), false);
                            casedPropertyNames.Add(nameof(IDomeV3.AtHome), false);
                            casedPropertyNames.Add(nameof(IDomeV3.AtPark), false);
                            casedPropertyNames.Add(nameof(IDomeV3.Azimuth), false);
                            casedPropertyNames.Add(nameof(IDomeV3.ShutterStatus), false);
                            casedPropertyNames.Add(nameof(IDomeV3.Slewing), false);
                            break;

                        case DeviceTypes.FilterWheel:
                            casedPropertyNames.Add(nameof(IFilterWheelV3.Position), false);
                            break;

                        case DeviceTypes.Focuser:
                            casedPropertyNames.Add(nameof(IFocuserV4.IsMoving), false);
                            casedPropertyNames.Add(nameof(IFocuserV4.Position), false);
                            casedPropertyNames.Add(nameof(IFocuserV4.Temperature), false);
                            break;

                        case DeviceTypes.ObservingConditions:
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.CloudCover), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.DewPoint), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.Humidity), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.Pressure), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.RainRate), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.SkyBrightness), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.SkyQuality), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.SkyTemperature), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.StarFWHM), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.Temperature), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.WindDirection), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.WindGust), false);
                            casedPropertyNames.Add(nameof(IObservingConditionsV2.WindSpeed), false);
                            break;

                        case DeviceTypes.Rotator:
                            casedPropertyNames.Add(nameof(IRotatorV4.IsMoving), false);
                            casedPropertyNames.Add(nameof(IRotatorV4.MechanicalPosition), false);
                            casedPropertyNames.Add(nameof(IRotatorV4.Position), false);
                            break;

                        case DeviceTypes.SafetyMonitor:
                            casedPropertyNames.Add(nameof(ISafetyMonitorV3.IsSafe), false);
                            break;

                        case DeviceTypes.Switch:

                            // Try to get the MaxSwitch property
                            short maxSwitch = 0;
                            LogCallToDriver("DeviceState", "About to get MaxSwitch property");
                            try { maxSwitch = (baseClassDevice as ISwitchV3).MaxSwitch; }
                            catch (Exception ex)
                            {
                                LogIssue("DeviceState", $"MaxSwitch exception: {ex}");
                            }
                            LogInfo("DeviceState", $"MaxSwitch: {maxSwitch}");
                            for (int i = 0; i < maxSwitch; i++)
                            {
                                casedPropertyNames.Add($"GetSwitch{i}", false);
                                casedPropertyNames.Add($"GetSwitchValue{i}", false);
                                casedPropertyNames.Add($"StateChangeComplete{i}", false);
                            }
                            break;

                        case DeviceTypes.Telescope:
                            casedPropertyNames.Add(nameof(ITelescopeV4.Altitude), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.AtHome), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.AtPark), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.Azimuth), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.Declination), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.IsPulseGuiding), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.RightAscension), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.SideOfPier), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.SiderealTime), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.Slewing), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.Tracking), false);
                            casedPropertyNames.Add(nameof(ITelescopeV4.UTCDate), false);
                            break;

                        case DeviceTypes.Video:
                            casedPropertyNames.Add(nameof(IVideoV2.CameraState), false);
                            break;

                        default:
                            LogError("DeviceState", $"Unknown device type: {settings.DeviceType}.");
                            break;
                    }

                    // Create a case insensitive dictionary containing property names to detect mis-casing issues
                    Dictionary<string, bool> caseInsensitivePropertyNames = new(casedPropertyNames, StringComparer.OrdinalIgnoreCase);

                    foreach (IStateValue property in deviceState)
                    {
                        // Test whether the property name is one that we are expecting
                        if ((casedPropertyNames.ContainsKey(property.Name)) | ((property.Name == "TimeStamp"))) // Property name is expected
                        {
                            casedPropertyNames[property.Name] = true;
                            LogOk("DeviceState", $"  {property.Name} = {property.Value}");
                        }
                        else // An unexpected name was found
                        {
                            // Test whether this has the correct spelling but incorrect casing
                            if (caseInsensitivePropertyNames.ContainsKey(property.Name)) // Property name was found
                            {
                                string key = casedPropertyNames.Where(pair => pair.Key.Equals(property.Name, StringComparison.InvariantCultureIgnoreCase)).Select(pair => pair.Key).FirstOrDefault();
                                LogIssue("DeviceState", $"The {key} property name is mis-cased: {property.Name}. The correct casing is: {key}");
                            }
                            else // An unexpected name was found
                            {
                                LogIssue("DeviceState", $"A non-operational property was included in the DeviceState response: {property.Name} = {property.Value}");
                            }
                        }
                    }

                    // Test whether all expected values were returned
                    List<KeyValuePair<string, bool>> missingProperties = casedPropertyNames.Where(x => x.Value == false).ToList();

                    // Test whether any properties were not found
                    if (missingProperties.Count != 0) // One or more properties were missing
                        foreach (KeyValuePair<string, bool> item in missingProperties)
                        {
                            LogInfo("DeviceState", $"Operational property {item.Key} was not included in the DeviceState response.");
                        }
                    else // All properties were supplied
                    {
                        LogOk("DeviceState", $"Found all expected operational properties");
                    }

                }
                catch (Exception ex)
                {
                    HandleException("DeviceState", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
            {
                LogInfo("DeviceState", "DeviceState tests omitted - DeviceState is not a member of this interface version."); //"Skipping test because Description is not a member of this interface version.");
            }

            if (ApplicationCancellationToken.IsCancellationRequested)
                return;

            LogNewLine();
        }

        public int GetInterfaceVersion()
        {
            try
            {
                if (baseInterfaceVersion.HasValue)
                    return baseInterfaceVersion.Value;

                LogCallToDriver("InterfaceVersion", "About to get property InterfaceVersion");
                baseInterfaceVersion = baseClassDevice.InterfaceVersion;
                LogDebug("GetInterfaceVersion", $"Device interface version: {baseInterfaceVersion}");

                return baseInterfaceVersion.Value;
            }
            catch (Exception ex) // Failed to get the interface version
            {
                LogDebug("GetInterfaceVersion", $"Exception getting interface version: {ex.Message}\r\n{ex}");

                // Check whether InterfaceVersion is implemented in the version 1 interface
                if (!DeviceCapabilities.InterfaceHasMember(MemberNames.InterfaceVersion, baseClassDeviceType, 1)) // Interface version is NOT implemented in the version 1 interface
                {
                    // Assume that this is a version 1 interface to enable further tests to proceed
                    baseInterfaceVersion = 1;
                    LogInfo("InterfaceVersion", $"#####");
                    LogInfo("InterfaceVersion", $"##### InterfaceVersion returned an error: {ex.Message}. Assuming interface version 1");
                    LogInfo("InterfaceVersion", $"#####");

                    // Return interface version 1
                    return baseInterfaceVersion.Value;
                }
                else // Report that an error occurred when trying to get the interface version
                {
                    throw;
                }
            }
        }

        public void Connect()
        {
            CheckDevice("Connect");

            SetTest("Connect");
            SetAction("Waiting for Connected to become 'true'");

            // Set an internal property indicating whether this is a Platform 7 or later device
            IsPlatform7OrLater = DeviceCapabilities.IsPlatform7OrLater(settings.DeviceType, GetInterfaceVersion());

            // Use Connect /Disconnect if present
            if (DeviceCapabilities.HasConnectAndDeviceState(baseClassDeviceType, GetInterfaceVersion()))
            {
                // First make sure that the Connecting property exists
                try
                {
                    bool connecting = baseClassDevice.Connecting;
                    // If we get there the property is implemented and can continue with connection tests
                }
                catch (Exception ex)
                {
                    // Check whether we got a missing member exception
                    if (ex is RuntimeBinderException | ((ex is COMException) & (ex.HResult == DISP_E_UNKNOWNNAME))) // This is a missing member exception (different exceptions for the native and driver access components)
                        throw new MissingMemberException("The Connecting property is not present in the device interface.");
                    else // Some other error occurred so just throw it to the test manager
                        throw;
                }

                LogCallToDriver("Connect", "About to get Connecting property");
                if (!baseClassDevice.Connecting) // No connection / disconnection is in progress
                {
                    // First make sure that Connecting is correctly implemented
                    SetAction("Waiting for Connected to become True");
                    LogCallToDriver("Connected", "About to set Connected property true");
                    baseClassDevice.Connected = true;

                    LogCallToDriver("Connect", "About to get Connected property");
                    if (baseClassDevice.Connected != true)
                    {
                        throw new ASCOM.InvalidOperationException($"Set Connected True - The device connected without error but Connected Get returned False.");
                    }

                    LogOk("Connected", "Connected to device successfully using Connected = True");

                    //Wait for a short time
                    WaitFor(500, 100);

                    // Make sure that we can disconnect as well
                    SetAction("Waiting for Connected to become False");
                    LogCallToDriver("Connected", "About to set Connected property False");
                    baseClassDevice.Connected = false;
                    if (baseClassDevice.Connected != false)
                    {
                        throw new ASCOM.InvalidOperationException($"Set Connected False - The device disconnected without error but Connected Get returned True.");
                    }

                    LogOk("Connected", "Disconnected from device successfully using Connected = False");

                    // Call the Connect method and wait for the device to connect
                    SetAction("Waiting for the Connect method to complete");
                    LogCallToDriver("Connect", "About to call Connect() method");
                    TimeMethod("Connect", () => baseClassDevice.Connect(), TargetTime.Standard);
                    LogCallToDriver("Connect", "About to get Connecting property repeatedly");
                    WaitWhile("Connecting to device", () => baseClassDevice.Connecting, SLEEP_TIME, settings.ConnectDisconnectTimeout);
                }
                else // Connection already in progress so ignore this connect request
                {
                    LogInfo("Connect", "Ignoring this request because a Connect() or Disconnect() operation is already in progress.");
                }
            }
            else // Historic synchronous behaviour
            {
                // Handle connection to devices 
                if (DeviceCapabilities.HasConnected(baseClassDeviceType, GetInterfaceVersion())) // All devices and interfaces except IFocuserV1
                {
                    SetAction("Waiting for Connected to become True");
                    LogCallToDriver("Connected", "About to set Connected property");
                    baseClassDevice.Connected = true;
                    SetAction("Device has connected");
                }
                else // IFocuserV1, which we can't test because the Link method is not available in ASCOM Library methods.
                {
                    LogError("Connected", $"The focuser returned interface version: {GetInterfaceVersion()}, Conform Universal can only test focusers that implement IFocuserV2 or later.");
                    throw new InvalidOperationException($"Unable to test focuser IFocuserV{GetInterfaceVersion()} devices.");
                }
            }

            // Make sure that the value set is reflected in Connected GET
            LogCallToDriver("Connect", "About to get Connected property");

            if (baseClassDevice.Connected != true)
            {
                throw new ASCOM.InvalidOperationException($"The device connected without error but Connected Get returned False.");
            }

            if (DeviceCapabilities.HasConnectAndDeviceState(settings.DeviceType, GetInterfaceVersion()))
                tl.LogMessage("Connect", MessageLevel.OK, "Connected to device successfully using Connect()");
            else
                LogOk("Connected", "Connected to device successfully using Connected = True");

            ResetTestActionStatus();
            tl.LogMessage("", MessageLevel.TestOnly, "");
        }

        public void ValidateInterfaceVersion()
        {
            // Try to get the device's interface version
            int interfaceVersion = GetInterfaceVersion();

            // Validate the interface version
            if (DeviceCapabilities.IsValidAscomInterface(settings.DeviceType, interfaceVersion))
                LogOk("CreateDevice", $"Found a valid interface version: {interfaceVersion}");
            else
            {
                if (interfaceVersion < 1)
                    throw new InvalidValueException($"The returned interface version was: {interfaceVersion}, ASCOM interface versions cannot be zero or negative.");
                else
                    throw new InvalidValueException($"The returned interface version was: {interfaceVersion}, this is above the highest supported ASCOM interface version.");
            }
        }

        public void Disconnect()
        {
            CheckDevice("Disconnect");
            SetTest("Disconnect");

            // Try to get the device's interface version
            LogDebug("Disconnect", $"Interface version: {GetInterfaceVersion()}");

            // Use Connect /Disconnect if present
            if (DeviceCapabilities.HasConnectAndDeviceState(settings.DeviceType, GetInterfaceVersion()))
            {
                LogCallToDriver("Disconnect", "About to get Connecting property");
                if (!baseClassDevice.Connecting) // No connection / disconnection is in progress
                {
                    // Call the DIsconnect method and wait for the device to disconnect
                    SetAction("Waiting for the Disconnect method to complete");
                    LogCallToDriver("Disconnect", "About to call Disconnect() method");
                    TimeMethod("Disconnect", () => baseClassDevice.Disconnect(), TargetTime.Standard);
                    LogCallToDriver("Disconnect", "About to get Connecting property repeatedly");
                    WaitWhile("Disconnecting from device", () => baseClassDevice.Connecting, SLEEP_TIME, settings.ConnectDisconnectTimeout);
                }
                else // Connection already in progress so ignore this connect request
                {
                    LogInfo("Disconnect", "Ignoring this request because a Connect() or Disconnect() operation is already in progress.");
                }
            }
            else // Historic synchronous behaviour
            {
                SetAction("Waiting for Connected to become False");
                LogCallToDriver("Connected", "About to set Connected property");
                baseClassDevice.Connected = false;
            }

            // Make sure that the value set is reflected in Connected GET
            LogCallToDriver("Disconnect", "About to get Connected property");
            if (baseClassDevice.Connected != false)
            {
                throw new ASCOM.InvalidOperationException($"The device disconnected without error but Connected Get returned True.");
            }

            if (settings.DeviceType != null && DeviceCapabilities.HasConnectAndDeviceState(settings.DeviceType, (short)GetInterfaceVersion()))
                tl.LogMessage("Disconnect", MessageLevel.OK, "Disconnected from device successfully using Disconnect()");
            else
                tl.LogMessage("Connected", MessageLevel.OK, "False");

            ResetTestActionStatus();
            tl.LogMessage("", MessageLevel.TestOnly, "");
        }

        public bool Connected
        {
            get
            {
                LogCallToDriver("ConformanceCheck", "About to get Connected");
                return baseClassDevice.Connected;
            }
        }

        public virtual bool HasCanProperties
        {
            get
            {
                return hasCanProperties;
            }

            set
            {
                hasCanProperties = value;
            }
        }

        public virtual bool HasProperties
        {
            get
            {
                return hasProperties;
            }

            set
            {
                hasProperties = value;
            }
        }

        public virtual bool HasMethods
        {
            get
            {
                return hasMethods;
            }

            set
            {
                hasMethods = value;
            }
        }

        public virtual bool HasPreConnectCheck
        {
            get
            {
                return hasPreConnectCheck;
            }

            set
            {
                hasPreConnectCheck = value;
            }
        }

        public virtual bool HasPreRunCheck
        {
            get
            {
                return hasPreRunCheck;
            }

            set
            {
                hasPreRunCheck = value;
            }
        }

        public virtual bool HasPostRunCheck
        {
            get
            {
                return hasPostRunCheck;
            }

            set
            {
                hasPostRunCheck = value;
            }
        }

        public virtual bool HasPerformanceCheck
        {
            get
            {
                return hasPerformanceCheck;
            }

            set
            {
                hasPerformanceCheck = value;
            }
        }

        internal bool IsPlatform7OrLater { get; private set; }

        private void CheckDevice(string method)
        {
            if (baseClassDevice is null)
                throw new InvalidOperationException($"DeviceTesterbaseClass.CheckDevice A device instance does not exist when called by the {method} method.");
        }

        /// <summary>
        /// Get error codes.
        /// </summary>
        /// <param name="p_ProgID">The ProgID.</param>
        public virtual void InitialiseTest()
        {
            SetTest(""); // Clear status messages
            SetAction("");
            SetStatus("");

            DateTime lastModifiedTime = DateTime.MinValue;
            try
            {
                LogDebug("ConformanceCheck", $"About to get executing assembly...");
                string assemblyName;
                if (OperatingSystem.IsWindows()) assemblyName = "ConformU.exe";
                else assemblyName = "conformu";
                LogDebug("ConformanceCheck", $"Assembly name: {assemblyName}");
                string baseDirectory = AppContext.BaseDirectory;
                LogDebug("ConformanceCheck", $"Base directory: {baseDirectory}");
                string assemblyPath = Path.Combine(baseDirectory, assemblyName);
                var fileInfo = new System.IO.FileInfo(assemblyPath);
                LogDebug("ConformanceCheck", $"Last write time: {fileInfo.LastWriteTime}");
                lastModifiedTime = fileInfo.LastWriteTime;
                LogDebug("", ""); // Blank line
            }
            catch (Exception ex)
            {
                LogTestAndMessage("ConformanceCheck", $"Exception while trying to determine the last modified time: {ex}");
            }

            LogTestOnly($"ASCOM Universal Device Conformance Checker Version {this.GetType().Assembly.GetName().Version}, Build time: {lastModifiedTime:ddd dd MMMM yyyy HH:mm:ss}");
            LogNewLine(); // Blank line

            // Set the architecture and bitness
            string architecture = ((RuntimeInformation.ProcessArchitecture == Architecture.Arm64) | (RuntimeInformation.ProcessArchitecture == Architecture.Arm)) ? "ARM" : "Intel/AMD";
            string osBitness = Environment.Is64BitOperatingSystem ? "64" : "32";
            string processBitness = Environment.Is64BitProcess ? "64" : "32";

            // Set the OS name depending on whether or not it is Windows
            string osName;
#if WINDOWS
            osName = ASCOM.Com.PlatformUtilities.OSBuildName();
#else
            osName = RuntimeInformation.OSDescription;
#endif
            LogTestOnly($"Operating system is {osName}, Processor is {architecture} {osBitness}bit, Application is {processBitness}bit.");
            LogNewLine(); // Blank line

            switch (settings.DeviceTechnology)
            {
                case DeviceTechnology.Alpaca:
                    LogTestOnly($"Alpaca device: {settings.AlpacaDevice.AscomDeviceName} ({settings.AlpacaDevice.IpAddress}:{settings.AlpacaDevice.IpPort} {settings.AlpacaDevice.AscomDeviceType}/{settings.AlpacaDevice.AlpacaDeviceNumber})");
                    if (!settings.AlpacaConfiguration.StrictCasing) LogIssue("ConformanceCheck", "Alpaca strict casing has been disabled, this in only supported for testing devices.");

                    break;

                case DeviceTechnology.COM:
                    LogTestOnly($"COM Driver ProgID: {settings.ComDevice.ProgId}"); break;

                case DeviceTechnology.NotSelected:
                    throw new InvalidValueException($"CheckInitialise - 'NotSelected' is not a technology type, please select Alpaca or COM.");

                default:
                    throw new InvalidValueException($"CheckInitialise - Unknown technology type: {settings.DeviceTechnology}");
            }

            LogNewLine(); // Blank line
        }

        public virtual void CreateDevice()
        {
            LogIssue("CreateDevice", "DeviceTester base Class warning message");
        }

        public virtual void PreConnectChecks()
        {
            LogIssue("PreConnectChecks", "DeviceTester base Class warning message");
        }

        public virtual void ReadCanProperties()
        {
            LogIssue("ReadCanProperties", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PreRunCheck()
        {
            LogIssue("PreSafetyCheck", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckProperties()
        {
            LogIssue("CheckProperties", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckMethods()
        {
            LogIssue("CheckMethods", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckPerformance()
        {
            LogIssue("CheckPerformance", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckConfiguration()
        {
            LogIssue("CheckConfiguration", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PostRunCheck()
        {
            LogIssue("PostSafetyCheck", "DeviceTester base Class warning message, you should not see this message!");
        }

        #endregion

        #region Support code common to all device types

        public void SynchronousBehaviourInformation(string operation, string completionFalse, string completionTrue)
        {
            LogInfo(testName, $"As an I{settings.DeviceType}V{GetInterfaceVersion()} device, {operation} should have operated asynchronously: Returning quickly (less than {Globals.STANDARD_TARGET_RESPONSE_TIME} second) after setting {completionFalse}.");
            LogInfo(testName, $"The {settings.DeviceType} device should then continue the operation in the background and set {completionTrue} when it has completed.");

        }

        /// <summary>
        ///Set the test, action and status in one call
        ///</summary>
        ///<param name="newTestName">Name of the test being conducted</param>
        ///<param name="newTestAction">Specific action within the test</param>
        ///<param name="newTestStatus">Status of the action</param>
        ///<remarks></remarks>
        public void SetFullStatus(string newTestName, string newTestAction, string newTestStatus)
        {
            testName = newTestName;
            testAction = newTestAction;
            testStatus = newTestStatus;

            string displayText = testName;
            if (!string.IsNullOrEmpty(newTestAction)) displayText += $" - {newTestAction}";
            if (!string.IsNullOrEmpty(newTestStatus)) displayText += $" - {newTestStatus}";

            tl?.SetStatusMessage(displayText);
        }

        public void SetTest(string test)
        {
            SetFullStatus(test, "", "");
        }

        public string GetAction()
        {
            return testAction;
        }

        public void SetAction(string action)
        {
            SetFullStatus(testName, action, "");
        }

        public void SetStatus(string status)
        {
            SetFullStatus(testName, testAction, status);
        }

        /// <summary>
        ///Clear all status fields
        ///</summary>
        ///<remarks></remarks>
        public void ClearStatus()
        {
            SetFullStatus(testName, testAction, "");
        }

        public void ResetTestActionStatus()
        {
            SetFullStatus("", "", "");
        }

        /// <summary>
        /// Delay execution for the given time period in milliseconds
        /// </summary>
        /// <param name="waitDuration">Period to wait in milliseconds</param>
        /// <param name="updateInterval">Optional interval between status updates(Default 500ms)</param>
        /// <remarks></remarks>
        internal void WaitFor(int waitDuration, int updateInterval = WAITFOR_UPDATE_INTERVAL)
        {
            if (waitDuration > 0)
            {
                // Ensure that we don't wait more than the expected duration
                if (updateInterval > waitDuration) updateInterval = waitDuration;

                // Only update the status field if the application is not cancelled
                if (!ApplicationCancellationToken.IsCancellationRequested)
                {
                    // Initialise the status message status field 
                    SetStatus($"0.0 / {Convert.ToDouble(waitDuration) / 1000.0:0.0} seconds");
                }

                // Start the loop timing stopwatch
                Stopwatch sw = Stopwatch.StartNew();

                // Wait for p_Duration milliseconds
                do
                {
                    // Calculate the current loop number (starts at 1 given that the timer's elapsed time will be zero or very low on the first loop)
                    int currentLoopNumber = ((int)sw.ElapsedMilliseconds + 50) / updateInterval;

                    // Calculate the sleep time required to start the next loop at a multiple of the poll interval
                    int sleeptime = updateInterval * (currentLoopNumber + 1) - (int)sw.ElapsedMilliseconds;

                    // Ensure that we don't over-wait on the last cycle
                    int remainingWaitTime = waitDuration - (int)sw.ElapsedMilliseconds;
                    if (remainingWaitTime < 0) remainingWaitTime = 0;
                    if (remainingWaitTime < updateInterval) sleeptime = remainingWaitTime;

                    // Sleep until it is time for the next completion function poll
                    Thread.Sleep(sleeptime);

                    // Only update the status field if the application is not cancelled
                    if (!ApplicationCancellationToken.IsCancellationRequested)
                    {
                        // Set the status message status field to the elapsed time
                        SetStatus($"{Math.Round(Convert.ToDouble(currentLoopNumber + 1) * updateInterval / 1000.0, 1):0.0} / {Convert.ToDouble(waitDuration) / 1000.0:0.0} seconds");
                    }
                }
                while ((sw.ElapsedMilliseconds <= waitDuration) & !ApplicationCancellationToken.IsCancellationRequested);
            }
        }

#nullable enable

        /// <summary>
        /// Call the wait function every poll interval milliseconds and delay until the wait function becomes false
        /// </summary>
        /// <param name="actionName">Text to set in the status Action field</param>
        /// <param name="waitFunction">Completion Func that returns false when the process is complete</param>
        /// <param name="pollInterval">Interval between calls of the completion function in milliseconds</param>
        /// <param name="timeoutSeconds">Number of seconds before the operation times out</param>
        /// <exception cref="InvalidValueException"></exception>
        /// <exception cref="TimeoutException">If the operation takes longer than the timeout value</exception>
        internal void WaitWhile(string actionName, Func<bool> waitFunction, int pollInterval, int timeoutSeconds, Func<string>? statusString = null)
        {
            // Validate the supplied poll interval
            if (pollInterval < 100) throw new InvalidValueException($"The poll interval must be >=100ms: {pollInterval}");

            // Only update the status if the task has not been cancelled
            if (!ApplicationCancellationToken.IsCancellationRequested)
            {
                // Set the status message action field to the supplied action name if it has not already been set
                if ((!string.IsNullOrEmpty(actionName)) & (actionName != GetAction()))
                {
                    SetAction(actionName);
                }

                // Initialise the status message
                SetStatus(statusString is null ? $"0.0 / {timeoutSeconds:0.0} seconds" : statusString());
            }

            // Create a timeout cancellation token source that times out after the required timeout period
            CancellationTokenSource timeoutCts = new();
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(Convert.ToDouble(timeoutSeconds)));

            // Combine the provided cancellation token parameter with the new timeout cancellation token
            CancellationTokenSource combinedCts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, ApplicationCancellationToken);

            // Wait for the completion function to return false
            Stopwatch sw = Stopwatch.StartNew(); // Start the loop timing stopwatch
                                                 //do
            while (waitFunction() & !combinedCts.Token.IsCancellationRequested)
            {
                // Calculate the current loop number (starts at 0 given that the timer's elapsed time will be zero or very low on the first loop)
                int currentLoopNumber = ((int)(sw.ElapsedMilliseconds) + 50) / pollInterval; // Add a small positive offset (50) because integer division always rounds down

                // Calculate the sleep time required to start the next loop at a multiple of the poll interval
                int sleeptime = pollInterval * (currentLoopNumber + 1) - (int)sw.ElapsedMilliseconds;

                // Sleep until it is time for the next completion function poll
                Thread.Sleep(sleeptime);

                // Only update the status field if not cancelled
                if (!combinedCts.IsCancellationRequested)
                {
                    // Set the status message status field
                    if (statusString is null) // No status string function was provided so display an elapsed time message
                    {
                        double elapsedTime = Math.Min(Math.Round(Convert.ToDouble(currentLoopNumber + 1) * pollInterval / 1000.0, 1), timeoutSeconds);
                        SetStatus($"{elapsedTime:0.0} / {timeoutSeconds:0.0} seconds");
                    }
                    else // Display the supplied message instead of the elapsed time message
                        SetStatus(statusString());

                }
            }

            bool currentWaitFunctionValue = waitFunction();

            LogDebug(nameof(WaitWhile), $"The Wait completed - Wait was successful: {!currentWaitFunctionValue}, " +
                $"Timeout cancelled: {timeoutCts.Token.IsCancellationRequested}, Application cancelled: {ApplicationCancellationToken.IsCancellationRequested}, " +
                $"Combined cancelled: {combinedCts.Token.IsCancellationRequested}.");
            if (currentWaitFunctionValue)
            {
                // Test whether the operation timed out
                if (timeoutCts.IsCancellationRequested) // The operation did time out
                {
                    // Make a final test of the wait function in case the operation completed between poll intervals
                    if (waitFunction()) // The wait function remains True
                    {
                        //  The wait function is still True so log the timeout and throw an exception to cancel the operation
                        LogDebug(nameof(WaitWhile), $"The {actionName} operation timed out after {timeoutSeconds} seconds.");
                        throw new TimeoutException($"The \"{actionName}\" operation exceeded its {timeoutSeconds} second timeout.");
                    }
                    else // The wait function is False
                    {
                        // The wait function transitioned to False after the last poll interval and before the timeout so this is a successful wait
                        // No action required
                    }
                }
            }
            else
            {
                //LogDebug(nameof(WaitWhile), $"WaitFunction is FALSE - Operation completed successfully");
            }

            SetStatus("");
        }

#nullable disable

        internal void LogNewLine()
        {
            LogTestOnly("");
        }

        internal void LogTestOnly(string test)
        {
            tl?.LogMessage(test, MessageLevel.TestOnly, "");
        }

        internal void LogTestAndMessage(string test, string message)
        {
            tl?.LogMessage(test, MessageLevel.TestAndMessage, message);
        }

        internal void LogOk(string test, string message)
        {
            tl?.LogMessage(test, MessageLevel.OK, message);
        }

        internal void LogDebug(string test, string message)
        {
            tl?.LogMessage(test, MessageLevel.Debug, message);
        }

        internal void LogInfo(string test, string message)
        {
            tl?.LogMessage(test, MessageLevel.Info, message);
        }

        internal void LogIssue(string test, string message)
        {
            conformResults.Issues.Add(new KeyValuePair<string, string>(test, message));
            tl?.LogMessage(test, MessageLevel.Issue, message);
        }

        internal void LogTiming(string test, string message)
        {
            conformResults.Timings.Add(new KeyValuePair<string, string>(test, message));
        }

        internal void LogError(string test, string message)
        {
            conformResults.Errors.Add(new System.Collections.Generic.KeyValuePair<string, string>(test, message));
            tl?.LogMessage(test, MessageLevel.Error, message);
        }

        internal void LogMsg(string testName, MessageLevel messageLevel, string message)
        {
            tl?.LogMessage(testName, messageLevel, message);
        }

        internal static void LogConfigurationAlert(string message)
        {
            conformResults.ConfigurationAlerts.Add(new System.Collections.Generic.KeyValuePair<string, string>("Conform configuration", message));
        }

        /// <summary>
        /// Test a supplied exception for whether it is a MethodNotImplemented type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotImplemented type</returns>
        /// <remarks>Different tests are applied for COM and MethodNotImplemmented exceptions</remarks>
        protected bool IsMethodNotImplementedException(Exception deviceException)
        {
            bool isMethodNotImplementedExceptionRet = false; // Set false default value
            try
            {
                switch (deviceException)
                {
                    // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                    case COMException exception:
                        if (exception.ErrorCode == ExNotImplemented | exception.ErrorCode == ErrorCodes.NotImplemented) // This is a not implemented exception
                            isMethodNotImplementedExceptionRet = true;
                        break;

                    case MethodNotImplementedException:
                        isMethodNotImplementedExceptionRet = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError("IsMethodNotImplementedException", $"Unexpected exception: {ex}");
            }

            return isMethodNotImplementedExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is a NotImplemented type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotImplemented type</returns>
        /// <remarks>Different tests are applied for COM and .NET exceptions</remarks>
        protected bool IsNotImplementedException(Exception deviceException)
        {
            bool isNotImplementedExceptionRet = false; // Set false default value
            try
            {
                switch (deviceException)
                {
                    // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                    case COMException exception:
                        if (exception.ErrorCode == ExNotImplemented | exception.ErrorCode == ErrorCodes.NotImplemented) // This is a not implemented exception
                            isNotImplementedExceptionRet = true;
                        break;

                    case NotImplementedException:
                        isNotImplementedExceptionRet = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError("IsNotImplementedException", $"Unexpected exception: {ex}");
            }

            return isNotImplementedExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is a PropertyNotImplementedException type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotImplemented type</returns>
        /// <remarks>Different tests are applied for COM and PropertyNotImplemmented exceptions</remarks>
        protected bool IsPropertyNotImplementedException(Exception deviceException)
        {
            bool isPropertyNotImplementedExceptionRet = false; // Set false default value
            try
            {
                switch (deviceException)
                {
                    // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                    case COMException exception:
                        if (exception.ErrorCode == ExNotImplemented | exception.ErrorCode == ErrorCodes.NotImplemented) // This is a not implemented exception
                            isPropertyNotImplementedExceptionRet = true;
                        break;

                    case PropertyNotImplementedException:
                        isPropertyNotImplementedExceptionRet = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError("IsPropertyNotImplementedException", $"Unexpected exception: {ex}");
            }

            return isPropertyNotImplementedExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is an InvalidValueException type
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a InvalidValue type</returns>
        /// <remarks>Different tests are applied for COM and InvalidValueException exceptions</remarks>
        protected bool IsInvalidValueException(string memberName, Exception deviceException)
        {
            bool isInvalidValueExceptionRet = false; // Set false default value
            try
            {
                switch (deviceException)
                {
                    // This is a COM exception so test whether the error code indicates that it is an invalid value exception
                    case COMException exception:
                        if ((exception.ErrorCode == ErrorCodes.InvalidValue) |
                            (exception.ErrorCode == ExInvalidValue1) |
                            (exception.ErrorCode == ExInvalidValue2) |
                            (exception.ErrorCode == ExInvalidValue3) |
                            (exception.ErrorCode == ExInvalidValue4) |
                            (exception.ErrorCode == ExInvalidValue5) |
                            (exception.ErrorCode == ExInvalidValue6)) // This is an invalid value exception
                            isInvalidValueExceptionRet = true;
                        break;

                    case InvalidValueException:
                        isInvalidValueExceptionRet = true;
                        break;

                    case ASCOM.InvalidOperationException:
                        LogIssue(memberName, "Received ASCOM.InvalidOperationException rather than ASCOM.InvalidValueException");
                        break;

                    case InvalidOperationException:
                        LogIssue(memberName, "Received System.InvalidOperationException rather than ASCOM.InvalidValueException");
                        break;

                    case DriverException exception1:
                        if (exception1.Number == ErrorCodes.InvalidValue) // This is an invalid value exception
                            LogIssue(memberName, $"Received ASCOM.DriverException(0x{ErrorCodes.InvalidValue:X8}), please use ASCOM.InvalidValueException to report invalid values");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError("IsInvalidValueException", $"Unexpected exception: {ex}");
            }

            return isInvalidValueExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is an InvalidValueException type
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a InvalidValue type</returns>
        /// <remarks>Different tests are applied for COM and InvalidValueException exceptions</remarks>
        protected bool IsInvalidOperationException(string memberName, Exception deviceException)
        {
            bool isInvalidOperationExceptionRet = false; // Set false default value
            try
            {
                switch (deviceException)
                {
                    // This is a COM exception so test whether the error code indicates that it is an invalid operation exception
                    case COMException exception:
                        if (exception.ErrorCode == ErrorCodes.InvalidOperationException) // This is an invalid operation exception
                            isInvalidOperationExceptionRet = true;
                        break;

                    case ASCOM.InvalidOperationException:
                        isInvalidOperationExceptionRet = true;
                        break;

                    case DriverException exception1:
                        if (exception1.Number == ErrorCodes.InvalidOperationException) // This is an invalid operation exception
                            LogIssue(memberName, $"Received ASCOM.DriverException(0x{ErrorCodes.InvalidOperationException:X8}), please use ASCOM.InvalidOperationException to report invalid operations");
                        break;

                    case InvalidOperationException:
                        LogIssue(memberName, "Received System.InvalidOperationException rather than ASCOM.InvalidOperationException");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError("IsInvalidOperationException", $"Unexpected exception: {ex}");
            }

            return isInvalidOperationExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is a NotSetException type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotSet type</returns>
        /// <remarks>Different tests are applied for COM and ValueNotSetException exceptions</remarks>
        protected bool IsNotSetException(Exception deviceException)
        {
            bool isNotSetExceptionRet = false; // Set false default value
            try
            {
                switch (deviceException)
                {
                    // This is a COM exception so test whether the error code indicates that it is a not set exception
                    case COMException exception:
                        if (exception.ErrorCode == ExNotSet1) // This is a not set exception
                            isNotSetExceptionRet = true;
                        break;

                    case ValueNotSetException:
                        isNotSetExceptionRet = true;
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError("IsNotSetException", $"Unexpected exception: {ex}");
            }

            return isNotSetExceptionRet;
        }

        /// <summary>
        /// Provides messaging when an exception is thrown by a member
        /// </summary>
        /// <param name="memberName">The name of the member throwing the exception</param>
        /// <param name="typeOfMember">Flag indicating whether the member is a property or a method</param>
        /// <param name="isRequired">Flag indicating whether the member is optional or mandatory</param>
        /// <param name="ex">The exception received from the device</param>
        /// <param name="userMessage">The member specific message to report</param>
        /// <remarks></remarks>
        protected void HandleException(string memberName, MemberType typeOfMember, Required isRequired, Exception ex, string userMessage)
        {

            // Handle PropertyNotImplemented exceptions from properties and MethodNotImplemented exceptions from methods and NotImplementedExceptions from Alpaca devices
            if (IsPropertyNotImplementedException(ex) & typeOfMember == MemberType.Property |
                IsMethodNotImplementedException(ex) & typeOfMember == MemberType.Method |
                IsNotImplementedException(ex) & settings.DeviceTechnology == DeviceTechnology.Alpaca)
            {
                switch (isRequired)
                {
                    case Required.Mandatory:
                        LogIssue(memberName, $"This member is mandatory but returned a {GetExceptionName(ex, typeOfMember)} error, it must function per the ASCOM specification.");
                        break;

                    case Required.MustNotBeImplemented:
                        LogOk(memberName, $"{userMessage} and a {GetExceptionName(ex, typeOfMember)} error was generated as expected");
                        break;

                    case Required.MustBeImplemented:
                        LogIssue(memberName, $"{userMessage} and a {GetExceptionName(ex, typeOfMember)} error was returned, this method must function per the ASCOM specification.");
                        break;

                    case Required.Optional:
                        LogOk(memberName, $"Optional member returned a {GetExceptionName(ex, typeOfMember)} error.");
                        break;

                    default:
                        LogError(memberName, $"CONFORM ERROR! - Received unexpected member of 'Required' enum: {isRequired}");
                        break;
                }
            }

            // Handle wrong type of not implemented exceptions
            else if (ex is MethodNotImplementedException & typeOfMember == MemberType.Property) // We got a MethodNotImplementedException so this is an error
            {
                LogIssue(memberName, "Received a MethodNotImplementedException instead of a PropertyNotImplementedException");
            }
            else if (ex is PropertyNotImplementedException & typeOfMember == MemberType.Method) // We got a PropertyNotImplementedException so this is an error
            {
                LogIssue(memberName, "Received a PropertyNotImplementedException instead of a MethodNotImplementedException");
            }
            else if (ex is NotImplementedException)
            {
                // Determine whether we are testing an Alpaca or a COM device
                if (settings.DeviceTechnology == DeviceTechnology.Alpaca) // We are testing an ALpaca device
                {
                    // NotImplementedException is expected from an Alpaca device so flag as OK
                    LogOk(memberName, "Received a NotImplementedException from an Alpaca device as expected");
                }
                else // We are testing a COM device
                {
                    // ASCOM.NotImplementedException is expected if received from the cross platform library DriverAccess module, otherwise it is an issue, so test for this condition.
                    if ((settings.ComConfiguration.ComAccessMechanic == ComAccessMechanic.DriverAccess) & (settings.DeviceTechnology == DeviceTechnology.COM)) // We are testing a COM device using the cross platform DriverAccess module so report OK.
                    {
                        LogOk(memberName, "Received a NotImplementedException from DriverAccess as expected");
                    }
                    else // We are NOT testing a COM device using the cross platform DriverAccess module so report an issue.
                    {
                        LogIssue(memberName, $"Received a NotImplementedException instead of a {((typeOfMember == MemberType.Property) ? "PropertyNotImplementedException" : "MethodNotImplementedException")}");
                    }
                }
            }
            else if (ex is System.NotImplementedException)
            {
                LogIssue(memberName,
                    $"Received a System.NotImplementedException instead of an ASCOM.{((typeOfMember == MemberType.Property) ? "PropertyNotImplementedException" : "MethodNotImplementedException")}");
            }

            // Handle all other types of error
            else
            {
                LogIssue(memberName, $"Unexpected error{(string.IsNullOrEmpty(userMessage) ? ":" : $" - {userMessage}:")} {ex.Message}");
            }

            LogDebug(memberName, $"Exception detail: {ex}");
        }

        protected void HandleInvalidValueExceptionAsOk(string memberName, MemberType typeOfMember, Required isRequired, Exception ex, string userAction, string message)
        {
            if (IsInvalidValueException(memberName, ex))
                LogOk(memberName, message);
            else
                HandleException(memberName, typeOfMember, isRequired, ex, userAction);
        }

        protected void HandleInvalidValueExceptionAsInfo(string memberName, MemberType typeOfMember, Required isRequired, Exception ex, string userAction, string message)
        {
            if (IsInvalidValueException(memberName, ex))
                LogInfo(memberName, message);
            else
                HandleException(memberName, typeOfMember, isRequired, ex, userAction);
        }

        protected void HandleInvalidOperationExceptionAsOk(string memberName, MemberType typeOfMember, Required isRequired, Exception ex, string userAction, string message)
        {
            if (IsInvalidOperationException(memberName, ex))
                LogOk(memberName, message);
            else
                HandleException(memberName, typeOfMember, isRequired, ex, userAction);
        }

        /// <summary>
        /// Get an exception name (and number if a COM or Driver exception)
        /// </summary>
        /// <param name="clientException">Exception whose name is required</param>
        /// <param name="memberType">Property or Method.</param>
        /// <returns>String exception name</returns>
        /// <remarks></remarks>
        protected static string GetExceptionName(Exception clientException, MemberType memberType)
        {
            string retVal;

            // Treat ASCOM exceptions specially
            if (clientException.GetType().FullName.Contains("ASCOM", StringComparison.InvariantCultureIgnoreCase))
            {
                if (clientException.GetType().FullName.Contains("DRIVEREXCEPTION", StringComparison.InvariantCultureIgnoreCase)) // We have a driver exception so add its number
                {
                    DriverException driverEx = (DriverException)clientException;
                    retVal = $"DriverException(0x{driverEx.Number:X8})";
                }
                else // Otherwise just use the ASCOM exception's name
                {
                    retVal = clientException.GetType().Name;
                }
            }
            else if (clientException is COMException comEx) // Handle XOM exceptions with their error code
            {
                try
                {
                    string exceptionName = ErrorCodes.GetExceptionName(comEx);

                    retVal = $"{(String.IsNullOrEmpty(exceptionName) ? comEx.GetType().Name : exceptionName)} (COM Error: 0x{comEx.ErrorCode:X8})";
                    if (comEx.ErrorCode == ErrorCodes.NotImplemented)
                    {
                        retVal = $"{memberType}{retVal}";
                    }
                }
                catch (Exception)
                {
                    retVal = $"COMException(0x{comEx.ErrorCode:X8})";
                }
            }
            else // We got something else so report it
            {
                retVal = $"{clientException.GetType().FullName} exception";
            }

            return retVal;
        }

        /// <summary>
        /// Logs a call to a driver if enabled within Conform's configuration
        /// </summary>
        /// <param name="test">Name of the current test</param>
        /// <param name="memberName">Name of member being called</param>
        protected void LogCallToDriver(string test, string memberName)
        {
            if (settings.DisplayMethodCalls)
                LogTestAndMessage(test, memberName);
        }

        #endregion

        #region Timing methods and functions

        /// <summary>
        /// Execute an interface method that has no parameters and report the timing
        /// </summary>
        /// <param name="methodName">Name of the method within the Action</param>
        /// <param name="method">Action representing the method being called</param>
        protected void TimeMethod(string methodName, Action method, TargetTime targetTime)
        {
            // Time the method call
            Stopwatch sw = Stopwatch.StartNew();
            method();
            sw.Stop();

            // Report the timing outcome
            ReportTiming(methodName, sw.Elapsed.TotalSeconds, targetTime);
        }

        /// <summary>
        /// Execute an interface method that has one parameter and report the timing
        /// </summary>
        /// <typeparam name="T">Parameter's type</typeparam>
        /// <param name="methodName">Name of the method within the Action</param>
        /// <param name="method">Action representing the method being called</param>
        /// <param name="param">Method's parameter</param>
        protected void TimeMethodOneParam<T>(string methodName, Action<T> method, T param, TargetTime targetTime)
        {
            // Time the method call
            Stopwatch sw = Stopwatch.StartNew();
            method(param);
            sw.Stop();

            // Report the timing outcome
            ReportTiming(methodName, sw.Elapsed.TotalSeconds, targetTime);
        }

        /// <summary>
        /// Execute an interface method that has two parameters and report the timing
        /// </summary>
        /// <typeparam name="T1">First parameter's type</typeparam>
        /// <typeparam name="T2">Second parameter's type</typeparam>
        /// <param name="methodName">Name of the method within the Action</param>
        /// <param name="method">Action representing the method being called</param>
        /// <param name="param1">Method's first parameter</param>
        /// <param name="param2">Method's second parameter</param>
        protected void TimeMethodTwoParams<T1, T2>(string methodName, Action<T1, T2> method, T1 param1, T2 param2, TargetTime targetTime)
        {
            // Time the method call
            Stopwatch sw = Stopwatch.StartNew();
            method(param1, param2);
            sw.Stop();

            // Report the timing outcome
            ReportTiming(methodName, sw.Elapsed.TotalSeconds, targetTime);
        }

        /// <summary>
        /// Execute an interface property or function that has no parameters and report the timing
        /// </summary>
        /// <typeparam name="TResult">The property or function's return type</typeparam>
        /// <param name="methodName">Name of the property or function within the Func</param>
        /// <param name="function">Func<TResult> representing the property or function being called</param>
        /// <returns>The property or function result.</returns>
        internal TResult TimeFunc<TResult>(string methodName, Func<TResult> function, TargetTime targetTime)
        {
            // Time the method call
            Stopwatch sw = Stopwatch.StartNew();
            TResult response = function();
            sw.Stop();

            // Report the timing outcome
            ReportTiming(methodName, sw.Elapsed.TotalSeconds, targetTime);

            // Return the device response
            return response;
        }

        /// <summary>
        /// Execute an interface property or function that has one parameter and report the timing
        /// </summary>
        /// <typeparam name="T1">First parameter's type</typeparam>
        /// <typeparam name="TResult">The property or function's return type</typeparam>
        /// <param name="methodName">Name of the property or function within the Func</param>
        /// <param name="function">Func<TResult> representing the property or function being called</param>
        /// <param name="param1">Method's first parameter</param>
        /// <returns>The property or function result.</returns>
        internal TResult TimeFuncOneParam<T1, TResult>(string methodName, Func<T1, TResult> function, T1 param1, TargetTime targetTime)
        {
            // Time the method call
            Stopwatch sw = Stopwatch.StartNew();
            TResult response = function(param1);
            sw.Stop();

            // Report the timing outcome
            ReportTiming(methodName, sw.Elapsed.TotalSeconds, targetTime);

            // Return the device response
            return response;
        }

        /// <summary>
        /// Execute an interface property or function that has two parameters and report the timing
        /// </summary>
        /// <typeparam name="T1">First parameter's type</typeparam>
        /// <typeparam name="T2">Second parameter's type</typeparam>
        /// <typeparam name="TResult">The property or function's return type</typeparam>
        /// <param name="methodName">Name of the property or function within the Func</param>
        /// <param name="function">Func<TResult> representing the property or function being called</param>
        /// <param name="param1">Method's first parameter</param>
        /// <param name="param2">Method's second parameter</param>
        /// <returns>The property or function result.</returns>
        internal TResult TimeFuncTwoParams<T1, T2, TResult>(string methodName, Func<T1, T2, TResult> function, T1 param1, T2 param2, TargetTime targetTime)
        {
            // Time the method call
            Stopwatch sw = Stopwatch.StartNew();
            TResult response = function(param1, param2);
            sw.Stop();

            // Report the timing outcome
            ReportTiming(methodName, sw.Elapsed.TotalSeconds, targetTime);

            // Return the device response
            return response;
        }

        /// <summary>
        /// Report the timing test outcome
        /// </summary>
        /// <param name="methodName">Name of the property or function within the Func</param>
        /// <param name="elapsedTime">Elapsed time running the member (seconds)</param>
        private void ReportTiming(string methodName, double elapsedTime, TargetTime targetTime)
        {
            switch (targetTime)
            {
                case TargetTime.Fast:
                    ReportTiming(methodName, elapsedTime, "FAST", Globals.FAST_TARGET_RESPONSE_TIME);
                    break;

                case TargetTime.Standard:
                    ReportTiming(methodName, elapsedTime, "STANDARD", Globals.STANDARD_TARGET_RESPONSE_TIME);
                    break;

                case TargetTime.Extended:
                    ReportTiming(methodName, elapsedTime, "EXTENDED", Globals.EXTENDED_TARGET_RESPONSE_TIME);
                    break;

                default:
                    throw new InvalidValueException($"DeviceTesterBaseClass-ReportTiming - Unsupported target time: {targetTime}");
            }
        }

        private void ReportTiming(string memberName, double elapsedTime, string targetName, double targetTime)
        {
            // Increment the count of timing measurements
            conformResults.TimingCount++;

            int memberNamePadWidth = 24; // Set a default method name column width

            switch (baseClassDeviceType)
            {
                case DeviceTypes.Camera:
                    break;

                case DeviceTypes.CoverCalibrator:
                    memberNamePadWidth = 20;
                    break;

                case DeviceTypes.Dome:
                    break;

                case DeviceTypes.FilterWheel:
                    break;

                case DeviceTypes.Focuser:
                    break;

                case DeviceTypes.ObservingConditions:
                    memberNamePadWidth = 33;
                    break;

                case DeviceTypes.Rotator:
                    break;

                case DeviceTypes.SafetyMonitor:
                    break;

                case DeviceTypes.Switch:
                    memberNamePadWidth = 25;
                    break;

                case DeviceTypes.Telescope:
                    memberNamePadWidth = 42;
                    break;

                case DeviceTypes.Video:
                    break;

                default:
                    break;
            }

            // Determine whether the member run within the target time
            if (elapsedTime <= targetTime) // Member completed within the target time
            {
                if (settings.ReportGoodTimings)
                {
                    LogTiming($"{memberName}", $"At {DateTime.Now:HH:mm:ss.fff} {memberName.PadRight(memberNamePadWidth)} {elapsedTime:0.000} seconds. {((char)0x2713)} ({targetName})"); // 0x2713 is the UTF16 tick character.
                }
            }
            else // Member took longer than the target time
            {
                if (settings.ReportBadTimings)
                {
                    LogTiming($"{memberName}", $"At {DateTime.Now:HH:mm:ss.fff} {memberName.PadRight(memberNamePadWidth)} {elapsedTime:0.000} seconds. OUTSIDE {targetName} RESPONSE TIME TARGET: {targetTime:0.0} seconds.");
                    conformResults.TimingIssuesCount++;
                }
            }
        }

        #endregion

        #region Base class support Code

        private static bool IncludeMethod(MandatoryMethod method, DeviceTypes deviceType, int interfaceVersion)
        {
            // This mechanic allows individual tests for particular devices to be skipped. It is no longer required because this is handled by DriverAccess
            // The code is left in place in case it is ever needed in the future

            bool retVal = true; // Default to true as most methods will be tested , we just list the exceptions to this below

            // Matrix controlling what tests
            switch (deviceType)
            {
                case DeviceTypes.Telescope:
                    switch (interfaceVersion)
                    {
                        case 1: // Telescope interface V1 does not have Driver Version
                            if (method == MandatoryMethod.DriverVersion)
                                retVal = false;
                            break;

                        default:
                            retVal = true; // All methods in all interface versions are mandatory
                            break;
                    }

                    break;

                case DeviceTypes.Camera:
                    retVal = true;
                    break;
            }

            return retVal;
        }

        public void SetDevice(IAscomDeviceV2 device, DeviceTypes deviceType)
        {
            this.baseClassDevice = device;
            this.baseClassDeviceType = deviceType;
        }

        #endregion

        #region Private tests

        public void SpecialTests(SpecialTest test)
        {
            switch (test)
            {
                case SpecialTest.TelescopeSideOfPier:
                    SpecialTelescopeSideOfPier();
                    break;

                case SpecialTest.TelescopeDestinationSideOfPier:
                    SpecialTelescopeDestinationSideOfPier();
                    break;

                case SpecialTest.TelescopeSideOfPierAnalysis:
                    SpecialTelescopeSideOfPierAnalysis();
                    break;

                case SpecialTest.TelescopeCommands:
                    SpecialTelescopeCommands();
                    break;

                default:
                    LogIssue("DeviceTesterBaseClass:SpecialTests", $"Unknown test: {test}");
                    break;
            }
        }

        protected virtual void SpecialTelescopeSideOfPier()
        {
            LogIssue("SpecialTelescopeSideOfPier", "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeDestinationSideOfPier()
        {
            LogIssue("SpecialTelescopeDestinationSideOfPier", "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeSideOfPierAnalysis()
        {
            LogIssue("SpecialTelescopeFlipRange", "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeCommands()
        {
            LogIssue("SpecialTelescopeCommands", "DeviceTester base Class warning message, you should not see this message!");
        }

        #endregion

    }
}
// Base class from which particular device testers are derived
// Put all common elements in here
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Threading;
using ASCOM;
using static ConformU.Globals;
using System.IO;

namespace ConformU
{

    /// <summary>
    /// Base class for device tester classes. Contains common code and placeholders for the 
    /// methods that must be implemented in the device tester class
    /// </summary>
    internal class DeviceTesterBaseClass : IDisposable
    {
        // Implements IDisposable

        #region Variables and Constants
        #region Constants
        internal const string NOT_IMP_NET = ".NET - Feature not implemented";
        internal const string NOT_IMP_COM = "COM - Feature not implemented";
        internal const string EX_DRV_NET = ".NET - Driver Exception: ";
        internal const string EX_NET = ".NET - Exception: ";
        internal const string EX_COM = "COM - Exception: ";
        internal const double PERF_LOOP_TIME = 5.0; // Performance loop run time in seconds
        internal const int SLEEP_TIME = 200; // Loop time for testing whether slewing has completed
        internal const int CAMERA_SLEEP_TIME = 10; // Loop time for testing whether camera events have completed
        internal const int DEVICE_DESTROY_WAIT = 500; // Time to wait after destroying a device before continuing

        #endregion

        internal int g_InterfaceVersion; // Variable to held interface version of the current device

        private bool l_Connected, l_HasProperties, l_HasCanProperties, l_HasMethods, l_HasPreRunCheck, l_HasPostRunCheck, l_HasPerformanceCheck;
        private bool l_HasPreConnectCheck;
        internal dynamic baseClassDevice; // IAscomDriverV1

        private string test, action, status;

        private readonly ConformLogger TL;
        private readonly CancellationToken cancellationToken;

        private readonly Settings settings;

        internal int g_ExNotImplemented, g_ExNotSet1, g_ExNotSet2;
        internal int g_ExInvalidValue1, g_ExInvalidValue2, g_ExInvalidValue3, g_ExInvalidValue4, g_ExInvalidValue5, g_ExInvalidValue6;

        #endregion

        #region Enums
        private enum CommandType
        {
            tstCommandString = 1,
            tstCommandBool = 2,
            tstCommandBlind = 3,
            tstCommandStringRaw = 4,
            tstCommandBoolRaw = 5,
            tstCommandBlindRaw = 6
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

        #region Enums

        internal enum StatusType
        {
            staTest = 1,
            staAction = 2,
            staStatus = 3
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
            CommandXXX = 5
        }
        #endregion



        #endregion

        #region New and Dispose
        public DeviceTesterBaseClass() : base()
        {
            l_HasPreConnectCheck = false;
            l_Connected = false;
            l_HasPreRunCheck = false;
            l_HasCanProperties = false;
            l_HasProperties = true;
            l_HasMethods = true;
            l_HasPostRunCheck = false;
            l_HasPerformanceCheck = true;
            ClearStatus();
        }

        /// <summary>
        /// Initialise device tester base class
        /// </summary>
        /// <param name="HasCanProperties">Device has Can properties</param>
        /// <param name="HasProperties">Device has properties</param>
        /// <param name="HasMethods">Device has methods</param>
        /// <param name="HasPreRunCheck">Device requires a pre run safety check </param>
        /// <param name="HasPreConnectCheck">Device requires a pre connection check</param>
        /// <param name="HasPerformanceCheck">Device has a performance test</param>
        /// <param name="HasPostRunCheck">Device requires a post run safety check</param>
        /// <remarks></remarks>
        public DeviceTesterBaseClass(bool HasCanProperties, bool HasProperties, bool HasMethods, bool HasPreRunCheck, bool HasPreConnectCheck, bool HasPerformanceCheck, bool HasPostRunCheck, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken cancellationToken) : this()
        {
            l_HasPreConnectCheck = HasPreConnectCheck;
            l_Connected = false;
            l_HasPreRunCheck = HasPreRunCheck;
            l_HasCanProperties = HasCanProperties;
            l_HasProperties = HasProperties;
            l_HasMethods = HasMethods;
            l_HasPostRunCheck = HasPostRunCheck;
            l_HasPerformanceCheck = HasPerformanceCheck;
            TL = logger;
            this.cancellationToken = cancellationToken;
            settings = conformConfiguration.Settings;
        }

        private bool disposedValue = false;        // To detect redundant calls
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    TL.LogMessage("DeviceTesterbaseClass","Dispose called");
                    try { baseClassDevice.Dispose(); } catch (Exception ex){ TL.LogMessage("DeviceTesterbaseClass", $"Exception\r\n{ex}"); }
                    baseClassDevice = null;
                    TL.LogMessage("DeviceTesterbaseClass", "Dispose finished");
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
            baseClassDevice.SetupDialog();
        }

        public void CheckCommonMethods(object p_DeviceObject, DeviceType p_DeviceType)
        {
            string m_DriverVersion, m_DriverInfo, m_Description, m_Name; // , m_LastResult As String
            IList SA;

            // Dim m_Configuration, SC() As String
            bool m_Connected;
            LogTestOnly("Common Driver Methods");
            baseClassDevice = p_DeviceObject; // CType(DeviceObject, IAscomDriverV1)

            // InterfaceVersion - Required
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("InterfaceVersion", "About to get property InterfaceVersion");
                g_InterfaceVersion = baseClassDevice.InterfaceVersion;
                switch (g_InterfaceVersion)
                {
                    case var @case when @case < 1:
                        {
                            LogIssue("InterfaceVersion", "InterfaceVersion must be 1 or greater but driver returned: " + g_InterfaceVersion.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK("InterfaceVersion", g_InterfaceVersion.ToString());
                            break;
                        }
                }

            }
            catch (Exception ex)
            {
                HandleException("InterfaceVersion", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (cancellationToken.IsCancellationRequested)
                return;

            // Connected - Required
            if (IncludeMethod(MandatoryMethod.Connected, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("Connected", "About to get property Connected");
                    m_Connected = baseClassDevice.Connected;
                    LogOK("Connected", m_Connected.ToString());
                }
                catch (COMException ex)
                {
                    LogError("Connected", EX_COM + ex.Message + " " + ex.ErrorCode.ToString("X8"));
                }
                catch (Exception ex)
                {
                    LogError("Connected", ex.Message);
                }

                if (cancellationToken.IsCancellationRequested)
                    return;
            }

            // Description - Required
            if (IncludeMethod(MandatoryMethod.Description, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("Description", "About to get property Description");
                    m_Description = baseClassDevice.Description;
                    switch (m_Description ?? "")
                    {
                        case var case1 when case1 == "":
                            {
                                LogInfo("Description", "No description string");
                                break;
                            }

                        default:
                            {
                                if (m_Description.Length > 68 & p_DeviceType == DeviceType.Camera)
                                {
                                    LogError("Description", "Maximum number of characters is 68 for compatibility with FITS headers, found: " + m_Description.Length + " characters: " + m_Description);
                                }
                                else
                                {
                                    LogOK("Description", m_Description.ToString());
                                }

                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Description", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (cancellationToken.IsCancellationRequested)
                    return;
            }

            // DriverInfo - Required
            if (IncludeMethod(MandatoryMethod.DriverInfo, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("DriverInfo", "About to get property DriverInfo");
                    m_DriverInfo = baseClassDevice.DriverInfo;
                    switch (m_DriverInfo ?? "")
                    {
                        case var case2 when case2 == "":
                            {
                                LogInfo("DriverInfo", "No DriverInfo string");
                                break;
                            }

                        default:
                            {
                                LogOK("DriverInfo", m_DriverInfo.ToString());
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("DriverInfo", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (cancellationToken.IsCancellationRequested)
                    return;
            }

            // DriverVersion - Required
            if (IncludeMethod(MandatoryMethod.DriverVersion, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("DriverVersion", "About to get property DriverVersion");
                    m_DriverVersion = baseClassDevice.DriverVersion;
                    switch (m_DriverVersion ?? "")
                    {
                        case var case3 when case3 == "":
                            {
                                LogInfo("DriverVersion", "No DriverVersion string");
                                break;
                            }

                        default:
                            {
                                LogOK("DriverVersion", m_DriverVersion.ToString());
                                break;
                            }
                    }
                }
                catch (COMException ex)
                {
                    LogError("DriverVersion", EX_COM + ex.Message + " " + ex.ErrorCode.ToString("X8"));
                }
                catch (Exception ex)
                {
                    LogError("DriverVersion", ex.Message);
                }

                if (cancellationToken.IsCancellationRequested)
                    return;
            }
            else
            {
                LogInfo("DriverVersion", "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // Name - Required
            if (IncludeMethod(MandatoryMethod.Name, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogTestAndMessage("Name", "About to get property Name");
                    m_Name = baseClassDevice.Name;
                    switch (m_Name ?? "")
                    {
                        case var case4 when case4 == "":
                            {
                                LogInfo("Name", "Name is empty");
                                break;
                            }

                        default:
                            {
                                LogOK("Name", m_Name);
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Name", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (cancellationToken.IsCancellationRequested)
                    return;
            }

            // CommandXXX tests - Optional
            //if (IncludeMethod(MandatoryMethod.CommandXXX, p_DeviceType, g_InterfaceVersion))
            //{

            //    // Validate that the g_TelescopeTests collection is in a healthy state
            //    try
            //    {
            //        LogMsgDebug("CommandXXX Tests", $"Test collection is null: {Information.IsNothing(g_TelescopeTests)}");
            //        LogMsgDebug("CommandXXX Tests", $"Test collection size: {g_TelescopeTests.Count}");
            //        foreach (KeyValuePair<string, CheckState> kvp in g_TelescopeTests)
            //            LogMsgDebug("CommandXXX Tests", $"Found key: {kvp.Key} = {g_TelescopeTests[kvp.Key]}");
            //    }
            //    catch (Exception ex)
            //    {
            //        LogMsgDebug("CommandXXX Tests", $"Exception: {ex}");
            //    }

            //    if (g_TelescopeTests[TELTEST_COMMANDXXX] == CheckState.Checked)
            //    {
            //        LogMsgInfo("CommandString", "Conform cannot test the CommandString method");
            //        LogMsgInfo("CommandBlind", "Conform cannot test the CommandBlind method");
            //        LogMsgInfo("CommandBool", "Conform cannot test the CommandBool method");
            //    }
            //    else
            //    {
            //        LogMsgInfo(TELTEST_COMMANDXXX, "Tests skipped");
            //    }
            //}

            // Action - optional but cannot be tested
            LogInfo("Action", "Conform cannot test the Action method");

            // Supported actions - Optional but Required through DriverAccess
            try
            {
                if (settings.DisplayMethodCalls)
                    LogTestAndMessage("SupportedActions", "About to call method SupportedActions");
                SA = (IList)baseClassDevice.SupportedActions;
                if (SA.Count == 0)
                {
                    LogOK("SupportedActions", "Driver returned an empty action list");
                }
                else
                {
                    var i = default(int);
                    foreach (object Action in SA)
                    {
                        i += 1;
                        if (Action.GetType().Name == "String")
                        {
                            string ActionString = Action.ToString();
                            string result;
                            const string TEST_PARAMETERS = "Conform test parameters";
                            switch (ActionString ?? "")
                            {
                                case var case5 when case5 == "":
                                    {
                                        LogError("SupportedActions", "Supported action " + i + " Is an empty string"); // List the action that was found
                                        break;
                                    }

                                default:
                                    {
                                        LogOK("SupportedActions", "Found action: " + ActionString);

                                        // Carry out the following Action tests only when we are testing the Observing Conditions Hub and it is configured to use the Switch and OC simulators
                                        if (p_DeviceType == DeviceType.ObservingConditions & settings.DeviceTechnology == DeviceTechnology.COM & settings.ComDevice.ProgId.ToUpper() == "ASCOM.OCH.OBSERVINGCONDITIONS")
                                        {
                                            if (ActionString.ToUpperInvariant().StartsWith("//OCSIMULATOR:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SupportedActions", "About to call method Action");
                                                    result = baseClassDevice.Action(ActionString, TEST_PARAMETERS);
                                                    LogOK("SupportedActions", string.Format("OC simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogError("SupportedActions", string.Format("Exception calling OCH simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//ASCOM.SIMULATOR.OBSERVINGCONDITIONS:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SupportedActions", "About to call method Action");
                                                    result = baseClassDevice.Action(ActionString, TEST_PARAMETERS);
                                                    LogOK("SupportedActions", string.Format("OC simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogError("SupportedActions", string.Format("Exception calling OCH simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//SWITCHSIMULATOR:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SupportedActions", "About to call method Action");
                                                    result = baseClassDevice.Action(ActionString, TEST_PARAMETERS);
                                                    LogOK("SupportedActions", string.Format("Switch simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogError("SupportedActions", string.Format("Exception calling switch simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//ASCOM.SIMULATOR.SWITCH:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogTestAndMessage("SupportedActions", "About to call method Action");
                                                    result = baseClassDevice.Action(ActionString, TEST_PARAMETERS);
                                                    LogOK("SupportedActions", string.Format("Switch simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogError("SupportedActions", string.Format("Exception calling switch simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                        }

                                        break;
                                    }
                            }
                        }
                        else
                        {
                            LogError("SupportedActions", "Actions must be strings. The type of action " + i + " " + Action.ToString() + " is: " + Action.GetType().Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (p_DeviceType == DeviceType.Switch & ReferenceEquals(ex.GetType(), typeof(MissingMemberException)))
                {
                    LogOK("SupportedActions", "Switch V1 Driver does not have SupportedActions");
                }
                else
                {
                    HandleException("SupportedActions", MemberType.Property, Required.Optional, ex, "");
                    LogError("SupportedActions", ex.Message);
                }
            }

            if (cancellationToken.IsCancellationRequested)
                return;
            LogNewLine();
        }

        public virtual void CheckCommonMethods()
        {
            LogError("CheckCommonMethods", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual bool HasCanProperties
        {
            get
            {
                return l_HasCanProperties;
            }

            set
            {
                l_HasCanProperties = value;
            }
        }

        public virtual bool HasProperties
        {
            get
            {
                return l_HasProperties;
            }

            set
            {
                l_HasProperties = value;
            }
        }

        public virtual bool HasMethods
        {
            get
            {
                return l_HasMethods;
            }

            set
            {
                l_HasMethods = value;
            }
        }

        public virtual bool HasPreConnectCheck
        {
            get
            {
                return l_HasPreConnectCheck;
            }

            set
            {
                l_HasPreConnectCheck = value;
            }
        }

        public virtual bool HasPreRunCheck
        {
            get
            {
                return l_HasPreRunCheck;
            }

            set
            {
                l_HasPreRunCheck = value;
            }
        }

        public virtual bool HasPostRunCheck
        {
            get
            {
                return l_HasPostRunCheck;
            }

            set
            {
                l_HasPostRunCheck = value;
            }
        }

        public virtual bool HasPerformanceCheck
        {
            get
            {
                return l_HasPerformanceCheck;
            }

            set
            {
                l_HasPerformanceCheck = value;
            }
        }

        /// <summary>
        /// Get error codes.
        /// </summary>
        /// <param name="p_ProgID">The ProgID.</param>
        internal void CheckInitialise()
        {
            Status(StatusType.staTest, ""); // Clear status messages
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");

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

            switch (settings.DeviceTechnology)
            {
                case DeviceTechnology.Alpaca:
                    LogTestOnly($"Alpaca device: {settings.AlpacaDevice.AscomDeviceName} ({settings.AlpacaDevice.IpAddress}:{settings.AlpacaDevice.IpPort} {settings.AlpacaDevice.AscomDeviceType}/{settings.AlpacaDevice.AlpacaDeviceNumber})");
                    if (!settings.StrictCasing) LogIssue("ConformanceCheck", "Alpaca strict casing has been disabled, this in only supported for testing devices.");

                    break;

                case DeviceTechnology.COM:
                    LogTestOnly($"COM Driver ProgID: {settings.ComDevice.ProgId}"); break;

                default:
                    throw new InvalidValueException($"CheckInitialise - Unknown technology type: {settings.DeviceTechnology}");
            }

            LogNewLine(); // Blank line
        }

        public virtual void CreateDevice()
        {
            LogError("CreateDevice", "DeviceTester base Class warning message");
        }

        public virtual void PreConnectChecks()
        {
            LogError("PreConnectChecks", "DeviceTester base Class warning message");
        }

        public virtual bool Connected
        {
            get
            {
                return l_Connected;
            }

            set
            {
                l_Connected = value;
            }
        }

        public virtual void ReadCanProperties()
        {
            LogError("ReadCanProperties", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PreRunCheck()
        {
            LogError("PreSafetyCheck", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckProperties()
        {
            LogError("CheckProperties", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckMethods()
        {
            LogError("CheckMethods", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckPerformance()
        {
            LogError("CheckPerformance", "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PostRunCheck()
        {
            LogError("PostSafetyCheck", "DeviceTester base Class warning message, you should not see this message!");
        }

        #endregion

        #region Common methods for all device tester classes

        /// <summary>
        /// Update the status display
        /// </summary>
        /// <param name="Status">Type of message to set, Test, Action or Status</param>
        /// <param name="Message">Message text</param>
        /// <remarks></remarks>
        internal void Status(StatusType Status, string message)
        {
            switch (Status)
            {
                case StatusType.staTest:
                    {
                        test = message;
                        break;
                    }

                case StatusType.staAction:
                    {
                        action = message;
                        break;
                    }

                case StatusType.staStatus:
                    {
                        status = message;
                        break;
                    }
            }

            TL?.OnStatusChanged($"{test} {action} {status}");
        }

        /// <summary>
        ///Set the test, action and status in one call
        ///</summary>
        ///<param name="newTest">Name of the test being conducted</param>
        ///<param name="newAction">Specific action within the test</param>
        ///<param name="newStatus">Status of the action</param>
        ///<remarks></remarks>
        public void SetStatus(string newTest, string newAction, string newStatus)
        {
            test = newTest;
            action = newAction;
            status = newStatus;

            TL?.OnStatusChanged($"{test} {action} {status}");
        }

        /// <summary>
        ///Clear all status fields
        ///</summary>
        ///<remarks></remarks>
        public void ClearStatus()
        {
            test = "";
            action = "";
            status = "";

            TL?.OnStatusChanged($"{test} {action} {status}");
        }

        /// <summary>
        /// Delays execution for the given time period in milliseconds
        /// </summary>
        /// <param name="p_Duration">Delay duration in milliseconds</param>
        /// <remarks></remarks>
        internal void WaitFor(int p_Duration)
        {
            DateTime l_StartTime;
            int WaitDuration;
            WaitDuration = (int)Math.Round(p_Duration / 100d);
            if (WaitDuration > SLEEP_TIME)
                WaitDuration = SLEEP_TIME;
            if (WaitDuration < 1)
                WaitDuration = 1;
            // Wait for p_Duration milliseconds
            l_StartTime = DateTime.Now; // Save start time
            do
            {
                Thread.Sleep(WaitDuration);
                //Application.DoEvents();
            }
            while ((DateTime.Now.Subtract(l_StartTime).TotalMilliseconds <= p_Duration) & !cancellationToken.IsCancellationRequested);
        }

        internal void LogNewLine()
        {
            LogTestOnly("");
        }

        internal void LogTestOnly(string p_Test)
        {
            TL?.LogMessage(p_Test, MessageLevel.TestOnly, "");
        }

        internal void LogTestAndMessage(string p_Test, string p_Msg)
        {
            TL?.LogMessage(p_Test, MessageLevel.TestAndMessage, p_Msg);
        }

        internal void LogOK(string p_Test, string p_Msg)
        {
            TL?.LogMessage(p_Test, MessageLevel.OK, p_Msg);
        }

        internal void LogDebug(string p_Test, string p_Msg)
        {
            TL?.LogMessage(p_Test, MessageLevel.Debug, p_Msg);
        }

        internal void LogInfo(string p_Test, string p_Msg)
        {
            TL?.LogMessage(p_Test, MessageLevel.Info, p_Msg);
        }

        internal void LogIssue(string p_Test, string p_Msg)
        {
            conformResults.Issues.Add(new System.Collections.Generic.KeyValuePair<string, string>(p_Test, p_Msg));
            TL?.LogMessage(p_Test, MessageLevel.Issue, p_Msg);
        }

        internal void LogError(string p_Test, string p_Msg)
        {
            conformResults.Errors.Add(new System.Collections.Generic.KeyValuePair<string, string>(p_Test, p_Msg));
            TL?.LogMessage(p_Test, MessageLevel.Error, p_Msg);
        }

        internal void LogMsg(string testName, MessageLevel messageLevel, string message)
        {
            TL?.LogMessage(testName, messageLevel, message);
        }

        /// <summary>
        /// Test a supplied exception for whether it is a MethodNotImplemented type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotImplemented type</returns>
        /// <remarks>Different tests are applied for COM and MethodNotImplemmented exceptions</remarks>
        protected bool IsMethodNotImplementedException(Exception deviceException)
        {
            COMException COMException;
            bool IsMethodNotImplementedExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException exception) // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                {
                    COMException = exception;
                    if (COMException.ErrorCode == g_ExNotImplemented | COMException.ErrorCode == ErrorCodes.NotImplemented) // This is a not implemented exception
                    {
                        IsMethodNotImplementedExceptionRet = true;
                    }
                }

                if (deviceException is MethodNotImplementedException)
                {
                    IsMethodNotImplementedExceptionRet = true;
                }
            }
            catch (Exception ex)
            {
                LogIssue("IsMethodNotImplementedException", "Unexpected exception: " + ex.ToString());
            }

            return IsMethodNotImplementedExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is a NotImplemented type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotImplemented type</returns>
        /// <remarks>Different tests are applied for COM and .NET exceptions</remarks>
        protected bool IsNotImplementedException(Exception deviceException)
        {
            COMException COMException;
            bool IsNotImplementedExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException exception) // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                {
                    COMException = exception;
                    if (COMException.ErrorCode == g_ExNotImplemented | COMException.ErrorCode == ErrorCodes.NotImplemented) // This is a not implemented exception
                    {
                        IsNotImplementedExceptionRet = true;
                    }
                }

                if (deviceException is ASCOM.NotImplementedException)
                {
                    IsNotImplementedExceptionRet = true;
                }
            }
            catch (Exception ex)
            {
                LogIssue("IsNotImplementedException", "Unexpected exception: " + ex.ToString());
            }

            return IsNotImplementedExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is a PropertyNotImplementedException type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotImplemented type</returns>
        /// <remarks>Different tests are applied for COM and PropertyNotImplemmented exceptions</remarks>
        protected bool IsPropertyNotImplementedException(Exception deviceException)
        {
            COMException COMException;
            bool IsPropertyNotImplementedExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException exception) // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                {
                    COMException = exception;
                    if (COMException.ErrorCode == g_ExNotImplemented | COMException.ErrorCode == ErrorCodes.NotImplemented) // This is a not implemented exception
                    {
                        IsPropertyNotImplementedExceptionRet = true;
                    }
                }

                if (deviceException is PropertyNotImplementedException)
                {
                    IsPropertyNotImplementedExceptionRet = true;
                }
            }
            catch (Exception ex)
            {
                LogIssue("IsPropertyNotImplementedException", "Unexpected exception: " + ex.ToString());
            }

            return IsPropertyNotImplementedExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is an InvalidValueException type
        /// </summary>
        /// <param name="MemberName"></param>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a InvalidValue type</returns>
        /// <remarks>Different tests are applied for COM and InvalidValueException exceptions</remarks>
        protected bool IsInvalidValueException(string MemberName, Exception deviceException)
        {
            COMException COMException;
            DriverException DriverException;
            bool IsInvalidValueExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException exception) // This is a COM exception so test whether the error code indicates that it is an invalid value exception
                {
                    COMException = exception;
                    if (COMException.ErrorCode == ErrorCodes.InvalidValue | COMException.ErrorCode == g_ExInvalidValue1 | COMException.ErrorCode == g_ExInvalidValue2 | COMException.ErrorCode == g_ExInvalidValue3 | COMException.ErrorCode == g_ExInvalidValue4 | COMException.ErrorCode == g_ExInvalidValue5 | COMException.ErrorCode == g_ExInvalidValue6) // This is an invalid value exception
                    {
                        IsInvalidValueExceptionRet = true;
                    }
                }

                if (deviceException is InvalidValueException)
                {
                    IsInvalidValueExceptionRet = true;
                }
                else if (deviceException is DriverException exception1)
                {
                    DriverException = exception1;
                    if (DriverException.Number == ErrorCodes.InvalidValue) // This is an invalid value exception
                    {
                        LogIssue(MemberName, "Received ASCOM.DriverException(0x" + ErrorCodes.InvalidValue.ToString("X8") + "), please use ASCOM.InvalidValueException to report invalid values");
                    }
                }

                if (deviceException is System.InvalidOperationException)
                {
                    LogIssue(MemberName, "Received System.InvalidOperationException rather than ASCOM.InvalidValueException");
                }
            }
            catch (Exception ex)
            {
                LogIssue("IsInvalidValueException", "Unexpected exception: " + ex.ToString());
            }

            return IsInvalidValueExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is an InvalidValueException type
        /// </summary>
        /// <param name="MemberName"></param>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a InvalidValue type</returns>
        /// <remarks>Different tests are applied for COM and InvalidValueException exceptions</remarks>
        protected bool IsInvalidOperationException(string MemberName, Exception deviceException)
        {
            COMException COMException;
            DriverException DriverException;
            bool IsInvalidOperationExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException exception) // This is a COM exception so test whether the error code indicates that it is an invalid operation exception
                {
                    COMException = exception;
                    if (COMException.ErrorCode == ErrorCodes.InvalidOperationException) // This is an invalid operation exception
                    {
                        IsInvalidOperationExceptionRet = true;
                    }
                }

                if (deviceException is ASCOM.InvalidOperationException)
                {
                    IsInvalidOperationExceptionRet = true;
                }
                else if (deviceException is DriverException exception1)
                {
                    DriverException = exception1;
                    if (DriverException.Number == ErrorCodes.InvalidOperationException) // This is an invalid operation exception
                    {
                        LogIssue(MemberName, "Received ASCOM.DriverException(0x" + ErrorCodes.InvalidOperationException.ToString("X8") + "), please use ASCOM.InvalidOperationException to report invalid operations");
                    }
                }

                if (deviceException is System.InvalidOperationException)
                {
                    LogIssue(MemberName, "Received System.InvalidOperationException rather than ASCOM.InvalidOperationException");
                }
            }
            catch (Exception ex)
            {
                LogIssue("IsInvalidOperationException", "Unexpected exception: " + ex.ToString());
            }

            return IsInvalidOperationExceptionRet;
        }

        /// <summary>
        /// Test a supplied exception for whether it is a NotSetException type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotSet type</returns>
        /// <remarks>Different tests are applied for COM and ValueNotSetException exceptions</remarks>
        protected bool IsNotSetException(Exception deviceException)
        {
            COMException COMException;
            bool IsNotSetExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException exception) // This is a COM exception so test whether the error code indicates that it is a not set exception
                {
                    COMException = exception;
                    if (COMException.ErrorCode == g_ExNotSet1) // This is a not set exception
                    {
                        IsNotSetExceptionRet = true;
                    }
                }

                if (deviceException is ValueNotSetException)
                {
                    IsNotSetExceptionRet = true;
                }
            }
            catch (Exception ex)
            {
                LogIssue("IsNotSetException", "Unexpected exception: " + ex.ToString());
            }

            return IsNotSetExceptionRet;
        }

        /// <summary>
        /// Provides messaging when an exception is thrown by a member
        /// </summary>
        /// <param name="MemberName">The name of the member throwing the exception</param>
        /// <param name="TypeOfMember">Flag indicating whether the member is a property or a method</param>
        /// <param name="IsRequired">Flag indicating whether the member is optional or mandatory</param>
        /// <param name="ex">The exception received from the device</param>
        /// <param name="UserMessage">The member specific message to report</param>
        /// <remarks></remarks>
        protected void HandleException(string MemberName, MemberType TypeOfMember, Required IsRequired, Exception ex, string UserMessage)
        {

            // Handle PropertyNotImplemented exceptions from properties and MethodNotImplemented exceptions from methods
            if (IsPropertyNotImplementedException(ex) & TypeOfMember == MemberType.Property | IsMethodNotImplementedException(ex) & TypeOfMember == MemberType.Method)
            {
                switch (IsRequired)
                {
                    case Required.Mandatory:
                        {
                            LogIssue(MemberName, "This member is mandatory but threw a " + GetExceptionName(ex) + " exception, it must function per the ASCOM specification.");
                            break;
                        }

                    case Required.MustNotBeImplemented:
                        {
                            LogOK(MemberName, UserMessage + " and a " + GetExceptionName(ex) + " exception was generated as expected");
                            break;
                        }

                    case Required.MustBeImplemented:
                        {
                            LogIssue(MemberName, UserMessage + " and a " + GetExceptionName(ex) + " exception was thrown, this method must function per the ASCOM specification.");
                            break;
                        }

                    case Required.Optional:
                        {
                            LogOK(MemberName, "Optional member threw a " + GetExceptionName(ex) + " exception.");
                            break;
                        }

                    default:
                        {
                            LogError(MemberName, "CONFORM ERROR! - Received unexpected member of 'Required' enum: " + IsRequired.ToString());
                            break;
                        }
                }
            }

            // Handle wrong type of not implemented exceptions
            else if (ex is MethodNotImplementedException & TypeOfMember == MemberType.Property) // We got a MethodNotImplementedException so this is an error
            {
                LogIssue(MemberName, "Received a MethodNotImplementedException instead of a PropertyNotImplementedException");
            }
            else if (ex is PropertyNotImplementedException & TypeOfMember == MemberType.Method) // We got a PropertyNotImplementedException so this is an error
            {
                LogIssue(MemberName, "Received a PropertyNotImplementedException instead of a MethodNotImplementedException");
            }
            else if (ex is ASCOM.NotImplementedException)
            {
                LogIssue(MemberName, "Received a NotImplementedException instead of a " + ((TypeOfMember == MemberType.Property) ? "PropertyNotImplementedException" : "MethodNotImplementedException"));
            }
            else if (ex is System.NotImplementedException)
            {
                LogIssue(MemberName, "Received a System.NotImplementedException instead of an ASCOM." + ((TypeOfMember == MemberType.Property) ? "PropertyNotImplementedException" : "MethodNotImplementedException"));
            }

            // Handle all other types of error
            else
            {
                LogError(MemberName, "Unexpected " + GetExceptionName(ex) + ", " + UserMessage + ": " + ex.Message);
            }

            LogDebug(MemberName, "Exception: " + ex.ToString());
        }

        protected void HandleInvalidValueExceptionAsOK(string MemberName, MemberType TypeOfMember, Required IsRequired, Exception ex, string UserAction, string Message)
        {
            if (IsInvalidValueException(MemberName, ex))
            {
                LogOK(MemberName, Message);
            }
            else
            {
                HandleException(MemberName, TypeOfMember, IsRequired, ex, UserAction);
            }
        }

        protected void HandleInvalidValueExceptionAsInfo(string MemberName, MemberType TypeOfMember, Required IsRequired, Exception ex, string UserAction, string Message)
        {
            if (IsInvalidValueException(MemberName, ex))
            {
                LogInfo(MemberName, Message);
            }
            else
            {
                HandleException(MemberName, TypeOfMember, IsRequired, ex, UserAction);
            }
        }

        protected void HandleInvalidOperationExceptionAsOK(string MemberName, MemberType TypeOfMember, Required IsRequired, Exception ex, string UserAction, string Message)
        {
            if (IsInvalidOperationException(MemberName, ex))
            {
                LogOK(MemberName, Message);
            }
            else
            {
                HandleException(MemberName, TypeOfMember, IsRequired, ex, UserAction);
            }
        }

        /// <summary>
        /// Get an exception name (and number if a COM or Driver exception)
        /// </summary>
        /// <param name="ex">Exception whose name is required</param>
        /// <returns>String exception name</returns>
        /// <remarks></remarks>
        protected static string GetExceptionName(Exception ex)
        {
            COMException ComEx;
            DriverException DriverEx;
            string RetVal;

            // Treat ASCOM exceptions specially
            if (ex.GetType().FullName.ToUpper().Contains("ASCOM"))
            {
                if (ex.GetType().FullName.ToUpper().Contains("DRIVEREXCEPTION")) // We have a driver exception so add its number
                {
                    DriverEx = (DriverException)ex;
                    RetVal = "DriverException(0x" + DriverEx.Number.ToString("X8") + ")";
                }
                else // Otherwise just use the ASCOM exception's name
                {
                    RetVal = ex.GetType().Name;
                }
            }
            else if (ex is COMException exception) // Handle XOM exceptions with their error code
            {
                ComEx = exception;
                RetVal = "COMException(0x" + ComEx.ErrorCode.ToString("X8") + ")";
            }
            else // We got something else so report it
            {
                RetVal = ex.GetType().FullName + " exception";
            }

            return RetVal;
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

        internal void WaitForAbsolute(int p_Duration, string p_Message)
        {
            LogDebug("WaitForAbsolute", p_Duration + " " + p_Message);
            for (int i = 0, loopTo = (int)Math.Round(p_Duration / 100d); i <= loopTo; i++)
            {
                Thread.Sleep(100);
                SetStatus(p_Message, ((p_Duration / 100d - i) / 10d).ToString(), "");
            }

            ClearStatus();
        }

        #endregion

        #region Base class support Code

        private static bool IncludeMethod(MandatoryMethod p_Method, DeviceType p_DeviceType, int p_InterfaceVersion)
        {
            // This mechanic allows individual tests for particular devices to be skipped. It is no longer required because this is handled by DriverAccess
            // The code is left in place in case it is ever needed in the future

            bool RetVal = true; // Default to true as most methods will be tested , we just list the exceptions to this below

            // Matrix controlling what tests
            switch (p_DeviceType)
            {
                case DeviceType.Telescope:
                    {
                        switch (p_InterfaceVersion)
                        {
                            case 1: // Telescope interface V1 does not have Driver Version
                                {
                                    if (p_Method == MandatoryMethod.DriverVersion)
                                        RetVal = false;
                                    break;
                                }

                            default:
                                {
                                    RetVal = true; // All methods in all interface versions are mandatory
                                    break;
                                }
                        }

                        break;
                    }

                case DeviceType.Camera:
                    {
                        RetVal = true;
                        break;
                    }
            }

            return RetVal;
        }

        //private void CommandTest(CommandType p_Type, string p_Name)
        //{
        //    string l_CmdSent = "!! Unknown Value !!";
        //    string m_CommandString;
        //    bool m_CommandBool;
        //    g_Status.Test = p_Name;
        //    try
        //    {
        //        switch (p_Type)
        //        {
        //            case CommandType.tstCommandBlind:
        //                {
        //                    if (g_CmdStrings.CommandBlind is object)
        //                    {
        //                        l_CmdSent = g_CmdStrings.CommandBlind;
        //                        Device.CommandBlind(l_CmdSent, false);
        //                        LogMsgOK(p_Name, "Sent string \"" + g_CmdStrings.CommandBlind + "\"");
        //                    }
        //                    else
        //                    {
        //                        LogMsgInfo(p_Name, "Skipped CommandBlind test");
        //                    }

        //                    break;
        //                }

        //            case CommandType.tstCommandBool:
        //                {
        //                    if (g_CmdStrings.CommandBool is object)
        //                    {
        //                        l_CmdSent = g_CmdStrings.CommandBool;
        //                        m_CommandBool = Conversions.ToBoolean(Device.CommandBool(l_CmdSent, (object)false));
        //                        if (m_CommandBool == g_CmdStrings.ReturnBool)
        //                        {
        //                            LogMsgOK(p_Name, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received expected return value: " + m_CommandBool);
        //                        }
        //                        else
        //                        {
        //                            LogMsgError(p_Name, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsgInfo(p_Name, "Skipped CommandBool test");
        //                    }

        //                    break;
        //                }

        //            case CommandType.tstCommandString:
        //                {
        //                    if (g_CmdStrings.CommandString is object)
        //                    {
        //                        l_CmdSent = g_CmdStrings.CommandString;
        //                        m_CommandString = Conversions.ToString(Device.CommandString(l_CmdSent, (object)false));
        //                        if (g_CmdStrings.ReturnString is object) // Valid return string to test
        //                        {
        //                            if ((m_CommandString ?? "") == (g_CmdStrings.ReturnString ?? ""))
        //                            {
        //                                LogMsgOK(p_Name, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
        //                            }
        //                            else
        //                            {
        //                                LogMsgError(p_Name, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStrings.ReturnString + "\"");
        //                            }
        //                        }
        //                        else // Skip the return string test
        //                        {
        //                            LogMsgOK(p_Name, "Sent string \"" + g_CmdStrings.CommandString + "\" - Return string test skipped");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsgInfo(p_Name, "Skipped CommandString test");
        //                    }

        //                    break;
        //                }

        //            case CommandType.tstCommandBlindRaw:
        //                {
        //                    if (g_CmdStringsRaw.CommandBlind is object)
        //                    {
        //                        l_CmdSent = g_CmdStringsRaw.CommandBlind;
        //                        Device.CommandBlind(l_CmdSent, true);
        //                        LogMsgOK(p_Name, "Sent string \"" + g_CmdStringsRaw.CommandBlind + "\"");
        //                    }
        //                    else
        //                    {
        //                        LogMsgInfo(p_Name, "Skipped CommandBlind Raw test");
        //                    }

        //                    break;
        //                }

        //            case CommandType.tstCommandBoolRaw:
        //                {
        //                    if (g_CmdStringsRaw.CommandBool is object)
        //                    {
        //                        l_CmdSent = g_CmdStringsRaw.CommandBool;
        //                        m_CommandBool = Conversions.ToBoolean(Device.CommandBool(l_CmdSent, (object)true));
        //                        if (m_CommandBool == g_CmdStringsRaw.ReturnBool)
        //                        {
        //                            LogMsgOK(p_Name, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received expected return value: " + m_CommandBool);
        //                        }
        //                        else
        //                        {
        //                            LogMsgError(p_Name, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsgInfo(p_Name, "Skipped CommandBool Raw test");
        //                    }

        //                    break;
        //                }

        //            case CommandType.tstCommandStringRaw:
        //                {
        //                    if (g_CmdStringsRaw.CommandString is object)
        //                    {
        //                        l_CmdSent = g_CmdStringsRaw.CommandString;
        //                        m_CommandString = Conversions.ToString(Device.CommandString(l_CmdSent, (object)true));
        //                        if (g_CmdStringsRaw.ReturnString is object) // Valid return string to test
        //                        {
        //                            if ((m_CommandString ?? "") == (g_CmdStringsRaw.ReturnString ?? ""))
        //                            {
        //                                LogMsgOK(p_Name, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
        //                            }
        //                            else
        //                            {
        //                                LogMsgError(p_Name, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStringsRaw.ReturnString + "\"");
        //                            }
        //                        }
        //                        else // Skip the return string test
        //                        {
        //                            LogMsgOK(p_Name, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Return string test skipped");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsgInfo(p_Name, "Skipped CommandString Raw test");
        //                    }

        //                    break;
        //                }

        //            default:
        //                {
        //                    LogMsgError(p_Name, "Conform:CommandTest: Unknown test type " + p_Type.ToString());
        //                    break;
        //                }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        HandleException(p_Name, MemberType.Method, Required.Optional, ex, "");
        //    }

        //    g_Status.Clear();  // Clear status messages
        //}

        #endregion

        #region Private tests
        public void SpecialTests(SpecialTest p_Test)
        {
            switch (p_Test)
            {
                case SpecialTest.TelescopeSideOfPier:
                    {
                        SpecialTelescopeSideOfPier();
                        break;
                    }

                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        SpecialTelescopeDestinationSideOfPier();
                        break;
                    }

                case SpecialTest.TelescopeSideOfPierAnalysis:
                    {
                        SpecialTelescopeSideOfPierAnalysis();
                        break;
                    }

                case SpecialTest.TelescopeCommands:
                    {
                        SpecialTelescopeCommands();
                        break;
                    }

                default:
                    {
                        LogError("DeviceTesterBaseClass:SpecialTests", $"Unknown test: {p_Test}");
                        break;
                    }
            }
        }

        protected virtual void SpecialTelescopeSideOfPier()
        {
            LogError("SpecialTelescopeSideOfPier", "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeDestinationSideOfPier()
        {
            LogError("SpecialTelescopeDestinationSideOfPier", "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeSideOfPierAnalysis()
        {
            LogError("SpecialTelescopeFlipRange", "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeCommands()
        {
            LogError("SpecialTelescopeCommands", "DeviceTester base Class warning message, you should not see this message!");
        }
        #endregion

    }
}
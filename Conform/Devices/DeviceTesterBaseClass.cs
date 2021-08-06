// Base class from which particular device testers are derived
// Put all common elements in here
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ASCOM;
using ConformU;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using static Conform.GlobalVarsAndCode;
using static ConformU.ConformConstants;

namespace Conform
{

    /// <summary>
    /// Base class for device tester classes. Contains common code and placeholders for the 
    /// methods that must be implemented in the device tester class
    /// </summary>
    internal class DeviceTesterBaseClass
    {
        // Implements IDisposable

        #region Variables and Constants
        private bool l_Connected, l_HasProperties, l_HasCanProperties, l_HasMethods, l_HasPreRunCheck, l_HasPostRunCheck, l_HasPerformanceCheck;
        private bool l_HasPreConnectCheck;
        private dynamic Device; // IAscomDriverV1

        private string test, action, status;

        private ConformU.ConformanceTestManager parentClass;
        private ConformU.ConformLogger TL;

        private Settings settings;
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
        public DeviceTesterBaseClass(bool HasCanProperties, bool HasProperties, bool HasMethods, bool HasPreRunCheck, bool HasPreConnectCheck, bool HasPerformanceCheck, bool HasPostRunCheck, ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger) : base()
        {
            l_HasPreConnectCheck = HasPreConnectCheck;
            l_Connected = false;
            l_HasPreRunCheck = HasPreRunCheck;
            l_HasCanProperties = HasCanProperties;
            l_HasProperties = HasProperties;
            l_HasMethods = HasMethods;
            l_HasPostRunCheck = HasPostRunCheck;
            l_HasPerformanceCheck = HasPerformanceCheck;
            parentClass = parent;
            TL = logger;
            settings = conformConfiguration.Settings;
        }

        private bool disposedValue = false;        // To detect redundant calls
                                                   // IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                }
            }

            disposedValue = true;
        }

        // This code added by Visual Basic to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code.  Put clean-up code in Dispose(ByVal disposing As Boolean) above.
            Dispose(true);
        }
        #endregion

        #region Code

        public void CheckCommonMethods(object p_DeviceObject, DeviceType p_DeviceType)
        {
            string m_DriverVersion, m_DriverInfo, m_Description, m_Name; // , m_LastResult As String
            IList SA;

            // Dim m_Configuration, SC() As String
            bool m_Connected;
            LogMsg("Common Driver Methods", MessageLevel.Always, "");
            Device = p_DeviceObject; // CType(DeviceObject, IAscomDriverV1)

            // InterfaceVersion - Required
            try
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("InterfaceVersion", MessageLevel.Comment, "About to get property InterfaceVersion");
                g_InterfaceVersion = Conversions.ToInteger(Device.InterfaceVersion);
                switch (g_InterfaceVersion)
                {
                    case var @case when @case < 1:
                        {
                            LogMsg("InterfaceVersion", MessageLevel.Issue, "InterfaceVersion must be 1 or greater but driver returned: " + g_InterfaceVersion.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg("InterfaceVersion", MessageLevel.OK, g_InterfaceVersion.ToString());
                            break;
                        }
                }

            }
            catch (Exception ex)
            {
                HandleException("InterfaceVersion", MemberType.Property, Required.Mandatory, ex, "");
            }

            if (TestStop())
                return;

            // Connected - Required
            if (IncludeMethod(MandatoryMethod.Connected, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("Connected", MessageLevel.Comment, "About to get property Connected");
                    m_Connected = Device.Connected;
                    LogMsg("Connected", MessageLevel.OK, m_Connected.ToString());
                }
                catch (COMException ex)
                {
                    LogMsg("Connected", MessageLevel.Error, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                }
                catch (Exception ex)
                {
                    LogMsg("Connected", MessageLevel.Error, ex.Message);
                }

                if (TestStop())
                    return;
            }

            // Description - Required
            if (IncludeMethod(MandatoryMethod.Description, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("Description", MessageLevel.Comment, "About to get property Description");
                    m_Description = Conversions.ToString(Device.Description);
                    switch (m_Description ?? "")
                    {
                        case var case1 when case1 == "":
                            {
                                LogMsg("Description", MessageLevel.Info, "No description string");
                                break;
                            }

                        default:
                            {
                                if (m_Description.Length > 68 & p_DeviceType == DeviceType.Camera)
                                {
                                    LogMsg("Description", MessageLevel.Error, "Maximum number of characters is 68 for compatibility with FITS headers, found: " + m_Description.Length + " characters: " + m_Description);
                                }
                                else
                                {
                                    LogMsg("Description", MessageLevel.OK, m_Description.ToString());
                                }

                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Description", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (TestStop())
                    return;
            }

            // DriverInfo - Required
            if (IncludeMethod(MandatoryMethod.DriverInfo, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("DriverInfo", MessageLevel.Comment, "About to get property DriverInfo");
                    m_DriverInfo = Conversions.ToString(Device.DriverInfo);
                    switch (m_DriverInfo ?? "")
                    {
                        case var case2 when case2 == "":
                            {
                                LogMsg("DriverInfo", MessageLevel.Info, "No DriverInfo string");
                                break;
                            }

                        default:
                            {
                                LogMsg("DriverInfo", MessageLevel.OK, m_DriverInfo.ToString());
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("DriverInfo", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (TestStop())
                    return;
            }

            // DriverVersion - Required
            if (IncludeMethod(MandatoryMethod.DriverVersion, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("DriverVersion", MessageLevel.Comment, "About to get property DriverVersion");
                    m_DriverVersion = Conversions.ToString(Device.DriverVersion);
                    switch (m_DriverVersion ?? "")
                    {
                        case var case3 when case3 == "":
                            {
                                LogMsg("DriverVersion", MessageLevel.Info, "No DriverVersion string");
                                break;
                            }

                        default:
                            {
                                LogMsg("DriverVersion", MessageLevel.OK, m_DriverVersion.ToString());
                                break;
                            }
                    }
                }
                catch (COMException ex)
                {
                    LogMsg("DriverVersion", MessageLevel.Error, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                }
                catch (Exception ex)
                {
                    LogMsg("DriverVersion", MessageLevel.Error, ex.Message);
                }

                if (TestStop())
                    return;
            }
            else
            {
                LogMsg("DriverVersion", MessageLevel.Info, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // Name - Required
            if (IncludeMethod(MandatoryMethod.Name, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (settings.DisplayMethodCalls)
                        LogMsg("Name", MessageLevel.Comment, "About to get property Name");
                    m_Name = Conversions.ToString(Device.Name);
                    switch (m_Name ?? "")
                    {
                        case var case4 when case4 == "":
                            {
                                LogMsg("Name", MessageLevel.Info, "Name is empty");
                                break;
                            }

                        default:
                            {
                                LogMsg("Name", MessageLevel.OK, m_Name);
                                break;
                            }
                    }
                }
                catch (Exception ex)
                {
                    HandleException("Name", MemberType.Property, Required.Mandatory, ex, "");
                }

                if (TestStop())
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
            //        LogMsg("CommandString", MessageLevel.Info, "Conform cannot test the CommandString method");
            //        LogMsg("CommandBlind", MessageLevel.Info, "Conform cannot test the CommandBlind method");
            //        LogMsg("CommandBool", MessageLevel.Info, "Conform cannot test the CommandBool method");
            //    }
            //    else
            //    {
            //        LogMsg(TELTEST_COMMANDXXX, MessageLevel.Info, "Tests skipped");
            //    }
            //}

            // Action - optional but cannot be tested
            LogMsg("Action", MessageLevel.Info, "Conform cannot test the Action method");

            // Supported actions - Optional but Required through DriverAccess
            try
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("SupportedActions", MessageLevel.Comment, "About to call method SupportedActions");
                SA = (IList)Device.SupportedActions;
                if (SA.Count == 0)
                {
                    LogMsg("SupportedActions", MessageLevel.OK, "Driver returned an empty action list");
                }
                else
                {
                    var i = default(int);
                    foreach (object Action in SA)
                    {
                        i += 1;
                        if (Action.GetType().Name == "String")
                        {
                            string ActionString = Conversions.ToString(Action);
                            string result;
                            const string TEST_PARAMETERS = "Conform test parameters";
                            switch (ActionString ?? "")
                            {
                                case var case5 when case5 == "":
                                    {
                                        LogMsg("SupportedActions", MessageLevel.Error, "Supported action " + i + " Is an empty string"); // List the action that was found
                                        break;
                                    }

                                default:
                                    {
                                        LogMsg("SupportedActions", MessageLevel.OK, "Found action: " + ActionString);

                                        // Carry out the following Action tests only when we are testing the Observing Conditions Hub and it is configured to use the Switch and OC simulators
                                        if (p_DeviceType == DeviceType.ObservingConditions & g_ObservingConditionsProgID.ToUpperInvariant() == "ASCOM.OCH.OBSERVINGCONDITIONS")
                                        {
                                            if (ActionString.ToUpperInvariant().StartsWith("//OCSIMULATOR:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.Comment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.OK, string.Format("OC simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.Error, string.Format("Exception calling OCH simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//ASCOM.SIMULATOR.OBSERVINGCONDITIONS:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.Comment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.OK, string.Format("OC simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.Error, string.Format("Exception calling OCH simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//SWITCHSIMULATOR:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.Comment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.OK, string.Format("Switch simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.Error, string.Format("Exception calling switch simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//ASCOM.SIMULATOR.SWITCH:"))
                                            {
                                                try
                                                {
                                                    if (settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.Comment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.OK, string.Format("Switch simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.Error, string.Format("Exception calling switch simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                        }

                                        break;
                                    }
                            }
                        }
                        else
                        {
                            LogMsg("SupportedActions", MessageLevel.Error, "Actions must be strings. The type of action " + i + " " + Action.ToString() + " is: " + Action.GetType().Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (p_DeviceType == DeviceType.Switch & ReferenceEquals(ex.GetType(), typeof(MissingMemberException)))
                {
                    LogMsg("SupportedActions", MessageLevel.OK, "Switch V1 Driver does not have SupportedActions");
                }
                else
                {
                    HandleException("SupportedActions", MemberType.Property, Required.Optional, ex, "");
                    LogMsg("SupportedActions", MessageLevel.Error, ex.Message);
                }
            }

            if (TestStop())
                return;
            LogMsg("", MessageLevel.Always, "");
        }

        public virtual void CheckCommonMethods()
        {
            LogMsg("CheckCommonMethods", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
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

        public virtual void CheckInitialise()
        {
            LogMsg("ConformanceCheckInitialise", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        /// <summary>
        /// Get error codes.
        /// </summary>
        /// <param name="p_ProgID">The ProgID.</param>
        protected void CheckInitialise(string p_ProgID)
        {
            Status(StatusType.staTest, ""); // Clear status messages
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
            g_Stop = true; // Initialise stop flag to stop
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var fileInfo = new System.IO.FileInfo(assembly.Location);
            var lastModified = fileInfo.LastWriteTime;
            LogMsg("", MessageLevel.Always, ""); // Blank line
            LogMsg("ConformanceCheck", MessageLevel.Always, "ASCOM Device Conformance Checker Version " + this.GetType().Assembly.GetName().Version.ToString() + ", Build time: " + lastModified.ToString());
            //LogMsg("ConformanceCheck", MessageLevel.Always, "Running on: " + Prof.GetProfile("Platform", "Platform Name", "Unknown") + " " + Prof.GetProfile("Platform", "Platform Version", "Unknown"));
            LogMsg("", MessageLevel.Always, ""); // Blank line
            LogMsg("ConformanceCheck", MessageLevel.Always, DRIVER_PROGID + p_ProgID);
            LogMsg("", MessageLevel.Always, ""); // Blank line
            LogMsg("Error handling", MessageLevel.Always, "");
            LogMsg("Error", MessageLevel.Always, "number for \"Not Implemented\" is: " + Conversion.Hex(g_ExNotImplemented));
            LogMsg("Error", MessageLevel.Always, "number for \"Invalid Value 1\" is: " + Conversion.Hex(g_ExInvalidValue1));
            if (g_ExInvalidValue2 != 0 & g_ExInvalidValue2 != g_ExInvalidValue1)
                LogMsg("Error", MessageLevel.Always, "number for \"Invalid Value 2\" is: " + Conversion.Hex(g_ExInvalidValue2));
            if (g_ExInvalidValue3 != 0 & g_ExInvalidValue3 != g_ExInvalidValue2)
                LogMsg("Error", MessageLevel.Always, "number for \"Invalid Value 3\" is: " + Conversion.Hex(g_ExInvalidValue3));
            if (g_ExInvalidValue4 != 0 & g_ExInvalidValue4 != g_ExInvalidValue3)
                LogMsg("Error", MessageLevel.Always, "number for \"Invalid Value 4\" is: " + Conversion.Hex(g_ExInvalidValue4));
            if (g_ExInvalidValue5 != 0 & g_ExInvalidValue5 != g_ExInvalidValue4)
                LogMsg("Error", MessageLevel.Always, "number for \"Invalid Value 5\" is: " + Conversion.Hex(g_ExInvalidValue5));
            if (g_ExInvalidValue6 != 0 & g_ExInvalidValue6 != g_ExInvalidValue5)
                LogMsg("Error", MessageLevel.Always, "number for \"Invalid Value 6\" is: " + Conversion.Hex(g_ExInvalidValue6));
            LogMsg("Error", MessageLevel.Always, "number for \"Value Not Set 1\" is: " + Conversion.Hex(ErrorCodes.ValueNotSet));
            LogMsg("Error", MessageLevel.Always, "number for \"Value Not Set 2\" is: " + Conversion.Hex(g_ExNotSet1));
            if (g_ExNotSet2 != 0 & g_ExNotSet2 != g_ExNotSet1)
                LogMsg("Error", MessageLevel.Always, "number for \"Value Not Set 3\" is: " + Conversion.Hex(g_ExNotSet2));
            if (settings.InterpretErrorMessages)
            {
                LogMsg("Error", MessageLevel.Always, "messages will be interpreted to infer state.");
            }
            else
            {
                LogMsg("Error", MessageLevel.Always, "messages will not be interpreted to infer state.");
            }

            LogMsg("", MessageLevel.Always, "");
        }

        public virtual void CheckAccessibility()
        {
            LogMsg("ConformanceCheckAccessibility", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected void CheckAccessibility(string p_ProgId, DeviceType p_DeviceType)
        {
            dynamic l_DeviceObject;
            Type l_Type;
            var l_TryCount = default(int);
            string l_ErrMsg = "";
            LogMsg("Driver Access Checks", MessageLevel.OK, "");

            // Try late binding as an object
            l_DeviceObject = null;
            do
            {
                l_TryCount += 1;
                try
                {
#if DEBUG
                    if (settings.DisplayMethodCalls) LogMsg("CreateObject", MessageLevel.Comment, "About to create instance using CreateObject");
                    l_Type = Type.GetTypeFromProgID(p_ProgId);
                    l_DeviceObject = Activator.CreateInstance(l_Type);
                    LogMsg("AccessChecks", MessageLevel.Debug, "Successfully created driver using CreateObject");
#else

                    l_Type = Type.GetTypeFromProgID(p_ProgId);
                    if (settings.DisplayMethodCalls) LogMsg("AccessChecks", MessageLevel.Comment, "About to create instance using Activator.CreateInstance");
                    l_DeviceObject = Activator.CreateInstance(l_Type);
                    LogMsg("AccessChecks", MessageLevel.Debug, "Successfully created driver using Activator.CreateInstance");

#endif
                    WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver initialisation");
                    LogMsg("AccessChecks", MessageLevel.OK, "Successfully created driver using late binding");
                    try
                    {
                        switch (p_DeviceType)
                        {
                            case DeviceType.Focuser: // Focuser uses link to connect
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.Comment, "About to set Link property true");
                                    l_DeviceObject.Link = true;
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.Comment, "About to set Link property false");
                                    l_DeviceObject.Link = false; // Everything else uses connect!
                                    break;
                                }

                            default:
                                {
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.Comment, "About to set Connected property true");
                                    l_DeviceObject.Connected = true;
                                    if (settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.Comment, "About to set Connected property false");
                                    try
                                    {
                                        l_DeviceObject.Connected = false;
                                    }
                                    catch (Exception ex)
                                    {
                                        LogMsg("AccessChecks", MessageLevel.Error, "Error disconnecting from late bound driver: " + ex.Message);
                                        LogMsg("AccessChecks", MessageLevel.Debug, "Exception: " + ex.ToString());
                                    }

                                    break;
                                }
                        }

                        LogMsg("AccessChecks", MessageLevel.OK, "Successfully connected using late binding");
                        try
                        {
                            if (l_DeviceObject.GetType().IsCOMObject)
                            {
                                LogMsg("AccessChecks", MessageLevel.Info, "The driver is a COM object");
                            }
                            else
                            {
                                LogMsg("AccessChecks", MessageLevel.Info, "The driver is a .NET object");
                                LogMsg("AccessChecks", MessageLevel.Info, "The AssemblyQualifiedName is: " + Strings.Left(l_DeviceObject.GetType().AssemblyQualifiedName.ToString(), 76));
                            }

                            foreach (var currentL_Type in l_DeviceObject.GetType().GetInterfaces())
                            {
                                l_Type = currentL_Type;
                                LogMsg("AccessChecks", MessageLevel.Info, "The driver implements interface: " + l_Type.FullName);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMsg("AccessChecks", MessageLevel.Error, "Error reading driver characteristics: " + ex.Message);
                            LogMsg("", MessageLevel.Always, "");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMsg("AccessChecks", MessageLevel.Error, "Error connecting to driver using late binding: " + ex.ToString());
                        LogMsg("", MessageLevel.Always, "");
                    }
                }
                catch (Exception ex)
                {
                    l_ErrMsg = ex.ToString();
                    LogMsg("DeviceTesterBaseClass", MessageLevel.Debug, l_ErrMsg);
                }

                if (l_DeviceObject is null)
                    WaitFor(200);
            }
            while (!(l_TryCount == 3 | l_DeviceObject is object));
            if (l_DeviceObject is null)
            {
                LogMsg("AccessChecks", MessageLevel.Error, "Error creating driver object using late binding: " + l_ErrMsg);
                LogMsg("", MessageLevel.Always, "");
            }

            // Clean up
            try
            {
                DisposeAndReleaseObject("AccessChecks", l_DeviceObject);
            }
            catch (Exception ex)
            {
                LogMsg("AccessChecks", MessageLevel.Debug, "Error releasing driver object using ReleaseCOMObject: " + ex.ToString());
            }

            l_DeviceObject = null;
            LogMsg("AccessChecks", MessageLevel.Debug, "Collecting garbage");
            GC.Collect();
            LogMsg("AccessChecks", MessageLevel.Debug, "Collecting garbage complete");
            GC.WaitForPendingFinalizers();
            LogMsg("AccessChecks", MessageLevel.Debug, "Finished waiting for pending finalisers");
            WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system"); // Wait to allow device to complete destruction
        }

        public virtual void CreateDevice()
        {
            LogMsg("CreateDevice", MessageLevel.Error, "DeviceTester base Class warning message");
        }

        public virtual void PreConnectChecks()
        {
            LogMsg("PreConnectChecks", MessageLevel.Error, "DeviceTester base Class warning message");
        }

        public virtual bool Connected
        {
            get
            {
                bool ConnectedRet = default;
                ConnectedRet = l_Connected;
                return ConnectedRet;
            }

            set
            {
                l_Connected = value;
            }
        }

        public virtual void ReadCanProperties()
        {
            LogMsg("ReadCanProperties", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PreRunCheck()
        {
            LogMsg("PreSafetyCheck", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckProperties()
        {
            LogMsg("CheckProperties", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckMethods()
        {
            LogMsg("CheckMethods", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckPerformance()
        {
            LogMsg("CheckPerformance", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PostRunCheck()
        {
            LogMsg("PostSafetyCheck", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

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
                        //MessageBox.Show("DeviceTesterBaseClass:SpecialTests - Unknown test: " + p_Test.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }
        }

        protected virtual void SpecialTelescopeSideOfPier()
        {
            LogMsg("SpecialTelescopeSideOfPier", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeDestinationSideOfPier()
        {
            LogMsg("SpecialTelescopeDestinationSideOfPier", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeSideOfPierAnalysis()
        {
            LogMsg("SpecialTelescopeFlipRange", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeCommands()
        {
            LogMsg("SpecialTelescopeCommands", MessageLevel.Error, "DeviceTester base Class warning message, you should not see this message!");
        }

        #endregion

        #region Base class support Code

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

            parentClass.OnStatusChanged($"{test} - {action} - {status}");
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

            parentClass.OnStatusChanged($"{test} - {action} - {status}");
        }

        private bool IncludeMethod(MandatoryMethod p_Method, DeviceType p_DeviceType, int p_InterfaceVersion)
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
        //                        LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStrings.CommandBlind + "\"");
        //                    }
        //                    else
        //                    {
        //                        LogMsg(p_Name, MessageLevel.Info, "Skipped CommandBlind test");
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
        //                            LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received expected return value: " + m_CommandBool);
        //                        }
        //                        else
        //                        {
        //                            LogMsg(p_Name, MessageLevel.Error, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsg(p_Name, MessageLevel.Info, "Skipped CommandBool test");
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
        //                                LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
        //                            }
        //                            else
        //                            {
        //                                LogMsg(p_Name, MessageLevel.Error, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStrings.ReturnString + "\"");
        //                            }
        //                        }
        //                        else // Skip the return string test
        //                        {
        //                            LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStrings.CommandString + "\" - Return string test skipped");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsg(p_Name, MessageLevel.Info, "Skipped CommandString test");
        //                    }

        //                    break;
        //                }

        //            case CommandType.tstCommandBlindRaw:
        //                {
        //                    if (g_CmdStringsRaw.CommandBlind is object)
        //                    {
        //                        l_CmdSent = g_CmdStringsRaw.CommandBlind;
        //                        Device.CommandBlind(l_CmdSent, true);
        //                        LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStringsRaw.CommandBlind + "\"");
        //                    }
        //                    else
        //                    {
        //                        LogMsg(p_Name, MessageLevel.Info, "Skipped CommandBlind Raw test");
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
        //                            LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received expected return value: " + m_CommandBool);
        //                        }
        //                        else
        //                        {
        //                            LogMsg(p_Name, MessageLevel.Error, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsg(p_Name, MessageLevel.Info, "Skipped CommandBool Raw test");
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
        //                                LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
        //                            }
        //                            else
        //                            {
        //                                LogMsg(p_Name, MessageLevel.Error, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStringsRaw.ReturnString + "\"");
        //                            }
        //                        }
        //                        else // Skip the return string test
        //                        {
        //                            LogMsg(p_Name, MessageLevel.OK, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Return string test skipped");
        //                        }
        //                    }
        //                    else
        //                    {
        //                        LogMsg(p_Name, MessageLevel.Info, "Skipped CommandString Raw test");
        //                    }

        //                    break;
        //                }

        //            default:
        //                {
        //                    LogMsg(p_Name, MessageLevel.Error, "Conform:CommandTest: Unknown test type " + p_Type.ToString());
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

        #region Common methods for all device tester classes

        internal void LogMsgOK(string p_Test, string p_Msg)
        {
            LogMsg(p_Test, MessageLevel.OK, p_Msg);
        }

        internal void LogMsgDebug(string p_Test, string p_Msg)
        {
            LogMsg(p_Test, MessageLevel.Debug, p_Msg);
        }

        internal void LogMsgInfo(string p_Test, string p_Msg)
        {
            LogMsg(p_Test, MessageLevel.Info, p_Msg);
        }

        internal void LogMsgIssue(string p_Test, string p_Msg)
        {
            LogMsg(p_Test, MessageLevel.Issue, p_Msg);
        }

        internal void LogMsgError(string p_Test, string p_Msg)
        {
            LogMsg(p_Test, MessageLevel.Error, p_Msg);
        }

        internal void LogMsgWarning(string p_Test, string p_Msg)
        {
            LogMsg(p_Test, MessageLevel.Warning, p_Msg);
        }

        internal void LogMsg(string p_Test, MessageLevel p_MsgLevel, string p_Msg)
        {
            const int c_Spacing = 52;
            const int TEST_NAME_WIDTH = 34;
            string l_MsgLevelFormatted, l_Msg, l_TestFormatted, l_MsgFormatted;
            int i, j;
            l_Msg = "";
            try
            {
                if (p_MsgLevel >= g_LogLevel)
                {
                    l_TestFormatted = p_Test.PadRight(TEST_NAME_WIDTH);
                    l_TestFormatted = Strings.Left(l_TestFormatted, TEST_NAME_WIDTH);
                    l_MsgLevelFormatted = "        ";
                    i = 1;
                    l_MsgFormatted = p_Msg;

                    // Remove CRLF from the message text
                    j = Strings.InStr(i, l_MsgFormatted, Microsoft.VisualBasic.Constants.vbCrLf);
                    while (j > 0)
                    {
                        l_MsgFormatted = Strings.Left(l_MsgFormatted, j + 1) + Strings.StrDup(c_Spacing, " ") + Strings.Mid(l_MsgFormatted, j + 2);
                        i = j + c_Spacing + 2;
                        j = Strings.InStr(i, l_MsgFormatted, Microsoft.VisualBasic.Constants.vbCrLf);
                    }

                    switch (p_MsgLevel)
                    {
                        case MessageLevel.None:
                            {
                                l_MsgLevelFormatted = "        ";
                                break;
                            }

                        case MessageLevel.Debug:
                            {
                                l_MsgLevelFormatted = "DEBUG   ";
                                break;
                            }

                        case MessageLevel.Comment:
                            {
                                l_MsgLevelFormatted = "        ";
                                break;
                            }

                        case MessageLevel.Info:
                            {
                                l_MsgLevelFormatted = "INFO    ";
                                break;
                            }

                        case MessageLevel.OK:
                            {
                                l_MsgLevelFormatted = "OK      ";
                                break;
                            }

                        case MessageLevel.Warning:
                            {
                                l_MsgLevelFormatted = "WARNING ";
                                g_CountWarning += 1;
                                break;
                            }

                        case MessageLevel.Issue:
                            {
                                l_MsgLevelFormatted = "ISSUE   ";
                                g_CountIssue += 1;
                                break;
                            }

                        case MessageLevel.Error:
                            {
                                l_MsgLevelFormatted = "ERROR   ";
                                g_CountError += 1;
                                break;
                            }

                        case MessageLevel.Always:
                            {
                                l_MsgLevelFormatted = "        ";
                                break;
                            }

                        default:
                            {
                                //MessageBox.Show("Conform:LogMsg - UNEXPECTED LogMessageType: " + p_MsgLevel.ToString() + " " + p_Msg);
                                break;
                            }
                    }

                    //if (My.MyProject.Forms.FrmConformMain.txtStatus.TextLength > 50000) // Limit lines to a sensible number
                    //{
                    //    My.MyProject.Forms.FrmConformMain.txtStatus.Text = Strings.Right(My.MyProject.Forms.FrmConformMain.txtStatus.Text, 28000);
                    //}

                    switch (p_MsgLevel)
                    {
                        case MessageLevel.Always:
                            {
                                l_Msg = p_Test + " " + p_Msg;
                                break;
                            }

                        default:
                            {
                                l_Msg = Strings.Format(DateAndTime.Now, "HH:mm:ss.fff") + " " + l_TestFormatted + l_MsgLevelFormatted + " " + l_MsgFormatted;
                                break;
                            }
                    }
                    parentClass.OnLogMessageChanged("LogMessage", $"{l_Msg}");
                    TL.LogMessage(p_Test, l_MsgLevelFormatted + " " + l_MsgFormatted);
                }
            }
            catch (Exception)
            {
                // MsgBox(Len(l_Msg) & " *" & l_Msg & "* " & ex.ToString, MsgBoxStyle.Critical)
            }
        }

        internal void ExTest(string p_TestName, string p_EXMessage, string p_LogMessage)
        {
            string l_Msg;
            l_Msg = Strings.UCase(p_EXMessage);
            if ((l_Msg.Contains("NOT") | l_Msg.Contains("DOESN'T")) & (l_Msg.Contains("SET") | l_Msg.Contains("IMPLEMENTED") | l_Msg.Contains("SUPPORTED") | l_Msg.Contains("PRESENT")) | l_Msg.Contains("INVALID") | l_Msg.Contains("SUPPORT"))
            {
                // 3.0.0.12 - removed next two lines and added third to make this an OK message
                if (settings.InterpretErrorMessages) // We are interpreting error messages so report this as OK (This was Conform's behaviour prior to 6.0.0.37)
                {
                    LogMsg(p_TestName, MessageLevel.OK, p_LogMessage);
                }
                else // As of v6.0.0.37 default behaviour is not to interpret error messages
                {
                    LogMsg(p_TestName, MessageLevel.Info, "The following Issue can be changed to OK by setting \"Interpret error messages\" in Conform's setup dialogue");
                    LogMsg(p_TestName, MessageLevel.Issue, p_LogMessage);
                }
            }
            else
            {
                LogMsg(p_TestName, MessageLevel.Issue, p_LogMessage);
            }
        }





        public void DisposeAndReleaseObject(string driverName, dynamic ObjectToRelease)
        {
            Type ObjectType;
            int RemainingObjectCount, LoopCount;
            LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  About to release {driverName} driver instance");
            if (settings.DisplayMethodCalls)
                LogMsg("DisposeAndReleaseObject", MessageLevel.Comment, $"About to release {driverName} driver instance");
            try
            {
                ObjectType = ObjectToRelease.GetType();
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  Unmarshalling {ObjectType.Name} -  {ObjectType.FullName}");
            }
            catch (Exception ex1)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  GetType Exception: " + ex1.Message);
            }

            try
            {
                if (settings.DisplayMethodCalls)
                    LogMsg("DisposeAndReleaseObject", MessageLevel.Comment, "About to set Connected property");
                ObjectToRelease.Connected = false;
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  Connected successfully set to False");
            }
            catch (Exception ex1)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Exception setting Connected = False: " + ex1.Message);
            }

            try
            {
                ObjectToRelease.Dispose();
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, $"  Successfully called Dispose()");
            }
            catch (Exception ex1)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Dispose Exception: " + ex1.Message);
            }

            try
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Releasing COM object");
                LoopCount = 0;
                do
                {
                    LoopCount += 1;
                    RemainingObjectCount = Marshal.ReleaseComObject(ObjectToRelease);
                    LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Remaining object count: " + RemainingObjectCount + ", LoopCount: " + LoopCount);
                }
                while (!(RemainingObjectCount <= 0 | LoopCount == 20));
            }
            catch (Exception ex2)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  ReleaseComObject Exception: " + ex2.Message);
            }

            try
            {
                ObjectToRelease = null;
                GC.Collect();
            }
            catch (Exception ex3)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  Set to nothing Exception: " + ex3.Message);
            }

            LogMsg("DisposeAndReleaseObject", MessageLevel.Debug, "  End of ReleaseCOMObject");
        }

        /// <summary>
        /// Test a supplied exception for whether it is a MethodNotImplemented type
        /// </summary>
        /// <param name="deviceException">The exception sent by the driver</param>
        /// <returns>True if the exception is a NotImplemented type</returns>
        /// <remarks>Different tests are applied for COM and MethodNotImplemmented exceptions</remarks>
        protected bool IsMethodNotImplementedException(Exception deviceException)
        {
            bool IsMethodNotImplementedExceptionRet = default;
            COMException COMException;
            IsMethodNotImplementedExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException) // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                {
                    COMException = (COMException)deviceException;
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
                LogMsg("IsMethodNotImplementedException", MessageLevel.Warning, "Unexpected exception: " + ex.ToString());
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
            bool IsNotImplementedExceptionRet = default;
            COMException COMException;
            IsNotImplementedExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException) // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                {
                    COMException = (COMException)deviceException;
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
                LogMsg("IsNotImplementedException", MessageLevel.Warning, "Unexpected exception: " + ex.ToString());
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
            bool IsPropertyNotImplementedExceptionRet = default;
            COMException COMException;
            IsPropertyNotImplementedExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException) // This is a COM exception so test whether the error code indicates that it is a not implemented exception
                {
                    COMException = (COMException)deviceException;
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
                LogMsg("IsPropertyNotImplementedException", MessageLevel.Warning, "Unexpected exception: " + ex.ToString());
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
            bool IsInvalidValueExceptionRet = default;
            COMException COMException;
            DriverException DriverException;
            IsInvalidValueExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException) // This is a COM exception so test whether the error code indicates that it is an invalid value exception
                {
                    COMException = (COMException)deviceException;
                    if (COMException.ErrorCode == ErrorCodes.InvalidValue | COMException.ErrorCode == g_ExInvalidValue1 | COMException.ErrorCode == g_ExInvalidValue2 | COMException.ErrorCode == g_ExInvalidValue3 | COMException.ErrorCode == g_ExInvalidValue4 | COMException.ErrorCode == g_ExInvalidValue5 | COMException.ErrorCode == g_ExInvalidValue6) // This is an invalid value exception
                    {
                        IsInvalidValueExceptionRet = true;
                    }
                }

                if (deviceException is InvalidValueException)
                {
                    IsInvalidValueExceptionRet = true;
                }
                else if (deviceException is DriverException)
                {
                    DriverException = (DriverException)deviceException;
                    if (DriverException.Number == ErrorCodes.InvalidValue) // This is an invalid value exception
                    {
                        LogMsg(MemberName, MessageLevel.Issue, "Received ASCOM.DriverException(0x" + ErrorCodes.InvalidValue.ToString("X8") + "), please use ASCOM.InvalidValueException to report invalid values");
                    }
                }

                if (deviceException is System.InvalidOperationException)
                {
                    LogMsg(MemberName, MessageLevel.Issue, "Received System.InvalidOperationException rather than ASCOM.InvalidValueException");
                }
            }
            catch (Exception ex)
            {
                LogMsg("IsInvalidValueException", MessageLevel.Warning, "Unexpected exception: " + ex.ToString());
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
            bool IsInvalidOperationExceptionRet = default;
            COMException COMException;
            DriverException DriverException;
            IsInvalidOperationExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException) // This is a COM exception so test whether the error code indicates that it is an invalid operation exception
                {
                    COMException = (COMException)deviceException;
                    if (COMException.ErrorCode == ErrorCodes.InvalidOperationException) // This is an invalid operation exception
                    {
                        IsInvalidOperationExceptionRet = true;
                    }
                }

                if (deviceException is ASCOM.InvalidOperationException)
                {
                    IsInvalidOperationExceptionRet = true;
                }
                else if (deviceException is DriverException)
                {
                    DriverException = (DriverException)deviceException;
                    if (DriverException.Number == ErrorCodes.InvalidOperationException) // This is an invalid operation exception
                    {
                        LogMsg(MemberName, MessageLevel.Issue, "Received ASCOM.DriverException(0x" + ErrorCodes.InvalidOperationException.ToString("X8") + "), please use ASCOM.InvalidOperationException to report invalid operations");
                    }
                }

                if (deviceException is System.InvalidOperationException)
                {
                    LogMsg(MemberName, MessageLevel.Issue, "Received System.InvalidOperationException rather than ASCOM.InvalidOperationException");
                }
            }
            catch (Exception ex)
            {
                LogMsg("IsInvalidOperationException", MessageLevel.Warning, "Unexpected exception: " + ex.ToString());
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
            bool IsNotSetExceptionRet = default;
            COMException COMException;
            IsNotSetExceptionRet = false; // Set false default value
            try
            {
                if (deviceException is COMException) // This is a COM exception so test whether the error code indicates that it is a not set exception
                {
                    COMException = (COMException)deviceException;
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
                LogMsg("IsNotSetException", MessageLevel.Warning, "Unexpected exception: " + ex.ToString());
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
                            LogMsg(MemberName, MessageLevel.Issue, "This member is mandatory but threw a " + GetExceptionName(ex) + " exception, it must function per the ASCOM specification.");
                            break;
                        }

                    case Required.MustNotBeImplemented:
                        {
                            LogMsg(MemberName, MessageLevel.OK, UserMessage + " and a " + GetExceptionName(ex) + " exception was generated as expected");
                            break;
                        }

                    case Required.MustBeImplemented:
                        {
                            LogMsg(MemberName, MessageLevel.Issue, UserMessage + " and a " + GetExceptionName(ex) + " exception was thrown, this method must function per the ASCOM specification.");
                            break;
                        }

                    case Required.Optional:
                        {
                            LogMsg(MemberName, MessageLevel.OK, "Optional member threw a " + GetExceptionName(ex) + " exception.");
                            break;
                        }

                    default:
                        {
                            LogMsg(MemberName, MessageLevel.Error, (Conversions.ToDouble("CONFORM ERROR! - Received unexpected member of 'Required' enum: ") + (double)IsRequired).ToString());
                            break;
                        }
                }
            }

            // Handle wrong type of not implemented exceptions
            else if (ex is MethodNotImplementedException & TypeOfMember == MemberType.Property) // We got a MethodNotImplementedException so this is an error
            {
                LogMsg(MemberName, MessageLevel.Issue, "Received a MethodNotImplementedException instead of a PropertyNotImplementedException");
            }
            else if (ex is PropertyNotImplementedException & TypeOfMember == MemberType.Method) // We got a PropertyNotImplementedException so this is an error
            {
                LogMsg(MemberName, MessageLevel.Issue, "Received a PropertyNotImplementedException instead of a MethodNotImplementedException");
            }
            else if (ex is ASCOM.NotImplementedException)
            {
                LogMsg(MemberName, MessageLevel.Issue, Conversions.ToString(Operators.ConcatenateObject("Received a NotImplementedException instead of a ", Interaction.IIf(TypeOfMember == MemberType.Property, "PropertyNotImplementedException", "MethodNotImplementedException"))));
            }
            else if (ex is System.NotImplementedException)
            {
                LogMsg(MemberName, MessageLevel.Issue, Conversions.ToString(Operators.ConcatenateObject("Received a System.NotImplementedException instead of an ASCOM.", Interaction.IIf(TypeOfMember == MemberType.Property, "PropertyNotImplementedException", "MethodNotImplementedException"))));
            }

            // Handle all other types of error
            else
            {
                LogMsg(MemberName, MessageLevel.Error, "Unexpected " + GetExceptionName(ex) + ", " + UserMessage + ": " + ex.Message);
            }

            LogMsg(MemberName, MessageLevel.Debug, "Exception: " + ex.ToString());
        }

        protected void HandleInvalidValueExceptionAsOK(string MemberName, MemberType TypeOfMember, Required IsRequired, Exception ex, string UserAction, string Message)
        {
            if (IsInvalidValueException(MemberName, ex))
            {
                LogMsg(MemberName, MessageLevel.OK, Message);
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
                LogMsg(MemberName, MessageLevel.Info, Message);
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
                LogMsg(MemberName, MessageLevel.OK, Message);
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
        protected string GetExceptionName(Exception ex)
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
            else if (ex is COMException) // Handle XOM exceptions with their error code
            {
                ComEx = (COMException)ex;
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
                LogMsg(test, MessageLevel.Comment, memberName);
        }

        #endregion
    }
}
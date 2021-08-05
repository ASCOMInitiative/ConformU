// Base class from which particular device testers are derived
// Put all common elements in here
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ASCOM;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using static Conform.GlobalVarsAndCode;

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
        public DeviceTesterBaseClass(bool HasCanProperties, bool HasProperties, bool HasMethods, bool HasPreRunCheck, bool HasPreConnectCheck, bool HasPerformanceCheck, bool HasPostRunCheck) : base()
        {
            l_HasPreConnectCheck = HasPreConnectCheck;
            l_Connected = false;
            l_HasPreRunCheck = HasPreRunCheck;
            l_HasCanProperties = HasCanProperties;
            l_HasProperties = HasProperties;
            l_HasMethods = HasMethods;
            l_HasPostRunCheck = HasPostRunCheck;
            l_HasPerformanceCheck = HasPerformanceCheck;
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
            ArrayList SA;

            // Dim m_Configuration, SC() As String
            bool m_Connected;
            LogMsg("Common Driver Methods", MessageLevel.msgAlways, "");
            Device = p_DeviceObject; // CType(DeviceObject, IAscomDriverV1)

            // InterfaceVersion - Required
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("InterfaceVersion", MessageLevel.msgComment, "About to get property InterfaceVersion");
                g_InterfaceVersion = Conversions.ToInteger(Device.InterfaceVersion);
                switch (g_InterfaceVersion)
                {
                    case var @case when @case < 1:
                        {
                            LogMsg("InterfaceVersion", MessageLevel.msgIssue, "InterfaceVersion must be 1 or greater but driver returned: " + g_InterfaceVersion.ToString());
                            break;
                        }

                    default:
                        {
                            LogMsg("InterfaceVersion", MessageLevel.msgOK, g_InterfaceVersion.ToString());
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
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("Connected", MessageLevel.msgComment, "About to get property Connected");
                    m_Connected = Device.Connected;
                    LogMsg("Connected", MessageLevel.msgOK, m_Connected.ToString());
                }
                catch (COMException ex)
                {
                    LogMsg("Connected", MessageLevel.msgError, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                }
                catch (Exception ex)
                {
                    LogMsg("Connected", MessageLevel.msgError, ex.Message);
                }

                if (TestStop())
                    return;
            }

            // Description - Required
            if (IncludeMethod(MandatoryMethod.Description, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("Description", MessageLevel.msgComment, "About to get property Description");
                    m_Description = Conversions.ToString(Device.Description);
                    switch (m_Description ?? "")
                    {
                        case var case1 when case1 == "":
                            {
                                LogMsg("Description", MessageLevel.msgInfo, "No description string");
                                break;
                            }

                        default:
                            {
                                if (m_Description.Length > 68 & p_DeviceType == DeviceType.Camera)
                                {
                                    LogMsg("Description", MessageLevel.msgError, "Maximum number of characters is 68 for compatibility with FITS headers, found: " + m_Description.Length + " characters: " + m_Description);
                                }
                                else
                                {
                                    LogMsg("Description", MessageLevel.msgOK, m_Description.ToString());
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
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("DriverInfo", MessageLevel.msgComment, "About to get property DriverInfo");
                    m_DriverInfo = Conversions.ToString(Device.DriverInfo);
                    switch (m_DriverInfo ?? "")
                    {
                        case var case2 when case2 == "":
                            {
                                LogMsg("DriverInfo", MessageLevel.msgInfo, "No DriverInfo string");
                                break;
                            }

                        default:
                            {
                                LogMsg("DriverInfo", MessageLevel.msgOK, m_DriverInfo.ToString());
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
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("DriverVersion", MessageLevel.msgComment, "About to get property DriverVersion");
                    m_DriverVersion = Conversions.ToString(Device.DriverVersion);
                    switch (m_DriverVersion ?? "")
                    {
                        case var case3 when case3 == "":
                            {
                                LogMsg("DriverVersion", MessageLevel.msgInfo, "No DriverVersion string");
                                break;
                            }

                        default:
                            {
                                LogMsg("DriverVersion", MessageLevel.msgOK, m_DriverVersion.ToString());
                                break;
                            }
                    }
                }
                catch (COMException ex)
                {
                    LogMsg("DriverVersion", MessageLevel.msgError, EX_COM + ex.Message + " " + Conversion.Hex(ex.ErrorCode));
                }
                catch (Exception ex)
                {
                    LogMsg("DriverVersion", MessageLevel.msgError, ex.Message);
                }

                if (TestStop())
                    return;
            }
            else
            {
                LogMsg("DriverVersion", MessageLevel.msgInfo, "Skipping test as this method is not supported in interface V" + g_InterfaceVersion);
            }

            // Name - Required
            if (IncludeMethod(MandatoryMethod.Name, p_DeviceType, g_InterfaceVersion))
            {
                try
                {
                    if (g_Settings.DisplayMethodCalls)
                        LogMsg("Name", MessageLevel.msgComment, "About to get property Name");
                    m_Name = Conversions.ToString(Device.Name);
                    switch (m_Name ?? "")
                    {
                        case var case4 when case4 == "":
                            {
                                LogMsg("Name", MessageLevel.msgInfo, "Name is empty");
                                break;
                            }

                        default:
                            {
                                LogMsg("Name", MessageLevel.msgOK, m_Name);
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
            if (IncludeMethod(MandatoryMethod.CommandXXX, p_DeviceType, g_InterfaceVersion))
            {

                // Validate that the g_TelescopeTests collection is in a healthy state
                try
                {
                    LogMsgDebug("CommandXXX Tests", $"Test collection is null: {Information.IsNothing(g_TelescopeTests)}");
                    LogMsgDebug("CommandXXX Tests", $"Test collection size: {g_TelescopeTests.Count}");
                    foreach (KeyValuePair<string, CheckState> kvp in g_TelescopeTests)
                        LogMsgDebug("CommandXXX Tests", $"Found key: {kvp.Key} = {g_TelescopeTests[kvp.Key]}");
                }
                catch (Exception ex)
                {
                    LogMsgDebug("CommandXXX Tests", $"Exception: {ex}");
                }

                if (g_TelescopeTests[TELTEST_COMMANDXXX] == CheckState.Checked)
                {
                    LogMsg("CommandString", MessageLevel.msgInfo, "Conform cannot test the CommandString method");
                    LogMsg("CommandBlind", MessageLevel.msgInfo, "Conform cannot test the CommandBlind method");
                    LogMsg("CommandBool", MessageLevel.msgInfo, "Conform cannot test the CommandBool method");
                }
                else
                {
                    LogMsg(TELTEST_COMMANDXXX, MessageLevel.msgInfo, "Tests skipped");
                }
            }

            // Action - optional but cannot be tested
            LogMsg("Action", MessageLevel.msgInfo, "Conform cannot test the Action method");

            // Supported actions - Optional but Required through DriverAccess
            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("SupportedActions", MessageLevel.msgComment, "About to call method SupportedActions");
                SA = (ArrayList)Device.SupportedActions;
                if (SA.Count == 0)
                {
                    LogMsg("SupportedActions", MessageLevel.msgOK, "Driver returned an empty action list");
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
                                        LogMsg("SupportedActions", MessageLevel.msgError, "Supported action " + i + " Is an empty string"); // List the action that was found
                                        break;
                                    }

                                default:
                                    {
                                        LogMsg("SupportedActions", MessageLevel.msgOK, "Found action: " + ActionString);

                                        // Carry out the following Action tests only when we are testing the Observing Conditions Hub and it is configured to use the Switch and OC simulators
                                        if (p_DeviceType == DeviceType.ObservingConditions & g_ObservingConditionsProgID.ToUpperInvariant() == "ASCOM.OCH.OBSERVINGCONDITIONS")
                                        {
                                            if (ActionString.ToUpperInvariant().StartsWith("//OCSIMULATOR:"))
                                            {
                                                try
                                                {
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.msgComment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.msgOK, string.Format("OC simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.msgError, string.Format("Exception calling OCH simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//ASCOM.SIMULATOR.OBSERVINGCONDITIONS:"))
                                            {
                                                try
                                                {
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.msgComment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.msgOK, string.Format("OC simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.msgError, string.Format("Exception calling OCH simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//SWITCHSIMULATOR:"))
                                            {
                                                try
                                                {
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.msgComment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.msgOK, string.Format("Switch simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.msgError, string.Format("Exception calling switch simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                            else if (ActionString.ToUpperInvariant().StartsWith("//ASCOM.SIMULATOR.SWITCH:"))
                                            {
                                                try
                                                {
                                                    if (g_Settings.DisplayMethodCalls)
                                                        LogMsg("SupportedActions", MessageLevel.msgComment, "About to call method Action");
                                                    result = Conversions.ToString(Device.Action(ActionString, TEST_PARAMETERS));
                                                    LogMsg("SupportedActions", MessageLevel.msgOK, string.Format("Switch simulator action {0} gave result: {1}", ActionString, result));
                                                }
                                                catch (Exception ex1)
                                                {
                                                    LogMsg("SupportedActions", MessageLevel.msgError, string.Format("Exception calling switch simulator action {0}: {1}", ActionString, ex1.Message));
                                                }
                                            }
                                        }

                                        break;
                                    }
                            }
                        }
                        else
                        {
                            LogMsg("SupportedActions", MessageLevel.msgError, "Actions must be strings. The type of action " + i + " " + Action.ToString() + " is: " + Action.GetType().Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (p_DeviceType == DeviceType.Switch & ReferenceEquals(ex.GetType(), typeof(MissingMemberException)))
                {
                    LogMsg("SupportedActions", MessageLevel.msgOK, "Switch V1 Driver does not have SupportedActions");
                }
                else
                {
                    HandleException("SupportedActions", MemberType.Property, Required.Optional, ex, "");
                    LogMsg("SupportedActions", MessageLevel.msgError, ex.Message);
                }
            }

            if (TestStop())
                return;
            LogMsg("", MessageLevel.msgAlways, "");
        }

        public virtual void CheckCommonMethods()
        {
            LogMsg("CheckCommonMethods", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
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
            LogMsg("ConformanceCheckInitialise", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        /// <summary>
        /// Get error codes.
        /// </summary>
        /// <param name="p_ProgID">The ProgID.</param>
        protected void CheckInitialise(string p_ProgID)
        {
            ASCOM.Utilities.RegistryAccess Prof;
            Prof = new ASCOM.Utilities.RegistryAccess();
            Status(StatusType.staTest, ""); // Clear status messages
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
            g_Stop = true; // Initialise stop flag to stop
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            var fileInfo = new System.IO.FileInfo(assembly.Location);
            var lastModified = fileInfo.LastWriteTime;
            LogMsg("", MessageLevel.msgAlways, ""); // Blank line
            LogMsg("ConformanceCheck", MessageLevel.msgAlways, "ASCOM Device Conformance Checker Version " + My.MyProject.Application.Info.Version.ToString() + ", Build time: " + lastModified.ToString());
            LogMsg("ConformanceCheck", MessageLevel.msgAlways, "Running on: " + Prof.GetProfile("Platform", "Platform Name", "Unknown") + " " + Prof.GetProfile("Platform", "Platform Version", "Unknown"));
            Prof.Dispose();
            LogMsg("", MessageLevel.msgAlways, ""); // Blank line
            LogMsg("ConformanceCheck", MessageLevel.msgAlways, DRIVER_PROGID + p_ProgID);
            LogMsg("", MessageLevel.msgAlways, ""); // Blank line
            LogMsg("Error handling", MessageLevel.msgAlways, "");
            LogMsg("Error", MessageLevel.msgAlways, "number for \"Not Implemented\" is: " + Conversion.Hex(g_ExNotImplemented));
            LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 1\" is: " + Conversion.Hex(g_ExInvalidValue1));
            if (g_ExInvalidValue2 != 0 & g_ExInvalidValue2 != g_ExInvalidValue1)
                LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 2\" is: " + Conversion.Hex(g_ExInvalidValue2));
            if (g_ExInvalidValue3 != 0 & g_ExInvalidValue3 != g_ExInvalidValue2)
                LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 3\" is: " + Conversion.Hex(g_ExInvalidValue3));
            if (g_ExInvalidValue4 != 0 & g_ExInvalidValue4 != g_ExInvalidValue3)
                LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 4\" is: " + Conversion.Hex(g_ExInvalidValue4));
            if (g_ExInvalidValue5 != 0 & g_ExInvalidValue5 != g_ExInvalidValue4)
                LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 5\" is: " + Conversion.Hex(g_ExInvalidValue5));
            if (g_ExInvalidValue6 != 0 & g_ExInvalidValue6 != g_ExInvalidValue5)
                LogMsg("Error", MessageLevel.msgAlways, "number for \"Invalid Value 6\" is: " + Conversion.Hex(g_ExInvalidValue6));
            LogMsg("Error", MessageLevel.msgAlways, "number for \"Value Not Set 1\" is: " + Conversion.Hex(ErrorCodes.ValueNotSet));
            LogMsg("Error", MessageLevel.msgAlways, "number for \"Value Not Set 2\" is: " + Conversion.Hex(g_ExNotSet1));
            if (g_ExNotSet2 != 0 & g_ExNotSet2 != g_ExNotSet1)
                LogMsg("Error", MessageLevel.msgAlways, "number for \"Value Not Set 3\" is: " + Conversion.Hex(g_ExNotSet2));
            if (g_Settings.InterpretErrorMessages)
            {
                LogMsg("Error", MessageLevel.msgAlways, "messages will be interpreted to infer state.");
            }
            else
            {
                LogMsg("Error", MessageLevel.msgAlways, "messages will not be interpreted to infer state.");
            }

            LogMsg("", MessageLevel.msgAlways, "");
        }

        public virtual void CheckAccessibility()
        {
            LogMsg("ConformanceCheckAccessibility", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected void CheckAccessibility(string p_ProgId, DeviceType p_DeviceType)
        {
            dynamic l_DeviceObject;
            Type l_Type;
            var l_TryCount = default(int);
            string l_ErrMsg = "";
            LogMsg("Driver Access Checks", MessageLevel.msgOK, "");

            // Try late binding as an object
            l_DeviceObject = null;
            do
            {
                l_TryCount += 1;
                try
                {
#if DEBUG
                    if (g_Settings.DisplayMethodCalls) LogMsg("CreateObject", MessageLevel.msgComment, "About to create instance using CreateObject");
                    l_Type = Type.GetTypeFromProgID(p_ProgId);
                    l_DeviceObject = Activator.CreateInstance(l_Type);
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using CreateObject");
#else

                    l_Type = Type.GetTypeFromProgID(p_ProgId);
                    if (g_Settings.DisplayMethodCalls) LogMsg("AccessChecks", MessageLevel.msgComment, "About to create instance using Activator.CreateInstance");
                    l_DeviceObject = Activator.CreateInstance(l_Type);
                    LogMsg("AccessChecks", MessageLevel.msgDebug, "Successfully created driver using Activator.CreateInstance");

#endif
                    WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver initialisation");
                    LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully created driver using late binding");
                    try
                    {
                        switch (p_DeviceType)
                        {
                            case DeviceType.Focuser: // Focuser uses link to connect
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Link property true");
                                    l_DeviceObject.Link = true;
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Link property false");
                                    l_DeviceObject.Link = false; // Everything else uses connect!
                                    break;
                                }

                            default:
                                {
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property true");
                                    l_DeviceObject.Connected = true;
                                    if (g_Settings.DisplayMethodCalls)
                                        LogMsg("AccessChecks", MessageLevel.msgComment, "About to set Connected property false");
                                    try
                                    {
                                        l_DeviceObject.Connected = false;
                                    }
                                    catch (Exception ex)
                                    {
                                        LogMsg("AccessChecks", MessageLevel.msgError, "Error disconnecting from late bound driver: " + ex.Message);
                                        LogMsg("AccessChecks", MessageLevel.msgDebug, "Exception: " + ex.ToString());
                                    }

                                    break;
                                }
                        }

                        LogMsg("AccessChecks", MessageLevel.msgOK, "Successfully connected using late binding");
                        try
                        {
                            if (l_DeviceObject.GetType().IsCOMObject)
                            {
                                LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver is a COM object");
                            }
                            else
                            {
                                LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver is a .NET object");
                                LogMsg("AccessChecks", MessageLevel.msgInfo, "The AssemblyQualifiedName is: " + Strings.Left(l_DeviceObject.GetType().AssemblyQualifiedName.ToString(), 76));
                            }

                            foreach (var currentL_Type in l_DeviceObject.GetType().GetInterfaces())
                            {
                                l_Type = currentL_Type;
                                LogMsg("AccessChecks", MessageLevel.msgInfo, "The driver implements interface: " + l_Type.FullName);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMsg("AccessChecks", MessageLevel.msgError, "Error reading driver characteristics: " + ex.Message);
                            LogMsg("", MessageLevel.msgAlways, "");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMsg("AccessChecks", MessageLevel.msgError, "Error connecting to driver using late binding: " + ex.ToString());
                        LogMsg("", MessageLevel.msgAlways, "");
                    }
                }
                catch (Exception ex)
                {
                    l_ErrMsg = ex.ToString();
                    LogMsg("DeviceTesterBaseClass", MessageLevel.msgDebug, l_ErrMsg);
                }

                if (l_DeviceObject is null)
                    WaitFor(200);
            }
            while (!(l_TryCount == 3 | l_DeviceObject is object));
            if (l_DeviceObject is null)
            {
                LogMsg("AccessChecks", MessageLevel.msgError, "Error creating driver object using late binding: " + l_ErrMsg);
                LogMsg("", MessageLevel.msgAlways, "");
            }

            // Clean up
            try
            {
                DisposeAndReleaseObject("AccessChecks", l_DeviceObject);
            }
            catch (Exception ex)
            {
                LogMsg("AccessChecks", MessageLevel.msgDebug, "Error releasing driver object using ReleaseCOMObject: " + ex.ToString());
            }

            l_DeviceObject = null;
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Collecting garbage");
            GC.Collect();
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Collecting garbage complete");
            GC.WaitForPendingFinalizers();
            LogMsg("AccessChecks", MessageLevel.msgDebug, "Finished waiting for pending finalisers");
            WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for device driver to be cleaned up by operating system"); // Wait to allow device to complete destruction
        }

        public virtual void CreateDevice()
        {
            LogMsg("CreateDevice", MessageLevel.msgError, "DeviceTester base Class warning message");
        }

        public virtual void PreConnectChecks()
        {
            LogMsg("PreConnectChecks", MessageLevel.msgError, "DeviceTester base Class warning message");
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
            LogMsg("ReadCanProperties", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PreRunCheck()
        {
            LogMsg("PreSafetyCheck", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckProperties()
        {
            LogMsg("CheckProperties", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckMethods()
        {
            LogMsg("CheckMethods", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void CheckPerformance()
        {
            LogMsg("CheckPerformance", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        public virtual void PostRunCheck()
        {
            LogMsg("PostSafetyCheck", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
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
                        MessageBox.Show("DeviceTesterBaseClass:SpecialTests - Unknown test: " + p_Test.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }
        }

        protected virtual void SpecialTelescopeSideOfPier()
        {
            LogMsg("SpecialTelescopeSideOfPier", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeDestinationSideOfPier()
        {
            LogMsg("SpecialTelescopeDestinationSideOfPier", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeSideOfPierAnalysis()
        {
            LogMsg("SpecialTelescopeFlipRange", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        protected virtual void SpecialTelescopeCommands()
        {
            LogMsg("SpecialTelescopeCommands", MessageLevel.msgError, "DeviceTester base Class warning message, you should not see this message!");
        }

        #endregion

        #region Base class support Code

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

        private void CommandTest(CommandType p_Type, string p_Name)
        {
            string l_CmdSent = "!! Unknown Value !!";
            string m_CommandString;
            bool m_CommandBool;
            g_Status.Test = p_Name;
            try
            {
                switch (p_Type)
                {
                    case CommandType.tstCommandBlind:
                        {
                            if (g_CmdStrings.CommandBlind is object)
                            {
                                l_CmdSent = g_CmdStrings.CommandBlind;
                                Device.CommandBlind(l_CmdSent, false);
                                LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandBlind + "\"");
                            }
                            else
                            {
                                LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBlind test");
                            }

                            break;
                        }

                    case CommandType.tstCommandBool:
                        {
                            if (g_CmdStrings.CommandBool is object)
                            {
                                l_CmdSent = g_CmdStrings.CommandBool;
                                m_CommandBool = Conversions.ToBoolean(Device.CommandBool(l_CmdSent, (object)false));
                                if (m_CommandBool == g_CmdStrings.ReturnBool)
                                {
                                    LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received expected return value: " + m_CommandBool);
                                }
                                else
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStrings.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
                                }
                            }
                            else
                            {
                                LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBool test");
                            }

                            break;
                        }

                    case CommandType.tstCommandString:
                        {
                            if (g_CmdStrings.CommandString is object)
                            {
                                l_CmdSent = g_CmdStrings.CommandString;
                                m_CommandString = Conversions.ToString(Device.CommandString(l_CmdSent, (object)false));
                                if (g_CmdStrings.ReturnString is object) // Valid return string to test
                                {
                                    if ((m_CommandString ?? "") == (g_CmdStrings.ReturnString ?? ""))
                                    {
                                        LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
                                    }
                                    else
                                    {
                                        LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStrings.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStrings.ReturnString + "\"");
                                    }
                                }
                                else // Skip the return string test
                                {
                                    LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStrings.CommandString + "\" - Return string test skipped");
                                }
                            }
                            else
                            {
                                LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandString test");
                            }

                            break;
                        }

                    case CommandType.tstCommandBlindRaw:
                        {
                            if (g_CmdStringsRaw.CommandBlind is object)
                            {
                                l_CmdSent = g_CmdStringsRaw.CommandBlind;
                                Device.CommandBlind(l_CmdSent, true);
                                LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandBlind + "\"");
                            }
                            else
                            {
                                LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBlind Raw test");
                            }

                            break;
                        }

                    case CommandType.tstCommandBoolRaw:
                        {
                            if (g_CmdStringsRaw.CommandBool is object)
                            {
                                l_CmdSent = g_CmdStringsRaw.CommandBool;
                                m_CommandBool = Conversions.ToBoolean(Device.CommandBool(l_CmdSent, (object)true));
                                if (m_CommandBool == g_CmdStringsRaw.ReturnBool)
                                {
                                    LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received expected return value: " + m_CommandBool);
                                }
                                else
                                {
                                    LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStringsRaw.CommandBool + "\" - Received unexpected return value: " + m_CommandBool);
                                }
                            }
                            else
                            {
                                LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandBool Raw test");
                            }

                            break;
                        }

                    case CommandType.tstCommandStringRaw:
                        {
                            if (g_CmdStringsRaw.CommandString is object)
                            {
                                l_CmdSent = g_CmdStringsRaw.CommandString;
                                m_CommandString = Conversions.ToString(Device.CommandString(l_CmdSent, (object)true));
                                if (g_CmdStringsRaw.ReturnString is object) // Valid return string to test
                                {
                                    if ((m_CommandString ?? "") == (g_CmdStringsRaw.ReturnString ?? ""))
                                    {
                                        LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\"");
                                    }
                                    else
                                    {
                                        LogMsg(p_Name, MessageLevel.msgError, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Received string: \"" + m_CommandString + "\" - Expected string: \"" + g_CmdStringsRaw.ReturnString + "\"");
                                    }
                                }
                                else // Skip the return string test
                                {
                                    LogMsg(p_Name, MessageLevel.msgOK, "Sent string \"" + g_CmdStringsRaw.CommandString + "\" - Return string test skipped");
                                }
                            }
                            else
                            {
                                LogMsg(p_Name, MessageLevel.msgInfo, "Skipped CommandString Raw test");
                            }

                            break;
                        }

                    default:
                        {
                            LogMsg(p_Name, MessageLevel.msgError, "Conform:CommandTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(p_Name, MemberType.Method, Required.Optional, ex, "");
            }

            g_Status.Clear();  // Clear status messages
        }

        #endregion

        #region Common methods for all device tester classes
        public static void DisposeAndReleaseObject(string driverName, dynamic ObjectToRelease)
        {
            Type ObjectType;
            int RemainingObjectCount, LoopCount;
            LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, $"  About to release {driverName} driver instance");
            if (g_Settings.DisplayMethodCalls)
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgComment, $"About to release {driverName} driver instance");
            try
            {
                ObjectType = ObjectToRelease.GetType();
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, $"  Unmarshalling {ObjectType.Name} -  {ObjectType.FullName}");
            }
            catch (Exception ex1)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  GetType Exception: " + ex1.Message);
            }

            try
            {
                if (g_Settings.DisplayMethodCalls)
                    LogMsg("DisposeAndReleaseObject", MessageLevel.msgComment, "About to set Connected property");
                ObjectToRelease.Connected = false;
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, $"  Connected successfully set to False");
            }
            catch (Exception ex1)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  Exception setting Connected = False: " + ex1.Message);
            }

            try
            {
                ObjectToRelease.Dispose();
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, $"  Successfully called Dispose()");
            }
            catch (Exception ex1)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  Dispose Exception: " + ex1.Message);
            }

            try
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  Releasing COM object");
                LoopCount = 0;
                do
                {
                    LoopCount += 1;
                    RemainingObjectCount = Marshal.ReleaseComObject(ObjectToRelease);
                    LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  Remaining object count: " + RemainingObjectCount + ", LoopCount: " + LoopCount);
                }
                while (!(RemainingObjectCount <= 0 | LoopCount == 20));
            }
            catch (Exception ex2)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  ReleaseComObject Exception: " + ex2.Message);
            }

            try
            {
                ObjectToRelease = null;
                GC.Collect();
            }
            catch (Exception ex3)
            {
                LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  Set to nothing Exception: " + ex3.Message);
            }

            LogMsg("DisposeAndReleaseObject", MessageLevel.msgDebug, "  End of ReleaseCOMObject");
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
                LogMsg("IsMethodNotImplementedException", MessageLevel.msgWarning, "Unexpected exception: " + ex.ToString());
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
                LogMsg("IsNotImplementedException", MessageLevel.msgWarning, "Unexpected exception: " + ex.ToString());
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
                LogMsg("IsPropertyNotImplementedException", MessageLevel.msgWarning, "Unexpected exception: " + ex.ToString());
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
                        LogMsg(MemberName, MessageLevel.msgIssue, "Received ASCOM.DriverException(0x" + ErrorCodes.InvalidValue.ToString("X8") + "), please use ASCOM.InvalidValueException to report invalid values");
                    }
                }

                if (deviceException is System.InvalidOperationException)
                {
                    LogMsg(MemberName, MessageLevel.msgIssue, "Received System.InvalidOperationException rather than ASCOM.InvalidValueException");
                }
            }
            catch (Exception ex)
            {
                LogMsg("IsInvalidValueException", MessageLevel.msgWarning, "Unexpected exception: " + ex.ToString());
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
                        LogMsg(MemberName, MessageLevel.msgIssue, "Received ASCOM.DriverException(0x" + ErrorCodes.InvalidOperationException.ToString("X8") + "), please use ASCOM.InvalidOperationException to report invalid operations");
                    }
                }

                if (deviceException is System.InvalidOperationException)
                {
                    LogMsg(MemberName, MessageLevel.msgIssue, "Received System.InvalidOperationException rather than ASCOM.InvalidOperationException");
                }
            }
            catch (Exception ex)
            {
                LogMsg("IsInvalidOperationException", MessageLevel.msgWarning, "Unexpected exception: " + ex.ToString());
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
                LogMsg("IsNotSetException", MessageLevel.msgWarning, "Unexpected exception: " + ex.ToString());
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
                            LogMsg(MemberName, MessageLevel.msgIssue, "This member is mandatory but threw a " + GetExceptionName(ex) + " exception, it must function per the ASCOM specification.");
                            break;
                        }

                    case Required.MustNotBeImplemented:
                        {
                            LogMsg(MemberName, MessageLevel.msgOK, UserMessage + " and a " + GetExceptionName(ex) + " exception was generated as expected");
                            break;
                        }

                    case Required.MustBeImplemented:
                        {
                            LogMsg(MemberName, MessageLevel.msgIssue, UserMessage + " and a " + GetExceptionName(ex) + " exception was thrown, this method must function per the ASCOM specification.");
                            break;
                        }

                    case Required.Optional:
                        {
                            LogMsg(MemberName, MessageLevel.msgOK, "Optional member threw a " + GetExceptionName(ex) + " exception.");
                            break;
                        }

                    default:
                        {
                            LogMsg(MemberName, MessageLevel.msgError, (Conversions.ToDouble("CONFORM ERROR! - Received unexpected member of 'Required' enum: ") + (double)IsRequired).ToString());
                            break;
                        }
                }
            }

            // Handle wrong type of not implemented exceptions
            else if (ex is MethodNotImplementedException & TypeOfMember == MemberType.Property) // We got a MethodNotImplementedException so this is an error
            {
                LogMsg(MemberName, MessageLevel.msgIssue, "Received a MethodNotImplementedException instead of a PropertyNotImplementedException");
            }
            else if (ex is PropertyNotImplementedException & TypeOfMember == MemberType.Method) // We got a PropertyNotImplementedException so this is an error
            {
                LogMsg(MemberName, MessageLevel.msgIssue, "Received a PropertyNotImplementedException instead of a MethodNotImplementedException");
            }
            else if (ex is ASCOM.NotImplementedException)
            {
                LogMsg(MemberName, MessageLevel.msgIssue, Conversions.ToString(Operators.ConcatenateObject("Received a NotImplementedException instead of a ", Interaction.IIf(TypeOfMember == MemberType.Property, "PropertyNotImplementedException", "MethodNotImplementedException"))));
            }
            else if (ex is System.NotImplementedException)
            {
                LogMsg(MemberName, MessageLevel.msgIssue, Conversions.ToString(Operators.ConcatenateObject("Received a System.NotImplementedException instead of an ASCOM.", Interaction.IIf(TypeOfMember == MemberType.Property, "PropertyNotImplementedException", "MethodNotImplementedException"))));
            }

            // Handle all other types of error
            else
            {
                LogMsg(MemberName, MessageLevel.msgError, "Unexpected " + GetExceptionName(ex) + ", " + UserMessage + ": " + ex.Message);
            }

            LogMsg(MemberName, MessageLevel.msgDebug, "Exception: " + ex.ToString());
        }

        protected void HandleInvalidValueExceptionAsOK(string MemberName, MemberType TypeOfMember, Required IsRequired, Exception ex, string UserAction, string Message)
        {
            if (IsInvalidValueException(MemberName, ex))
            {
                LogMsg(MemberName, MessageLevel.msgOK, Message);
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
                LogMsg(MemberName, MessageLevel.msgInfo, Message);
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
                LogMsg(MemberName, MessageLevel.msgOK, Message);
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
            if (g_Settings.DisplayMethodCalls)
                LogMsg(test, MessageLevel.msgComment, memberName);
        }

        #endregion
    }
}
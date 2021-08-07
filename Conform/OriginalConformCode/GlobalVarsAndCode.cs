// Option Strict On
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using ASCOM.Standard.Utilities;
using Microsoft.VisualBasic;
using static ConformU.ConformConstants;

namespace Conform
{
    static class GlobalVarsAndCode
    {

        #region Enums

        //
        // Summary:
        //     Specifies the state of a control, such as a check box, that can be checked, unchecked,
        //     or set to an indeterminate state.
        public enum CheckState
        {
            //
            // Summary:
            //     The control is unchecked.
            Unchecked = 0,
            //
            // Summary:
            //     The control is checked.
            Checked = 1,
            //
            // Summary:
            //     The control is indeterminate. An indeterminate control generally has a shaded
            //     appearance.
            Indeterminate = 2
        }

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

        #region Constants
        internal const string DRIVER_PROGID = "Driver ProgID: ";
        internal const string NOT_IMP_NET = ".NET - Feature not implemented";
        internal const string NOT_IMP_COM = "COM - Feature not implemented";
        // Const INV_VAL_NET As String = ".NET - Invalid value: "
        internal const string EX_DRV_NET = ".NET - Driver Exception: ";
        internal const string EX_NET = ".NET - Exception: ";
        internal const string EX_COM = "COM - Exception: ";
        internal const int PERF_LOOP_TIME = 5; // Performance loop run time in seconds
        internal const int SLEEP_TIME = 200; // Loop time for testing whether slewing has completed
        internal const int CAMERA_SLEEP_TIME = 10; // Loop time for testing whether camera events have completed
        internal const int DEVICE_DESTROY_WAIT = 500; // Time to wait after destroying a device before continuing

        // TelescopeTest Constants
        internal const string TELTEST_ABORT_SLEW = "AbortSlew";
        internal const string TELTEST_AXIS_RATE = "AxisRate";
        internal const string TELTEST_CAN_MOVE_AXIS = "CanMoveAxis";
        internal const string TELTEST_COMMANDXXX = "CommandXXX";
        internal const string TELTEST_DESTINATION_SIDE_OF_PIER = "DestinationSideOfPier";
        internal const string TELTEST_FIND_HOME = "FindHome";
        internal const string TELTEST_MOVE_AXIS = "MoveAxis";
        internal const string TELTEST_PARK_UNPARK = "Park/Unpark";
        internal const string TELTEST_PULSE_GUIDE = "PulseGuide";
        internal const string TELTEST_SLEW_TO_ALTAZ = "SlewToAltAz";
        internal const string TELTEST_SLEW_TO_ALTAZ_ASYNC = "SlewToAltAzAsync";
        internal const string TELTEST_SLEW_TO_TARGET = "SlewToTarget";
        internal const string TELTEST_SLEW_TO_TARGET_ASYNC = "SlewToTargetAsync";
        internal const string TELTEST_SYNC_TO_ALTAZ = "SyncToAltAz";
        internal const string TELTEST_SLEW_TO_COORDINATES = "SlewToCoordinates";
        internal const string TELTEST_SLEW_TO_COORDINATES_ASYNC = "SlewToCoordinatesAsync";
        internal const string TELTEST_SYNC_TO_COORDINATES = "SyncToCoordinates";
        internal const string TELTEST_SYNC_TO_TARGET = "SyncToTarget";
        internal const int FWTEST_IS_MOVING = -1;
        internal const int FWTEST_TIMEOUT = 30;

        #endregion

        #region Global Variables
        internal static MessageLevel g_LogLevel;
        internal static int g_CountError, g_CountWarning, g_CountIssue;
        internal static System.IO.StreamWriter g_LogFile;
        internal static System.IO.StreamWriter g_ValidationTempFile;
        //internal static ConformCommandStrings g_CmdStrings = null;
        //internal static ConformCommandStrings g_CmdStringsRaw = null;
        internal static object g_DeviceObject;
        //internal static ApplicationSettings g_Settings = new ApplicationSettings();
        internal static int g_InterfaceVersion; // Variable to held interface version of the current device

        // Exception number variables
        internal static int g_ExNotImplemented, g_ExNotSet1, g_ExNotSet2;
        internal static int g_ExInvalidValue1, g_ExInvalidValue2, g_ExInvalidValue3, g_ExInvalidValue4, g_ExInvalidValue5, g_ExInvalidValue6;
        internal static bool g_Stop = false;

        // Helper variables
        internal static Utilities g_Util;
        // Friend g_Util As ASCOM.Utilities.Util

        // Device ProgID variables
        internal static string g_SafetyMonitorProgID;
        internal static string g_SwitchProgID;
        internal static string g_FilterWheelProgID;
        internal static string g_FocuserProgID;
        internal static string g_RotatorProgID;
        internal static string g_CameraProgID;
        internal static string g_VideoCameraProgID;
        internal static string g_DomeProgID;
        internal static string g_ObservingConditionsProgID = "";
        internal static string g_CoverCalibratorProgID;

        // Status update class
        //internal static Dictionary<string, CheckState> g_TelescopeTests = new Dictionary<string, CheckState>();
        //internal static Dictionary<string, CheckState> g_TelescopeTestsMaster = new Dictionary<string, CheckState>();

        internal static void SetTelescopeTestOptions()
        {
            // Populate the master list of Telescope Tests that are called from main and setup forms on-load events
            //g_TelescopeTestsMaster.Clear();
            //g_TelescopeTestsMaster.Add(TELTEST_CAN_MOVE_AXIS, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_PARK_UNPARK, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_ABORT_SLEW, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_AXIS_RATE, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_COMMANDXXX, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_FIND_HOME, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_MOVE_AXIS, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_PULSE_GUIDE, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SLEW_TO_COORDINATES, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SLEW_TO_COORDINATES_ASYNC, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SLEW_TO_TARGET, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SLEW_TO_TARGET_ASYNC, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_DESTINATION_SIDE_OF_PIER, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SLEW_TO_ALTAZ, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SLEW_TO_ALTAZ_ASYNC, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SYNC_TO_COORDINATES, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SYNC_TO_TARGET, CheckState.Checked);
            //g_TelescopeTestsMaster.Add(TELTEST_SYNC_TO_ALTAZ, CheckState.Checked);
        }

        #endregion

        #region Code
        internal static bool TestStop()
        {
            //Application.DoEvents();
            Thread.Sleep(10);
            //Application.DoEvents();
            Thread.Sleep(10);
            //Application.DoEvents();
            bool TestStopRet = g_Stop;
            return TestStopRet;
        }

        /// <summary>
        /// Delays execution for the given time period in milliseconds
        /// </summary>
        /// <param name="p_Duration">Delay duration in milliseconds</param>
        /// <remarks></remarks>
        internal static void WaitFor(int p_Duration)
        {
            DateTime l_StartTime;
            int WaitDuration;
            WaitDuration = (int)Math.Round(p_Duration / 100d);
            if (WaitDuration > SLEEP_TIME)
                WaitDuration = SLEEP_TIME;
            if (WaitDuration < 1)
                WaitDuration = 1;
            // Wait for p_Duration milliseconds
            l_StartTime = DateAndTime.Now; // Save start time
            do
            {
                Thread.Sleep(WaitDuration);
                //Application.DoEvents();
            }
            while (!(DateAndTime.Now.Subtract(l_StartTime).TotalMilliseconds > p_Duration | TestStop()));
        }

        internal static void WaitForAbsolute(int p_Duration, string p_Message)
        {
            //LogMsg("WaitForAbsolute", MessageLevel.Debug, p_Duration + " " + p_Message);
            for (int i = 0, loopTo = (int)Math.Round(p_Duration / 100d); i <= loopTo; i++)
            {
                Thread.Sleep(100);
                //Application.DoEvents();
                //SetStatus(p_Message, ((p_Duration / 100d - i) / 10d).ToString(), "");
            }

            //SetStatus("", "", "");
        }


        #endregion

    }
}
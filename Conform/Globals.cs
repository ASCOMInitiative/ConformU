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
    static class Globals
    {

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

        #region Constants
        internal const string NOT_IMP_NET = ".NET - Feature not implemented";
        internal const string NOT_IMP_COM = "COM - Feature not implemented";
        internal const string EX_DRV_NET = ".NET - Driver Exception: ";
        internal const string EX_NET = ".NET - Exception: ";
        internal const string EX_COM = "COM - Exception: ";
        internal const int PERF_LOOP_TIME = 5; // Performance loop run time in seconds
        internal const int SLEEP_TIME = 200; // Loop time for testing whether slewing has completed
        internal const int CAMERA_SLEEP_TIME = 10; // Loop time for testing whether camera events have completed
        internal const int DEVICE_DESTROY_WAIT = 500; // Time to wait after destroying a device before continuing

        // TelescopeTest constants
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

        // Filter wheel constants
        internal const int FWTEST_IS_MOVING = -1;
        internal const int FWTEST_TIMEOUT = 30;

        #endregion

        #region Global Variables

        internal static MessageLevel g_LogLevel;
        internal static int g_CountError, g_CountWarning, g_CountIssue;
        internal static int g_InterfaceVersion; // Variable to held interface version of the current device

        // Exception number variables
        internal static int g_ExNotImplemented, g_ExNotSet1, g_ExNotSet2;
        internal static int g_ExInvalidValue1, g_ExInvalidValue2, g_ExInvalidValue3, g_ExInvalidValue4, g_ExInvalidValue5, g_ExInvalidValue6;
        internal static bool g_Stop = false;

        // Helper variables
        internal static Utilities g_Util;

        #endregion

    }
}
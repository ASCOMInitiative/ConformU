using System;

namespace ConformU
{
    internal static class Globals
    {
        #region Global constants

        internal const string TECHNOLOGY_ALPACA = "Alpaca";
        internal const string TECHNOLOGY_COM = "COM";

        internal const string ASCOM_PROFILE_KEY = @"SOFTWARE\ASCOM";

        internal const string NO_DEVICE_SELECTED = "No device selected"; // Text indicating that no device has been sleected

        internal const string COMMAND_OPTION_SETTINGS = "ConformSettings";
        internal const string COMMAND_OPTION_LOGFILENAME = "ConformLogFileName";
        internal const string COMMAND_OPTION_LOGFILEPATH = "ConformLogFilePath";
        internal const string COMMAND_OPTION_SHOW_DISCOVERY = "ShowDiscovery";

        internal const int TEST_NAME_WIDTH = 45;
        
        #endregion

        #region Global Variables

        // Variables shared between the test manager and device testers        
        internal static int g_CountError, g_CountWarning, g_CountIssue;
        #endregion



    }

    public enum ComAccessMechanic
    {
        Native = 0,
        DriverAccess = 1
    }

    public enum DeviceTechnology
    {
        NotSelected = 0,
        Alpaca = 1,
        COM = 2
    }

    public enum DeviceType
    {
        NoDeviceType = 0,
        Telescope = 1,
        Camera = 2,
        Dome = 3,
        FilterWheel = 4,
        Focuser = 5,
        ObservingConditions = 6,
        Rotator = 7,
        Switch = 8,
        SafetyMonitor = 9,
        Video = 10,
        CoverCalibrator = 11
    }

    public enum MessageLevel
    {
        Debug = 0,
        Comment = 1,
        Info = 2,
        OK = 3,
        Warning = 4,
        Issue = 5,
        Error = 6,
        TestAndMessage = 7,
        TestOnly=8
    }

    // Must be valid service types because they are used as values in Alpaca access code i.e. ServiceType.http.ToString()
    public enum ServiceType
    {
        Http = 0,
        Https = 1
    }

}
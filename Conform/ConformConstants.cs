using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public enum ComAccessMechanic
    {
        Native = 0,
        DriverAccess = 1
    }

    public enum DeviceTechnology
    {
        NotSelected=0,
        Alpaca = 1,
        COM = 2
    }

    public enum DeviceType
    {
        NoDeviceType=0,
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
        None = 0,
        msgDebug = 1,
        msgComment = 2,
        msgInfo = 3,
        msgOK = 4,
        msgWarning = 5,
        msgIssue = 6,
        msgError = 7,
        msgAlways = 8
    }

    // Must be valid service types because they are used as values in Alpaca access code i.e. ServiceType.http.ToString()
    public enum ServiceType
    {
        Http = 0,
        Https = 1
    }

    public static class ConformConstants
    {
        public const string TECHNOLOGY_ALPACA = "Alpaca";
        public const string TECHNOLOGY_COM = "COM";

        public const string ASCOM_PROFILE_KEY = @"SOFTWARE\ASCOM";

        public const string NO_DEVICE_SELECTED = "No device selected"; // Text indicating that no device has been sleected

        public const string COMMAND_OPTION_SETTINGS = "ConformSettings";
        public const string COMMAND_OPTION_LOGFILENAME = "ConformLogFileName";
        public const string COMMAND_OPTION_LOGFILEPATH = "ConformLogFilePath";
        public const string COMMAND_OPTION_SHOW_DISCOVERY = "ShowDiscovery";
    }

}
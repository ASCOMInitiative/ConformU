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
        Alpaca = 0,
        COM = 1
    }

    public enum DeviceType
    {
        Telescope = 0,
        Camera = 1,
        Dome = 2,
        FilterWheel = 3,
        Focuser = 4,
        ObservingConditions = 5,
        Rotator = 6,
        Switch = 7,
        SafetyMonitor = 8,
        Video = 9,
        CoverCalibrator = 10
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

    }

}
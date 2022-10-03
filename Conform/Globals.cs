namespace ConformU
{
    internal static class Globals
    {

        #region Global constants

        internal const string TECHNOLOGY_ALPACA = "Alpaca";
        internal const string TECHNOLOGY_COM = "COM";

        internal const string ASCOM_PROFILE_KEY = @"SOFTWARE\ASCOM";

        internal const string NO_DEVICE_SELECTED = "No device selected"; // Text indicating that no device has been selected

        internal const string COMMAND_OPTION_SETTINGS = "ConformSettings";
        internal const string COMMAND_OPTION_LOGFILENAME = "ConformLogFileName";
        internal const string COMMAND_OPTION_LOGFILEPATH = "ConformLogFilePath";
        internal const string COMMAND_OPTION_DEBUG_DISCOVERY = "DebugDiscovery";
        internal const string COMMAND_OPTION_DEBUG_STARTUP = "DebugStartup";
        internal const string COMMAND_OPTION_RESULTS_FILENAME = "ResultsFileName";

        internal const int MESSAGE_LEVEL_WIDTH = 8; // Width to which the message level will be padded
        internal const int TEST_NAME_WIDTH = 35; // Width allowed for test names in screen display and log files

        internal const double UI_SMALL_TRANSITION_SIZE = 767.98;

        #endregion

        #region Global Variables

        // Variables shared between the test manager and device testers        
        internal static ConformResults conformResults;

        #endregion

        #region Static classes
        internal static bool IsSmall(double width, double height)
        {
            return (width < UI_SMALL_TRANSITION_SIZE) | (height < UI_SMALL_TRANSITION_SIZE);
        }

        #endregion

    }

    #region Global Enums

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

    public enum MessageLevel
    {
        Debug = 0,
        Info = 1,
        OK = 2,
        Issue = 3,
        Error = 4,
        TestAndMessage = 5,
        TestOnly = 6
    }

    #endregion

}
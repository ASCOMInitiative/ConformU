using System;

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

        internal const int TEST_NAME_WIDTH = 35; // Width allowed for test names in screen display and log files

        #endregion

        #region Global Variables

        // Variables shared between the test manager and device testers        
        internal static ConformResults conformResults;

        #endregion

        /// <summary>
        /// Static method to print to screen and log file
        /// </summary>
        /// <param name="testName"></param>
        /// <param name="messageLevel"></param>
        /// <param name="message"></param>
        /// <param name="debug"></param>
        /// <param name="parentClass"></param>
        /// <param name="TL"></param>
        /// <remarks>Used by DevicetesterbaseClass and ConformtestManager</remarks>
        internal static void LogMessage(string testName, MessageLevel messageLevel, string message, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            string testNameFormatted, messageLevelFormatted, screenMessage, logFileMessage;
            MessageLevel logLevel;

            if (debug)
            {
                logLevel = MessageLevel.Debug;
            }
            else
            {
                logLevel = MessageLevel.Info;
            }

            try
            {
                if (messageLevel >= logLevel)
                {
                    testNameFormatted = testName.PadRight(TEST_NAME_WIDTH).Substring(0, TEST_NAME_WIDTH); // Pad right to required length and limit to the required length
                    messageLevelFormatted = "        ";

                    switch (messageLevel)
                    {
                        case MessageLevel.Debug:
                            {
                                messageLevelFormatted = "DEBUG   ";
                                break;
                            }

                        case MessageLevel.Info:
                            {
                                messageLevelFormatted = "INFO    ";
                                break;
                            }

                        case MessageLevel.OK:
                            {
                                messageLevelFormatted = "OK      ";
                                break;
                            }

                        case MessageLevel.Issue:
                            {
                                messageLevelFormatted = "ISSUE   ";
                                break;
                            }

                        case MessageLevel.Error:
                            {
                                messageLevelFormatted = "ERROR   ";
                                break;
                            }

                        case MessageLevel.TestOnly:
                        case MessageLevel.TestAndMessage:
                            {
                                messageLevelFormatted = "        ";
                                break;
                            }
                        default:
                            {
                                throw new ASCOM.InvalidValueException($"LogMsg - Unknown message level: {messageLevel}.");
                            }
                    }

                    // Cater for screen display, which requires test name, message level and message in one string and log file, which requires test name separate from message level and message.
                    switch (messageLevel)
                    {
                        case MessageLevel.TestOnly:
                            {
                                screenMessage = testName;
                                logFileMessage = "";
                                break;
                            }
                        default:
                            {
                                screenMessage = $"{testNameFormatted} {messageLevelFormatted} {message}";
                                logFileMessage = $"{messageLevelFormatted} {message}";
                                break;
                            }
                    }
                    parentClass.OnLogMessageChanged("LogMessage", $"{DateTime.Now:HH:mm:ss.fff} {screenMessage}");
                    TL.LogMessage(testName, logFileMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception in DeviceTesterbaseClass.LogMsg method: \r\n{ex}");
            }
        }

        internal static void LogNewLine(bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            LogTestOnly("", debug, parentClass, TL);
        }

        internal static void LogTestOnly(string p_Test, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            LogMessage(p_Test, MessageLevel.TestOnly, "", debug, parentClass, TL);
        }

        internal static void LogTestAndMessage(string p_Test, string p_Msg, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            LogMessage(p_Test, MessageLevel.TestAndMessage, p_Msg, debug, parentClass, TL);
        }

        internal static void LogOK(string p_Test, string p_Msg, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            LogMessage(p_Test, MessageLevel.OK, p_Msg, debug, parentClass, TL);
        }

        internal static void LogDebug(string p_Test, string p_Msg, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            LogMessage(p_Test, MessageLevel.Debug, p_Msg, debug, parentClass, TL);
        }

        internal static void LogInfo(string p_Test, string p_Msg, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            LogMessage(p_Test, MessageLevel.Info, p_Msg, debug, parentClass, TL);
        }

        internal static void LogIssue(string p_Test, string p_Msg, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            conformResults.Issues.Add(new System.Collections.Generic.KeyValuePair<string, string>(p_Test, p_Msg));
            LogMessage(p_Test, MessageLevel.Issue, p_Msg, debug, parentClass, TL);
        }

        internal static void LogError(string p_Test, string p_Msg, bool debug, ConformanceTestManager parentClass, ConformLogger TL)
        {
            conformResults.Errors.Add(new System.Collections.Generic.KeyValuePair<string, string>(p_Test, p_Msg));
            LogMessage(p_Test, MessageLevel.Error, p_Msg, debug, parentClass, TL);
        }

    }

    #region Enums

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
        Info = 1,
        OK = 2,
        Issue = 3,
        Error = 4,
        TestAndMessage = 5,
        TestOnly = 6
    }

    // Must be valid service types because they are used as values in Alpaca access code i.e. ServiceType.http.ToString()
    public enum ServiceType
    {
        Http = 0,
        Https = 1
    }

    #endregion

}
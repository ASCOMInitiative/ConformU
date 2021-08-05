using System;
using System.Collections.Generic;
using ASCOM.Standard.Interfaces;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.Win32;
using static Conform.GlobalVarsAndCode;

namespace Conform
{

    // Class to manage state storage between Conform runs

    // To add a new saved value:
    // 1) Decide on the variable name and its default value
    // 2) Create appropriately named constants similar to those below
    // 3) Create a property of the relevant type
    // 4) Create Get and Set code based on the patterns already implemented
    // 5) If the property is of a type not already handled,you will need to create a GetXXX function in the Utility code region

    internal class ApplicationSettings
    {

        #region Constants

        private const string REGISTRY_CONFORM_FOLDER = @"Software\ASCOM\Conform";
        private const string REGISTRY_SIDEOFPIER_FOLDER = REGISTRY_CONFORM_FOLDER + @"\Side Of Pier";
        private const string REGISTRY_DESTINATIONSIDEOFPIER_FOLDER = REGISTRY_CONFORM_FOLDER + @"\Destination Side Of Pier";
        private const string REGISTRY_TELESCOPE_TESTS = REGISTRY_CONFORM_FOLDER + @"\Telescope Tests";
        private const string FLIP_TEST_HA_START = "Flip Test HA Start";
        private const double FLIP_TEST_HA_START_DEFAULT = 3.0d;
        private const string FLIP_TEST_HA_END = "Flip Test HA End";
        private const double FLIP_TEST_HA_END_DEFAULT = -3.0d;
        private const string FLIP_TEST_DEC_START = "Flip Test DEC Start";
        private const double FLIP_TEST_DEC_START_DEFAULT = 0.0d;
        private const string FLIP_TEST_DEC_END = "Flip Test DEC End";
        private const double FLIP_TEST_DEC_END_DEFAULT = 0.0d;
        private const string FLIP_TEST_DEC_STEP = "Flip Test DEC Step";
        private const double FLIP_TEST_DEC_STEP_DEFAULT = 10.0d;
        private const string DSOP_SIDE = "Destination Side Of Pier Side";
        private const string DSOP_SIDE_DEFAULT = "pierWest only";
        private const string SOP_SIDE = "Side Of Pier Side";
        private const string SOP_SIDE_DEFAULT = "pierWest only";
        private const string COMMAND_BLIND = "Command Blind";
        private const bool COMMAND_BLIND_DEFAULT = true;
        private const string COMMAND_BLIND_RAW = "Command Blind Raw";
        private const bool COMMAND_BLIND_RAW_DEFAULT = false;
        private const string COMMAND_BOOL = "Command Bool";
        private const bool COMMAND_BOOL_DEFAULT = true;
        private const string COMMAND_BOOL_RAW = "Command Bool Raw";
        private const bool COMMAND_BOOL_RAW_DEFAULT = false;
        private const string COMMAND_STRING = "Command String";
        private const bool COMMAND_STRING_DEFAULT = true;
        private const string COMMAND_STRING_RAW = "Command String Raw";
        private const bool COMMAND_STRING_RAW_DEFAULT = false;
        private const string CREATE_VALIDATION_FILE = "Create Validation File";
        private const bool CREATE_VALIDATION_FILE_DEFAULT = false;
        private const string DE_BUG = "Debug";
        private const bool DE_BUG_DEFAULT = false;
        private const string DEVICE_CAMERA = "Device Camera";
        private const string DEVICE_CAMERA_DEFAULT = "CCDSimulator.Camera";
        private const string DEVICE_VIDEO = "Device Video Camera";
        private const string DEVICE_VIDEO_DEFAULT = "ASCOM.Simulator.Video";
        private const string DEVICE_COVER_CALIBRATOR = "Device Cover Calibrator";
        private const string DEVICE_COVER_CALIBRATOR_DEFAULT = "ASCOM.Simulator.CoverCalibrator";
        private const string DEVICE_DOME = "Device Dome";
        private const string DEVICE_DOME_DEFAULT = "DomeSim.Dome";
        private const string DEVICE_FILTER_WHEEL = "Device Filter Wheel";
        private const string DEVICE_FILTER_WHEEL_DEFAULT = "FilterWheelSim.FilterWheel";
        private const string DEVICE_FOCUSER = "Device Focuser";
        private const string DEVICE_FOCUSER_DEFAULT = "FocusSim.Focuser";
        private const string DEVICE_OBSERVINGCONDITIONS = "Device Observing Conditions";
        private const string DEVICE_OBSERVINGCONDITIONS_DEFAULT = "ASCOM.OCH.ObservingConditions";
        private const string DEVICE_ROTATOR = "Device Rotator";
        private const string DEVICE_ROTATOR_DEFAULT = "ASCOM.Simulator.Rotator";
        private const string DEVICE_SAFETY_MONITOR = "Device Safety Monitor";
        private const string DEVICE_SAFETY_MONITOR_DEFAULT = "ASCOM.Simulator.SafetyMonitor";
        private const string DEVICE_SWITCH = "Device Switch";
        private const string DEVICE_SWITCH_DEFAULT = "SwitchSim.Switch";
        private const string DEVICE_TELESCOPE = "Device Telescope";
        private const string DEVICE_TELESCOPE_DEFAULT = "ScopeSim.Telescope";
        private const string CURRENT_DEVICE_TYPE = "Current Device Type";
        private const DeviceType CURRENT_DEVICE_TYPE_DEFAULT = DeviceType.Telescope;
        private const string LOG_FILES_DIRECTORY = "Log File Directory";
        private const string LOG_FILES_DIRECTORY_DEFAULT = @"\ASCOM";
        private const string MESSAGE_LEVEL = "Message Level";
        private const MessageLevel MESSAGE_LEVEL_DEFAULT = MessageLevel.msgInfo;
        private const string TEST_METHODS = "Test Methods";
        private const bool TEST_METHODS_DEFAULT = true;
        private const string TEST_PERFORMANCE = "Test Performance";
        private const bool TEST_PERFORMANCE_DEFAULT = false;
        private const string TEST_PROPERTIES = "Test Properties";
        private const bool TEST_PROPERTIES_DEFAULT = true;
        private const string UPDATE_CHECK = "Update Check";
        private const bool UPDATE_CHECK_DEFAULT = true;
        private const string UPDATE_CHECK_DATE = "Update Check Date";
        private static DateTime UPDATE_CHECK_DATE_DEFAULT = DateTime.Parse("2008-01-01 01:00:00");
        private const string WARNING_MESSAGE = "Warning Message Platform 6";
        private const bool WARNING_MESSAGE_DEFAULT = false; // Updated for Platform 6 to force a redisplay as the words have changed
        private const string TEST_SIDEOFPIER_READ = "Test SideOfPier Read";
        private const bool TEST_SIDEOFPIER_READ_DEFAULT = true;
        private const string TEST_SIDEOFPIER_WRITE = "Test SideOfPier Write";
        private const bool TEST_SIDEOFPIER_WRITE_DEFAULT = true;
        private const string RUN_AS_THIRTYTWO_BITS = "Run As 64Bit";
        private const bool RUN_AS_THIRTYTWO_BITS_DEFAULT = false;
        private const string INTERPRET_ERROR_MESSAGES = "Interpret Error Messages";
        private const bool INTERPRET_ERROR_MESSAGES_DEFAULT = false;
        private const string SWITCH_WARNING_MESSAGE = "Switch Warning Message";
        private const bool SWITCH_WARNING_MESSAGE_DEFAULT = false; // Updated for Platform 6 to force a redisplay as the words have changed
        private const string USE_DRIVERACCESS = "Use DriverAccess";
        private const bool USE_DRIVERACCESS_DEFAULT = true;
        private const string DISPLAY_METHOD_CALLS = "Display Method Calls";
        private const bool DISPLAY_METHOD_CALLS_DEFAULT = false;

        // Dome device variables
        private const string DOME_SHUTTER = "Dome Shutter";
        private const bool DOME_SHUTTER_DEFAULT = false;
        private const string DOME_SHUTTER_TIMEOUT = "Dome Shutter Timeout";
        private const double DOME_SHUTTER_TMEOUT_DEFAULT = 240.0d; // Timeout for dome to open or close shutter
        private const string DOME_AZIMUTH_TIMEOUT = "Dome Azimuth Timeout";
        private const double DOME_AZIMUTH_TIMEOUT_DEFAULT = 240.0d; // Timeout for dome to move to azimuth
        private const string DOME_ALTITUDE_TIMEOUT = "Dome Altitude Timeout";
        private const double DOME_ALTITUDE_TIMEOUT_DEFAULT = 240.0d; // Timeout for dome to move to altitude
        private const string DOME_STABILISATION_WAIT = "Dome Stabilisation Wait";
        private const double DOME_STABILISATION_WAIT_DEFAULT = 10.0d; // Time to wait for Dome to stabilise after a command

        // Switch device variables
        private const string SWITCH_SET = "Switch Set";
        private const bool SWITCH_SET_DEFAULT = false;
        private const string SWITCH_READ_DELAY = "Switch Read Delay";
        private const int SWITCH_READ_DELAY_DEFAULT = 500; // Switch delay after a set before undertaking a read (ms)
        private const string SWITCH_WRITE_DELAY = "Switch Write Delay";
        private const int SWITCH_WRITE_DELAY_DEFAULT = 3000; // Switch write delay between changing states (ms)
        private const string EXTENDED_SWITCH_NUMBER_TEST_RANGE = "Extended Switch Number Test Range";
        private const int EXTENDED_SWITCH_NUMBER_TEST_RANGE_DEFAULT = 100; // Switch write delay between changing states (ms)

        // ObservingConditions device variables
        private const string OBSERVINGCONDITIONS_RETRIES = "ObservingConditions Retries";
        private const int OBSERVINGCONDITIONS_MAX_RETRIES_DEFAULT = 5;
        private const string OBSERVINGCONDITIONS_RETRY_TIME = "ObservingConditions Retry Time";
        private const int OBSERVINGCONDITIONS_RETRY_TIME_DEFAULT = 1;

        // Camera device variables
        private const string CAMERA_MAX_BIN_X = "Camera Max Bin X";
        private const int CAMERA_MAX_BIN_X_DEFAULT = 0;
        private const string CAMERA_MAX_BIN_Y = "Camera Max Bin Y";
        private const int CAMERA_MAX_BIN_Y_DEFAULT = 0;

        #endregion

        #region Variables

        private RegistryKey m_HKCU, m_SettingsKey, m_SideOfPierKey, m_DestinationSideOfPierKey, m_TelescopeTestsKey;

        #endregion

        #region New and Finalize
        public ApplicationSettings()
        {
            m_HKCU = Registry.CurrentUser;
            m_HKCU.CreateSubKey(REGISTRY_CONFORM_FOLDER);
            m_HKCU.CreateSubKey(REGISTRY_SIDEOFPIER_FOLDER);
            m_HKCU.CreateSubKey(REGISTRY_DESTINATIONSIDEOFPIER_FOLDER);
            m_HKCU.CreateSubKey(REGISTRY_TELESCOPE_TESTS);
            m_SettingsKey = m_HKCU.OpenSubKey(REGISTRY_CONFORM_FOLDER, true);
            m_SideOfPierKey = m_HKCU.OpenSubKey(REGISTRY_SIDEOFPIER_FOLDER, true);
            m_DestinationSideOfPierKey = m_HKCU.OpenSubKey(REGISTRY_DESTINATIONSIDEOFPIER_FOLDER, true);
            m_TelescopeTestsKey = m_HKCU.OpenSubKey(REGISTRY_TELESCOPE_TESTS, true);
        }

        ~ApplicationSettings()
        {
            m_SideOfPierKey.Flush();
            m_SideOfPierKey.Close();
            m_SideOfPierKey = null;
            m_DestinationSideOfPierKey.Flush();
            m_DestinationSideOfPierKey.Close();
            m_DestinationSideOfPierKey = null;
            m_SettingsKey.Flush();
            m_SettingsKey.Close();
            m_SettingsKey = null;
            m_HKCU.Flush();
            m_HKCU.Close();
            m_HKCU = null;
        }
        #endregion

        #region Parameters

        // Dome
        public bool DomeShutter
        {
            get
            {
                return GetBool(DOME_SHUTTER, DOME_SHUTTER_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DOME_SHUTTER, value.ToString());
            }
        }

        public double DomeStabilisationWait
        {
            get
            {
                return GetDouble(m_SettingsKey, DOME_STABILISATION_WAIT, DOME_STABILISATION_WAIT_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DOME_STABILISATION_WAIT, value.ToString());
            }
        }

        public double DomeShutterTimeout
        {
            get
            {
                return GetDouble(m_SettingsKey, DOME_SHUTTER_TIMEOUT, DOME_SHUTTER_TMEOUT_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DOME_SHUTTER_TIMEOUT, value.ToString());
            }
        }

        public double DomeAzimuthTimeout
        {
            get
            {
                return GetDouble(m_SettingsKey, DOME_AZIMUTH_TIMEOUT, DOME_AZIMUTH_TIMEOUT_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DOME_AZIMUTH_TIMEOUT, value.ToString());
            }
        }

        public double DomeAltitudeTimeout
        {
            get
            {
                return GetDouble(m_SettingsKey, DOME_ALTITUDE_TIMEOUT, DOME_ALTITUDE_TIMEOUT_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DOME_ALTITUDE_TIMEOUT, value.ToString());
            }
        }

        // Telescope
        //public Dictionary<string, CheckState> TeleScopeTests
        //{
        //    get
        //    {
        //        var RetVal = new Dictionary<string, CheckState>();
        //        CheckState testState;
        //        string registryValue;
        //        foreach (KeyValuePair<string, CheckState> kvp in g_TelescopeTestsMaster)
        //        {
        //            try
        //            {
        //                //LogMsgDebug("TeleScopeTests", $"Retrieving key: {kvp.Key}");
        //                registryValue = m_TelescopeTestsKey.GetValue(kvp.Key).ToString();
        //                //LogMsgDebug("TeleScopeTests", $"Retrieved registry value: {registryValue} for {kvp.Key}");
        //                testState = (CheckState)Conversions.ToInteger(Enum.Parse(typeof(CheckState), registryValue, true));
        //                //LogMsgDebug("TeleScopeTests", $"Retrieved checked state: {testState} for {kvp.Key}");
        //                RetVal.Add(kvp.Key, testState);
        //            }
        //            catch (System.IO.IOException ex) // Value doesn't exist so create it
        //            {
        //                //LogMsgDebug("TeleScopeTests", $"IOException for key {kvp.Key}: {ex}");
        //                SetName(m_TelescopeTestsKey, kvp.Key, CheckState.Checked.ToString());
        //                RetVal.Add(kvp.Key, CheckState.Checked);
        //            }
        //            catch (Exception ex)
        //            {
        //                //LogMsgDebug("TeleScopeTests", $"Unexpected exception for key {kvp.Key}: {ex}");
        //                SetName(m_TelescopeTestsKey, kvp.Key, CheckState.Checked.ToString());
        //                RetVal.Add(kvp.Key, CheckState.Checked);
        //            }
        //        }

        //        //LogMsgDebug("TeleScopeTests", $"Returning {RetVal.Count} values.");
        //        return RetVal;
        //    }

        //    set
        //    {
        //        //LogMsgDebug("TeleScopeTests", $"Setting {value.Count} values.");
        //        foreach (KeyValuePair<string, CheckState> kvp in value)
        //            SetName(m_TelescopeTestsKey, kvp.Key, kvp.Value.ToString());
        //    }
        //}

        public bool TestSideOfPierRead
        {
            get
            {
                return GetBool(TEST_SIDEOFPIER_READ, TEST_SIDEOFPIER_READ_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, TEST_SIDEOFPIER_READ, value.ToString());
            }
        }

        public bool TestSideOfPierWrite
        {
            get
            {
                return GetBool(TEST_SIDEOFPIER_WRITE, TEST_SIDEOFPIER_WRITE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, TEST_SIDEOFPIER_WRITE, value.ToString());
            }
        }

        public string DSOPSide
        {
            get
            {
                return GetString(DSOP_SIDE, DSOP_SIDE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DSOP_SIDE, value);
            }
        }

        public string SOPSide
        {
            get
            {
                return GetString(SOP_SIDE, SOP_SIDE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, SOP_SIDE, value);
            }
        }

        public double get_FlipTestHAStart(SpecialTest p_TestType, PointingState p_PierSide)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                                {
                                    return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_START, FLIP_TEST_HA_START_DEFAULT);
                                }

                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_START, -FLIP_TEST_HA_START_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAStart Get - Unexpected pier side: " + p_PierSide.ToString(), "",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                                {
                                    return GetDouble(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_START, FLIP_TEST_HA_START_DEFAULT);
                                }

                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_START, -FLIP_TEST_HA_START_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAStart Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest HAStart - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }

            return default;
        }

        public void set_FlipTestHAStart(SpecialTest p_TestType, PointingState p_PierSide, double value)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_START, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAStart Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_START, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAStart Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest HAStart - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }
        }

        public double get_FlipTestHAEnd(SpecialTest p_TestType, PointingState p_PierSide)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                                {
                                    return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_END, FLIP_TEST_HA_END_DEFAULT);
                                }

                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_END, -FLIP_TEST_HA_END_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAEnd Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                                {
                                    return GetDouble(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_END, FLIP_TEST_HA_END_DEFAULT);
                                }

                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_END, -FLIP_TEST_HA_END_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAEnd Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest HAEnd - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }

            return default;
        }

        public void set_FlipTestHAEnd(SpecialTest p_TestType, PointingState p_PierSide, double value)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_END, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAEnd Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_HA_END, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest HAEnd Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest HAEnd - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }
        }

        public double get_FlipTestDECStart(SpecialTest p_TestType, PointingState p_PierSide)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_START, FLIP_TEST_DEC_START_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStart Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_START, FLIP_TEST_DEC_START_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStart Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest DECStart - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }

            return default;
        }

        public void set_FlipTestDECStart(SpecialTest p_TestType, PointingState p_PierSide, double value)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_START, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStart Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_START, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStart Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest DECStart - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }
        }

        public double get_FlipTestDECEnd(SpecialTest p_TestType, PointingState p_PierSide)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_END, FLIP_TEST_DEC_END_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECEnd Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_END, FLIP_TEST_DEC_END_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECEnd Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest DECStart - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }

            return default;
        }

        public void set_FlipTestDECEnd(SpecialTest p_TestType, PointingState p_PierSide, double value)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_END, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECEnd Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_END, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECEnd Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest DECEnd - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }
        }

        public double get_FlipTestDECStep(SpecialTest p_TestType, PointingState p_PierSide)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_STEP, FLIP_TEST_DEC_STEP_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStep Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    return GetDouble(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_STEP, FLIP_TEST_DEC_STEP_DEFAULT);
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStep Get - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest DECStep - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }

            return default;
        }

        public void set_FlipTestDECStep(SpecialTest p_TestType, PointingState p_PierSide, double value)
        {
            switch (p_TestType)
            {
                case SpecialTest.TelescopeDestinationSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_DestinationSideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_STEP, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStep Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                case SpecialTest.TelescopeSideOfPier:
                    {
                        switch (p_PierSide)
                        {
                            case PointingState.Normal:
                            case PointingState.ThroughThePole:
                                {
                                    SetName(m_SideOfPierKey, p_PierSide.ToString() + " " + FLIP_TEST_DEC_STEP, value.ToString());
                                    break;
                                }

                            default:
                                {
                                    // MessageBox.Show("FlipTest DECStep Set - Unexpected pier side: " + p_PierSide.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                        }

                        break;
                    }

                default:
                    {
                        // MessageBox.Show("FlipTest DECStep - Unexpected test type: " + p_TestType.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
            }
        }

        // Common
        public bool CommandBlind
        {
            get
            {
                return GetBool(COMMAND_BLIND, COMMAND_BLIND_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, COMMAND_BLIND, value.ToString());
            }
        }

        public bool CommandBlindRaw
        {
            get
            {
                return GetBool(COMMAND_BLIND_RAW, COMMAND_BLIND_RAW_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, COMMAND_BLIND_RAW, value.ToString());
            }
        }

        public bool CommandBool
        {
            get
            {
                return GetBool(COMMAND_BOOL, COMMAND_BOOL_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, COMMAND_BOOL, value.ToString());
            }
        }

        public bool CommandBoolRaw
        {
            get
            {
                return GetBool(COMMAND_BOOL_RAW, COMMAND_BOOL_RAW_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, COMMAND_BOOL_RAW, value.ToString());
            }
        }

        public bool CommandString
        {
            get
            {
                return GetBool(COMMAND_STRING, COMMAND_STRING_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, COMMAND_STRING, value.ToString());
            }
        }

        public bool CommandStringRaw
        {
            get
            {
                return GetBool(COMMAND_STRING_RAW, COMMAND_STRING_RAW_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, COMMAND_STRING_RAW, value.ToString());
            }
        }

        public bool CreateValidationFile
        {
            get
            {
                return GetBool(CREATE_VALIDATION_FILE, CREATE_VALIDATION_FILE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, CREATE_VALIDATION_FILE, value.ToString());
            }
        }

        // Internal
        public DeviceType CurrentDeviceType
        {
            get
            {
                return GetDeviceType(CURRENT_DEVICE_TYPE, CURRENT_DEVICE_TYPE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, CURRENT_DEVICE_TYPE, value.ToString());
            }
        }

        public bool Debug
        {
            get
            {
                return GetBool(DE_BUG, DE_BUG_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DE_BUG, value.ToString());
            }
        }

        public string DeviceCamera
        {
            get
            {
                return GetString(DEVICE_CAMERA, DEVICE_CAMERA_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_CAMERA, value.ToString());
            }
        }

        public string DeviceVideo
        {
            get
            {
                return GetString(DEVICE_VIDEO, DEVICE_VIDEO_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_VIDEO, value.ToString());
            }
        }

        public string DeviceCoverCalibrator
        {
            get
            {
                return GetString(DEVICE_COVER_CALIBRATOR, DEVICE_COVER_CALIBRATOR_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_COVER_CALIBRATOR, value.ToString());
            }
        }

        public string DeviceDome
        {
            get
            {
                return GetString(DEVICE_DOME, DEVICE_DOME_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_DOME, value.ToString());
            }
        }

        public string DeviceFilterWheel
        {
            get
            {
                return GetString(DEVICE_FILTER_WHEEL, DEVICE_FILTER_WHEEL_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_FILTER_WHEEL, value.ToString());
            }
        }

        public string DeviceFocuser
        {
            get
            {
                return GetString(DEVICE_FOCUSER, DEVICE_FOCUSER_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_FOCUSER, value.ToString());
            }
        }

        public string DeviceObservingConditions
        {
            get
            {
                return GetString(DEVICE_OBSERVINGCONDITIONS, DEVICE_OBSERVINGCONDITIONS_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_OBSERVINGCONDITIONS, value.ToString());
            }
        }

        public string DeviceRotator
        {
            get
            {
                return GetString(DEVICE_ROTATOR, DEVICE_ROTATOR_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_ROTATOR, value.ToString());
            }
        }

        public string DeviceSafetyMonitor
        {
            get
            {
                return GetString(DEVICE_SAFETY_MONITOR, DEVICE_SAFETY_MONITOR_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_SAFETY_MONITOR, value.ToString());
            }
        }

        public string DeviceSwitch
        {
            get
            {
                return GetString(DEVICE_SWITCH, DEVICE_SWITCH_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_SWITCH, value.ToString());
            }
        }

        public string DeviceTelescope
        {
            get
            {
                return GetString(DEVICE_TELESCOPE, DEVICE_TELESCOPE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DEVICE_TELESCOPE, value.ToString());
            }
        }

        public string LogFileDirectory
        {
            get
            {
                return GetString(LOG_FILES_DIRECTORY, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + LOG_FILES_DIRECTORY_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, LOG_FILES_DIRECTORY, value.ToString());
            }
        }

        public MessageLevel MessageLevel
        {
            get
            {
                return GetMessageLevel(MESSAGE_LEVEL, MESSAGE_LEVEL_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, MESSAGE_LEVEL, value.ToString());
            }
        }

        public bool TestMethods
        {
            get
            {
                return GetBool(TEST_METHODS, TEST_METHODS_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, TEST_METHODS, value.ToString());
            }
        }

        public bool TestPerformance
        {
            get
            {
                return GetBool(TEST_PERFORMANCE, TEST_PERFORMANCE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, TEST_PERFORMANCE, value.ToString());
            }
        }

        public bool TestProperties
        {
            get
            {
                return GetBool(TEST_PROPERTIES, TEST_PROPERTIES_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, TEST_PROPERTIES, value.ToString());
            }
        }

        public bool UpdateCheck
        {
            get
            {
                return GetBool(UPDATE_CHECK, UPDATE_CHECK_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, UPDATE_CHECK, value.ToString());
            }
        }

        public DateTime UpdateCheckDate
        {
            get
            {
                return GetDate(UPDATE_CHECK_DATE, UPDATE_CHECK_DATE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, UPDATE_CHECK_DATE, value.ToString());
            }
        }

        public bool WarningMessage
        {
            get
            {
                return GetBool(WARNING_MESSAGE, WARNING_MESSAGE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, WARNING_MESSAGE, value.ToString());
            }
        }

        public bool RunAs32Bit
        {
            get
            {
                return GetBool(RUN_AS_THIRTYTWO_BITS, RUN_AS_THIRTYTWO_BITS_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, RUN_AS_THIRTYTWO_BITS, value.ToString());
            }
        }

        public bool InterpretErrorMessages
        {
            get
            {
                return GetBool(INTERPRET_ERROR_MESSAGES, INTERPRET_ERROR_MESSAGES_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, INTERPRET_ERROR_MESSAGES, value.ToString());
            }
        }

        public bool SwitchWarningMessage
        {
            get
            {
                return GetBool(SWITCH_WARNING_MESSAGE, SWITCH_WARNING_MESSAGE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, SWITCH_WARNING_MESSAGE, value.ToString());
            }
        }

        public bool UseDriverAccess
        {
            get
            {
                return GetBool(USE_DRIVERACCESS, USE_DRIVERACCESS_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, USE_DRIVERACCESS, value.ToString());
            }
        }

        public bool DisplayMethodCalls
        {
            get
            {
                return GetBool(DISPLAY_METHOD_CALLS, DISPLAY_METHOD_CALLS_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, DISPLAY_METHOD_CALLS, value.ToString());
            }
        }

        // Switch
        public bool SwitchSet
        {
            get
            {
                return GetBool(SWITCH_SET, SWITCH_SET_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, SWITCH_SET, value.ToString());
            }
        }

        public int SwitchReadDelay
        {
            get
            {
                return GetInteger(SWITCH_READ_DELAY, SWITCH_READ_DELAY_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, SWITCH_READ_DELAY, value.ToString());
            }
        }

        public int SwitchWriteDelay
        {
            get
            {
                return GetInteger(SWITCH_WRITE_DELAY, SWITCH_WRITE_DELAY_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, SWITCH_WRITE_DELAY, value.ToString());
            }
        }

        public int ExtendedSwitchNumberTestRange
        {
            get
            {
                return GetInteger(EXTENDED_SWITCH_NUMBER_TEST_RANGE, EXTENDED_SWITCH_NUMBER_TEST_RANGE_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, EXTENDED_SWITCH_NUMBER_TEST_RANGE, value.ToString());
            }
        }

        // ObservingConditions
        public int ObservingConditionsRetryTime
        {
            get
            {
                return GetInteger(OBSERVINGCONDITIONS_RETRY_TIME, OBSERVINGCONDITIONS_RETRY_TIME_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, OBSERVINGCONDITIONS_RETRY_TIME, value.ToString());
            }
        }

        public int ObservingConditionsMaxRetries
        {
            get
            {
                return GetInteger(OBSERVINGCONDITIONS_RETRIES, OBSERVINGCONDITIONS_MAX_RETRIES_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, OBSERVINGCONDITIONS_RETRIES, value.ToString());
            }
        }

        // Camera
        public int CameraMaxBinX
        {
            get
            {
                return GetInteger(CAMERA_MAX_BIN_X, CAMERA_MAX_BIN_X_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, CAMERA_MAX_BIN_X, value.ToString());
            }
        }

        public int CameraMaxBinY
        {
            get
            {
                return GetInteger(CAMERA_MAX_BIN_Y, CAMERA_MAX_BIN_Y_DEFAULT);
            }

            set
            {
                SetName(m_SettingsKey, CAMERA_MAX_BIN_Y, value.ToString());
            }
        }


        #endregion

        #region Utility Code

        private bool GetBool(RegistryKey p_Key, string p_Name, bool p_DefaultValue)
        {
            var l_Value = default(bool);
            try
            {
                if (p_Key.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = Conversions.ToBoolean(p_Key.GetValue(p_Name));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(p_Key, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetBool Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private bool GetBool(string p_Name, bool p_DefaultValue)
        {
            var l_Value = default(bool);
            try
            {
                if (m_SettingsKey.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = Conversions.ToBoolean(m_SettingsKey.GetValue(p_Name));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetBool Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private string GetString(string p_Name, string p_DefaultValue)
        {
            string l_Value;
            l_Value = "";
            try
            {
                if (m_SettingsKey.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = m_SettingsKey.GetValue(p_Name).ToString();
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetString Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private int GetInteger(string p_Name, int p_DefaultValue)
        {
            var l_Value = default(int);
            try
            {
                if (m_SettingsKey.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = Conversions.ToInteger(m_SettingsKey.GetValue(p_Name));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetInteger Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private int GetInteger(RegistryKey p_Key, string p_Name, int p_DefaultValue)
        {
            var l_Value = default(int);
            Console.WriteLine($"GetInteger {p_Name} {p_DefaultValue}");
            try
            {
                if (p_Key.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = Conversions.ToInteger(p_Key.GetValue(p_Name));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(p_Key, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetInteger Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private double GetDouble(RegistryKey p_Key, string p_Name, double p_DefaultValue)
        {
            var l_Value = default(double);
            Console.WriteLine($"GetDouble {p_Name} {p_DefaultValue}");
            try
            {
                if (p_Key.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = Conversions.ToDouble(p_Key.GetValue(p_Name));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(p_Key, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetDouble Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private DateTime GetDate(string p_Name, DateTime p_DefaultValue)
        {
            var l_Value = default(DateTime);
            try
            {
                if (m_SettingsKey.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = Conversions.ToDate(m_SettingsKey.GetValue(p_Name));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetDate Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private DeviceType GetDeviceType(string p_Name, DeviceType p_DefaultValue)
        {
            var l_Value = default(DeviceType);
            try
            {
                if (m_SettingsKey.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = (DeviceType)Conversions.ToInteger(Enum.Parse(typeof(DeviceType), m_SettingsKey.GetValue(p_Name).ToString(), true));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetDeviceType Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private MessageLevel GetMessageLevel(string p_Name, MessageLevel p_DefaultValue)
        {
            var l_Value = default(MessageLevel);
            try
            {
                if (m_SettingsKey.GetValueKind(p_Name) == RegistryValueKind.String) // Value does exist
                {
                    l_Value = (MessageLevel)Conversions.ToInteger(Enum.Parse(typeof(MessageLevel), m_SettingsKey.GetValue(p_Name).ToString(), true));
                }
            }
            catch (System.IO.IOException) // Value doesn't exist so create it
            {
                SetName(m_SettingsKey, p_Name, p_DefaultValue.ToString());
                l_Value = p_DefaultValue;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"GetMessageLevel Unexpected exception: {ex}");
                l_Value = p_DefaultValue;
            }

            return l_Value;
        }

        private void SetName(RegistryKey p_Key, string p_Name, string p_Value)
        {
            p_Key.SetValue(p_Name, p_Value.ToString(), RegistryValueKind.String);
            p_Key.Flush();
        }
        #endregion

    }
}
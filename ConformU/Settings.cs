//using AlpacaDiscovery;
using ASCOM;
using ASCOM.Alpaca.Discovery;
using ASCOM.Common;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;

namespace ConformU
{
    public class Settings
    {
        // Current version number for this settings class. Only needs to be incremented when there are breaking changes!
        // For example this can be left at its current level when adding new settings that have usable default values.

        // This value is set when values are actually persisted in ConformConfiguration.PersistSettings in order not to overwrite the value that is retrieved from the current settings file when it is read.
        internal const int SETTINGS_COMPATIBILTY_VERSION = 1;

        private AscomDevice ascomDevice = new();
        private ComDevice comDevice = new();
        private DeviceTechnology deviceTechnology = DeviceTechnology.NotSelected;

        private const string NO_DEVICE_SELECTED = "No device selected";
        public Settings() { }

        #region Application behaviour

        public int SettingsCompatibilityVersion { get; set; } = 0; // Default is zero so that versions prior to introduction of the settings compatibility version number can be detected.
        public bool GoHomeOnDeviceSelected { get; set; } = true;
        public double ConnectionTimeout { get; set; } = 2.0;
        public bool RunAs32Bit { get; set; } = false;
        public bool RiskAcknowledged { get; set; } = false;

        #endregion

        #region Conform test configuration

        // Conform application configuration 
        [FullTestValue(false)]
        public bool DisplayMethodCalls { get; set; } = false;

        public bool UpdateCheck { get; set; } = true;
        public int ApplicationPort { get; set; } = 0;

        // Debug output switches
        [FullTestValue(false)]
        public bool Debug { get; set; } = false;

        [FullTestValue(false)]
        public bool TraceDiscovery { get; set; } = false;

        [FullTestValue(false)]
        public bool TraceAlpacaCalls { get; set; } = false;

        // Conformance test configuration 
        [FullTestValue(true)]
        public bool TestProperties { get; set; } = true;

        [FullTestValue(true)]
        public bool TestMethods { get; set; } = true;

        public bool TestPerformance { get; set; } = false;

        /// <summary>
        /// Details of the currently selected Alpaca device
        /// </summary>
        public AscomDevice AlpacaDevice
        {
            get
            {
                return ascomDevice;
            }
            set
            {
                ascomDevice = value;
                DeviceName = ascomDevice.AscomDeviceName;
            }
        }

        public AlpacaConfiguration AlpacaConfiguration { get; set; } = new AlpacaConfiguration();

        /// <summary>
        /// Details of the currently selected Alpaca device
        /// </summary>
        public ComDevice ComDevice
        {
            get
            {
                return comDevice;
            }
            set
            {
                comDevice = value;
                DeviceName = comDevice.DisplayName;
            }
        }

        public ComConfiguration ComConfiguration { get; set; } = new ComConfiguration();

        /// <summary>
        /// Descriptive name of the current device
        /// </summary>
        public string DeviceName { get; private set; } = NO_DEVICE_SELECTED;

        /// <summary>
        /// Technology of the current device: Alpaca or COM
        /// </summary>
        public DeviceTechnology DeviceTechnology
        {
            get
            {
                return deviceTechnology;
            }
            set
            {
                deviceTechnology = value;
                DeviceName = deviceTechnology switch
                {
                    DeviceTechnology.NotSelected => "No device selected",
                    DeviceTechnology.Alpaca => AlpacaDevice.AscomDeviceName,
                    DeviceTechnology.COM => ComDevice.DisplayName,
                    _ => throw new InvalidValueException($"Unknown technology type: {value}"),
                };
            }
        }

        /// <summary>
        /// ASCOM Device type of the current device
        /// </summary>
        [JsonConverter(typeof(JsonStringEnumConverter))]
        public DeviceTypes? DeviceType { get; set; } = null;

        #endregion

        #region Device test configuration

        // Telescope test configuration
        /// <summary>
        /// List of telescope tests that can be enabled / disabled through configuration
        /// </summary>
        public Dictionary<string, bool> TelescopeTests { get; set; } = new()
        {
            { TelescopeTester.TELTEST_CAN_MOVE_AXIS, true },
            { TelescopeTester.TELTEST_PARK_UNPARK, true },
            { TelescopeTester.TELTEST_ABORT_SLEW, true },
            { TelescopeTester.TELTEST_AXIS_RATE, true },
            { TelescopeTester.TELTEST_FIND_HOME, true },
            { TelescopeTester.TELTEST_MOVE_AXIS, true },
            { TelescopeTester.TELTEST_PULSE_GUIDE, true },
            { TelescopeTester.TELTEST_SLEW_TO_COORDINATES, true },
            { TelescopeTester.TELTEST_SLEW_TO_COORDINATES_ASYNC, true },
            { TelescopeTester.TELTEST_SLEW_TO_TARGET, true },
            { TelescopeTester.TELTEST_SLEW_TO_TARGET_ASYNC, true },
            { TelescopeTester.TELTEST_DESTINATION_SIDE_OF_PIER, true },
            { TelescopeTester.TELTEST_SLEW_TO_ALTAZ, true },
            { TelescopeTester.TELTEST_SLEW_TO_ALTAZ_ASYNC, true },
            { TelescopeTester.TELTEST_SYNC_TO_COORDINATES, true },
            { TelescopeTester.TELTEST_SYNC_TO_TARGET, true },
            { TelescopeTester.TELTEST_SYNC_TO_ALTAZ, true }
        };

        [FullTestValue(true)]
        public bool TelescopeExtendedRateOffsetTests { get; set; } = true;

        [FullTestValue(true)]
        public bool TelescopeFirstUseTests { get; set; } = true;

        [FullTestValue(true)]
        public bool TestSideOfPierRead { get; set; } = false;

        [FullTestValue(true)]
        public bool TestSideOfPierWrite { get; set; } = false;

        [FullTestValue(true)]
        public bool TelescopeExtendedPulseGuideTests { get; set; } = true;

        [FullTestValue(true)]
        public bool TelescopeExtendedMoveAxisTests { get; set; } = true;

        [FullTestValue(true)]
        public bool TelescopeExtendedSiteTests { get; set; } = true;

        public double TelescopePulseGuideTolerance { get; set; } = 1; // Arc-seconds
        public double TelescopeSlewTolerance { get; set; } = 10.0; // Arc-seconds
        public int TelescopeMaximumSlewTime { get; set; } = 300; // Seconds

        // Camera test configuration
        public int CameraMaxBinX { get; set; } = 0;
        public int CameraMaxBinY { get; set; } = 0;
        [FullTestValue(true)]
        public bool CameraFirstUseTests { get; set; } = true;
        [FullTestValue(true)]
        public bool CameraTestImageArrayVariant { get; set; } = true;
        public double CameraExposureDuration { get; set; } = 2.0;
        public int CameraXMax { get; set; } = 0;
        public int CameraYMax { get; set; } = 0;
        public int CameraWaitTimeout { get; set; } = 10;

        // Dome test configuration
        public int DomeShutterMovementTimeout { get; set; } = 240;
        public int DomeAzimuthMovementTimeout { get; set; } = 240;
        public int DomeAltitudeMovementTimeout { get; set; } = 240;

        /// <summary>
        /// Dome stabilisation time (seconds)
        /// </summary>
        public int DomeStabilisationWaitTime { get; set; } = 10;
        [FullTestValue(true)]
        public bool DomeOpenShutter { get; set; } = false;
        public double DomeSlewTolerance { get; set; } = 1.0; // Degrees

        // Filter wheel test configuration
        public int FilterWheelTimeout { get; set; } = 30;

        // Focuser test configuration
        public int FocuserTimeout { get; set; } = 60;
        public int FocuserMoveTolerance { get; set; } = 2;

        // ObservingConditions test configuration
        public int ObservingConditionsRetryTime { get; set; } = 1; // Seconds
        public int ObservingConditionsMaxRetries { get; set; } = 5;

        // Rotator test configuration
        public int RotatorTimeout { get; set; } = 60;

        // Switch test configuration
        [FullTestValue(true)]
        public bool SwitchEnableSet { get; set; } = false;
        public int SwitchReadDelay { get; set; } = 500; // Milliseconds
        public int SwitchWriteDelay { get; set; } = 3000; // Milliseconds
        public int SwitchExtendedNumberTestRange { get; set; } = 100;

        #endregion

        #region Internal state variables - not persisted

        /// <summary>
        /// Flag indicating that a long running process is underway
        /// </summary>
        internal bool OperationInProgress { get; set; }

        /// <summary>
        /// Output filename for the results file as specified on the command line
        /// </summary>
        internal string ResultsFileName { get; set; } = "";

        #endregion

    }
}
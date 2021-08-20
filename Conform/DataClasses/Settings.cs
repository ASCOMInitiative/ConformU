using AlpacaDiscovery;
using ASCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;
using static ConformU.ConformConstants;

namespace ConformU
{
    public class Settings
    {
        private AscomDevice ascomDevice = new();
        private ComDevice comDevice = new();
        private DeviceTechnology deviceTechnology = DeviceTechnology.NotSelected;

        private const string NO_DEVICE_SELECTED = "No device selected";
        public Settings() { }

        // Conform application configuration 
        public bool Debug { get; set; } = false;
        public bool DisplayMethodCalls { get; set; } = false;
        public bool UpdateCheck { get; set; } = true;
        public DateTime UpdateDate { get; set; } = DateTime.MinValue;
        public bool WarningMessageDisplayed { get; set; } = false;

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
                switch (deviceTechnology)
                {
                    case DeviceTechnology.NotSelected:
                        DeviceName = "No device selected";
                        break;
                    case DeviceTechnology.Alpaca:
                        DeviceName = AlpacaDevice.AscomDeviceName;
                        break;
                    case DeviceTechnology.COM:
                        DeviceName = ComDevice.DisplayName;
                        break;
                    default:
                        throw new InvalidValueException($"Unknown technology type: {value}");
                }
            }
        }

        /// <summary>
        /// ASCOM Device type of the current device
        /// </summary>
        public DeviceType DeviceType { get; set; } = DeviceType.NoDeviceType;

        // Telescope test configuration
        /// <summary>
        /// List of telescope tests that can be enabled / disabled through configuration
        /// </summary>
        public Dictionary<string, bool> TelescopeTests { get; set; } = new()
        {
            { "CanMoveAxis", true },
            { "Park/Unpark", true },
            { "AbortSlew", true },
            { "AxisRate", true },
            { "FindHome", true },
            { "MoveAxis", true },
            { "PulseGuide", true },
            { "SlewToCoordinates", true },
            { "SlewToCoordinatesAsync", true },
            { "SlewToTarget", true },
            { "SlewToTargetAsync", true },
            { "DestinationSideOfPier", true },
            { "SlewToAltAz", true },
            { "SlewToAltAzAsync", true },
            { "SyncToCoordinates", true },
            { "SyncToTarget", true },
            { "SyncToAltAz", true }
        };

        // Camera test configuration
        public int CameraMaxBinX { get; set; }
        public int CameraMaxBinY { get; set; }

        // Conformance test configuration 
        public bool TestProperties { get; set; } = true;
        public bool TestMethods { get; set; } = true;
        public bool TestPerformance { get; set; } = false;
        public bool TestSideOfPierRead { get; set; } = false;
        public bool TestSideOfPierWrite { get; set; } = false;

        // Dome test configuration
        public int DomeShutterTimeout { get; set; } = 240;
        public int DomeAzimuthTimeout { get; set; } = 240;
        public int DomeAltitudeTimeout { get; set; } = 240;
        public int DomeStabilisationWaitTime { get; set; } = 10;
        public bool DomeOpenShutter { get; set; } = false;

        // ObservingConditions test configuration
        public int ObservingConditionsRetryTime { get; set; } = 1;
        public int ObservingConditionsMaxRetries { get; set; } = 5;

        // Switch test configuration
        public bool SwitchEnableSet { get; set; } = false;
        public int SwitchReadDelay { get; set; } = 500;
        public int SwitchWriteDelay { get; set; } = 3000;
        public int SwitchExtendedNumberTestRange { get; set; } = 100;

        internal CancellationToken ApplicationCancellationToken { get; set; }
    }
}
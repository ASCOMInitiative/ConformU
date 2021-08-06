using AlpacaDiscovery;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using static ConformU.ConformConstants;

namespace ConformU
{
    public class Settings
    {

        public  const string NO_DEVICE_SELECTED = "No device selected";
        public Settings()
        {
            TelescopeTests.Add("CanMoveAxis", true);
            TelescopeTests.Add("Park/Unpark", true);
            TelescopeTests.Add("AbortSlew", true);
            TelescopeTests.Add("AxisRate", true);
            TelescopeTests.Add("FindHome", true);
            TelescopeTests.Add("MoveAxis", true);
            TelescopeTests.Add("PulseGuide", true);
            TelescopeTests.Add("SlewToCoordinates", true);
            TelescopeTests.Add("SlewToCoordinatesAsync", true);
            TelescopeTests.Add("SlewToTarget", true);
            TelescopeTests.Add("SlewToTargetAsync", true);
            TelescopeTests.Add("DestinationSideOfPier", true);
            TelescopeTests.Add("SlewToAltAz", true);
            TelescopeTests.Add("SlewToAltAzAsync", true);
            TelescopeTests.Add("SyncToCoordinates", true);
            TelescopeTests.Add("SyncToTarget", true);
            TelescopeTests.Add("SyncToAltAz", true);
        }

        /// <summary>
        /// Details of the currently selected Alpaca device
        /// </summary>
        public AscomDevice CurrentAlpacaDevice { get; set; } = new();
        /// <summary>
        /// Details of the currently selected Alpaca device
        /// </summary>
        public ComDevice CurrentComDevice { get; set; } = new();
        /// <summary>
        /// Descriptive name of the current device
        /// </summary>
        public string CurrentDeviceName { get; set; } = NO_DEVICE_SELECTED;
        /// <summary>
        /// ProgID of the current device (COM only)
        /// </summary>
        public string CurrentDeviceProgId { get; set; } = "";
        /// <summary>
        /// Technology of the current device: Alpaca or COM
        /// </summary>
        public string CurrentDeviceTechnology { get; set; } = ConformConstants.TECHNOLOGY_ALPACA;
        /// <summary>
        /// ASCOM Device type of the current device
        /// </summary>
        public DeviceType CurrentDeviceType { get; set; } = DeviceType.Telescope;

        public Dictionary<string, bool> TelescopeTests { get; set; } = new();

        public int CameraTestMaxBinX { get; set; }
        public int CameraTestMaxBinY { get; set; }

        public bool Debug { get; set; } = false;
        public string DeviceCamera { get; set; } = "";
        public string DeviceCoverCalibrator { get; set; }
        public string DeviceDome { get; set; } = "";
        public string DeviceFilterWheel { get; set; } = "";
        public string DeviceFocuser { get; set; } = "";
        public string DeviceObservingConditions { get; set; } = "";
        public string DeviceRotator { get; set; } = "";
        public string DeviceSafetyMonitor { get; set; } = "";
        public string DeviceSwitch { get; set; } = "";
        public string DeviceTelescope { get; set; } = "";
        public string DeviceVideo { get; set; } = "";
        public bool DisplayMethodCalls { get; set; } = false;
        public string LogFileFolder { get; set; } = "";
        public bool TestProperties { get; set; } = true;
        public bool TestMethods { get; set; } = true;
        public bool TestPerformance { get; set; } = false;
        public bool TestSideOfPierRead { get; set; } = false;
        public bool TestSideOfPierWrite { get; set; } = false;
        public bool UpdateCheck { get; set; } = true;
        public DateTime UpdateDate { get; set; } = DateTime.MinValue;
        public bool WarningMessageDisplayed { get; set; } = false;
        public bool UseDriverAccess { get; set; } = false;
       
        public bool InterpretErrorMessages { get; set; } = false;

    }
}
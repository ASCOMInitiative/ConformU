﻿using AlpacaDiscovery;
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
        AscomDevice ascomDevice = new();
        ComDevice comDevice = new();

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

        /// <summary>
        /// Descriptive name of the current device
        /// </summary>
        public string DeviceName { get; private set; } = NO_DEVICE_SELECTED;

        /// <summary>
        /// Technology of the current device: Alpaca or COM
        /// </summary>
        public DeviceTechnology DeviceTechnology { get; set; } = DeviceTechnology.Alpaca;

        /// <summary>
        /// ASCOM Device type of the current device
        /// </summary>
        public DeviceType DeviceType { get; set; } = DeviceType.Telescope;

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
        public int CameraTestMaxBinX { get; set; }
        public int CameraTestMaxBinY { get; set; }

        // Conformance test configuration 
        public bool TestProperties { get; set; } = true;
        public bool TestMethods { get; set; } = true;
        public bool TestPerformance { get; set; } = false;
        public bool TestSideOfPierRead { get; set; } = false;
        public bool TestSideOfPierWrite { get; set; } = false;

    }
}
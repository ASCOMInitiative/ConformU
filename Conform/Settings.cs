using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

public class Settings
{
    public bool Debug { get; set; } = false;
    public string CurrentDeviceType { get; set; } = "";
    public string CurrentDeviceTechnology { get; set; } = "Alpaca";
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

}

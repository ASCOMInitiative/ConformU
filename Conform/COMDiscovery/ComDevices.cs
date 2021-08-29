using ASCOM.Standard.Interfaces;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class ComDevices
    {
        public ComDevices()
        {
        }

        public static Dictionary<string, string> GetRegisteredDrivers(string requiredDeviceType, ConformLogger TL)
        {
            Dictionary<string, string> result = new();
            if (!OperatingSystem.IsWindows()) throw new InvalidOperationException("Conform.ComDevices.GetRegisteredDrivers can only be used on a Windows operating system");

            RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32);
            RegistryKey profile = baseKey.OpenSubKey(Globals.ASCOM_PROFILE_KEY);

            string[] keyNames = profile.GetSubKeyNames();
            foreach (string keyName in keyNames)
            {
                TL?.LogMessage("GetRegisteredDrivers", $"Found key name {keyName}");
                if (keyName.ToUpperInvariant() == $"{requiredDeviceType} Drivers".ToUpperInvariant())
                {
                    TL?.LogMessage("GetRegisteredDrivers", $"Found DRIVERS of type {keyName}");
                    RegistryKey drivers = profile.OpenSubKey($"{requiredDeviceType} Drivers");
                    string[] driverNames = drivers.GetSubKeyNames();
                    foreach (string driverName in driverNames)
                    {
                        TL?.LogMessage("GetRegisteredDrivers", $"Found Driver: {driverName}");
                        RegistryKey driver = drivers.OpenSubKey(driverName);
                        string description = (string)driver.GetValue("");
                        result.Add(driverName, description);
                        TL?.LogMessage("GetRegisteredDrivers", $"Added Driver: {driverName} {description}");
                    }

                }
            }

            return result;
        }
    }
}
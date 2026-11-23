using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    internal static class DeviceCapabilitySupport
    {
        internal static bool HasConnectAndDeviceState(DeviceTypes? deviceType, int interfaceVersion)
        {
            if (deviceType == DeviceTypes.Aligner)
                return interfaceVersion == 1;

            return DeviceCapabilities.HasConnectAndDeviceState(deviceType, interfaceVersion);
        }

        internal static bool IsPlatform7OrLater(DeviceTypes? deviceType, int interfaceVersion)
        {
            if (deviceType == DeviceTypes.Aligner)
                return interfaceVersion == 1;

            return DeviceCapabilities.IsPlatform7OrLater(deviceType, interfaceVersion);
        }

        internal static bool TryGetLatestPlatform6Interface(DeviceTypes deviceType, out short interfaceVersion)
        {
            return DeviceCapabilities.LatestPlatform6Interface.TryGetValue(deviceType, out interfaceVersion);
        }
    }
}

//using AlpacaDiscovery;
using ASCOM.Alpaca.Discovery;

namespace ConformU
{
    public class AlpacaDeviceListItem
    {
        public AlpacaDeviceListItem(string displayName, AscomDevice ascomDevice)
        {
            DisplayName = displayName;
            AscomDevice = ascomDevice;
        }

        public string DisplayName { get; set; }
        public AscomDevice AscomDevice { get; set; }
    }


}

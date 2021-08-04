using AlpacaDiscovery;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

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

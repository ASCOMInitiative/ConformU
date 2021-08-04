using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class ComDeviceListItem
    {
        public ComDeviceListItem(string displayName, ComDevice comDevice)
        {
            DisplayName = displayName;
            ComDevice = comDevice;
        }

        public string DisplayName { get; set; }
        public ComDevice ComDevice { get; set; }
    }
}

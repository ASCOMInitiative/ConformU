using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class ComDevice
    {
        public ComDevice(string displayName, string progId)
        {
            DisplayName = displayName;
            ProgId = progId;
        }

        public string DisplayName { get; set; }
        public string ProgId { get; set; }
    }
}
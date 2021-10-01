using ASCOM.Common.DeviceInterfaces;
/* Unmerged change from project 'ConformU (net5.0)'
Before:
using System.Runtime.InteropServices;
using ASCOM.Alpaca.Clients;
After:
using System.Runtime.InteropServices;
*/


namespace ConformU
{
    public class SafetyMonitorFacade : FacadeBaseClass, ISafetyMonitor
    {

        // Create the test device in the facade base class
        public SafetyMonitorFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public bool IsSafe => (bool)FunctionNoParameters(() => driver.IsSafe);

        #endregion

    }
}

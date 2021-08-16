using ASCOM;
using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using ASCOM.Standard.AlpacaClients;

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

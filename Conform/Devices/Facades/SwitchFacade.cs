using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class SwitchFacade : FacadeBaseClass, ISwitchV2
    {
        // Create the test device in the facade base class
        public SwitchFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public short MaxSwitch => driver.MaxSwitch;

        public bool CanWrite(short id)
        {
            return driver.CanWrite(id);
        }

        public bool GetSwitch(short id)
        {
            return driver.GetSwitch(id);
        }

        public string GetSwitchDescription(short id)
        {
            return driver.GetSwitchDescription(id);
        }

        public string GetSwitchName(short id)
        {
            return driver.GetSwitchName(id);
        }

        public double GetSwitchValue(short id)
        {
            return driver.GetSwitchValue(id);
        }

        public double MaxSwitchValue(short id)
        {
            return driver.MaxSwitchValue(id);
        }

        public double MinSwitchValue(short id)
        {
            return driver.MinSwitchValue(id);
        }

        public void SetSwitch(short id, bool state)
        {
            driver.SetSwitch(id, state);
        }

        public void SetSwitchName(short id, string name)
        {
            driver.SetSwitchName(id, name);
        }

        public void SetSwitchValue(short id, double value)
        {
            driver.SetSwitchValue(id, value);
        }

        public double SwitchStep(short id)
        {
            return driver.SwitchStep(id);
        }

        #endregion

    }
}

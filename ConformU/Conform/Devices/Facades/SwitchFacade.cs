using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class SwitchFacade : FacadeBaseClass, ISwitchV2
    {
        // Create the test device in the facade base class
        public SwitchFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public short MaxSwitch
        {
            get
            {
                return (short)FunctionNoParameters(() => driver.MaxSwitch);
            }
        }

        public bool CanWrite(short id)
        {
            return Function1Parameter<bool>((i) => driver.CanWrite(i), id);
        }

        public bool GetSwitch(short id)
        {
            return Function1Parameter<bool>((i) => driver.GetSwitch(i), id);
        }

        public string GetSwitchDescription(short id)
        {
            return Function1Parameter<string>((i) => driver.GetSwitchDescription(i), id);
        }

        public string GetSwitchName(short id)
        {
            return Function1Parameter<string>((i) => driver.GetSwitchName(i), id);
        }

        public double GetSwitchValue(short id)
        {
            return Function1Parameter<double>((i) => driver.GetSwitchValue(i), id);
        }

        public double MaxSwitchValue(short id)
        {
            return Function1Parameter<double>((i) => driver.MaxSwitchValue(i), id);
        }

        public double MinSwitchValue(short id)
        {
            return Function1Parameter<double>((i) => driver.MinSwitchValue(i), id);
        }

        public void SetSwitch(short id, bool state)
        {
            Method2Parameters((i, j) => driver.SetSwitch(i, j), id, state);
        }

        public void SetSwitchName(short id, string name)
        {
            Method2Parameters((i, j) => driver.SetSwitchName(i, j), id, name);
        }

        public void SetSwitchValue(short id, double value)
        {
            Method2Parameters((i, j) => driver.SetSwitchValue(i, j), id, value);
        }

        public double SwitchStep(short id)
        {
            return Function1Parameter<double>((i) => driver.SwitchStep(i), id);
        }

        #endregion

    }
}

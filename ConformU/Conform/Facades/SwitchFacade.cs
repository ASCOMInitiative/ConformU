using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class SwitchFacade : FacadeBaseClass, ISwitchV3
    {
        // Create the test device in the facade base class
        public SwitchFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region ISwitchV2 implementation

        public short MaxSwitch
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.MaxSwitch);
            }
        }

        public bool CanWrite(short id)
        {
            return Function1Parameter<bool>((i) => Driver.CanWrite(i), id);
        }

        public bool GetSwitch(short id)
        {
            return Function1Parameter<bool>((i) => Driver.GetSwitch(i), id);
        }

        public string GetSwitchDescription(short id)
        {
            return Function1Parameter<string>((i) => Driver.GetSwitchDescription(i), id);
        }

        public string GetSwitchName(short id)
        {
            return Function1Parameter<string>((i) => Driver.GetSwitchName(i), id);
        }

        public double GetSwitchValue(short id)
        {
            return Function1Parameter<double>((i) => Driver.GetSwitchValue(i), id);
        }

        public double MaxSwitchValue(short id)
        {
            return Function1Parameter<double>((i) => Driver.MaxSwitchValue(i), id);
        }

        public double MinSwitchValue(short id)
        {
            return Function1Parameter<double>((i) => Driver.MinSwitchValue(i), id);
        }

        public void SetSwitch(short id, bool state)
        {
            Method2Parameters((i, j) => Driver.SetSwitch(i, j), id, state);
        }

        public void SetSwitchName(short id, string name)
        {
            Method2Parameters((i, j) => Driver.SetSwitchName(i, j), id, name);
        }

        public void SetSwitchValue(short id, double value)
        {
            Method2Parameters((i, j) => Driver.SetSwitchValue(i, j), id, value);
        }

        public double SwitchStep(short id)
        {
            return Function1Parameter<double>((i) => Driver.SwitchStep(i), id);
        }

        #endregion

        #region ISwitchV3 implementation

        public bool CanAsync(short id)
        {
            return Function1Parameter<bool>((i) => Driver.CanAsync(i), id);
        }

        public void SetAsync(short id, bool state)
        {
            Method2Parameters((i, j) => Driver.SetAsync(i, j), id, state);
        }

        public void SetAsyncValue(short id, double value)
        {
            Method2Parameters((i, j) => Driver.SetAsyncValue(i, j), id, value);
        }

        public void CancelAsync(short id)
        {
            Method1Parameter((i) => Driver.CancelAsync(i), id);
        }

        public bool StateChangeComplete(short id)
        {
            return Function1Parameter<bool>((i) => Driver.StateChangeComplete(i), id);
        }

        #endregion

    }
}

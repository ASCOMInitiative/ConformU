using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class FilterWheelFacade : FacadeBaseClass, IFilterWheelV2
    {
        // Create the test device in the facade base class
        public FilterWheelFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public int[] FocusOffsets
        {
            get
            {
                return FunctionNoParameters<int[]>(() => driver.FocusOffsets);
            }
        }

        public string[] Names
        {
            get
            {
                return FunctionNoParameters<string[]>(() => driver.Names);
            }
        }

        public short Position { get => FunctionNoParameters<short>(() => driver.Position); set => Method1Parameter((i) => driver.Position = i, value); }

        #endregion

    }
}

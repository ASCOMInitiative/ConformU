using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class FilterWheelFacade : FacadeBaseClass, IFilterWheelV3
    {
        // Create the test device in the facade base class
        public FilterWheelFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public int[] FocusOffsets
        {
            get
            {
                return FunctionNoParameters<int[]>(() => Driver.FocusOffsets);
            }
        }

        public string[] Names
        {
            get
            {
                return FunctionNoParameters<string[]>(() => Driver.Names);
            }
        }

        public short Position { get => FunctionNoParameters<short>(() => Driver.Position); set => Method1Parameter((i) => Driver.Position = i, value); }

        #endregion

    }
}

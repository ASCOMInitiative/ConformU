using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class FilterWheelFacade : FacadeBaseClass, IFilterWheelV2
    {
        // Create the test device in the facade base class
        public FilterWheelFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public int[] FocusOffsets => (int[])FunctionNoParameters(() => driver.FocusOffsets);

        public string[] Names => (string[])FunctionNoParameters(() => driver.Names);

        public short Position { get => (short)FunctionNoParameters(() => driver.Position); set => Method1Parameter((i) => driver.Position = i, value); }

        #endregion

    }
}

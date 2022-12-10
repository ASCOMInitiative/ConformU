using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class FocuserFacade : FacadeBaseClass, IFocuserV3
    {
        // Create the test device in the facade base class
        public FocuserFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public bool Absolute => (bool)FunctionNoParameters(() => driver.Absolute);

        public bool IsMoving => (bool)FunctionNoParameters(() => driver.IsMoving);

        public bool Link { get => (bool)FunctionNoParameters(() => driver.Link); set => Method1Parameter((i) => driver.Link = i, value); }

        public int MaxIncrement => (int)FunctionNoParameters(() => driver.MaxIncrement);

        public int MaxStep => (int)FunctionNoParameters(() => driver.MaxStep);

        public int Position => (int)FunctionNoParameters(() => driver.Position);

        public double StepSize => (double)FunctionNoParameters(() => driver.StepSize);

        public bool TempComp { get => (bool)FunctionNoParameters(() => driver.TempComp); set => Method1Parameter((i) => driver.TempComp = i, value); }

        public bool TempCompAvailable => (bool)FunctionNoParameters(() => driver.TempCompAvailable);

        public double Temperature => (double)FunctionNoParameters(() => driver.Temperature);

        public void Halt()
        {
            MethodNoParameters(() => driver.Halt());
        }

        public void Move(int Position)
        {
            Method1Parameter((i) => driver.Move(i), Position);
        }

        #endregion

    }
}

using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class FocuserFacade : FacadeBaseClass, IFocuserV3
    {
        // Create the test device in the facade base class
        public FocuserFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public bool Absolute
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.Absolute);
            }
        }

        public bool IsMoving
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.IsMoving);
            }
        }

        public bool Link { get => FunctionNoParameters<bool>(() => driver.Link); set => Method1Parameter((i) => driver.Link = i, value); }

        public int MaxIncrement
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.MaxIncrement);
            }
        }

        public int MaxStep
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.MaxStep);
            }
        }

        public int Position
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.Position);
            }
        }

        public double StepSize
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.StepSize);
            }
        }

        public bool TempComp { get => FunctionNoParameters<bool>(() => driver.TempComp); set => Method1Parameter((i) => driver.TempComp = i, value); }

        public bool TempCompAvailable
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.TempCompAvailable);
            }
        }

        public double Temperature
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Temperature);
            }
        }

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

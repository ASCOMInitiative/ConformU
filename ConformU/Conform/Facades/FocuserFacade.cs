using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class FocuserFacade : FacadeBaseClass, IFocuserV4
    {
        // Create the test device in the facade base class
        public FocuserFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public bool Absolute
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.Absolute);
            }
        }

        public bool IsMoving
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.IsMoving);
            }
        }

        public bool Link { get => FunctionNoParameters<bool>(() => Driver.Link); set => Method1Parameter((i) => Driver.Link = i, value); }

        public int MaxIncrement
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.MaxIncrement);
            }
        }

        public int MaxStep
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.MaxStep);
            }
        }

        public int Position
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.Position);
            }
        }

        public double StepSize
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.StepSize);
            }
        }

        public bool TempComp { get => FunctionNoParameters<bool>(() => Driver.TempComp); set => Method1Parameter((i) => Driver.TempComp = i, value); }

        public bool TempCompAvailable
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.TempCompAvailable);
            }
        }

        public double Temperature
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.Temperature);
            }
        }

        public void Halt()
        {
            MethodNoParameters(() => Driver.Halt());
        }

        public void Move(int position)
        {
            Method1Parameter((i) => Driver.Move(i), position);
        }

        #endregion

    }
}

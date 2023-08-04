using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class RotatorFacade : FacadeBaseClass, IRotatorV4
    {
        // Create the test device in the facade base class
        public RotatorFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public bool CanReverse
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanReverse);
            }
        }

        public bool IsMoving
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.IsMoving);
            }
        }

        public float Position
        {
            get
            {
                return FunctionNoParameters<float>(() => Driver.Position);
            }
        }

        public bool Reverse { get => FunctionNoParameters<bool>(() => Driver.Reverse); set => Method1Parameter((i) => Driver.Reverse = i, value); }

        public float StepSize
        {
            get
            {
                return FunctionNoParameters<float>(() => Driver.StepSize);
            }
        }

        public float TargetPosition
        {
            get
            {
                return FunctionNoParameters<float>(() => Driver.TargetPosition);
            }
        }

        public float MechanicalPosition
        {
            get
            {
                return FunctionNoParameters<float>(() => Driver.MechanicalPosition);
            }
        }

        public void Halt()
        {
            MethodNoParameters(() => Driver.Halt());
        }

        public void Move(float position)
        {
            Method1Parameter((i) => Driver.Move(i), position);
        }

        public void MoveAbsolute(float position)
        {
            Method1Parameter((i) => Driver.MoveAbsolute(i), position);
        }

        public void MoveMechanical(float position)
        {
            Method1Parameter((i) => Driver.MoveMechanical(i), position);
        }

        public void Sync(float position)
        {
            Method1Parameter((i) => Driver.Sync(i), position);
        }
    }
}

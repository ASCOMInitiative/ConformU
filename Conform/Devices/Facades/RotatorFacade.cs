using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class RotatorFacade : FacadeBaseClass, IRotatorV3
    {
        // Create the test device in the facade base class
        public RotatorFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public bool CanReverse => (bool)FunctionNoParameters(() => driver.CanReverse);

        public bool IsMoving => (bool)FunctionNoParameters(() => driver.IsMoving);

        public float Position => (float)FunctionNoParameters(() => driver.Position);

        public bool Reverse { get => (bool)FunctionNoParameters(() => driver.Reverse); set => Method1Parameter((i) => driver.Reverse = i, value); }

        public float StepSize => (float)FunctionNoParameters(() => driver.StepSize);

        public float TargetPosition => (float)FunctionNoParameters(() => driver.TargetPosition);

        public float MechanicalPosition => (float)FunctionNoParameters(() => driver.MechanicalPosition);

        public void Halt()
        {
            MethodNoParameters(() => driver.Halt());
        }

        public void Move(float Position)
        {
            Method1Parameter((i) => driver.Move(i), Position);
        }

        public void MoveAbsolute(float Position)
        {
            Method1Parameter((i) => driver.MoveAbsolute(i), Position);
        }

        public void MoveMechanical(float Position)
        {
            Method1Parameter((i) => driver.MoveMechanical(i), Position);
        }

        public void Sync(float Position)
        {
            Method1Parameter((i) => driver.Sync(i), Position);
        }
    }
}

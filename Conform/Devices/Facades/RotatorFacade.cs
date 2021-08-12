using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class RotatorFacade : FacadeBaseClass, IRotatorV3
    {
        // Create the test device in the facade base class
        public RotatorFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public bool CanReverse => driver.CanReverse;

        public bool IsMoving => driver.IsMoving;

        public float Position => driver.Position;

        public bool Reverse { get => driver.Reverse; set => driver.Reverse = value; }

        public float StepSize => driver.StepSize;

        public float TargetPosition => driver.TargetPosition;

        public float MechanicalPosition => driver.MechanicalPosition;

        public void Halt()
        {
            driver.Halt();
        }

        public void Move(float Position)
        {
            driver.Move(Position);
        }

        public void MoveAbsolute(float Position)
        {
            driver.MoveAbsolute(Position);
        }

        public void MoveMechanical(float Position)
        {
            driver.MoveMechanical(Position);
        }

        public void Sync(float Position)
        {
            driver.Sync(Position);
        }
    }
}

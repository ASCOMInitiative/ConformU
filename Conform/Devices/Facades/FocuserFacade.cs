using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class FocuserFacade : FacadeBaseClass, IFocuserV3
    {
        // Create the test device in the facade base class
        public FocuserFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public bool Absolute => driver.Absolute;

        public bool IsMoving => driver.IsMoving;

        public bool Link { get => driver.Link; set => driver.Link = value; }

        public int MaxIncrement => driver.MaxIncrement;

        public int MaxStep => driver.MaxStep;

        public int Position => driver.Position;

        public double StepSize => driver.StepSize;

        public bool TempComp { get => driver.TempComp; set => driver.TempComp = value; }

        public bool TempCompAvailable => driver.TempCompAvailable;

        public double Temperature => driver.Temperature;

        public void Halt()
        {
            driver.Halt();
        }

        public void Move(int Position)
        {
            driver.Move(Position);
        }

        #endregion

    }
}

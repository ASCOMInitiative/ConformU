using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class CoverCalibratorFacade : FacadeBaseClass, ICoverCalibratorV1
    {
        // Create the test device in the facade base class
        public CoverCalibratorFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public CoverStatus CoverState => driver.CoverState;

        public CalibratorStatus CalibratorState => driver.CalibratorState;

        public int Brightness => driver.Brightness;

        public int MaxBrightness => driver.MaxBrightness;

        public void CalibratorOff()
        {
            driver.CalibratorOff();
        }

        public void CalibratorOn(int Brightness)
        {
            driver.CalibratorOn(Brightness);
        }

        public void CloseCover()
        {
            driver.CloseCover();
        }

        public void HaltCover()
        {
            driver.HaltCover();
        }

        public void OpenCover()
        {
            driver.OpenCover();
        }

        #endregion

    }
}

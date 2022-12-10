using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class CoverCalibratorFacade : FacadeBaseClass, ICoverCalibratorV1
    {
        // Create the test device in the facade base class
        public CoverCalibratorFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public CoverStatus CoverState => (CoverStatus)FunctionNoParameters(() => driver.CoverState);

        public CalibratorStatus CalibratorState => (CalibratorStatus)FunctionNoParameters(() => driver.CalibratorState);

        public int Brightness => (int)FunctionNoParameters(() => driver.Brightness);

        public int MaxBrightness => (int)FunctionNoParameters(() => driver.MaxBrightness);

        public void CalibratorOff()
        {
            MethodNoParameters(() => driver.CalibratorOff());
        }

        public void CalibratorOn(int Brightness)
        {
            Method1Parameter((i) => driver.CalibratorOn(i), Brightness);
        }

        public void CloseCover()
        {
            MethodNoParameters(() => driver.CloseCover());
        }

        public void HaltCover()
        {
            MethodNoParameters(() => driver.HaltCover());
        }

        public void OpenCover()
        {
            MethodNoParameters(() => driver.OpenCover());
        }

        #endregion

    }
}

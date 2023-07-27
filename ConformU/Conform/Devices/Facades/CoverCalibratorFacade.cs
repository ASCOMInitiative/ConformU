using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class CoverCalibratorFacade : FacadeBaseClass, ICoverCalibratorV2
    {
        // Create the test device in the facade base class
        public CoverCalibratorFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region ICoverCalibratorV1  Interface implementation

        public CoverStatus CoverState
        {
            get
            {
                return (CoverStatus)FunctionNoParameters<object>(() => driver.CoverState);
            }
        }

        public CalibratorStatus CalibratorState
        {
            get
            {
                return (CalibratorStatus)FunctionNoParameters<object>(() => driver.CalibratorState);
            }
        }

        public int Brightness
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.Brightness);
            }
        }

        public int MaxBrightness
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.MaxBrightness);
            }
        }

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

        #region ICoverCalibratorV2 implementation

        public bool CalibratorChanging
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CalibratorChanging);
            }
        }

        public bool CoverMoving
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CoverMoving);
            }
        }

        #endregion
    }
}

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
                return (CoverStatus)FunctionNoParameters<object>(() => Driver.CoverState);
            }
        }

        public CalibratorStatus CalibratorState
        {
            get
            {
                return (CalibratorStatus)FunctionNoParameters<object>(() => Driver.CalibratorState);
            }
        }

        public int Brightness
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.Brightness);
            }
        }

        public int MaxBrightness
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.MaxBrightness);
            }
        }

        public void CalibratorOff()
        {
            MethodNoParameters(() => Driver.CalibratorOff());
        }

        public void CalibratorOn(int brightness)
        {
            Method1Parameter((i) => Driver.CalibratorOn(i), brightness);
        }

        public void CloseCover()
        {
            MethodNoParameters(() => Driver.CloseCover());
        }

        public void HaltCover()
        {
            MethodNoParameters(() => Driver.HaltCover());
        }

        public void OpenCover()
        {
            MethodNoParameters(() => Driver.OpenCover());
        }

        #endregion

        #region ICoverCalibratorV2 implementation

        public bool CalibratorReady
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CalibratorChanging);
            }
        }

        public bool CoverMoving
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CoverMoving);
            }
        }

        #endregion
    }
}

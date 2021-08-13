using ASCOM.Standard.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class CameraFacade : FacadeBaseClass, ICameraV3
    {

        // Create the test device in the facade base class
        public CameraFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public short BinX { get => driver.BinX; set => driver.BinX = value; }
        public short BinY { get => driver.BinY; set => driver.BinY = value; }

        public CameraState CameraState => (CameraState)driver.CameraState;

        public int CameraXSize => driver.CameraXSize;

        public int CameraYSize => driver.CameraYSize;

        public bool CanAbortExposure => driver.CanAbortExposure;

        public bool CanAsymmetricBin => driver.CanAsymmetricBin;

        public bool CanGetCoolerPower => driver.CanGetCoolerPower;

        public bool CanPulseGuide => driver.CanPulseGuide;

        public bool CanSetCCDTemperature => driver.CanSetCCDTemperature;

        public bool CanStopExposure => driver.CanStopExposure;

        public double CCDTemperature => driver.CCDTemperature;

        public bool CoolerOn { get => driver.CoolerOn; set => driver.CoolerOn = value; }

        public double CoolerPower => driver.CoolerPower;

        public double ElectronsPerADU => driver.ElectronsPerADU;

        public double FullWellCapacity => driver.FullWellCapacity;

        public bool HasShutter => driver.HasShutter;

        public double HeatSinkTemperature => driver.HeatSinkTemperature;

        public object ImageArray => driver.ImageArray;

        public object ImageArrayVariant => driver.ImageArrayVariant;

        public bool ImageReady => driver.ImageReady;

        public bool IsPulseGuiding => driver.IsPulseGuiding;

        public double LastExposureDuration => driver.LastExposureDuration;

        public string LastExposureStartTime => driver.LastExposureStartTime;

        public int MaxADU => driver.MaxADU;

        public short MaxBinX => driver.MaxBinX;

        public short MaxBinY => driver.MaxBinY;

        public int NumX { get => driver.NumX; set => driver.NumX = value; }
        public int NumY { get => driver.NumY; set => driver.NumY = value; }

        public double PixelSizeX => driver.PixelSizeX;

        public double PixelSizeY => driver.PixelSizeY;

        public double SetCCDTemperature { get => driver.SetCCDTemperature; set => driver.SetCCDTemperature = value; }
        public int StartX { get => driver.StartX; set => driver.StartX = value; }
        public int StartY { get => driver.StartY; set => driver.StartY = value; }

        public short BayerOffsetX => driver.BayerOffsetX;

        public short BayerOffsetY => driver.BayerOffsetY;

        public bool CanFastReadout => driver.CanFastReadout;

        public double ExposureMax => driver.ExposureMax;

        public double ExposureMin => driver.ExposureMin;

        public double ExposureResolution => driver.ExposureResolution;

        public bool FastReadout { get => driver.FastReadout; set => driver.FastReadout = value; }
        public short Gain { get => driver.Gain; set => driver.Gain = value; }

        public short GainMax => driver.GainMax;

        public short GainMin => driver.GainMin;

        public IList<string> Gains
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in driver.Gains)
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public short PercentCompleted => driver.PercentCompleted;

        public short ReadoutMode { get => driver.ReadoutMode; set => driver.ReadoutMode = value; }

        public IList<string> ReadoutModes
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in driver.ReadoutModes)
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public string SensorName => driver.SensorName;

        public SensorType SensorType => (SensorType)driver.SensorType;

        public int Offset { get => driver.Offset; set => driver.Offset = value; }

        public int OffsetMax => driver.OffsetMax;

        public int OffsetMin => driver.OffsetMin;

        public IList<string> Offsets
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in driver.Offsets)
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public double SubExposureDuration { get => driver.SubExposureDuration; set => driver.SubExposureDuration = value; }

        public void AbortExposure()
        {
            driver.AbortExposure();
        }

        public void PulseGuide(GuideDirection Direction, int Duration)
        {
            driver.PulseGuide(Direction, Duration);
        }

        public void StartExposure(double Duration, bool Light)
        {
            driver.StartExposure(Duration, Light);
        }

        public void StopExposure()
        {
            driver.StopExposure();
        }

        #region Interface implementation

        #endregion


    }
}

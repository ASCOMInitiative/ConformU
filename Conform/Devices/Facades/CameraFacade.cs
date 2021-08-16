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

        public short BinX { get => (short)FunctionNoParameters(() => driver.BinX); set => Method1Parameter((i) => driver.BinX = i, value); }

        public short BinY { get => (short)FunctionNoParameters(() => driver.BinY); set => Method1Parameter((i) => driver.BinY = i, value); }

        public CameraState CameraState => (CameraState)FunctionNoParameters(() => driver.CameraState);

        public int CameraXSize => (int)FunctionNoParameters(() => driver.CameraXSize);

        public int CameraYSize => (int)FunctionNoParameters(() => driver.CameraYSize);

        public bool CanAbortExposure => (bool)FunctionNoParameters(() => driver.CanAbortExposure);

        public bool CanAsymmetricBin => (bool)FunctionNoParameters(() => driver.CanAsymmetricBin);

        public bool CanGetCoolerPower => (bool)FunctionNoParameters(() => driver.CanGetCoolerPower);

        public bool CanPulseGuide => (bool)FunctionNoParameters(() => driver.CanPulseGuide);

        public bool CanSetCCDTemperature => (bool)FunctionNoParameters(() => driver.CanSetCCDTemperature);

        public bool CanStopExposure => (bool)FunctionNoParameters(() => driver.CanStopExposure);

        public double CCDTemperature => (double)FunctionNoParameters(() => driver.CCDTemperature);

        public bool CoolerOn { get => (bool)FunctionNoParameters(() => driver.CoolerOn); set => Method1Parameter((i) => driver.CoolerOn = i,value); }

        public double CoolerPower => (double)FunctionNoParameters(() => driver.CoolerPower);

        public double ElectronsPerADU => (double)FunctionNoParameters(() => driver.ElectronsPerADU);

        public double FullWellCapacity => (double)FunctionNoParameters(() => driver.FullWellCapacity);

        public bool HasShutter => (bool)FunctionNoParameters(() => driver.HasShutter);

        public double HeatSinkTemperature => (double)FunctionNoParameters(() => driver.HeatSinkTemperature);

        public object ImageArray => FunctionNoParameters(() => driver.ImageArray);

        public object ImageArrayVariant => FunctionNoParameters(() => driver.ImageArrayVariant);

        public bool ImageReady => (bool)FunctionNoParameters(() => driver.ImageReady);

        public bool IsPulseGuiding => (bool)FunctionNoParameters(() => driver.IsPulseGuiding);

        public double LastExposureDuration => (double)FunctionNoParameters(() => driver.LastExposureDuration);

        public string LastExposureStartTime => (string)FunctionNoParameters(() => driver.LastExposureStartTime);

        public int MaxADU => (int)FunctionNoParameters(() => driver.MaxADU);

        public short MaxBinX => (short)FunctionNoParameters(() => driver.MaxBinX);

        public short MaxBinY => (short)FunctionNoParameters(() => driver.MaxBinY);

        public int NumX { get => (int)FunctionNoParameters(() => driver.NumX); set => Method1Parameter((i) => driver.NumX = i,value); }

        public int NumY { get => (int)FunctionNoParameters(() => driver.NumY); set => Method1Parameter((i) => driver.NumY = i, value); }

        public double PixelSizeX => (double)FunctionNoParameters(() => driver.PixelSizeX);

        public double PixelSizeY => (double)FunctionNoParameters(() => driver.PixelSizeY);

        public double SetCCDTemperature { get => (double)FunctionNoParameters(() => driver.SetCCDTemperature); set => Method1Parameter((i) => driver.SetCCDTemperature = i, value); }

        public int StartX { get => (int)FunctionNoParameters(() => driver.StartX); set => Method1Parameter((i) => driver.StartX = i, value); }

        public int StartY { get => (int)FunctionNoParameters(() => driver.StartY); set => Method1Parameter((i) => driver.StartY = i, value); }

        public short BayerOffsetX => (short)FunctionNoParameters(() => driver.BayerOffsetX);

        public short BayerOffsetY => (short)FunctionNoParameters(() => driver.BayerOffsetY);

        public bool CanFastReadout => (bool)FunctionNoParameters(() => driver.CanFastReadout);

        public double ExposureMax => (double)FunctionNoParameters(() => driver.ExposureMax);

        public double ExposureMin => (double)FunctionNoParameters(() => driver.ExposureMin);

        public double ExposureResolution => (double)FunctionNoParameters(() => driver.ExposureResolution);

        public bool FastReadout { get => (bool)FunctionNoParameters(() => driver.FastReadout); set => Method1Parameter((i) => driver.FastReadout = i, value); }

        public short Gain { get => (short)FunctionNoParameters(() => driver.Gain); set => Method1Parameter((i) => driver.Gain = i, value); }

        public short GainMax => (short)FunctionNoParameters(() => driver.GainMax);

        public short GainMin => (short)FunctionNoParameters(() => driver.GainMin);

        public IList<string> Gains
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in (System.Collections.IEnumerable)FunctionNoParameters(() => driver.Gains))
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public short PercentCompleted => (short)FunctionNoParameters(() => driver.PercentCompleted);

        public short ReadoutMode { get => (short)FunctionNoParameters(() => driver.ReadoutMode); set => Method1Parameter((i) => driver.ReadoutMode = i, value); }

        public IList<string> ReadoutModes
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in (System.Collections.IEnumerable)FunctionNoParameters(() => driver.ReadoutModes))
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public string SensorName => (string)FunctionNoParameters(() => driver.SensorName);

        public SensorType SensorType => (SensorType)FunctionNoParameters(() => driver.SensorType);

        public int Offset { get => (int)FunctionNoParameters(() => driver.Offset); set => Method1Parameter((i) => driver.Offset = i, value); }

        public int OffsetMax => (int)FunctionNoParameters(() => driver.OffsetMax);

        public int OffsetMin => (int)FunctionNoParameters(() => driver.OffsetMin);

        public IList<string> Offsets
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in (System.Collections.IEnumerable)FunctionNoParameters(() => driver.Offsets))
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public double SubExposureDuration { get => (double)FunctionNoParameters(() => driver.SubExposureDuration); set => Method1Parameter((i) => driver.SubExposureDuration = i, value); }

        public void AbortExposure()
        {
            MethodNoParameters(() => driver.AbortExposure());

        }

        public void PulseGuide(GuideDirection Direction, int Duration)
        {
            Method2Parameters((i, j) => driver.PulseGuide(i, j), Direction, Duration);
        }

        public void StartExposure(double Duration, bool Light)
        {
            Method2Parameters((i, j) => driver.StartExposure(i, j), Duration, Light);

        }

        public void StopExposure()
        {
            MethodNoParameters(() => driver.StopExposure());

        }

    }
}

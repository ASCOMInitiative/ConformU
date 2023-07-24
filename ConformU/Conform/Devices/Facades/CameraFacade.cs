using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections.Generic;

namespace ConformU
{
    public class CameraFacade : FacadeBaseClass, ICameraV3, IDisposable
    {

        // Create the test device in the facade base class
        public CameraFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public short BinX { get => FunctionNoParameters<short>(() => driver.BinX); set => Method1Parameter((i) => driver.BinX = i, value); }

        public short BinY { get => FunctionNoParameters<short>(() => driver.BinY); set => Method1Parameter((i) => driver.BinY = i, value); }

        public CameraState CameraState
        {
            get
            {
                return FunctionNoParameters<CameraState>(() => driver.CameraState);
            }
        }

        public int CameraXSize
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.CameraXSize);
            }
        }

        public int CameraYSize
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.CameraYSize);
            }
        }

        public bool CanAbortExposure
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanAbortExposure);
            }
        }

        public bool CanAsymmetricBin
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanAsymmetricBin);
            }
        }

        public bool CanGetCoolerPower
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanGetCoolerPower);
            }
        }

        public bool CanPulseGuide
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanPulseGuide);
            }
        }

        public bool CanSetCCDTemperature
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanSetCCDTemperature);
            }
        }

        public bool CanStopExposure
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanStopExposure);
            }
        }

        public double CCDTemperature
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.CCDTemperature);
            }
        }

        public bool CoolerOn { get => FunctionNoParameters<bool>(() => driver.CoolerOn); set => Method1Parameter((i) => driver.CoolerOn = i, value); }

        public double CoolerPower
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.CoolerPower);
            }
        }

        public double ElectronsPerADU
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.ElectronsPerADU);
            }
        }

        public double FullWellCapacity
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.FullWellCapacity);
            }
        }

        public bool HasShutter
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.HasShutter);
            }
        }

        public double HeatSinkTemperature
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.HeatSinkTemperature);
            }
        }

        public object ImageArray
        {
            get
            {
                return FunctionNoParameters(() => driver.ImageArray);
            }
        }

        public object ImageArrayVariant
        {
            get
            {
                return FunctionNoParameters(() => driver.ImageArrayVariant);
            }
        }

        public bool ImageReady
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.ImageReady);
            }
        }

        public bool IsPulseGuiding
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.IsPulseGuiding);
            }
        }

        public double LastExposureDuration
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.LastExposureDuration);
            }
        }

        public string LastExposureStartTime
        {
            get
            {
                return FunctionNoParameters<string>(() => driver.LastExposureStartTime);
            }
        }

        public int MaxADU
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.MaxADU);
            }
        }

        public short MaxBinX
        {
            get
            {
                return FunctionNoParameters<short>(() => driver.MaxBinX);
            }
        }

        public short MaxBinY
        {
            get
            {
                return FunctionNoParameters<short>(() => driver.MaxBinY);
            }
        }

        public int NumX { get => FunctionNoParameters<int>(() => driver.NumX); set => Method1Parameter((i) => driver.NumX = i, value); }

        public int NumY { get => FunctionNoParameters<int>(() => driver.NumY); set => Method1Parameter((i) => driver.NumY = i, value); }

        public double PixelSizeX
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.PixelSizeX);
            }
        }

        public double PixelSizeY
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.PixelSizeY);
            }
        }

        public double SetCCDTemperature { get => FunctionNoParameters<double>(() => driver.SetCCDTemperature); set => Method1Parameter((i) => driver.SetCCDTemperature = i, value); }

        public int StartX { get => FunctionNoParameters<int>(() => driver.StartX); set => Method1Parameter((i) => driver.StartX = i, value); }

        public int StartY { get => FunctionNoParameters<int>(() => driver.StartY); set => Method1Parameter((i) => driver.StartY = i, value); }

        public short BayerOffsetX
        {
            get
            {
                return FunctionNoParameters<short>(() => driver.BayerOffsetX);
            }
        }

        public short BayerOffsetY
        {
            get
            {
                return FunctionNoParameters<short>(() => driver.BayerOffsetY);
            }
        }

        public bool CanFastReadout
        {
            get
            {
                return FunctionNoParameters<bool>(() => driver.CanFastReadout);
            }
        }

        public double ExposureMax
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.ExposureMax);
            }
        }

        public double ExposureMin
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.ExposureMin);
            }
        }

        public double ExposureResolution
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.ExposureResolution);
            }
        }

        public bool FastReadout { get => FunctionNoParameters<bool>(() => driver.FastReadout); set => Method1Parameter((i) => driver.FastReadout = i, value); }

        public short Gain { get => FunctionNoParameters<short>(() => driver.Gain); set => Method1Parameter((i) => driver.Gain = i, value); }

        public short GainMax
        {
            get
            {
                return FunctionNoParameters<short>(() => driver.GainMax);
            }
        }

        public short GainMin
        {
            get
            {
                return FunctionNoParameters<short>(() => driver.GainMin);
            }
        }

        public IList<string> Gains
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in FunctionNoParameters<System.Collections.IEnumerable>(() => driver.Gains))
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public short PercentCompleted
        {
            get
            {
                return FunctionNoParameters<short>(() => driver.PercentCompleted);
            }
        }

        public short ReadoutMode { get => FunctionNoParameters<short>(() => driver.ReadoutMode); set => Method1Parameter((i) => driver.ReadoutMode = i, value); }

        public IList<string> ReadoutModes
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in FunctionNoParameters<System.Collections.IEnumerable>(() => driver.ReadoutModes))
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public string SensorName
        {
            get
            {
                return FunctionNoParameters<string>(() => driver.SensorName);
            }
        }

        public SensorType SensorType
        {
            get
            {
                return FunctionNoParameters<SensorType>(() => driver.SensorType);
            }
        }

        public int Offset { get => FunctionNoParameters<int>(() => driver.Offset); set => Method1Parameter((i) => driver.Offset = i, value); }

        public int OffsetMax
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.OffsetMax);
            }
        }

        public int OffsetMin
        {
            get
            {
                return FunctionNoParameters<int>(() => driver.OffsetMin);
            }
        }

        public IList<string> Offsets
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in FunctionNoParameters<System.Collections.IEnumerable>(() => driver.Offsets))
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public double SubExposureDuration { get => FunctionNoParameters<double>(() => driver.SubExposureDuration); set => Method1Parameter((i) => driver.SubExposureDuration = i, value); }

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

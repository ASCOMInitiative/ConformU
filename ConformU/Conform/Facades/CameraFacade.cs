using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections.Generic;

namespace ConformU
{
    public class CameraFacade : FacadeBaseClass, ICameraV4, IDisposable
    {

        // Create the test device in the facade base class
        public CameraFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public short BinX { get => FunctionNoParameters<short>(() => Driver.BinX); set => Method1Parameter((i) => Driver.BinX = i, value); }

        public short BinY { get => FunctionNoParameters<short>(() => Driver.BinY); set => Method1Parameter((i) => Driver.BinY = i, value); }

        public CameraState CameraState
        {
            get
            {
                return (CameraState)FunctionNoParameters<object>(() => Driver.CameraState);
            }
        }

        public int CameraXSize
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.CameraXSize);
            }
        }

        public int CameraYSize
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.CameraYSize);
            }
        }

        public bool CanAbortExposure
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanAbortExposure);
            }
        }

        public bool CanAsymmetricBin
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanAsymmetricBin);
            }
        }

        public bool CanGetCoolerPower
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanGetCoolerPower);
            }
        }

        public bool CanPulseGuide
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanPulseGuide);
            }
        }

        public bool CanSetCCDTemperature
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanSetCCDTemperature);
            }
        }

        public bool CanStopExposure
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanStopExposure);
            }
        }

        public double CCDTemperature
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.CCDTemperature);
            }
        }

        public bool CoolerOn { get => FunctionNoParameters<bool>(() => Driver.CoolerOn); set => Method1Parameter((i) => Driver.CoolerOn = i, value); }

        public double CoolerPower
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.CoolerPower);
            }
        }

        public double ElectronsPerADU
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.ElectronsPerADU);
            }
        }

        public double FullWellCapacity
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.FullWellCapacity);
            }
        }

        public bool HasShutter
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.HasShutter);
            }
        }

        public double HeatSinkTemperature
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.HeatSinkTemperature);
            }
        }

        public object ImageArray
        {
            get
            {
                return FunctionNoParameters<object>(() => Driver.ImageArray);
            }
        }

        public object ImageArrayVariant
        {
            get
            {
                return FunctionNoParameters<object>(() => Driver.ImageArrayVariant);
            }
        }

        public bool ImageReady
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.ImageReady);
            }
        }

        public bool IsPulseGuiding
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.IsPulseGuiding);
            }
        }

        public double LastExposureDuration
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.LastExposureDuration);
            }
        }

        public string LastExposureStartTime
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.LastExposureStartTime);
            }
        }

        public int MaxADU
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.MaxADU);
            }
        }

        public short MaxBinX
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.MaxBinX);
            }
        }

        public short MaxBinY
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.MaxBinY);
            }
        }

        public int NumX { get => FunctionNoParameters<int>(() => Driver.NumX); set => Method1Parameter((i) => Driver.NumX = i, value); }

        public int NumY { get => FunctionNoParameters<int>(() => Driver.NumY); set => Method1Parameter((i) => Driver.NumY = i, value); }

        public double PixelSizeX
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.PixelSizeX);
            }
        }

        public double PixelSizeY
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.PixelSizeY);
            }
        }

        public double SetCCDTemperature { get => FunctionNoParameters<double>(() => Driver.SetCCDTemperature); set => Method1Parameter((i) => Driver.SetCCDTemperature = i, value); }

        public int StartX { get => FunctionNoParameters<int>(() => Driver.StartX); set => Method1Parameter((i) => Driver.StartX = i, value); }

        public int StartY { get => FunctionNoParameters<int>(() => Driver.StartY); set => Method1Parameter((i) => Driver.StartY = i, value); }

        public short BayerOffsetX
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.BayerOffsetX);
            }
        }

        public short BayerOffsetY
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.BayerOffsetY);
            }
        }

        public bool CanFastReadout
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanFastReadout);
            }
        }

        public double ExposureMax
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.ExposureMax);
            }
        }

        public double ExposureMin
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.ExposureMin);
            }
        }

        public double ExposureResolution
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.ExposureResolution);
            }
        }

        public bool FastReadout { get => FunctionNoParameters<bool>(() => Driver.FastReadout); set => Method1Parameter((i) => Driver.FastReadout = i, value); }

        public short Gain { get => FunctionNoParameters<short>(() => Driver.Gain); set => Method1Parameter((i) => Driver.Gain = i, value); }

        public short GainMax
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.GainMax);
            }
        }

        public short GainMin
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.GainMin);
            }
        }

        public IList<string> Gains
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in FunctionNoParameters<System.Collections.IEnumerable>(() => Driver.Gains))
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
                return FunctionNoParameters<short>(() => Driver.PercentCompleted);
            }
        }

        public short ReadoutMode { get => FunctionNoParameters<short>(() => Driver.ReadoutMode); set => Method1Parameter((i) => Driver.ReadoutMode = i, value); }

        public IList<string> ReadoutModes
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in FunctionNoParameters<System.Collections.IEnumerable>(() => Driver.ReadoutModes))
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
                return FunctionNoParameters<string>(() => Driver.SensorName);
            }
        }

        public SensorType SensorType
        {
            get
            {
                return (SensorType)FunctionNoParameters<object>(() => Driver.SensorType);
            }
        }

        public int Offset { get => FunctionNoParameters<int>(() => Driver.Offset); set => Method1Parameter((i) => Driver.Offset = i, value); }

        public int OffsetMax
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.OffsetMax);
            }
        }

        public int OffsetMin
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.OffsetMin);
            }
        }

        public IList<string> Offsets
        {
            get
            {
                List<string> returnValue = new();
                foreach (string gain in FunctionNoParameters<System.Collections.IEnumerable>(() => Driver.Offsets))
                {
                    returnValue.Add(gain);
                }
                return returnValue;
            }
        }

        public double SubExposureDuration { get => FunctionNoParameters<double>(() => Driver.SubExposureDuration); set => Method1Parameter((i) => Driver.SubExposureDuration = i, value); }

        public void AbortExposure()
        {
            MethodNoParameters(() => Driver.AbortExposure());

        }

        public void PulseGuide(GuideDirection direction, int duration)
        {
            Method2Parameters((i, j) => Driver.PulseGuide(i, j), direction, duration);
        }

        public void StartExposure(double duration, bool light)
        {
            Method2Parameters((i, j) => Driver.StartExposure(i, j), duration, light);

        }

        public void StopExposure()
        {
            MethodNoParameters(() => Driver.StopExposure());

        }

    }
}

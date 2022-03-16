
/* Unmerged change from project 'ConformU (net5.0)'
Before:
using System;
After:
using ASCOM.Common.DeviceInterfaces;
using System;
*/
using ASCOM.Common.DeviceInterfaces;
using System.
/* Unmerged change from project 'ConformU (net5.0)'
Before:
using System.Threading.Tasks;
using ASCOM.Common.DeviceInterfaces;
After:
using System.Threading.Tasks;
*/
Collections;
using System.Collections.Generic;

namespace ConformU
{
    public class VideoFacade : FacadeBaseClass, IVideo
    {
        // Create the test device in the facade base class
        public VideoFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public string VideoCaptureDeviceName => (string)FunctionNoParameters(() => driver.VideoCaptureDeviceName);

        public double ExposureMax => (double)FunctionNoParameters(() => driver.ExposureMax);

        public double ExposureMin => (double)FunctionNoParameters(() => driver.ExposureMin);

        public VideoCameraFrameRate FrameRate => (VideoCameraFrameRate)FunctionNoParameters(() => driver.FrameRate);

        public IList<double> SupportedIntegrationRates
        {
            get
            {
                List<double> returnValue = new();
                var gains = FunctionNoParameters(() => driver.SupportedIntegrationRates);
                foreach (double o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;
            }
        }

        public int IntegrationRate { get => (int)FunctionNoParameters(() => driver.IntegrationRate); set => Method1Parameter((i) => driver.IntegrationRate = i, value); }

        public IVideoFrame LastVideoFrame
        {
            get
            {
                dynamic returnValue = FunctionNoParameters(() => driver.LastVideoFrame);
                return returnValue;
            }
        }

        public string SensorName => (string)FunctionNoParameters(() => driver.SensorName);

        public SensorType SensorType => (SensorType)FunctionNoParameters(() => driver.SensorType);

        public int Width => (int)FunctionNoParameters(() => driver.Width);

        public int Height => (int)FunctionNoParameters(() => driver.Height);

        public double PixelSizeX => (double)FunctionNoParameters(() => driver.PixelSizeX);

        public double PixelSizeY => (double)FunctionNoParameters(() => driver.PixelSizeY);

        public int BitDepth => (int)FunctionNoParameters(() => driver.BitDepth);

        public string VideoCodec => (string)FunctionNoParameters(() => driver.VideoCodec);

        public string VideoFileFormat => (string)FunctionNoParameters(() => driver.VideoFileFormat);

        public int VideoFramesBufferSize => (int)FunctionNoParameters(() => driver.VideoFramesBufferSize);

        public VideoCameraState CameraState => (VideoCameraState)FunctionNoParameters(() => driver.CameraState);

        public short GainMax => (short)FunctionNoParameters(() => driver.GainMax);

        public short GainMin => (short)FunctionNoParameters(() => driver.GainMin);

        public short Gain { get => (short)FunctionNoParameters(() => driver.Gain); set => Method1Parameter((i) => driver.Gain = i, value); }

        public IList<string> Gains
        {
            get
            {
                List<string> returnValue = new();
                var gains = FunctionNoParameters(() => driver.Gains);
                foreach (string o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;
            }
        }

        public short GammaMax => (short)FunctionNoParameters(() => driver.GammaMax);

        public short GammaMin => (short)FunctionNoParameters(() => driver.GammaMin);

        public short Gamma { get => (short)FunctionNoParameters(() => driver.Gamma); set => Method1Parameter((i) => driver.Gamma = i, value); }

        public IList<string> Gammas
        {
            get
            {
                List<string> returnValue = new();
                var gains = FunctionNoParameters(() => driver.Gammas);
                foreach (string o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;

            }
        }

        public bool CanConfigureDeviceProperties => (bool)FunctionNoParameters(() => driver.CanConfigureDeviceProperties);

        public string StartRecordingVideoFile(string PreferredFileName)
        {
            return (string)Function1Parameter((i) => driver.StartRecordingVideoFile(i), PreferredFileName);
        }

        public void StopRecordingVideoFile()
        {
            MethodNoParameters(() => driver.StopRecordingVideoFile());
        }

        public void ConfigureDeviceProperties()
        {
            throw new ASCOM.MethodNotImplementedException($"Conform does not support testing the ConfigureDeviceProperties method.");
        }
    }
}

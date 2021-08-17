using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ASCOM.Standard.Interfaces;

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

        public ArrayList SupportedIntegrationRates
        {
            get
            {
                ArrayList returnValue = new ArrayList();
                var gains = FunctionNoParameters(() => driver.SupportedIntegrationRates);
                foreach (object o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;
            }
        }

        public int IntegrationRate { get => (int)FunctionNoParameters(() => driver.IntegrationRate); set => Method1Parameter((i) => driver.IntegrationRate = i, value); }

        public object LastVideoFrame
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

        public ArrayList Gains
        {
            get
            {
                ArrayList returnValue = new ArrayList();
                var gains = FunctionNoParameters(() => driver.Gains);
                foreach (object o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;
            }
        }

        public short GammaMax => (short)FunctionNoParameters(() => driver.GammaMax);

        public short GammaMin => (short)FunctionNoParameters(() => driver.GammaMin);

        public short Gamma { get => (short)FunctionNoParameters(() => driver.Gamma); set => Method1Parameter((i) => driver.Gamma = i, value); }

        public ArrayList Gammas
        {
            get
            {
                ArrayList returnValue = new();
                var gains = FunctionNoParameters(() => driver.Gammas);
                foreach (object o in (IList)gains)
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
    }
}

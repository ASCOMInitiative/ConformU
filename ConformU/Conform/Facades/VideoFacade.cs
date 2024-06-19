
using ASCOM.Common.DeviceInterfaces;
using System.Collections;
using System.Collections.Generic;

namespace ConformU
{
    public class VideoFacade : FacadeBaseClass, IVideoV2
    {
        // Create the test device in the facade base class
        public VideoFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        public string VideoCaptureDeviceName
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.VideoCaptureDeviceName);
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

        public VideoCameraFrameRate FrameRate
        {
            get
            {
                return (VideoCameraFrameRate)FunctionNoParameters<object>(() => Driver.FrameRate);
            }
        }

        public IList<double> SupportedIntegrationRates
        {
            get
            {
                List<double> returnValue = new();
                var gains = FunctionNoParameters<object>(() => Driver.SupportedIntegrationRates);
                foreach (double o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;
            }
        }

        public int IntegrationRate { get => FunctionNoParameters<int>(() => Driver.IntegrationRate); set => Method1Parameter((i) => Driver.IntegrationRate = i, value); }

        public IVideoFrame LastVideoFrame
        {
            get
            {
                VideoFrame frame = null;
                try
                {
                    // Get the last VideoFrame
                    dynamic lastFrame = FunctionNoParameters<dynamic>(() => Driver.LastVideoFrame);
                    Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Got last frame from driver");
                    // Create and populate the correct metadata return type
                    List<KeyValuePair<string, string>> imageMetaData = new();
                    foreach (var pair in lastFrame.ImageMetadata)
                    {
                        Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Found pair: {pair.Key()}, {pair.Value()}");
                        imageMetaData.Add(new KeyValuePair<string, string>(pair.Key(), pair.Value()));
                    }
                    // Create a new frame with the correct parameter types
                    Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Creating frame");

                    object imageArray;
                    byte[] previewBitmap;
                    long frameNumber;
                    double exposureDuration;
                    string exposureStartTime;

                    try
                    {
                        imageArray = lastFrame.ImageArray;
                    }
                    catch (System.Exception ex)
                    {
                        Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Exception when setting ImageArray, supplying default value: null\r\n{ex}");
                        imageArray = null;
                    }

                    try
                    {
                        previewBitmap = lastFrame.PreviewBitmap;
                    }
                    catch (System.Exception ex)
                    {
                        Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Exception when setting PreviewBitmap, supplying default value: byte[10]\r\n{ex}");
                        previewBitmap = new byte[10];
                    }

                    try
                    {
                        frameNumber = lastFrame.FrameNumber;
                    }
                    catch (System.Exception ex)
                    {
                        Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Exception when setting FrameNumber, supplying default value: 0\r\n{ex}");
                        frameNumber = 0;
                    }

                    try
                    {
                        exposureDuration = lastFrame.ExposureDuration;
                    }
                    catch (System.Exception ex)
                    {
                        Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Exception when setting ExposureDuration, supplying default value: 0.0\r\n{ex}");
                        exposureDuration = 0.0;
                    }

                    try
                    {
                        exposureStartTime = lastFrame.ExposureStartTime;
                    }
                    catch (System.Exception ex)
                    {
                        Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Exception when setting ExposureStartTime, supplying default value: empty string\r\n{ex}");
                        exposureStartTime = "";
                    }

                    frame = new(imageArray, previewBitmap, frameNumber, exposureDuration, exposureStartTime, imageMetaData);
                    Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Created frame");

                }
                catch (System.Exception ex)
                {
                    Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Exception\r\n{ex}");
                }
                Logger.LogMessage("VideoFacade.LastVideoFrame", MessageLevel.Debug, $"Returning frame");

                // Return the new frame in place of the original
                return frame;
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

        public int Width
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.Width);
            }
        }

        public int Height
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.Height);
            }
        }

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

        public int BitDepth
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.BitDepth);
            }
        }

        public string VideoCodec
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.VideoCodec);
            }
        }

        public string VideoFileFormat
        {
            get
            {
                return FunctionNoParameters<string>(() => Driver.VideoFileFormat);
            }
        }

        public int VideoFramesBufferSize
        {
            get
            {
                return FunctionNoParameters<int>(() => Driver.VideoFramesBufferSize);
            }
        }

        public VideoCameraState CameraState
        {
            get
            {
                return (VideoCameraState)FunctionNoParameters<object>(() => Driver.CameraState);
            }
        }

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

        public short Gain { get => FunctionNoParameters<short>(() => Driver.Gain); set => Method1Parameter((i) => Driver.Gain = i, value); }

        public IList<string> Gains
        {
            get
            {
                List<string> returnValue = new();
                var gains = FunctionNoParameters<object>(() => Driver.Gains);
                foreach (string o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;
            }
        }

        public short GammaMax
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.GammaMax);
            }
        }

        public short GammaMin
        {
            get
            {
                return FunctionNoParameters<short>(() => Driver.GammaMin);
            }
        }

        public short Gamma { get => FunctionNoParameters<short>(() => Driver.Gamma); set => Method1Parameter((i) => Driver.Gamma = i, value); }

        public IList<string> Gammas
        {
            get
            {
                List<string> returnValue = new();
                var gains = FunctionNoParameters<object>(() => Driver.Gammas);
                foreach (string o in (IList)gains)
                {
                    returnValue.Add(o);
                }

                return returnValue;

            }
        }

        public bool CanConfigureDeviceProperties
        {
            get
            {
                return FunctionNoParameters<bool>(() => Driver.CanConfigureDeviceProperties);
            }
        }

        public string StartRecordingVideoFile(string preferredFileName)
        {
            return Function1Parameter<string>((i) => Driver.StartRecordingVideoFile(i), preferredFileName);
        }

        public void StopRecordingVideoFile()
        {
            MethodNoParameters(() => Driver.StopRecordingVideoFile());
        }

        public void ConfigureDeviceProperties()
        {
            throw new ASCOM.MethodNotImplementedException($"Conform does not support testing the ConfigureDeviceProperties method.");
        }
    }
}

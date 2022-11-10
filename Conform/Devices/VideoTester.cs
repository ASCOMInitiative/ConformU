using ASCOM;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;

/* Unmerged change from project 'ConformU (net5.0)'
Before:
using System.Collections.Generic;
using System.Linq;
using ASCOM.Common.DeviceInterfaces;
using ASCOM.Com.DriverAccess;
using System.Threading;
using System.Collections;
After:
using ASCOM.Com.DriverAccess;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
*/
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConformU
{

    internal class VideoTester : DeviceTesterBaseClass
    {
        const int CAMERA_PULSE_DURATION = 2000; // Duration of camera pulse guide test (ms)
        const int CAMERA_PULSE_TOLERANCE = 300; // Tolerance for acceptable;e performance (ms)

        // Camera variables
        private bool CanConfigureDeviceProperties, CanReadSensorType, CanReadGainMax, CanReadGainMin;
        private bool CanReadGammaMin, CanReadGammaMax, CanReadIntegrationRate, CanReadSupportedIntegrationRates;
        private bool CanReadVideoFrame;
        private double PixelSizeX, PixelSizeY, ExposureMax, ExposureMin;
        private int BitDepth, Height, Width, IntegrationRate, VideoFramesBufferSize;
        private short Gain, GainMax, GainMin, Gamma, GammaMin, GammaMax;
        private dynamic Gains, Gammas, SupportedIntegrationRates;
        private string SensorName;
        private SensorType SensorType;
        private VideoCameraState CameraState;
        private VideoCameraFrameRate FrameRate;
        private dynamic LastVideoFrame;
        private string ExposureStartTime, VideoCaptureDeviceName, VideoCodec, VideoFileFormat;

        // VideoFrame properties
        private double ExposureDuration;
        private long FrameNumber;
        private object ImageArray;
        private Array ImageArrayAsArray;
        private dynamic ImageMetadata;
        private byte[] PreviewBitmap;

        private enum CanProperty
        {
            CanConfigureDeviceProperties = 1
        }
        private enum CameraPerformance : int
        {
            CameraState
        }
        private enum VideoProperty
        {
            // IVideo properties
            BitDepth,
            CameraState,
            CanConfigureDeviceProperties,
            ExposureMax,
            ExposureMin,
            CCDTemperature,
            FrameRate,
            Gain,
            GainMax,
            GainMin,
            Gains,
            Gamma,
            GammaMax,
            GammaMin,
            Gammas,
            Height,
            IntegrationRate,
            LastVideoFrame,
            PixelSizeX,
            PixelSizeY,
            SensorName,
            SensorType,
            SupportedIntegrationRates,
            VideoCaptureDeviceName,
            VideoCodec,
            VideoFileFormat,
            VideoFramesBufferSize,
            Width,

            // IVideoFrame Properties
            ExposureDuration,
            ExposureStartTime,
            FrameNumber,
            ImageArray,
            ImageMetadata,
            PreviewBitmap
        }

        // Helper variables
        private IVideo videoDevice;
        private readonly CancellationToken cancellationToken;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region New and Dispose
        public VideoTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(true, true, true, false, false, true, true, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        protected override void Dispose(bool disposing)
        {
            LogDebug("Dispose", "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (videoDevice is not null) videoDevice.Dispose();
                    videoDevice = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(videoDevice, DeviceTypes.Camera);
        }

        public new void CheckInitialise()
        {
            // Set the error type numbers according to the standards adopted by individual authors.
            // Unfortunately these vary between drivers so I have to allow for these here in order to give meaningful
            // messages to driver authors!
            unchecked
            {
                switch (settings.ComDevice.ProgId)
                {
                    default:
                        {
                            g_ExNotImplemented = (int)0x80040400;
                            g_ExInvalidValue1 = (int)0x80040405;
                            g_ExInvalidValue2 = (int)0x80040405;
                            g_ExNotSet1 = (int)0x80040403;
                            break;
                        }
                }
            }
            base.CheckInitialise();
        }

        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        LogIssue("CreateDevice", "The Alpaca implementation does not support video devices.");
                        throw new Exception("The Alpaca implementation does not support video devices.");

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                videoDevice = new VideoFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                videoDevice = new Video(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                baseClassDevice = videoDevice; // Assign the driver to the base class

                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);
            }
            catch (Exception ex)
            {
                LogDebug("CreateDevice", "Exception thrown: " + ex.Message);
                throw; // Re throw exception 
            }

        }

        public override bool Connected
        {
            get
            {
                LogCallToDriver("Connected", "About to get Connected property");
                return videoDevice.Connected;
            }
            set
            {
                LogCallToDriver("Connected", "About to set Connected property");
                videoDevice.Connected = value;
            }
        }

        public override void ReadCanProperties()
        {
            // IVideoV1 properties
            CameraCanTest(CanProperty.CanConfigureDeviceProperties, "CanConfigureDeviceProperties");
        }

        private void CameraCanTest(CanProperty p_Type, string p_Name)
        {
            try
            {
                switch (p_Type)
                {
                    case CanProperty.CanConfigureDeviceProperties:
                        {
                            LogCallToDriver(p_Type.ToString(), "About to get CanConfigureDeviceProperties property");
                            CanConfigureDeviceProperties = videoDevice.CanConfigureDeviceProperties;
                            LogOK(p_Name, CanConfigureDeviceProperties.ToString());
                            break;
                        }

                    default:
                        {
                            LogIssue(p_Name, "Conform:CanTest: Unknown test type " + p_Type.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogIssue(p_Name, $"Exception: {ex.Message}");
            }
        }

        public override void CheckProperties()
        {

            // BitDepth - Mandatory
            BitDepth = TestInteger(VideoProperty.BitDepth, 1, int.MaxValue, true); if (cancellationToken.IsCancellationRequested)
                return;

            // CameraState - Mandatory
            try
            {
                LogCallToDriver("CameraState", "About to get VideoCameraRunning property");
                CameraState = VideoCameraState.Running;
                CameraState = videoDevice.CameraState;
                LogOK("CameraState Read", CameraState.ToString());
            }
            catch (Exception ex)
            {
                HandleException("CameraState Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            // ExposureMin and ExpoosureMax Read - Mandatory
            ExposureMax = TestDouble(VideoProperty.ExposureMax, 0.0001, double.MaxValue, true);
            ExposureMin = TestDouble(VideoProperty.ExposureMin, 0.0, double.MaxValue, true);

            // Apply tests to resultant exposure values
            if (ExposureMin <= ExposureMax)
                LogOK("ExposureMin", "ExposureMin is less than or equal to ExposureMax");
            else
                LogIssue("ExposureMin", "ExposureMin is greater than ExposureMax");

            // FrameRate - Mandatory
            try
            {
                FrameRate = VideoCameraFrameRate.PAL;
                LogCallToDriver("FrameRate", "About to get FrameRate property");
                FrameRate = videoDevice.FrameRate;
                LogOK("FrameRate Read", FrameRate.ToString());
            }
            catch (Exception ex)
            {
                HandleException("FrameRate Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            // Read the Gain properties - Optional
            GainMax = TestShort(VideoProperty.GainMax, 0, short.MaxValue, false);
            GainMin = TestShort(VideoProperty.GainMin, 0, short.MaxValue, false);
            Gains = TestArrayList(VideoProperty.Gains, false, Type.GetType("System.String"));
            Gain = TestShort(VideoProperty.Gain, 0, short.MaxValue, false);

            // Now apply tests to the resultant Gain values
            if (CanReadGainMin ^ CanReadGainMax)
            {
                if (CanReadGainMin)
                    LogIssue("GainMinMax", "Can read GainMin but GainMax threw an exception");
                else
                    LogIssue("GainMinMax", "Can read GainMax but GainMin threw an exception");
            }
            else
                LogOK("GainMinMax", "Both GainMin and GainMax are readable or both throw exceptions");

            // Read the Gamma properties - Optional
            GammaMax = TestShort(VideoProperty.GammaMax, 0, short.MaxValue, false);
            GammaMin = TestShort(VideoProperty.GammaMin, 0, short.MaxValue, false);
            Gammas = TestArrayList(VideoProperty.Gammas, false, Type.GetType("System.String"));
            Gamma = TestShort(VideoProperty.Gamma, 0, short.MaxValue, false);

            // Now apply tests to the resultant Gamma values
            if (CanReadGammaMin ^ CanReadGammaMax)
            {
                if (CanReadGammaMin)
                    LogIssue("GammaMinMax", "Can read GammaMin but GammaMax threw an exception");
                else
                    LogIssue("GammaMinMax", "Can read GammaMax but GammaMin threw an exception");
            }
            else
                LogOK("GammaMinMax", "Both GammaMin and GammaMax are readable or both throw exceptions");

            // Height and width - Mandatory
            Height = TestInteger(VideoProperty.Height, 1, int.MaxValue, true);
            Width = TestInteger(VideoProperty.Width, 1, int.MaxValue, true);

            // Integration rates - Optional
            IntegrationRate = TestInteger(VideoProperty.IntegrationRate, 0, int.MaxValue, false);
            SupportedIntegrationRates = TestArrayList(VideoProperty.SupportedIntegrationRates, false, Type.GetType("System.Double"));

            // Now apply tests to the resultant integration rate values
            if (CanReadIntegrationRate ^ CanReadSupportedIntegrationRates)
            {
                if (CanReadIntegrationRate)
                    LogIssue("IntegrationRates", "Can read IntegrationRate but SupportedIntegrationRates threw an exception");
                else
                    LogIssue("IntegrationRates", "Can read SupportedIntegrationRates but IntegrationRate threw an exception");
            }
            else
                LogOK("IntegrationRates", "Both IntegrationRate and SupportedIntegrationRates are readable or both throw exceptions");

            // Pixel size - Mandatory
            PixelSizeX = TestDouble(VideoProperty.PixelSizeX, 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;
            PixelSizeY = TestDouble(VideoProperty.PixelSizeY, 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;

            // SensorName - Mandatory
            SensorName = TestString(VideoProperty.SensorName, 250, true);

            // SensorType - Mandatory
            try
            {
                CanReadSensorType = false;
                LogCallToDriver("SensorType", "About to get SensorType property");
                SensorType = videoDevice.SensorType;
                CanReadSensorType = true; // Set a flag to indicate that we have got a valid SensorType value
                                          // Successfully retrieved a value
                LogOK("SensorType", SensorType.ToString());
            }
            catch (Exception ex)
            {
                HandleException("SensorType", MemberType.Property, Required.Mandatory, ex, "");
            }

            // VideoCaptureDeviceName
            VideoCaptureDeviceName = TestString(VideoProperty.VideoCaptureDeviceName, 1000, false);

            // VideoCodec 
            VideoCodec = TestString(VideoProperty.VideoCodec, 1000, false);

            // VideoFileFormat 
            VideoFileFormat = TestString(VideoProperty.VideoFileFormat, 1000, true);

            // VideoFramesBufferSize 
            VideoFramesBufferSize = TestInteger(VideoProperty.VideoFramesBufferSize, 0, int.MaxValue, true);

            // LastVideoFrame
            LastVideoFrame = TestVideoFrame(VideoProperty.LastVideoFrame, true);

            // Check contents of received frame
            if (CanReadVideoFrame)
            {
                ExposureDuration = TestDouble(VideoProperty.ExposureDuration, 0.0, double.MaxValue, false);
                ExposureStartTime = TestString(VideoProperty.ExposureStartTime, int.MaxValue, false);
                FrameNumber = TestLong(VideoProperty.FrameNumber, 0, long.MaxValue, true);
                ImageMetadata = TestArrayList(VideoProperty.ImageMetadata, true, typeof(KeyValuePair));

                try
                {
                    ImageArray = LastVideoFrame.ImageArray;
                    try
                    {
                        LogOK("ImageArray", "Received an image object from the driver of type: " + ImageArray.GetType().Name);
                    }
                    catch (Exception)
                    {
                        LogInfo("ImageArray", "Received an image object from the driver of indeterminate type");
                    }

                    // Check image array dimensions
                    try
                    {
                        ImageArrayAsArray = (Array)ImageArray;
                        LogOK("ImageArray", "  Received an array of rank: " + ImageArrayAsArray.Rank + ", length: " + ImageArrayAsArray.LongLength.ToString("#,0") + " and type: " + ImageArrayAsArray.GetType().Name);

                        switch (ImageArrayAsArray.Rank)
                        {
                            case 1 // Rank 1
                           :
                                {
                                    if (ImageArrayAsArray.GetType().Equals(typeof(int[])))
                                        LogOK("ImageArray", "  Received a 1 dimension Integer array as expected.");
                                    else
                                        LogIssue("ImageArray", "  Did not receive a 1 dimension Integer array as expected. Received: " + ImageArrayAsArray.GetType().Name);
                                    break;
                                }

                            case 2 // Rank 2
                     :
                                {
                                    if (ImageArrayAsArray.GetType().Equals(typeof(int[,])))
                                        LogOK("ImageArray", "  Received a 2 dimension Integer array as expected.");
                                    else
                                        LogIssue("ImageArray", "  Did not receive a 2 dimension Integer array as expected. Received: " + ImageArrayAsArray.GetType().Name);
                                    break;
                                }

                            case 3 // Rank 3
                     :
                                {
                                    if (ImageArrayAsArray.GetType().Equals(typeof(int[,,])))
                                        LogOK("ImageArray", "  Received a 3 dimension Integer array as expected.");
                                    else
                                        LogIssue("ImageArray", "  Did not receive a 3 dimension Integer array as expected. Received: " + ImageArrayAsArray.GetType().Name);
                                    break;
                                }

                            default:
                                {
                                    LogIssue("ImageArray", "  Array rank is 0 or exceeds 3: " + ImageArrayAsArray.GetType().Name);
                                    break;
                                }
                        }

                        if (CanReadSensorType)
                        {
                            switch (SensorType)
                            {
                                case object _ when SensorType == SensorType.Color // This camera returns multiple image planes of colour information
                               :
                                    {
                                        switch (ImageArrayAsArray.Rank)
                                        {
                                            case 1 // Invalid configuration
                                           :
                                                {
                                                    LogIssue("ImageArray", "  The SensorType is Colour and the zero based array rank is 0. For a colour sensor the array rank must be 1 or 2.");
                                                    LogInfo("ImageArray", "  Please see the IVideoFrame.ImageArray entry in the Platform Help file for allowed combinations of SensorType and ImageArray format.");
                                                    break;
                                                }

                                            case 2 // NumPlanes x (Height * Width). NumPlanes should be 3
                                     :
                                                {
                                                    CheckImage(ImageArrayAsArray, 3, Height * Width, 0);
                                                    break;
                                                }

                                            case 3 // NumPlanes x Height x Width. NumPlanes should be 3
                                     :
                                                {
                                                    CheckImage(ImageArrayAsArray, 3, Height, Width);
                                                    break;
                                                }

                                            default:
                                                {
                                                    // This is an unsupported rank 0 or >3 so create an error
                                                    LogIssue("ImageArray", "  The zero based array rank must be 1, 2 or 3 . The returned array had rank: " + ImageArrayAsArray.Rank);
                                                    break;
                                                }
                                        }

                                        break;
                                    }

                                default:
                                    {
                                        // This camera returns just one plane that may be literally monochrome or may appear monochrome 
                                        // but contain encoded colour information e.g. Bayer RGGB format
                                        switch (ImageArrayAsArray.Rank)
                                        {
                                            case 1 // (Height * Width)
                                           :
                                                {
                                                    CheckImage(ImageArrayAsArray, Height * Width, 0, 0);
                                                    break;
                                                }

                                            case 2 // Height x Width.
                                     :
                                                {
                                                    CheckImage(ImageArrayAsArray, Height, Width, 0);
                                                    break;
                                                }

                                            case 3 // Invalid configuration
                                     :
                                                {
                                                    LogIssue("ImageArray", "  The SensorType is not Colour and the array rank is 3. For non-colour sensors the array rank must be 1 or 2.");
                                                    LogInfo("ImageArray", "  Please see the IVideoFrame.ImageArray entry in the Platform Help file for allowed combinations of SensorType and ImageArray format.");
                                                    break;
                                                }

                                            default:
                                                {
                                                    // This is an unsupported rank 0 or >3 so create an error
                                                    LogIssue("ImageArray", "  The ImageArray rank must be 1, 2 or 3. The returned array had rank: " + ImageArrayAsArray.Rank);
                                                    break;
                                                }
                                        }

                                        break;
                                    }
                            }
                        }
                        else
                            LogInfo("ImageArray", "SensorType could not be determined so ImageArray quality tests have been skipped");
                    }
                    catch (Exception ex)
                    {
                        LogIssue("ImageArray", "Unexpected exception when testing ImageArray: " + ex.ToString());
                    }
                }
                catch (Exception ex)
                {
                    HandleException("PreviewBitmap", MemberType.Property, Required.Mandatory, ex, "");
                }

                try
                {
                    PreviewBitmap = LastVideoFrame.PreviewBitmap;
                    LogOK("PreviewBitmap", "Received an array with " + PreviewBitmap.Length.ToString("#,#,#") + " entries");
                }
                catch (Exception ex)
                {
                    HandleException("PreviewBitmap", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
                LogInfo("", "Skipping VideoFrame contents check because of issue reading LastVideoFrame");
        }

        /// <summary>
        ///     ''' Reports whether the overall array size matches the expected size
        ///     ''' </summary>
        ///     ''' <param name="TestArray">The array to test</param>
        ///     ''' <param name="Dimension1">Size of the first array dimension</param>
        ///     ''' <param name="Dimension2">Size of the second dimension or 0 for not present</param>
        ///     ''' <param name="Dimension3">Size of the third dimension or 0 for not present</param>
        ///     ''' <remarks></remarks>
        private void CheckImage(Array TestArray, long Dimension1, long Dimension2, long Dimension3)
        {
            long Length;
            const string CommaFormat = "#,0";

            Length = Dimension1 * ((Dimension2 > 0) ? Dimension2 : 1) * ((Dimension3 > 0) ? Dimension3 : 1); // Calculate the overall expected size

            if (TestArray.LongLength == Length)
                LogOK("CheckImage", "  ImageArray has the expected total number of pixels: " + Length.ToString(CommaFormat));
            else
                LogIssue("CheckImage", "  ImageArray returned a total of " + TestArray.Length.ToString(CommaFormat) + " pixels instead of the expected number: " + Length.ToString(CommaFormat));

            if (Dimension1 >= 1)
            {
                if (Dimension2 > 0)
                {
                    if (Dimension3 >= 1)
                    {
                        if (TestArray.GetLongLength(0) == Dimension1)
                            LogOK("CheckImage", "  ImageArray dimension 1 has the expected length:: " + Dimension1.ToString(CommaFormat));
                        else
                            LogIssue("CheckImage", "  ImageArray dimension 1 does not has the expected length:: " + Dimension1.ToString(CommaFormat) + ", received: " + TestArray.GetLongLength(0).ToString(CommaFormat));
                        if (TestArray.GetLongLength(1) == Dimension2)
                            LogOK("CheckImage", "  ImageArray dimension 2 has the expected length:: " + Dimension2.ToString(CommaFormat));
                        else
                            LogIssue("CheckImage", "  ImageArray dimension 2 does not has the expected length:: " + Dimension2.ToString(CommaFormat) + ", received: " + TestArray.GetLongLength(1).ToString(CommaFormat));
                        if (TestArray.GetLongLength(2) == Dimension3)
                            LogOK("CheckImage", "  ImageArray dimension 3 has the expected length:: " + Dimension3.ToString(CommaFormat));
                        else
                            LogIssue("CheckImage", "  ImageArray dimension 3 does not has the expected length:: " + Dimension3.ToString(CommaFormat) + ", received: " + TestArray.GetLongLength(2).ToString(CommaFormat));
                    }
                    else
                    {
                        if (TestArray.GetLongLength(0) == Dimension1)
                            LogOK("CheckImage", "  ImageArray dimension 1 has the expected length:: " + Dimension1.ToString(CommaFormat));
                        else
                            LogIssue("CheckImage", "  ImageArray dimension 1 does not has the expected length:: " + Dimension1.ToString(CommaFormat) + ", received: " + TestArray.GetLongLength(0).ToString(CommaFormat));
                        if (TestArray.GetLongLength(1) == Dimension2)
                            LogOK("CheckImage", "  ImageArray dimension 2 has the expected length:: " + Dimension2.ToString(CommaFormat));
                        else
                            LogIssue("CheckImage", "  ImageArray dimension 2 does not has the expected length:: " + Dimension2.ToString(CommaFormat) + ", received: " + TestArray.GetLongLength(1).ToString(CommaFormat));
                    }
                }
                else if (TestArray.GetLongLength(0) == Dimension1)
                    LogOK("CheckImage", "  ImageArray dimension 1 has the expected length:: " + Dimension1.ToString(CommaFormat));
                else
                    LogIssue("CheckImage", "  ImageArray dimension 1 does not has the expected length:: " + Dimension1.ToString(CommaFormat) + ", received: " + TestArray.GetLongLength(0).ToString(CommaFormat));
            }
            else
                LogIssue("CheckImage", "  Dimension 1 is 0 it should never be!");
        }

        private short TestShort(VideoProperty p_Type, short p_Min, short p_Max, bool p_Mandatory)
        {
            string MethodName;
            short returnValue = 0;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            try
            {
                returnValue = 0;
                LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property");
                switch (p_Type)
                {
                    case VideoProperty.GainMax:
                        {
                            CanReadGainMax = false;
                            returnValue = videoDevice.GainMax;
                            CanReadGainMax = true;
                            break;
                        }

                    case VideoProperty.GainMin:
                        {
                            CanReadGainMin = false;
                            returnValue = videoDevice.GainMin;
                            CanReadGainMin = true;
                            break;
                        }

                    case VideoProperty.Gain:
                        {
                            returnValue = videoDevice.Gain;
                            break;
                        }

                    case VideoProperty.GammaMax:
                        {
                            CanReadGammaMax = false;
                            returnValue = videoDevice.GammaMax;
                            CanReadGammaMax = true;
                            break;
                        }

                    case VideoProperty.GammaMin:
                        {
                            CanReadGammaMin = false;
                            returnValue = videoDevice.GammaMin;
                            CanReadGammaMin = true;
                            break;
                        }

                    case VideoProperty.Gamma:
                        {
                            returnValue = videoDevice.Gamma;
                            break;
                        }

                    default:
                        {
                            LogIssue(MethodName, "TestShort: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK(MethodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        ///     ''' Test whether an integer is returned by a driver
        ///     ''' </summary>
        ///     ''' <param name="p_Type">Method to test</param>
        ///     ''' <param name="p_Min">Lowest valid value</param>
        ///     ''' <param name="p_Max">Highest valid value</param>
        ///     ''' <param name="p_Mandatory">Mandatory method</param>
        ///     ''' <returns>Integer value returned by the driver</returns>
        ///     ''' <remarks></remarks>
        private int TestInteger(VideoProperty p_Type, int p_Min, int p_Max, bool p_Mandatory)
        {
            string MethodName;
            int returnValue = 0;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            try
            {
                returnValue = 0;
                LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property");
                switch (p_Type)
                {
                    case VideoProperty.BitDepth:
                        {
                            returnValue = videoDevice.BitDepth;
                            break;
                        }

                    case VideoProperty.CameraState:
                        {
                            returnValue = (int)videoDevice.CameraState;
                            break;
                        }

                    case VideoProperty.Height:
                        {
                            returnValue = videoDevice.Height;
                            break;
                        }

                    case VideoProperty.IntegrationRate:
                        {
                            CanReadIntegrationRate = false;
                            returnValue = videoDevice.IntegrationRate;
                            CanReadIntegrationRate = true;
                            break;
                        }

                    case VideoProperty.Width:
                        {
                            returnValue = videoDevice.Width;
                            break;
                        }

                    case VideoProperty.VideoFramesBufferSize:
                        {
                            returnValue = videoDevice.VideoFramesBufferSize;
                            break;
                        }

                    default:
                        {
                            LogIssue(MethodName, "TestInteger: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK(MethodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }
        /// <summary>
        ///     ''' Test whether an integer is returned by a driver
        ///     ''' </summary>
        ///     ''' <param name="p_Type">Method to test</param>
        ///     ''' <param name="p_Min">Lowest valid value</param>
        ///     ''' <param name="p_Max">Highest valid value</param>
        ///     ''' <param name="p_Mandatory">Mandatory method</param>
        ///     ''' <returns>Integer value returned by the driver</returns>
        ///     ''' <remarks></remarks>
        private long TestLong(VideoProperty p_Type, long p_Min, long p_Max, bool p_Mandatory)
        {
            string MethodName;
            long returnValue = 0;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            try
            {
                returnValue = 0;
                switch (p_Type)
                {
                    case VideoProperty.FrameNumber:
                        {
                            returnValue = LastVideoFrame.FrameNumber;
                            break;
                        }

                    default:
                        {
                            LogIssue(MethodName, "TestInteger: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK(MethodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        ///     ''' Test whether an integer is returned by a driver
        ///     ''' </summary>
        ///     ''' <param name="p_Type">Method to test</param>
        ///     ''' <param name="p_Min">Lowest valid value</param>
        ///     ''' <param name="p_Max">Highest valid value</param>
        ///     ''' <param name="p_Mandatory">Mandatory method</param>
        ///     ''' <returns>Integer value returned by the driver</returns>
        ///     ''' <remarks></remarks>
        private dynamic TestVideoFrame(VideoProperty p_Type, bool p_Mandatory)
        {
            string MethodName;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            dynamic returnValue = null;

            try
            {
                LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property");
                switch (p_Type)
                {
                    case VideoProperty.LastVideoFrame:
                        {
                            CanReadVideoFrame = false;
                            returnValue = videoDevice.LastVideoFrame;
                            CanReadVideoFrame = true;
                            break;
                        }

                    default:
                        {
                            LogIssue(MethodName, "TestVideoFrame: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }

                // Successfully retrieved a value
                LogOK(MethodName, "Successfully received VideoFrame");
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        private double TestDouble(VideoProperty p_Type, double p_Min, double p_Max, bool p_Mandatory)
        {
            string MethodName;
            double returnValue = 0.0;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            try
            {
                returnValue = 0.0;
                LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property");
                switch (p_Type)
                {
                    case VideoProperty.PixelSizeX:
                        {
                            returnValue = videoDevice.PixelSizeX;
                            break;
                        }

                    case VideoProperty.PixelSizeY:
                        {
                            returnValue = videoDevice.PixelSizeY;
                            break;
                        }

                    case VideoProperty.ExposureMax:
                        {
                            returnValue = videoDevice.ExposureMax;
                            break;
                        }

                    case VideoProperty.ExposureMin:
                        {
                            returnValue = videoDevice.ExposureMin;
                            break;
                        }

                    case VideoProperty.ExposureDuration:
                        {
                            returnValue = LastVideoFrame.ExposureDuration;
                            break;
                        }

                    default:
                        {
                            LogIssue(MethodName, "TestDouble: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < p_Min // Lower than minimum value
                   :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    case object _ when returnValue > p_Max // Higher than maximum value
             :
                        {
                            LogIssue(MethodName, "Invalid value: " + returnValue.ToString());
                            break;
                        }

                    default:
                        {
                            LogOK(MethodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }
        private bool TestBoolean(VideoProperty p_Type, bool p_Mandatory)
        {
            string MethodName;
            bool returnValue = false;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            try
            {
                returnValue = false;
                switch (p_Type)
                {
                    default:
                        {
                            LogIssue(MethodName, "TestBoolean: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                LogOK(MethodName, returnValue.ToString());
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        ///     ''' Test whether the driver returns a valid ArrayList
        ///     ''' </summary>
        ///     ''' <param name="p_Type">Property</param>
        ///     ''' <param name="p_Mandatory"></param>
        ///     ''' <param name="p_ItemType"></param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        ///     
        private dynamic TestArrayList(VideoProperty p_Type, bool p_Mandatory, Type p_ItemType)
        {
            string MethodName;
            int Count;

            dynamic returnValue = new ArrayList();

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            try
            {
                LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property");
                switch (p_Type)
                {
                    case VideoProperty.Gains:
                        {
                            returnValue = videoDevice.Gains;
                            break;
                        }

                    case VideoProperty.Gammas:
                        {
                            returnValue = videoDevice.Gammas;
                            break;
                        }

                    case VideoProperty.SupportedIntegrationRates:
                        {
                            CanReadSupportedIntegrationRates = false;
                            returnValue = videoDevice.SupportedIntegrationRates;
                            CanReadSupportedIntegrationRates = true;
                            break;
                        }

                    case VideoProperty.ImageMetadata:
                        {
                            returnValue = LastVideoFrame.ImageMetadata;
                            break;
                        }

                    default:
                        {
                            LogIssue(MethodName, "TestArrayList: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }

                // Successfully retrieved a value
                Count = 0;
                LogOK(MethodName, "Received an array containing " + returnValue.Count + " items.");

                foreach (object ListItem in returnValue)
                {
                    if (ListItem.GetType().Equals(p_ItemType))
                        LogOK(MethodName + "(" + Count + ")", "  " + ListItem.ToString());
                    else
                        LogIssue(MethodName, "  Type of ArrayList item: " + ListItem.GetType().Name + " does not match expected type: " + p_ItemType.Name);
                    Count += 1;
                }
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }

        private string TestString(VideoProperty p_Type, int p_MaxLength, bool p_Mandatory)
        {
            string MethodName;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            string returnValue = "";
            try
            {
                LogCallToDriver(p_Type.ToString(), $"About to get {p_Type} property");
                switch (p_Type)
                {
                    case VideoProperty.SensorName:
                        {
                            returnValue = videoDevice.SensorName;
                            break;
                        }

                    case VideoProperty.ExposureStartTime:
                        {
                            returnValue = LastVideoFrame.ExposureStartTime;
                            break;
                        }

                    case VideoProperty.VideoCaptureDeviceName:
                        {
                            returnValue = videoDevice.VideoCaptureDeviceName;
                            break;
                        }

                    case VideoProperty.VideoCodec:
                        {
                            returnValue = videoDevice.VideoCodec;
                            break;
                        }

                    case VideoProperty.VideoFileFormat:
                        {
                            returnValue = videoDevice.VideoFileFormat;
                            break;
                        }

                    default:
                        {
                            LogIssue(MethodName, "TestString: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue == "":
                        {
                            LogOK(MethodName, "The driver returned an empty string");
                            break;
                        }

                    default:
                        {
                            if (returnValue.Length <= p_MaxLength)
                                LogOK(MethodName, returnValue);
                            else
                                LogIssue(MethodName, "String exceeds " + p_MaxLength + " characters maximum length - " + returnValue);
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Property, p_Mandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        /// Not currently used in ConformU
        /// </summary>
        /// <param name="p_Type"></param>
        /// <param name="p_Property"></param>
        /// <param name="p_TestOK"></param>
        /// <param name="p_TestLow"></param>
        /// <param name="p_TestHigh"></param>
        private void CameraPropertyWriteTest(VideoProperty p_Type, string p_Property, int p_TestOK)
        {
            try // OK value first
            {
                switch (p_Type)
                {
                    case VideoProperty.BitDepth:
                        break;
                }
                LogOK(p_Property + " write", "Successfully wrote " + p_TestOK);
            }
            catch (Exception ex)
            {
                LogIssue(p_Property + " write", "Exception generated when setting legal value: " + p_TestOK.ToString() + " - " + ex.Message);
            }
        }

        public override void CheckMethods()
        {
        }

        public override void CheckPerformance()
        {
            CameraPerformanceTest(CameraPerformance.CameraState, "CameraState");
        }
        private void CameraPerformanceTest(CameraPerformance p_Type, string p_Name)
        {
            DateTime l_StartTime;
            double l_Count, l_LastElapsedTime, l_ElapsedTime, l_Rate;
            SetAction(p_Name);
            try
            {
                l_StartTime = DateTime.Now;
                l_Count = 0.0;
                l_LastElapsedTime = 0.0;
                do
                {
                    l_Count += 1.0;
                    switch (p_Type)
                    {
                        case CameraPerformance.CameraState:
                            {
                                CameraState = videoDevice.CameraState;
                                break;
                            }

                        default:
                            {
                                LogIssue(p_Name, "Conform:PerformanceTest: Unknown test type " + p_Type.ToString());
                                break;
                            }
                    }

                    l_ElapsedTime = DateTime.Now.Subtract(l_StartTime).TotalSeconds;
                    if (l_ElapsedTime > l_LastElapsedTime + 1.0)
                    {
                        SetStatus(l_Count + " transactions in " + l_ElapsedTime.ToString("0") + " seconds");
                        l_LastElapsedTime = l_ElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (l_ElapsedTime <= PERF_LOOP_TIME);
                l_Rate = l_Count / l_ElapsedTime;
                switch (l_Rate)
                {
                    case object _ when l_Rate > 10.0:
                        {
                            LogInfo(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 2.0 <= l_Rate && l_Rate <= 10.0:
                        {
                            LogOK(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    case object _ when 1.0 <= l_Rate && l_Rate <= 2.0:
                        {
                            LogInfo(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(p_Name, "Transaction rate: " + l_Rate.ToString("0.0") + " per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(p_Name, "Unable to complete test: " + ex.ToString());
            }
        }

        public override void PostRunCheck()
        {
            try
            {
                videoDevice.StopRecordingVideoFile();
            }
            catch
            {
            }
            LogOK("PostRunCheck", "Camera returned to initial cooler temperature");
        }
    }

}

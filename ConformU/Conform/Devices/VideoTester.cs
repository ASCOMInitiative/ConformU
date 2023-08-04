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
        private const int CAMERA_PULSE_DURATION = 2000; // Duration of camera pulse guide test (ms)
        private const int CAMERA_PULSE_TOLERANCE = 300; // Tolerance for acceptable;e performance (ms)

        // Camera variables
        private bool canConfigureDeviceProperties, canReadSensorType, canReadGainMax, canReadGainMin;
        private bool canReadGammaMin, canReadGammaMax, canReadIntegrationRate, canReadSupportedIntegrationRates;
        private bool canReadVideoFrame;
        private double pixelSizeX, pixelSizeY, exposureMax, exposureMin;
        private int bitDepth, height, width, integrationRate, videoFramesBufferSize;
        private short gain, gainMax, gainMin, gamma, gammaMin, gammaMax;
        private dynamic gains, gammas, supportedIntegrationRates;
        private string sensorName;
        private SensorType sensorType;
        private VideoCameraState cameraState;
        private VideoCameraFrameRate frameRate;
        private dynamic lastVideoFrame;
        private string exposureStartTime, videoCaptureDeviceName, videoCodec, videoFileFormat;

        // VideoFrame properties
        private double exposureDuration;
        private long frameNumber;
        private object imageArray;
        private Array imageArrayAsArray;
        private dynamic imageMetadata;
        private byte[] previewBitmap;

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
            CcdTemperature,
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
        private IVideoV2 videoDevice;
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
            LogDebug("Dispose", $"Disposing of device: {disposing} {disposedValue}");
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
            base.CheckCommonMethods(videoDevice, DeviceTypes.Video);
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
                            GExNotImplemented = (int)0x80040400;
                            GExInvalidValue1 = (int)0x80040405;
                            GExInvalidValue2 = (int)0x80040405;
                            GExNotSet1 = (int)0x80040403;
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
                        switch (settings.ComConfiguration.ComAccessMechanic)
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
                                throw new InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                BaseClassDevice = videoDevice; // Assign the driver to the base class

                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);
            }
            catch (COMException exCom) when (exCom.ErrorCode == REGDB_E_CLASSNOTREG)
            {
                LogDebug("CreateDevice", $"Exception thrown: {exCom.Message}\r\n{exCom}");

                throw new Exception($"The driver is not registered as a {(Environment.Is64BitProcess ? "64bit" : "32bit")} driver");
            }
            catch (Exception ex)
            {
                LogDebug("CreateDevice", $"Exception thrown: {ex.Message}\r\n{ex}");
                throw; // Re throw exception 
            }
        }

        public override void ReadCanProperties()
        {
            // IVideoV1 properties
            CameraCanTest(CanProperty.CanConfigureDeviceProperties, "CanConfigureDeviceProperties");
        }

        private void CameraCanTest(CanProperty pType, string pName)
        {
            try
            {
                switch (pType)
                {
                    case CanProperty.CanConfigureDeviceProperties:
                        {
                            LogCallToDriver(pType.ToString(), "About to get CanConfigureDeviceProperties property");
                            canConfigureDeviceProperties = videoDevice.CanConfigureDeviceProperties;
                            LogOk(pName, canConfigureDeviceProperties.ToString());
                            break;
                        }

                    default:
                        {
                            LogIssue(pName, $"Conform:CanTest: Unknown test type {pType}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogIssue(pName, $"Exception: {ex.Message}");
            }
        }

        public override void CheckProperties()
        {

            // BitDepth - Mandatory
            bitDepth = TestInteger(VideoProperty.BitDepth, 1, int.MaxValue, true); if (cancellationToken.IsCancellationRequested)
                return;

            // CameraState - Mandatory
            try
            {
                LogCallToDriver("CameraState", "About to get VideoCameraRunning property");
                cameraState = VideoCameraState.Running;
                cameraState = videoDevice.CameraState;
                LogOk("CameraState Read", cameraState.ToString());
            }
            catch (Exception ex)
            {
                HandleException("CameraState Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            // ExposureMin and ExpoosureMax Read - Mandatory
            exposureMax = TestDouble(VideoProperty.ExposureMax, 0.0001, double.MaxValue, true);
            exposureMin = TestDouble(VideoProperty.ExposureMin, 0.0, double.MaxValue, true);

            // Apply tests to resultant exposure values
            if (exposureMin <= exposureMax)
                LogOk("ExposureMin", "ExposureMin is less than or equal to ExposureMax");
            else
                LogIssue("ExposureMin", "ExposureMin is greater than ExposureMax");

            // FrameRate - Mandatory
            try
            {
                frameRate = VideoCameraFrameRate.PAL;
                LogCallToDriver("FrameRate", "About to get FrameRate property");
                frameRate = videoDevice.FrameRate;
                LogOk("FrameRate Read", frameRate.ToString());
            }
            catch (Exception ex)
            {
                HandleException("FrameRate Read", MemberType.Property, Required.Mandatory, ex, "");
            }

            // Read the Gain properties - Optional
            gainMax = TestShort(VideoProperty.GainMax, 0, short.MaxValue, false);
            gainMin = TestShort(VideoProperty.GainMin, 0, short.MaxValue, false);
            gains = TestArrayList(VideoProperty.Gains, false, Type.GetType("System.String"));
            gain = TestShort(VideoProperty.Gain, 0, short.MaxValue, false);

            // Now apply tests to the resultant Gain values
            if (canReadGainMin ^ canReadGainMax)
            {
                if (canReadGainMin)
                    LogIssue("GainMinMax", "Can read GainMin but GainMax threw an exception");
                else
                    LogIssue("GainMinMax", "Can read GainMax but GainMin threw an exception");
            }
            else
                LogOk("GainMinMax", "Both GainMin and GainMax are readable or both throw exceptions");

            // Read the Gamma properties - Optional
            gammaMax = TestShort(VideoProperty.GammaMax, 0, short.MaxValue, false);
            gammaMin = TestShort(VideoProperty.GammaMin, 0, short.MaxValue, false);
            gammas = TestArrayList(VideoProperty.Gammas, false, Type.GetType("System.String"));
            gamma = TestShort(VideoProperty.Gamma, 0, short.MaxValue, false);

            // Now apply tests to the resultant Gamma values
            if (canReadGammaMin ^ canReadGammaMax)
            {
                if (canReadGammaMin)
                    LogIssue("GammaMinMax", "Can read GammaMin but GammaMax threw an exception");
                else
                    LogIssue("GammaMinMax", "Can read GammaMax but GammaMin threw an exception");
            }
            else
                LogOk("GammaMinMax", "Both GammaMin and GammaMax are readable or both throw exceptions");

            // Height and width - Mandatory
            height = TestInteger(VideoProperty.Height, 1, int.MaxValue, true);
            width = TestInteger(VideoProperty.Width, 1, int.MaxValue, true);

            // Integration rates - Optional
            integrationRate = TestInteger(VideoProperty.IntegrationRate, 0, int.MaxValue, false);
            supportedIntegrationRates = TestArrayList(VideoProperty.SupportedIntegrationRates, false, Type.GetType("System.Double"));

            // Now apply tests to the resultant integration rate values
            if (canReadIntegrationRate ^ canReadSupportedIntegrationRates)
            {
                if (canReadIntegrationRate)
                    LogIssue("IntegrationRates", "Can read IntegrationRate but SupportedIntegrationRates threw an exception");
                else
                    LogIssue("IntegrationRates", "Can read SupportedIntegrationRates but IntegrationRate threw an exception");
            }
            else
                LogOk("IntegrationRates", "Both IntegrationRate and SupportedIntegrationRates are readable or both throw exceptions");

            // Pixel size - Mandatory
            pixelSizeX = TestDouble(VideoProperty.PixelSizeX, 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;
            pixelSizeY = TestDouble(VideoProperty.PixelSizeY, 1.0, double.PositiveInfinity, false); if (cancellationToken.IsCancellationRequested)
                return;

            // SensorName - Mandatory
            sensorName = TestString(VideoProperty.SensorName, 250, true);

            // SensorType - Mandatory
            try
            {
                canReadSensorType = false;
                LogCallToDriver("SensorType", "About to get SensorType property");
                sensorType = videoDevice.SensorType;
                canReadSensorType = true; // Set a flag to indicate that we have got a valid SensorType value
                                          // Successfully retrieved a value
                LogOk("SensorType", sensorType.ToString());
            }
            catch (Exception ex)
            {
                HandleException("SensorType", MemberType.Property, Required.Mandatory, ex, "");
            }

            // VideoCaptureDeviceName
            videoCaptureDeviceName = TestString(VideoProperty.VideoCaptureDeviceName, 1000, false);

            // VideoCodec 
            videoCodec = TestString(VideoProperty.VideoCodec, 1000, false);

            // VideoFileFormat 
            videoFileFormat = TestString(VideoProperty.VideoFileFormat, 1000, true);

            // VideoFramesBufferSize 
            videoFramesBufferSize = TestInteger(VideoProperty.VideoFramesBufferSize, 0, int.MaxValue, true);

            // LastVideoFrame
            lastVideoFrame = TestVideoFrame(VideoProperty.LastVideoFrame, true);

            // Check contents of received frame
            if (canReadVideoFrame)
            {
                exposureDuration = TestDouble(VideoProperty.ExposureDuration, 0.0, double.MaxValue, false);
                exposureStartTime = TestString(VideoProperty.ExposureStartTime, int.MaxValue, false);
                frameNumber = TestLong(VideoProperty.FrameNumber, 0, long.MaxValue, true);
                imageMetadata = TestArrayList(VideoProperty.ImageMetadata, true, typeof(KeyValuePair));

                try
                {
                    imageArray = lastVideoFrame.ImageArray;
                    try
                    {
                        LogOk("ImageArray",
                            $"Received an image object from the driver of type: {imageArray.GetType().Name}");
                    }
                    catch (Exception)
                    {
                        LogInfo("ImageArray", "Received an image object from the driver of indeterminate type");
                    }

                    // Check image array dimensions
                    try
                    {
                        imageArrayAsArray = (Array)imageArray;
                        LogOk("ImageArray",
                            $"  Received an array of rank: {imageArrayAsArray.Rank}, length: {imageArrayAsArray.LongLength:#,0} and type: {imageArrayAsArray.GetType().Name}");

                        switch (imageArrayAsArray.Rank)
                        {
                            case 1 // Rank 1
                           :
                                {
                                    if (imageArrayAsArray.GetType().Equals(typeof(int[])))
                                        LogOk("ImageArray", "  Received a 1 dimension Integer array as expected.");
                                    else
                                        LogIssue("ImageArray",
                                            $"  Did not receive a 1 dimension Integer array as expected. Received: {imageArrayAsArray.GetType().Name}");
                                    break;
                                }

                            case 2 // Rank 2
                     :
                                {
                                    if (imageArrayAsArray.GetType().Equals(typeof(int[,])))
                                        LogOk("ImageArray", "  Received a 2 dimension Integer array as expected.");
                                    else
                                        LogIssue("ImageArray",
                                            $"  Did not receive a 2 dimension Integer array as expected. Received: {imageArrayAsArray.GetType().Name}");
                                    break;
                                }

                            case 3 // Rank 3
                     :
                                {
                                    if (imageArrayAsArray.GetType().Equals(typeof(int[,,])))
                                        LogOk("ImageArray", "  Received a 3 dimension Integer array as expected.");
                                    else
                                        LogIssue("ImageArray",
                                            $"  Did not receive a 3 dimension Integer array as expected. Received: {imageArrayAsArray.GetType().Name}");
                                    break;
                                }

                            default:
                                {
                                    LogIssue("ImageArray",
                                        $"  Array rank is 0 or exceeds 3: {imageArrayAsArray.GetType().Name}");
                                    break;
                                }
                        }

                        if (canReadSensorType)
                        {
                            switch (sensorType)
                            {
                                case object _ when sensorType == SensorType.Color // This camera returns multiple image planes of colour information
                               :
                                    {
                                        switch (imageArrayAsArray.Rank)
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
                                                    CheckImage(imageArrayAsArray, 3, height * width, 0);
                                                    break;
                                                }

                                            case 3 // NumPlanes x Height x Width. NumPlanes should be 3
                                     :
                                                {
                                                    CheckImage(imageArrayAsArray, 3, height, width);
                                                    break;
                                                }

                                            default:
                                                {
                                                    // This is an unsupported rank 0 or >3 so create an error
                                                    LogIssue("ImageArray",
                                                        $"  The zero based array rank must be 1, 2 or 3 . The returned array had rank: {imageArrayAsArray.Rank}");
                                                    break;
                                                }
                                        }

                                        break;
                                    }

                                default:
                                    {
                                        // This camera returns just one plane that may be literally monochrome or may appear monochrome 
                                        // but contain encoded colour information e.g. Bayer RGGB format
                                        switch (imageArrayAsArray.Rank)
                                        {
                                            case 1 // (Height * Width)
                                           :
                                                {
                                                    CheckImage(imageArrayAsArray, height * width, 0, 0);
                                                    break;
                                                }

                                            case 2 // Height x Width.
                                     :
                                                {
                                                    CheckImage(imageArrayAsArray, height, width, 0);
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
                                                    LogIssue("ImageArray",
                                                        $"  The ImageArray rank must be 1, 2 or 3. The returned array had rank: {imageArrayAsArray.Rank}");
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
                        LogIssue("ImageArray", $"Unexpected exception when testing ImageArray: {ex}");
                    }
                }
                catch (Exception ex)
                {
                    HandleException("PreviewBitmap", MemberType.Property, Required.Mandatory, ex, "");
                }

                try
                {
                    previewBitmap = lastVideoFrame.PreviewBitmap;
                    LogOk("PreviewBitmap", $"Received an array with {previewBitmap.Length:#,#,#} entries");
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
        ///     ''' <param name="testArray">The array to test</param>
        ///     ''' <param name="dimension1">Size of the first array dimension</param>
        ///     ''' <param name="dimension2">Size of the second dimension or 0 for not present</param>
        ///     ''' <param name="dimension3">Size of the third dimension or 0 for not present</param>
        ///     ''' <remarks></remarks>
        private void CheckImage(Array testArray, long dimension1, long dimension2, long dimension3)
        {
            long length;
            const string commaFormat = "#,0";

            length = dimension1 * ((dimension2 > 0) ? dimension2 : 1) * ((dimension3 > 0) ? dimension3 : 1); // Calculate the overall expected size

            if (testArray.LongLength == length)
                LogOk("CheckImage",
                    $"  ImageArray has the expected total number of pixels: {length.ToString(commaFormat)}");
            else
                LogIssue("CheckImage",
                    $"  ImageArray returned a total of {testArray.Length.ToString(commaFormat)} pixels instead of the expected number: {length.ToString(commaFormat)}");

            if (dimension1 >= 1)
            {
                if (dimension2 > 0)
                {
                    if (dimension3 >= 1)
                    {
                        if (testArray.GetLongLength(0) == dimension1)
                            LogOk("CheckImage",
                                $"  ImageArray dimension 1 has the expected length:: {dimension1.ToString(commaFormat)}");
                        else
                            LogIssue("CheckImage",
                                $"  ImageArray dimension 1 does not has the expected length:: {dimension1.ToString(commaFormat)}, received: {testArray.GetLongLength(0).ToString(commaFormat)}");
                        if (testArray.GetLongLength(1) == dimension2)
                            LogOk("CheckImage",
                                $"  ImageArray dimension 2 has the expected length:: {dimension2.ToString(commaFormat)}");
                        else
                            LogIssue("CheckImage",
                                $"  ImageArray dimension 2 does not has the expected length:: {dimension2.ToString(commaFormat)}, received: {testArray.GetLongLength(1).ToString(commaFormat)}");
                        if (testArray.GetLongLength(2) == dimension3)
                            LogOk("CheckImage",
                                $"  ImageArray dimension 3 has the expected length:: {dimension3.ToString(commaFormat)}");
                        else
                            LogIssue("CheckImage",
                                $"  ImageArray dimension 3 does not has the expected length:: {dimension3.ToString(commaFormat)}, received: {testArray.GetLongLength(2).ToString(commaFormat)}");
                    }
                    else
                    {
                        if (testArray.GetLongLength(0) == dimension1)
                            LogOk("CheckImage",
                                $"  ImageArray dimension 1 has the expected length:: {dimension1.ToString(commaFormat)}");
                        else
                            LogIssue("CheckImage",
                                $"  ImageArray dimension 1 does not has the expected length:: {dimension1.ToString(commaFormat)}, received: {testArray.GetLongLength(0).ToString(commaFormat)}");
                        if (testArray.GetLongLength(1) == dimension2)
                            LogOk("CheckImage",
                                $"  ImageArray dimension 2 has the expected length:: {dimension2.ToString(commaFormat)}");
                        else
                            LogIssue("CheckImage",
                                $"  ImageArray dimension 2 does not has the expected length:: {dimension2.ToString(commaFormat)}, received: {testArray.GetLongLength(1).ToString(commaFormat)}");
                    }
                }
                else if (testArray.GetLongLength(0) == dimension1)
                    LogOk("CheckImage",
                        $"  ImageArray dimension 1 has the expected length:: {dimension1.ToString(commaFormat)}");
                else
                    LogIssue("CheckImage",
                        $"  ImageArray dimension 1 does not has the expected length:: {dimension1.ToString(commaFormat)}, received: {testArray.GetLongLength(0).ToString(commaFormat)}");
            }
            else
                LogIssue("CheckImage", "  Dimension 1 is 0 it should never be!");
        }

        private short TestShort(VideoProperty pType, short pMin, short pMax, bool pMandatory)
        {
            string methodName;
            short returnValue = 0;

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            try
            {
                returnValue = 0;
                LogCallToDriver(pType.ToString(), $"About to get {pType} property");
                switch (pType)
                {
                    case VideoProperty.GainMax:
                        {
                            canReadGainMax = false;
                            returnValue = videoDevice.GainMax;
                            canReadGainMax = true;
                            break;
                        }

                    case VideoProperty.GainMin:
                        {
                            canReadGainMin = false;
                            returnValue = videoDevice.GainMin;
                            canReadGainMin = true;
                            break;
                        }

                    case VideoProperty.Gain:
                        {
                            returnValue = videoDevice.Gain;
                            break;
                        }

                    case VideoProperty.GammaMax:
                        {
                            canReadGammaMax = false;
                            returnValue = videoDevice.GammaMax;
                            canReadGammaMax = true;
                            break;
                        }

                    case VideoProperty.GammaMin:
                        {
                            canReadGammaMin = false;
                            returnValue = videoDevice.GammaMin;
                            canReadGammaMin = true;
                            break;
                        }

                    case VideoProperty.Gamma:
                        {
                            returnValue = videoDevice.Gamma;
                            break;
                        }

                    default:
                        {
                            LogIssue(methodName, $"TestShort: Unknown test type - {pType}");
                            break;
                        }
                }

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < pMin // Lower than minimum value
                   :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    case object _ when returnValue > pMax // Higher than maximum value
             :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    default:
                        {
                            LogOk(methodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        ///     ''' Test whether an integer is returned by a driver
        ///     ''' </summary>
        ///     ''' <param name="pType">Method to test</param>
        ///     ''' <param name="pMin">Lowest valid value</param>
        ///     ''' <param name="pMax">Highest valid value</param>
        ///     ''' <param name="pMandatory">Mandatory method</param>
        ///     ''' <returns>Integer value returned by the driver</returns>
        ///     ''' <remarks></remarks>
        private int TestInteger(VideoProperty pType, int pMin, int pMax, bool pMandatory)
        {
            string methodName;
            int returnValue = 0;

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            try
            {
                returnValue = 0;
                LogCallToDriver(pType.ToString(), $"About to get {pType} property");
                switch (pType)
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
                            canReadIntegrationRate = false;
                            returnValue = videoDevice.IntegrationRate;
                            canReadIntegrationRate = true;
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
                            LogIssue(methodName, $"TestInteger: Unknown test type - {pType}");
                            break;
                        }
                }

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < pMin // Lower than minimum value
                   :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    case object _ when returnValue > pMax // Higher than maximum value
             :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    default:
                        {
                            LogOk(methodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }
        /// <summary>
        ///     ''' Test whether an integer is returned by a driver
        ///     ''' </summary>
        ///     ''' <param name="pType">Method to test</param>
        ///     ''' <param name="pMin">Lowest valid value</param>
        ///     ''' <param name="pMax">Highest valid value</param>
        ///     ''' <param name="pMandatory">Mandatory method</param>
        ///     ''' <returns>Integer value returned by the driver</returns>
        ///     ''' <remarks></remarks>
        private long TestLong(VideoProperty pType, long pMin, long pMax, bool pMandatory)
        {
            string methodName;
            long returnValue = 0;

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            try
            {
                returnValue = 0;
                switch (pType)
                {
                    case VideoProperty.FrameNumber:
                        {
                            returnValue = lastVideoFrame.FrameNumber;
                            break;
                        }

                    default:
                        {
                            LogIssue(methodName, $"TestInteger: Unknown test type - {pType}");
                            break;
                        }
                }

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < pMin // Lower than minimum value
                   :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    case object _ when returnValue > pMax // Higher than maximum value
             :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    default:
                        {
                            LogOk(methodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        ///     ''' Test whether an integer is returned by a driver
        ///     ''' </summary>
        ///     ''' <param name="pType">Method to test</param>
        ///     ''' <param name="p_Min">Lowest valid value</param>
        ///     ''' <param name="p_Max">Highest valid value</param>
        ///     ''' <param name="pMandatory">Mandatory method</param>
        ///     ''' <returns>Integer value returned by the driver</returns>
        ///     ''' <remarks></remarks>
        private dynamic TestVideoFrame(VideoProperty pType, bool pMandatory)
        {
            string methodName;

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            dynamic returnValue = null;

            try
            {
                LogCallToDriver(pType.ToString(), $"About to get {pType} property");
                switch (pType)
                {
                    case VideoProperty.LastVideoFrame:
                        {
                            canReadVideoFrame = false;
                            returnValue = videoDevice.LastVideoFrame;
                            canReadVideoFrame = true;
                            break;
                        }

                    default:
                        {
                            LogIssue(methodName, $"TestVideoFrame: Unknown test type - {pType}");
                            break;
                        }
                }

                // Successfully retrieved a value
                LogOk(methodName, "Successfully received VideoFrame");
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        private double TestDouble(VideoProperty pType, double pMin, double pMax, bool pMandatory)
        {
            string methodName;
            double returnValue = 0.0;

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            try
            {
                returnValue = 0.0;
                LogCallToDriver(pType.ToString(), $"About to get {pType} property");
                switch (pType)
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
                            returnValue = lastVideoFrame.ExposureDuration;
                            break;
                        }

                    default:
                        {
                            LogIssue(methodName, $"TestDouble: Unknown test type - {pType}");
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue < pMin // Lower than minimum value
                   :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    case object _ when returnValue > pMax // Higher than maximum value
             :
                        {
                            LogIssue(methodName, $"Invalid value: {returnValue}");
                            break;
                        }

                    default:
                        {
                            LogOk(methodName, returnValue.ToString());
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }
        private bool TestBoolean(VideoProperty pType, bool pMandatory)
        {
            string methodName;
            bool returnValue = false;

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            try
            {
                returnValue = false;
                switch (pType)
                {
                    default:
                        {
                            LogIssue(methodName, $"TestBoolean: Unknown test type - {pType}");
                            break;
                        }
                }
                // Successfully retrieved a value
                LogOk(methodName, returnValue.ToString());
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        ///     ''' Test whether the driver returns a valid ArrayList
        ///     ''' </summary>
        ///     ''' <param name="pType">Property</param>
        ///     ''' <param name="pMandatory"></param>
        ///     ''' <param name="pItemType"></param>
        ///     ''' <returns></returns>
        ///     ''' <remarks></remarks>
        ///     
        private dynamic TestArrayList(VideoProperty pType, bool pMandatory, Type pItemType)
        {
            string methodName;
            int count;

            dynamic returnValue = new ArrayList();

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            try
            {
                LogCallToDriver(pType.ToString(), $"About to get {pType} property");
                switch (pType)
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
                            canReadSupportedIntegrationRates = false;
                            returnValue = videoDevice.SupportedIntegrationRates;
                            canReadSupportedIntegrationRates = true;
                            break;
                        }

                    case VideoProperty.ImageMetadata:
                        {
                            returnValue = lastVideoFrame.ImageMetadata;
                            break;
                        }

                    default:
                        {
                            LogIssue(methodName, $"TestArrayList: Unknown test type - {pType}");
                            break;
                        }
                }

                // Successfully retrieved a value
                count = 0;
                LogOk(methodName, "Received an array containing " + returnValue.Count + " items.");

                foreach (object listItem in returnValue)
                {
                    if (listItem.GetType().Equals(pItemType))
                        LogOk($"{methodName}({count})", $"  {listItem}");
                    else
                        LogIssue(methodName,
                            $"  Type of ArrayList item: {listItem.GetType().Name} does not match expected type: {pItemType.Name}");
                    count += 1;
                }
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }

            return returnValue;
        }

        private string TestString(VideoProperty pType, int pMaxLength, bool pMandatory)
        {
            string methodName;

            // Create a text version of the calling method name
            try
            {
                methodName = pType.ToString(); // & " Read"
            }
            catch (Exception)
            {
                methodName = "?????? Read";
            }

            string returnValue = "";
            try
            {
                LogCallToDriver(pType.ToString(), $"About to get {pType} property");
                switch (pType)
                {
                    case VideoProperty.SensorName:
                        {
                            returnValue = videoDevice.SensorName;
                            break;
                        }

                    case VideoProperty.ExposureStartTime:
                        {
                            returnValue = lastVideoFrame.ExposureStartTime;
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
                            LogIssue(methodName, $"TestString: Unknown test type - {pType}");
                            break;
                        }
                }
                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue == "":
                        {
                            LogOk(methodName, "The driver returned an empty string");
                            break;
                        }

                    default:
                        {
                            if (returnValue.Length <= pMaxLength)
                                LogOk(methodName, returnValue);
                            else
                                LogIssue(methodName,
                                    $"String exceeds {pMaxLength} characters maximum length - {returnValue}");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                HandleException(methodName, MemberType.Property, pMandatory ? Required.Mandatory : Required.Optional, ex, "");
            }
            return returnValue;
        }

        /// <summary>
        /// Not currently used in ConformU
        /// </summary>
        /// <param name="pType"></param>
        /// <param name="pProperty"></param>
        /// <param name="pTestOk"></param>
        /// <param name="p_TestLow"></param>
        /// <param name="p_TestHigh"></param>
        private void CameraPropertyWriteTest(VideoProperty pType, string pProperty, int pTestOk)
        {
            try // OK value first
            {
                switch (pType)
                {
                    case VideoProperty.BitDepth:
                        break;
                }
                LogOk($"{pProperty} write", $"Successfully wrote {pTestOk}");
            }
            catch (Exception ex)
            {
                LogIssue($"{pProperty} write",
                    $"Exception generated when setting legal value: {pTestOk} - {ex.Message}");
            }
        }

        public override void CheckMethods()
        {
        }

        public override void CheckPerformance()
        {
            CameraPerformanceTest(CameraPerformance.CameraState, "CameraState");
        }

        public override void CheckConfiguration()
        {
            try
            {
                // Common configuration
                if (!settings.TestProperties)
                    LogConfigurationAlert("Property tests were omitted due to Conform configuration.");

                if (!settings.TestMethods)
                    LogConfigurationAlert("Method tests were omitted due to Conform configuration.");

            }
            catch (Exception ex)
            {
                LogError("CheckConfiguration", $"Exception when checking Conform configuration: {ex.Message}");
                LogDebug("CheckConfiguration", $"Exception detail:\r\n:{ex}");
            }
        }

        private void CameraPerformanceTest(CameraPerformance pType, string pName)
        {
            DateTime lStartTime;
            double lCount, lLastElapsedTime, lElapsedTime, lRate;
            SetAction(pName);
            try
            {
                lStartTime = DateTime.Now;
                lCount = 0.0;
                lLastElapsedTime = 0.0;
                do
                {
                    lCount += 1.0;
                    switch (pType)
                    {
                        case CameraPerformance.CameraState:
                            {
                                cameraState = videoDevice.CameraState;
                                break;
                            }

                        default:
                            {
                                LogIssue(pName, $"Conform:PerformanceTest: Unknown test type {pType}");
                                break;
                            }
                    }

                    lElapsedTime = DateTime.Now.Subtract(lStartTime).TotalSeconds;
                    if (lElapsedTime > lLastElapsedTime + 1.0)
                    {
                        SetStatus($"{lCount} transactions in {lElapsedTime:0} seconds");
                        lLastElapsedTime = lElapsedTime;
                        if (cancellationToken.IsCancellationRequested)
                            return;
                    }
                }
                while (lElapsedTime <= PERF_LOOP_TIME);
                lRate = lCount / lElapsedTime;
                switch (lRate)
                {
                    case object _ when lRate > 10.0:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case object _ when 2.0 <= lRate && lRate <= 10.0:
                        {
                            LogOk(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    case object _ when 1.0 <= lRate && lRate <= 2.0:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }

                    default:
                        {
                            LogInfo(pName, $"Transaction rate: {lRate:0.0} per second");
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                LogInfo(pName, $"Unable to complete test: {ex}");
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
            LogOk("PostRunCheck", "Camera returned to initial cooler temperature");
        }
    }

}

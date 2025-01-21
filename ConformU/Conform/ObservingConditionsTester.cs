using ASCOM.Alpaca.Clients;
using ASCOM.Com.DriverAccess;
using ASCOM.Common;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

namespace ConformU
{

    internal class ObservingConditionsTester : DeviceTesterBaseClass
    {
        private double averageperiod, dewPoint, humidity, windDirection, windSpeed;

        // Variables to indicate whether each function is or is not implemented to that it is possible to check that for any given sensor, all three either are or are not implemented
        private readonly List<string> sensorIsImplemented = [];
        private readonly List<string> sensorHasDescription = [];
        private readonly List<string> sensorHasTimeOfLastUpdate = [];

        // Helper variables
        private IObservingConditionsV2 mObservingConditions;
        private readonly Settings settings;
        private readonly ConformLogger logger;

        #region Constants and Enums

        private const double ABSOLUTE_ZERO = -273.15;
        private const double WATER_BOILING_POINT = 100.0;
        private const double BAD_VALUE = double.NaN;

        // Valid sensor properties constants
        private const string PROPERTY_CLOUDCOVER = "CloudCover";
        private const string PROPERTY_DEWPOINT = "DewPoint";
        private const string PROPERTY_HUMIDITY = "Humidity";
        private const string PROPERTY_PRESSURE = "Pressure";
        private const string PROPERTY_RAINRATE = "RainRate";
        private const string PROPERTY_SKYBRIGHTNESS = "SkyBrightness";
        private const string PROPERTY_SKYQUALITY = "SkyQuality";
        private const string PROPERTY_SKYTEMPERATURE = "SkyTemperature";
        private const string PROPERTY_STARFWHM = "StarFWHM";
        private const string PROPERTY_TEMPERATURE = "Temperature";
        private const string PROPERTY_WINDDIRECTION = "WindDirection";
        private const string PROPERTY_WINDGUST = "WindGust";
        private const string PROPERTY_WINDSPEED = "WindSpeed";

        // Other property names
        private const string PROPERTY_AVERAGEPERIOD = "AveragePeriod";
        private const string PROPERTY_LATESTUPDATETIME = "LatestUpdateTime";
        private const string PROPERTY_TIMESINCELASTUPDATE = "TimeSinceLastUpdate";

        // List of valid ObservingConditions sensor properties
        private readonly List<string> validSensors = [
            PROPERTY_CLOUDCOVER,
            PROPERTY_DEWPOINT,
            PROPERTY_HUMIDITY,
            PROPERTY_PRESSURE,
            PROPERTY_RAINRATE,
            PROPERTY_SKYBRIGHTNESS,
            PROPERTY_SKYQUALITY,
            PROPERTY_SKYTEMPERATURE,
            PROPERTY_STARFWHM,
            PROPERTY_TEMPERATURE,
            PROPERTY_WINDDIRECTION,
            PROPERTY_WINDGUST,
            PROPERTY_WINDSPEED
        ];

        private enum ObservingConditionsProperty
        {
            // Primary properties
            AveragePeriod,
            CloudCover,
            DewPoint,
            Humidity,
            Pressure,
            RainRate,
            SkyBrightness,
            SkyQuality,
            SkyTemperature,
            StarFWHM,
            Temperature,
            WindDirection,
            WindGust,
            WindSpeed,

            // TimeSinceLastUpdate method
            TimeSinceLastUpdateLatest,
            TimeSinceLastUpdateCloudCover,
            TimeSinceLastUpdateDewPoint,
            TimeSinceLastUpdateHumidity,
            TimeSinceLastUpdatePressure,
            TimeSinceLastUpdateRainRate,
            TimeSinceLastUpdateSkyBrightness,
            TimeSinceLastUpdateSkyQuality,
            TimeSinceLastUpdateSkyTemperature,
            TimeSinceLastUpdateStarFwhm,
            TimeSinceLastUpdateTemperature,
            TimeSinceLastUpdateWindDirection,
            TimeSinceLastUpdateWindGust,
            TimeSinceLastUpdateWindSpeed,

            // SensorDescription method
            SensorDescriptionCloudCover,
            SensorDescriptionDewPoint,
            SensorDescriptionHumidity,
            SensorDescriptionPressure,
            SensorDescriptionRainRate,
            SensorDescriptionSkyBrightness,
            SensorDescriptionSkyQuality,
            SensorDescriptionSkyTemperature,
            SensorDescriptionStarFwhm,
            SensorDescriptionTemperature,
            SensorDescriptionWindDirection,
            SensorDescriptionWindGust,
            SensorDescriptionWindSpeed
        }

        #endregion

        #region New and Dispose
        public ObservingConditionsTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
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
                    mObservingConditions?.Dispose();
                    mObservingConditions = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

        #region Conform Process

        public override void InitialiseTest()
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
                            ExNotImplemented = (int)0x80040400;
                            ExInvalidValue1 = (int)0x80040405;
                            ExInvalidValue2 = (int)0x80040405;
                            ExNotSet1 = (int)0x80040403;
                            break;
                        }
                }
            }
            base.InitialiseTest();
        }

        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        LogInfo("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        mObservingConditions = new AlpacaObservingConditions(
                                                    settings.AlpacaConfiguration.AccessServiceType,
                                                    settings.AlpacaDevice.IpAddress,
                                                    settings.AlpacaDevice.IpPort,
                                                    settings.AlpacaDevice.AlpacaDeviceNumber,
                                                    settings.AlpacaConfiguration.EstablishConnectionTimeout,
                                                    settings.AlpacaConfiguration.StandardResponseTimeout,
                                                    settings.AlpacaConfiguration.LongResponseTimeout,
                                                    Globals.CLIENT_NUMBER_DEFAULT,
                                                    settings.AlpacaConfiguration.AccessUserName,
                                                    settings.AlpacaConfiguration.AccessPassword,
                                                    settings.AlpacaConfiguration.StrictCasing,
                                                    settings.TraceAlpacaCalls ? logger : null,
                                                    Globals.USER_AGENT_PRODUCT_NAME,
                                                    Assembly.GetExecutingAssembly().GetName().Version.ToString(4),
                                                    settings.AlpacaConfiguration.TrustUserGeneratedSslCertificates);

                        LogInfo("CreateDevice", $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComAccessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                mObservingConditions = new ObservingConditionsFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                mObservingConditions = new ObservingConditions(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComAccessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                SetDevice(mObservingConditions, DeviceTypes.ObservingConditions); // Assign the driver to the base class

                SetFullStatus("Create device", "Waiting for driver to stabilise", "");
                WaitFor(1000, 100);

                // Validate the interface version
                ValidateInterfaceVersion();
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

        public override void CheckProperties()
        {
            averageperiod = TestDouble(PROPERTY_AVERAGEPERIOD, ObservingConditionsProperty.AveragePeriod, 0.0, 100000.0, Required.Mandatory); // AveragePeriod is mandatory

            // AveragePeriod Write - Mandatory
            if (IsGoodValue(averageperiod))
            {
                // Try setting a bad value i.e. less than -1, this should be rejected
                try // Invalid low value
                {
                    LogCallToDriver("AveragePeriod Write", "About to set AveragePeriod property");
                    mObservingConditions.AveragePeriod = -2.0;
                    LogIssue("AveragePeriod Write", "No error generated on setting the average period < -1.0");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "", "Invalid Value exception generated as expected on set average period < -1.0");
                }

                // Try setting a good zero value, this should be accepted
                try // Invalid low value
                {
                    LogCallToDriver("AveragePeriod Write", "About to set AveragePeriod property");
                    mObservingConditions.AveragePeriod = 0.0;
                    LogOk("AveragePeriod Write", "Successfully set average period to 0.0");
                }
                catch (Exception ex)
                {
                    HandleException("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "");
                }

                // Try setting a good positive value, this should be accepted
                try // Invalid low value
                {
                    LogCallToDriver("AveragePeriod Write", "About to set AveragePeriod property");
                    mObservingConditions.AveragePeriod = 5.0;
                    LogOk("AveragePeriod Write", "Successfully set average period to 5.0");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOk("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "", "Invalid value exception thrown when average period is set to 5.0, which is permitted by the specification");
                }

                // Restore original value, this should be accepted
                try // Invalid low value
                {
                    LogCallToDriver("AveragePeriod Write", "About to set AveragePeriod property");
                    mObservingConditions.AveragePeriod = averageperiod;
                    LogOk("AveragePeriod Write", $"Successfully restored original average period: {averageperiod}");
                }
                catch (Exception ex)
                {
                    LogInfo("AveragePeriod Write", "Unable to restore original average period");
                    HandleException("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
                LogInfo("AveragePeriod Write", "Test skipped because AveragePerid cold not be read");

            TestDouble(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.CloudCover, 0.0, 100.0, Required.Optional);
            dewPoint = TestDouble(PROPERTY_DEWPOINT, ObservingConditionsProperty.DewPoint, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional);
            humidity = TestDouble(PROPERTY_HUMIDITY, ObservingConditionsProperty.Humidity, 0.0, 100.0, Required.Optional);

            if ((IsGoodValue(dewPoint) & IsGoodValue(humidity)))
                LogOk("DewPoint & Humidity", "Dew point and humidity are both implemented per the interface specification");
            else if ((!IsGoodValue(dewPoint) & !IsGoodValue(humidity)))
                LogOk("DewPoint & Humidity", "Dew point and humidity are both not implemented per the interface specification");
            else
                LogIssue("DewPoint & Humidity", "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            TestDouble(PROPERTY_PRESSURE, ObservingConditionsProperty.Pressure, 0.0, 1100.0, Required.Optional);
            TestDouble(PROPERTY_RAINRATE, ObservingConditionsProperty.RainRate, 0.0, 20000.0, Required.Optional);
            TestDouble(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.SkyBrightness, 0.0, 1000000.0, Required.Optional);
            TestDouble(PROPERTY_SKYQUALITY, ObservingConditionsProperty.SkyQuality, -20.0, 30.0, Required.Optional);
            TestDouble(PROPERTY_STARFWHM, ObservingConditionsProperty.StarFWHM, 0.0, 1000.0, Required.Optional);
            TestDouble(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.SkyTemperature, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional);
            TestDouble(PROPERTY_TEMPERATURE, ObservingConditionsProperty.Temperature, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional);
            windDirection = TestDouble(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.WindDirection, 0.0, 360.0, Required.Optional);
            TestDouble(PROPERTY_WINDGUST, ObservingConditionsProperty.WindGust, 0.0, 1000.0, Required.Optional);
            windSpeed = TestDouble(PROPERTY_WINDSPEED, ObservingConditionsProperty.WindSpeed, 0.0, 1000.0, Required.Optional);

            // Additional test to confirm that wind direction is reported as 0.0 if wind speed is reported as 0.0
            if ((windSpeed == 0.0)) // Wind speed is reported as 0.0
            {
                // Check whether wind direction is implemented
                if (IsGoodValue(windDirection)) // Wind direction is implemented
                {
                    // Check whether wind direction is reported as 0.0
                    if (windDirection == 0.0)
                        LogOk(PROPERTY_WINDSPEED, "Wind direction is reported as 0.0 when wind speed is 0.0");
                    else
                        LogIssue(PROPERTY_WINDSPEED, $"When wind speed is reported as 0.0, wind direction should also be reported as 0.0, it is actually reported as {windDirection}");
                }
            }
        }

        public override void CheckMethods()
        {
            // TimeSinceLastUpdate
            TestDouble(PROPERTY_LATESTUPDATETIME, ObservingConditionsProperty.TimeSinceLastUpdateLatest, -1.0, double.MaxValue, Required.Mandatory);

            TestDouble(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.TimeSinceLastUpdateCloudCover, -1.0, double.MaxValue, Required.Optional);
            double lastUpdateTimeDewPoint = TestDouble(PROPERTY_DEWPOINT, ObservingConditionsProperty.TimeSinceLastUpdateDewPoint, -1.0, double.MaxValue, Required.Optional);
            double lastUpdateTimeHumidity = TestDouble(PROPERTY_HUMIDITY, ObservingConditionsProperty.TimeSinceLastUpdateHumidity, -1.0, double.MaxValue, Required.Optional);

            if ((IsGoodValue(lastUpdateTimeDewPoint) & IsGoodValue(lastUpdateTimeHumidity)))
                LogOk("DewPoint & Humidity", "Dew point and humidity are both implemented per the interface specification");
            else if ((!IsGoodValue(lastUpdateTimeDewPoint) & !IsGoodValue(lastUpdateTimeHumidity)))
                LogOk("DewPoint & Humidity", "Dew point and humidity are both not implemented per the interface specification");
            else
                LogIssue("DewPoint & Humidity", "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            // Test property values
            TestDouble(PROPERTY_PRESSURE, ObservingConditionsProperty.TimeSinceLastUpdatePressure, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_RAINRATE, ObservingConditionsProperty.TimeSinceLastUpdateRainRate, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.TimeSinceLastUpdateSkyBrightness, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_SKYQUALITY, ObservingConditionsProperty.TimeSinceLastUpdateSkyQuality, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_STARFWHM, ObservingConditionsProperty.TimeSinceLastUpdateStarFwhm, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateSkyTemperature, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_TEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateTemperature, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.TimeSinceLastUpdateWindDirection, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_WINDGUST, ObservingConditionsProperty.TimeSinceLastUpdateWindGust, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_WINDSPEED, ObservingConditionsProperty.TimeSinceLastUpdateWindSpeed, double.MinValue, double.MaxValue, Required.Optional);

            // Test the Refresh method
            try
            {
                LogCallToDriver("AveragePeriod Write", "About to call Refresh method");
                mObservingConditions.Refresh();
                LogOk("Refresh", "Refreshed OK");
            }
            catch (Exception ex)
            {
                HandleException("Refresh", MemberType.Method, Required.Optional, ex, "");
            }

            // Test SensorDescrtiption
            TestSensorDescription(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.SensorDescriptionCloudCover, int.MaxValue, Required.Optional);
            string sensorDescriptionDewPoint = TestSensorDescription(PROPERTY_DEWPOINT, ObservingConditionsProperty.SensorDescriptionDewPoint, int.MaxValue, Required.Optional);
            string sensorDescriptionHumidity = TestSensorDescription(PROPERTY_HUMIDITY, ObservingConditionsProperty.SensorDescriptionHumidity, int.MaxValue, Required.Optional);

            if ((sensorDescriptionDewPoint is null) & (sensorDescriptionHumidity is null))
                LogOk("DewPoint & Humidity", "Dew point and humidity are both not implemented per the interface specification");
            else if ((sensorDescriptionDewPoint is not null) & (sensorDescriptionHumidity is not null))
                LogOk("DewPoint & Humidity", "Dew point and humidity are both implemented per the interface specification");
            else
                LogIssue("DewPoint & Humidity", "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            TestSensorDescription(PROPERTY_PRESSURE, ObservingConditionsProperty.SensorDescriptionPressure, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_RAINRATE, ObservingConditionsProperty.SensorDescriptionRainRate, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.SensorDescriptionSkyBrightness, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_SKYQUALITY, ObservingConditionsProperty.SensorDescriptionSkyQuality, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_STARFWHM, ObservingConditionsProperty.SensorDescriptionStarFwhm, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.SensorDescriptionSkyTemperature, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_TEMPERATURE, ObservingConditionsProperty.SensorDescriptionTemperature, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.SensorDescriptionWindDirection, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_WINDGUST, ObservingConditionsProperty.SensorDescriptionWindGust, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_WINDSPEED, ObservingConditionsProperty.SensorDescriptionWindSpeed, int.MaxValue, Required.Optional);

            // Now check that the sensor value, description and last updated time are all either implemented or not implemented
            foreach (string sensorName in validSensors)
            {
                LogDebug("Consistency", $"Sensor name: {sensorName}");
                if (sensorIsImplemented.Contains(sensorName) & sensorHasDescription.Contains(sensorName) & sensorHasTimeOfLastUpdate.Contains(sensorName))
                    LogOk($"Consistency - {sensorName}", "Sensor value, description and time since last update are all implemented as required by the specification");
                else if (!sensorIsImplemented.Contains(sensorName) & !sensorHasDescription.Contains(sensorName) & !sensorHasTimeOfLastUpdate.Contains(sensorName))
                    LogOk($"Consistency - {sensorName}", "Sensor value, description and time since last update are all not implemented as required by the specification");
                else
                {
                    LogIssue($"Consistency - {sensorName}",
                        $"Sensor {sensorName} value is implemented: {sensorIsImplemented.Contains(sensorName)}, Sensor description is implemented: {sensorHasDescription.Contains(sensorName)}, Sensor time since last update is implemented: {sensorHasTimeOfLastUpdate.Contains(sensorName)}");
                    LogInfo($"Consistency - {sensorName}", "The ASCOM specification requires that sensor value, description and time since last update must either all be implemented or all not be implemented.");

                    foreach (string sn in sensorIsImplemented)
                    {
                        LogDebug(sensorName, $"Sensor is implemented: {sn}");
                    }
                }
            }
        }

        public override void CheckPerformance()
        {
            SetTest("");
            SetAction("");
            SetStatus("");
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

        #endregion

        #region Support Code

        /// <summary>
        /// Tests whether a double has a good value or the NaN bad value indicator
        /// </summary>
        /// <param name="value">Variable to be tested</param>
        /// <returns>Returns True if the variable has a good value, otherwise returns False</returns>
        /// <remarks></remarks>
        private static bool IsGoodValue(double value)
        {
            return !double.IsNaN(value);
        }

        private double TestDouble(string propertyName, ObservingConditionsProperty enumName, double pMin, double pMax, Required pMandatory)
        {
            double returnValue;
            int retryCount = 0;
            bool readOk = false;
            MemberType methodType = MemberType.Property;
            bool unexpectedError = false;

            do
            {
                try
                {
                    returnValue = BAD_VALUE;

                    TimeMethod(enumName.ToString(), () =>
                    {
                        switch (enumName)
                        {
                            #region Primary properties

                            case ObservingConditionsProperty.AveragePeriod:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.AveragePeriod;
                                break;

                            case ObservingConditionsProperty.CloudCover:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.CloudCover;
                                break;

                            case ObservingConditionsProperty.DewPoint:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.DewPoint;
                                break;

                            case ObservingConditionsProperty.Humidity:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.Humidity;
                                break;

                            case ObservingConditionsProperty.Pressure:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.Pressure;
                                break;

                            case ObservingConditionsProperty.RainRate:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.RainRate;
                                break;

                            case ObservingConditionsProperty.SkyBrightness:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.SkyBrightness;
                                break;

                            case ObservingConditionsProperty.SkyQuality:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.SkyQuality;
                                break;

                            case ObservingConditionsProperty.StarFWHM:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.StarFWHM;
                                break;

                            case ObservingConditionsProperty.SkyTemperature:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.SkyTemperature;
                                break;

                            case ObservingConditionsProperty.Temperature:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.Temperature;
                                break;

                            case ObservingConditionsProperty.WindDirection:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.WindDirection;
                                break;

                            case ObservingConditionsProperty.WindGust:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.WindGust;
                                break;

                            case ObservingConditionsProperty.WindSpeed:
                                methodType = MemberType.Property;
                                returnValue = mObservingConditions.WindSpeed;
                                break;

                            #endregion

                            #region Time since last update

                            case ObservingConditionsProperty.TimeSinceLastUpdateLatest:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate("");
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateCloudCover:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_CLOUDCOVER);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateDewPoint:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_DEWPOINT);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateHumidity:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_HUMIDITY);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdatePressure:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_PRESSURE);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateRainRate:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_RAINRATE);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateSkyBrightness:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYBRIGHTNESS);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateSkyQuality:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYQUALITY);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateStarFwhm:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_STARFWHM);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateSkyTemperature:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYTEMPERATURE);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateTemperature:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_TEMPERATURE);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateWindDirection:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDDIRECTION);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateWindGust:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDGUST);
                                break;

                            case ObservingConditionsProperty.TimeSinceLastUpdateWindSpeed:
                                methodType = MemberType.Method;
                                returnValue = mObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDSPEED);
                                break;

                            #endregion

                            default:
                                LogError(propertyName, $"TestDouble - Unknown test type - {enumName}");
                                break;
                        }
                    }, TargetTime.Fast);
                    readOk = true;

                    // Successfully retrieved a value so check validity
                    switch (returnValue)
                    {
                        case var _ when returnValue < pMin:
                            LogIssue(propertyName, $"Invalid value (below minimum expected - {pMin}): {returnValue}");
                            break;

                        case var _ when returnValue > pMax:
                            LogIssue(propertyName, $"Invalid value (above maximum expected - {pMax}): {returnValue}");
                            break;

                        default:
                            LogOk(enumName.ToString(), returnValue.ToString());
                            break;
                    }

                    // Record that a value was returned
                    if (enumName.ToString().StartsWith(PROPERTY_TIMESINCELASTUPDATE))
                    {
                        sensorHasTimeOfLastUpdate.Add(propertyName);
                        LogDebug(enumName.ToString(), $"Added {propertyName} to sensorHasTimeOfLastUpdate list");
                    }
                    else
                    {
                        sensorIsImplemented.Add(propertyName);
                        LogDebug(propertyName, $"Added {propertyName} to sensorIsImplemented list");
                    }
                }
                catch (Exception ex)
                {
                    if (IsInvalidOperationException(propertyName, ex))
                    {
                        returnValue = BAD_VALUE;
                        retryCount += 1;
                        LogInfo(propertyName, $"Sensor not ready, received InvalidOperationException, waiting {settings.ObservingConditionsRetryTime} second to retry. Attempt {retryCount} out of {settings.ObservingConditionsMaxRetries}");
                        WaitFor(settings.ObservingConditionsRetryTime * 1000);
                    }
                    else
                    {
                        unexpectedError = true;
                        returnValue = BAD_VALUE;
                        HandleException(propertyName, methodType, pMandatory, ex, "");
                    }
                }
            }
            while (!readOk & (retryCount <= settings.ObservingConditionsMaxRetries) & !unexpectedError); // Lower than minimum value// Higher than maximum value

            if ((!readOk) & (!unexpectedError))
                LogInfo(propertyName, $"InvalidOperationException persisted for longer than {settings.ObservingConditionsMaxRetries * settings.ObservingConditionsRetryTime} seconds.");
            return returnValue;

        }

        private string TestSensorDescription(string propertyName, ObservingConditionsProperty enumName, int pMaxLength, Required pMandatory)
        {
            string returnValue = "";
            try
            {
                LogCallToDriver(enumName.ToString(), $"About to call SensorDescription({propertyName}) method");
                TimeMethod(enumName.ToString(), () => returnValue = mObservingConditions.SensorDescription(propertyName), TargetTime.Fast);

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case null:
                        {
                            LogIssue(enumName.ToString(), "The driver did not return any string at all: Nothing (VB), null (C#)");
                            break;
                        }

                    case "":
                        {
                            LogOk(enumName.ToString(), "The driver returned an empty string: \"\"");
                            break;
                        }

                    default:
                        {
                            if (returnValue.Length <= pMaxLength)
                                LogOk(enumName.ToString(), returnValue);
                            else
                                LogIssue(enumName.ToString(), $"String exceeds {pMaxLength} characters maximum length - {returnValue}");
                            break;
                        }
                }

                sensorHasDescription.Add(propertyName);
            }
            catch (Exception ex)
            {
                HandleException(enumName.ToString(), MemberType.Method, pMandatory, ex, "");
                returnValue = null;
            }
            return returnValue;
        }

        #endregion

    }
}

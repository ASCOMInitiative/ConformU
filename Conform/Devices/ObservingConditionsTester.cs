using System;
using System.Collections.Generic;
using ASCOM.Standard.Interfaces;
using ASCOM.Alpaca.Clients;
using ASCOM.Standard.COM.DriverAccess;
using System.Threading;

namespace ConformU
{

    internal class ObservingConditionsTester : DeviceTesterBaseClass
    {
        private double averageperiod, dewPoint, humidity, windDirection, windSpeed;

        // Variables to indicate whether each function is or is not implemented to that it is possible to check that for any given sensor, all three either are or are not implemented
        private readonly Dictionary<string, bool> sensorIsImplemented = new();
        private readonly Dictionary<string, bool> sensorHasDescription = new();
        private readonly Dictionary<string, bool> sensorHasTimeOfLastUpdate = new();

        const double ABSOLUTE_ZERO = -273.15;
        const double WATER_BOILING_POINT = 100.0;
        const double BAD_VALUE = double.NaN;

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
        private readonly List<string> ValidSensors = new() { PROPERTY_CLOUDCOVER, PROPERTY_DEWPOINT, PROPERTY_HUMIDITY, PROPERTY_PRESSURE, PROPERTY_RAINRATE, PROPERTY_SKYBRIGHTNESS, PROPERTY_SKYQUALITY, PROPERTY_SKYTEMPERATURE, PROPERTY_STARFWHM, PROPERTY_TEMPERATURE, PROPERTY_WINDDIRECTION, PROPERTY_WINDGUST, PROPERTY_WINDSPEED };

        // Helper variables
        private IObservingConditions m_ObservingConditions;
        private readonly Settings settings;
        private readonly ConformLogger logger;

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
            TimeSinceLastUpdateStarFWHM,
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
            SensorDescriptionStarFWHM,
            SensorDescriptionTemperature,
            SensorDescriptionWindDirection,
            SensorDescriptionWindGust,
            SensorDescriptionWindSpeed
        }

        #region New and Dispose
        public ObservingConditionsTester(ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            this.logger = logger;
            foreach (string sensorName in ValidSensors)
            {
                sensorIsImplemented[sensorName] = false;
                sensorHasDescription[sensorName] = false;
                sensorHasTimeOfLastUpdate[sensorName] = false;

            }

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
                    if (m_ObservingConditions is not null) m_ObservingConditions.Dispose();
                    m_ObservingConditions = null;
                }
            }

            // Call the DeviceTesterBaseClass dispose method
            base.Dispose(disposing);
            disposedValue = true;
        }

        #endregion

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
                        LogInfo("CreateDevice", $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_ObservingConditions = new AlpacaObservingConditions(settings.AlpacaConfiguration.AccessServiceType,
                            settings.AlpacaDevice.IpAddress,
                            settings.AlpacaDevice.IpPort,
                            settings.AlpacaDevice.AlpacaDeviceNumber,
                            settings.AlpacaConfiguration.StrictCasing,
                            settings.TraceAlpacaCalls ? logger : null);
                        LogInfo("CreateDevice", $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                LogInfo("CreateDevice", $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_ObservingConditions = new ObservingConditionsFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                LogInfo("CreateDevice", $"Creating DRIVERACCESS device: {settings.ComDevice.ProgId}");
                                m_ObservingConditions = new ObservingConditions(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogInfo("CreateDevice", "Successfully created driver");
                baseClassDevice = m_ObservingConditions; // Assign the driver to the base class

                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");

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
                LogCallToDriver("Absolute", "About to get Connected property");
                return m_ObservingConditions.Connected;
            }
            set
            {
                LogCallToDriver("Absolute", "About to set Connected property");
                m_ObservingConditions.Connected = value;
            }
        }

        public override void CheckCommonMethods()
        {
            base.CheckCommonMethods(m_ObservingConditions, DeviceType.ObservingConditions);
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
                    m_ObservingConditions.AveragePeriod = -2.0;
                    LogIssue("AveragePeriod Write", "No error generated on setting the average period < -1.0");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOK("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "", "Invalid Value exception generated as expected on set average period < -1.0");
                }

                // Try setting a good zero value, this should be accepted
                try // Invalid low value
                {
                    LogCallToDriver("AveragePeriod Write", "About to set AveragePeriod property");
                    m_ObservingConditions.AveragePeriod = 0.0;
                    LogOK("AveragePeriod Write", "Successfully set average period to 0.0");
                }
                catch (Exception ex)
                {
                    HandleException("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "");
                }

                // Try setting a good positive value, this should be accepted
                try // Invalid low value
                {
                    LogCallToDriver("AveragePeriod Write", "About to set AveragePeriod property");
                    m_ObservingConditions.AveragePeriod = 5.0;
                    LogOK("AveragePeriod Write", "Successfully set average period to 5.0");
                }
                catch (Exception ex)
                {
                    HandleInvalidValueExceptionAsOK("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "", "Invalid value exception thrown when average period is set to 5.0, which is permitted by the specification");
                }

                // Restore original value, this should be accepted
                try // Invalid low value
                {
                    LogCallToDriver("AveragePeriod Write", "About to set AveragePeriod property");
                    m_ObservingConditions.AveragePeriod = averageperiod;
                    LogOK("AveragePeriod Write", "Successfully restored original average period: " + averageperiod);
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
                LogOK("DewPoint & Humidity", "Dew point and humidity are both implemented per the interface specification");
            else if ((!IsGoodValue(dewPoint) & !IsGoodValue(humidity)))
                LogOK("DewPoint & Humidity", "Dew point and humidity are both not implemented per the interface specification");
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

            // Additional test to confirm that the reported direction is 0.0 if the wind speed is reported as 0.0
            if ((windSpeed == 0.0))
            {
                if ((windDirection == 0.0))
                    LogOK(PROPERTY_WINDSPEED, "Wind direction is reported as 0.0 when wind speed is 0.0");
                else
                    LogIssue(PROPERTY_WINDSPEED, string.Format("When wind speed is reported as 0.0, wind direction should also be reported as 0.0, it is actually reported as {0}", windDirection));
            }
        }

        public override void CheckMethods()
        {
            double LastUpdateTimeDewPoint, LastUpdateTimeHumidity;

            string SensorDescriptionDewPoint, SensorDescriptionHumidity;

            // TimeSinceLastUpdate
            TestDouble(PROPERTY_LATESTUPDATETIME, ObservingConditionsProperty.TimeSinceLastUpdateLatest, -1.0, double.MaxValue, Required.Mandatory);

            TestDouble(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.TimeSinceLastUpdateCloudCover, -1.0, double.MaxValue, Required.Optional);
            LastUpdateTimeDewPoint = TestDouble(PROPERTY_DEWPOINT, ObservingConditionsProperty.TimeSinceLastUpdateDewPoint, -1.0, double.MaxValue, Required.Optional);
            LastUpdateTimeHumidity = TestDouble(PROPERTY_HUMIDITY, ObservingConditionsProperty.TimeSinceLastUpdateHumidity, -1.0, double.MaxValue, Required.Optional);

            if ((IsGoodValue(LastUpdateTimeDewPoint) & IsGoodValue(LastUpdateTimeHumidity)))
                LogOK("DewPoint & Humidity", "Dew point and humidity are both implemented per the interface specification");
            else if ((!IsGoodValue(LastUpdateTimeDewPoint) & !IsGoodValue(LastUpdateTimeHumidity)))
                LogOK("DewPoint & Humidity", "Dew point and humidity are both not implemented per the interface specification");
            else
                LogIssue("DewPoint & Humidity", "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            TestDouble(PROPERTY_PRESSURE, ObservingConditionsProperty.TimeSinceLastUpdatePressure, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_RAINRATE, ObservingConditionsProperty.TimeSinceLastUpdateRainRate, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.TimeSinceLastUpdateSkyBrightness, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_SKYQUALITY, ObservingConditionsProperty.TimeSinceLastUpdateSkyQuality, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_STARFWHM, ObservingConditionsProperty.TimeSinceLastUpdateStarFWHM, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateSkyTemperature, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_TEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateTemperature, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.TimeSinceLastUpdateWindDirection, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_WINDGUST, ObservingConditionsProperty.TimeSinceLastUpdateWindGust, double.MinValue, double.MaxValue, Required.Optional);
            TestDouble(PROPERTY_WINDSPEED, ObservingConditionsProperty.TimeSinceLastUpdateWindSpeed, double.MinValue, double.MaxValue, Required.Optional);

            // Refresh
            try
            {
                LogCallToDriver("AveragePeriod Write", "About to call Refresh method");
                m_ObservingConditions.Refresh();
                LogOK("Refresh", "Refreshed OK");
            }
            catch (Exception ex)
            {
                HandleException("Refresh", MemberType.Method, Required.Optional, ex, "");
            }

            // SensorDescrtiption
            TestSensorDescription(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.SensorDescriptionCloudCover, int.MaxValue, Required.Optional);
            SensorDescriptionDewPoint = TestSensorDescription(PROPERTY_DEWPOINT, ObservingConditionsProperty.SensorDescriptionDewPoint, int.MaxValue, Required.Optional);
            SensorDescriptionHumidity = TestSensorDescription(PROPERTY_HUMIDITY, ObservingConditionsProperty.SensorDescriptionHumidity, int.MaxValue, Required.Optional);

            if (((SensorDescriptionDewPoint == null) & (SensorDescriptionHumidity == null)))
                LogOK("DewPoint & Humidity", "Dew point and humidity are both not implemented per the interface specification");
            else if (((!(SensorDescriptionDewPoint == null)) & (!(SensorDescriptionHumidity == null))))
                LogOK("DewPoint & Humidity", "Dew point and humidity are both implemented per the interface specification");
            else
                LogIssue("DewPoint & Humidity", "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            TestSensorDescription(PROPERTY_PRESSURE, ObservingConditionsProperty.SensorDescriptionPressure, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_RAINRATE, ObservingConditionsProperty.SensorDescriptionRainRate, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.SensorDescriptionSkyBrightness, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_SKYQUALITY, ObservingConditionsProperty.SensorDescriptionSkyQuality, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_STARFWHM, ObservingConditionsProperty.SensorDescriptionStarFWHM, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.SensorDescriptionSkyTemperature, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_TEMPERATURE, ObservingConditionsProperty.SensorDescriptionTemperature, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.SensorDescriptionWindDirection, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_WINDGUST, ObservingConditionsProperty.SensorDescriptionWindGust, int.MaxValue, Required.Optional);
            TestSensorDescription(PROPERTY_WINDSPEED, ObservingConditionsProperty.SensorDescriptionWindSpeed, int.MaxValue, Required.Optional);

            // Now check that the sensor value, description and last updated time are all either implemented or not implemented
            foreach (string sensorName in ValidSensors)
            {
                LogDebug("Consistency", "Sensor name: " + sensorName);
                if ((sensorIsImplemented[sensorName] & sensorHasDescription[sensorName] & sensorHasTimeOfLastUpdate[sensorName]))
                    LogOK("Consistency - " + sensorName, "Sensor value, description and time since last update are all implemented as required by the specification");
                else if (((!sensorIsImplemented[sensorName]) & (!sensorHasDescription[sensorName]) & (!sensorHasTimeOfLastUpdate[sensorName])))
                    LogOK("Consistency - " + sensorName, "Sensor value, description and time since last update are all not implemented as required by the specification");
                else
                {
                    LogIssue("Consistency - " + sensorName, "Sensor value is implemented: " + sensorIsImplemented[sensorName] + ", Sensor description is implemented: " + sensorHasDescription[sensorName] + ", Sensor time since last update is implemented: " + sensorHasTimeOfLastUpdate[sensorName]);
                    LogInfo("Consistency - " + sensorName, "The ASCOM specification requires that sensor value, description and time since last update must either all be implemented or all not be implemented.");
                }
            }
        }

        public override void CheckPerformance()
        {
            Status(StatusType.staTest, "");
            Status(StatusType.staAction, "");
            Status(StatusType.staStatus, "");
        }

        /// <summary>
        ///     ''' Tests whether a double has a good value or the NaN bad value indicator
        ///     ''' </summary>
        ///     ''' <param name="value">Variable to be tested</param>
        ///     ''' <returns>Returns True if the variable has a good value, otherwise returns False</returns>
        ///     ''' <remarks></remarks>
        private static bool IsGoodValue(double value)
        {
            return !double.IsNaN(value);
        }

        private double TestDouble(string p_Nmae, ObservingConditionsProperty p_Type, double p_Min, double p_Max, Required p_Mandatory)
        {
            string MethodName, SensorName;
            double returnValue;
            int retryCount = 0;
            bool readOK = false;
            MemberType methodType = MemberType.Property;
            bool unexpectedError = false;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }
            if (MethodName.StartsWith(PROPERTY_TIMESINCELASTUPDATE))
            {
                SensorName = MethodName[PROPERTY_TIMESINCELASTUPDATE.Length..];
                LogCallToDriver(MethodName, $"About to call TimeSinceLastUpdate({SensorName}) method");
            }
            else
            {
                SensorName = MethodName;
                LogCallToDriver(MethodName, $"About to get {SensorName} property");
            }
            LogDebug("returnValue", "methodName: " + MethodName + ", SensorName: " + SensorName);
            sensorHasTimeOfLastUpdate[SensorName] = false;

            do
            {
                try
                {
                    returnValue = BAD_VALUE;
                    switch (p_Type)
                    {
                        case ObservingConditionsProperty.AveragePeriod:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.AveragePeriod;
                                break;
                            }

                        case ObservingConditionsProperty.CloudCover:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.CloudCover;
                                break;
                            }

                        case ObservingConditionsProperty.DewPoint:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.DewPoint;
                                break;
                            }

                        case ObservingConditionsProperty.Humidity:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.Humidity;
                                break;
                            }

                        case ObservingConditionsProperty.Pressure:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.Pressure;
                                break;
                            }

                        case ObservingConditionsProperty.RainRate:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.RainRate;
                                break;
                            }

                        case ObservingConditionsProperty.SkyBrightness:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.SkyBrightness;
                                break;
                            }

                        case ObservingConditionsProperty.SkyQuality:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.SkyQuality;
                                break;
                            }

                        case ObservingConditionsProperty.StarFWHM:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.StarFWHM;
                                break;
                            }

                        case ObservingConditionsProperty.SkyTemperature:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.SkyTemperature;
                                break;
                            }

                        case ObservingConditionsProperty.Temperature:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.Temperature;
                                break;
                            }

                        case ObservingConditionsProperty.WindDirection:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.WindDirection;
                                break;
                            }

                        case ObservingConditionsProperty.WindGust:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.WindGust;
                                break;
                            }

                        case ObservingConditionsProperty.WindSpeed:
                            {
                                methodType = MemberType.Property; returnValue = m_ObservingConditions.WindSpeed;
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateLatest:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate("");
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateCloudCover:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_CLOUDCOVER);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateDewPoint:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_DEWPOINT);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateHumidity:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_HUMIDITY);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdatePressure:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_PRESSURE);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateRainRate:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_RAINRATE);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateSkyBrightness:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYBRIGHTNESS);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateSkyQuality:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYQUALITY);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateStarFWHM:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_STARFWHM);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateSkyTemperature:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_SKYTEMPERATURE);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateTemperature:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_TEMPERATURE);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateWindDirection:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDDIRECTION);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateWindGust:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDGUST);
                                break;
                            }

                        case ObservingConditionsProperty.TimeSinceLastUpdateWindSpeed:
                            {
                                methodType = MemberType.Method; returnValue = m_ObservingConditions.TimeSinceLastUpdate(PROPERTY_WINDSPEED);
                                break;
                            }

                        default:
                            {
                                LogIssue(MethodName, "returnValue: Unknown test type - " + p_Type.ToString());
                                break;
                            }
                    }

                    readOK = true;

                    // Successfully retrieved a value so check validity
                    switch (returnValue)
                    {
                        case object _ when returnValue < p_Min:
                            {
                                LogIssue(MethodName, "Invalid value (below minimum expected - " + p_Min.ToString() + "): " + returnValue.ToString());
                                break;
                            }

                        case object _ when returnValue > p_Max:
                            {
                                LogIssue(MethodName, "Invalid value (above maximum expected - " + p_Max.ToString() + "): " + returnValue.ToString());
                                break;
                            }

                        default:
                            {
                                LogOK(MethodName, returnValue.ToString());
                                break;
                            }
                    }

                    if (MethodName.StartsWith(PROPERTY_TIMESINCELASTUPDATE))
                        sensorHasTimeOfLastUpdate[SensorName] = true;
                    else
                        sensorIsImplemented[SensorName] = true;
                }
                catch (Exception ex)
                {
                    if (IsInvalidOperationException(p_Nmae, ex))
                    {
                        returnValue = BAD_VALUE;
                        retryCount += 1;
                        LogInfo(MethodName, "Sensor not ready, received InvalidOperationException, waiting " + settings.ObservingConditionsRetryTime + " second to retry. Attempt " + retryCount + " out of " + settings.ObservingConditionsMaxRetries);
                        WaitFor(settings.ObservingConditionsRetryTime * 1000);
                    }
                    else
                    {
                        unexpectedError = true;
                        returnValue = BAD_VALUE;
                        HandleException(MethodName, methodType, p_Mandatory, ex, "");
                    }
                }
            }
            while (!readOK & (retryCount <= settings.ObservingConditionsMaxRetries) & !unexpectedError); // Lower than minimum value// Higher than maximum value

            if ((!readOK) & (!unexpectedError))
                LogInfo(MethodName, "InvalidOperationException persisted for longer than " + settings.ObservingConditionsMaxRetries * settings.ObservingConditionsRetryTime + " seconds.");
            return returnValue;

        }

        private string TestSensorDescription(string p_Name, ObservingConditionsProperty p_Type, int p_MaxLength, Required p_Mandatory)
        {
            string MethodName, returnValue;

            // Create a text version of the calling method name
            try
            {
                MethodName = p_Type.ToString(); // & " Read"
            }
            catch (Exception)
            {
                MethodName = "?????? Read";
            }

            sensorHasDescription[p_Name] = false;

            returnValue = null;
            try
            {
                LogCallToDriver(MethodName, $"About to call SensorDescription({p_Name}) method");
                switch (p_Type)
                {
                    case ObservingConditionsProperty.SensorDescriptionCloudCover:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_CLOUDCOVER);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionDewPoint:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_DEWPOINT);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionHumidity:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_HUMIDITY);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionPressure:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_PRESSURE);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionRainRate:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_RAINRATE);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionSkyBrightness:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_SKYBRIGHTNESS);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionSkyQuality:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_SKYQUALITY);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionStarFWHM:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_STARFWHM);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionSkyTemperature:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_SKYTEMPERATURE);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionTemperature:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_TEMPERATURE);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionWindDirection:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_WINDDIRECTION);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionWindGust:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_WINDGUST);
                            break;
                        }

                    case ObservingConditionsProperty.SensorDescriptionWindSpeed:
                        {
                            returnValue = m_ObservingConditions.SensorDescription(PROPERTY_WINDSPEED);
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
                    case object _ when returnValue == null:
                        {
                            LogIssue(MethodName, "The driver did not return any string at all: Nothing (VB), null (C#)");
                            break;
                        }

                    case object _ when returnValue == "":
                        {
                            LogOK(MethodName, "The driver returned an empty string: \"\"");
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

                sensorHasDescription[p_Name] = true;
            }
            catch (Exception ex)
            {
                HandleException(MethodName, MemberType.Method, p_Mandatory, ex, "");
                returnValue = null;
            }
            return returnValue;
        }
    }

}

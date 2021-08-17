﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using Microsoft.VisualBasic;
using ASCOM.Standard.Interfaces;
using ASCOM.Standard.AlpacaClients;
using ASCOM.Standard.COM.DriverAccess;
using System.Threading;

namespace ConformU
{

    internal class ObservingConditionsTester : DeviceTesterBaseClass
    {
        private double averageperiod, cloudCover, dewPoint, humidity, pressure, rainRate, skyBrightness, skyQuality, starFWHM, skyTemperature, temperature, windDirection, windGust, windSpeed;

        // Variables to indicate whether each function is or is not implemented to that it is possible to check that for any given sensor, all three either are or are not implemented
        private readonly Dictionary<string, bool> sensorisImplemented = new();
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
        private const string PROPERTY_LASTUPDATETIME = "LastUpdateTime";

        // List of valid ObservingConditions sensor properties
        private List<string> ValidSensors = new List<string>() { PROPERTY_CLOUDCOVER, PROPERTY_DEWPOINT, PROPERTY_HUMIDITY, PROPERTY_PRESSURE, PROPERTY_RAINRATE, PROPERTY_SKYBRIGHTNESS, PROPERTY_SKYQUALITY, PROPERTY_SKYTEMPERATURE, PROPERTY_STARFWHM, PROPERTY_TEMPERATURE, PROPERTY_WINDDIRECTION, PROPERTY_WINDGUST, PROPERTY_WINDSPEED };

        // Helper variables
        private IObservingConditions m_ObservingConditions;
        private readonly CancellationToken cancellationToken;
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
        public ObservingConditionsTester(ConformanceTestManager parent, ConformConfiguration conformConfiguration, ConformLogger logger, CancellationToken conformCancellationToken) : base(false, true, true, false, false, true, false, parent, conformConfiguration, logger, conformCancellationToken) // Set flags for this device:  HasCanProperties, HasProperties, HasMethods, PreRunCheck, PreConnectCheck, PerformanceCheck, PostRunCheck
        {
            settings = conformConfiguration.Settings;
            cancellationToken = conformCancellationToken;
            this.logger = logger;
        }

        // IDisposable
        private bool disposedValue = false;        // To detect redundant calls

        protected override void Dispose(bool disposing)
        {
            LogMsg("Dispose", MessageLevel.msgDebug, "Disposing of device: " + disposing.ToString() + " " + disposedValue.ToString());
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

        public override void CheckInitialise()
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
            base.CheckInitialise(settings.ComDevice.ProgId);
        }

        public override void CreateDevice()
        {
            try
            {
                switch (settings.DeviceTechnology)
                {
                    case DeviceTechnology.Alpaca:
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating Alpaca device: IP address: {settings.AlpacaDevice.IpAddress}, IP Port: {settings.AlpacaDevice.IpPort}, Alpaca device number: {settings.AlpacaDevice.AlpacaDeviceNumber}");
                        m_ObservingConditions = new AlpacaObservingConditions(settings.AlpacaConfiguration.AccessServiceType.ToString(), settings.AlpacaDevice.IpAddress, settings.AlpacaDevice.IpPort, settings.AlpacaDevice.AlpacaDeviceNumber, logger);
                        logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Alpaca device created OK");
                        break;

                    case DeviceTechnology.COM:
                        switch (settings.ComConfiguration.ComACcessMechanic)
                        {
                            case ComAccessMechanic.Native:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating NATIVE COM device: {settings.ComDevice.ProgId}");
                                m_ObservingConditions = new ObservingConditionsFacade(settings, logger);
                                break;

                            case ComAccessMechanic.DriverAccess:
                                logger.LogMessage("CreateDevice", MessageLevel.msgDebug, $"Creating DriverAccess device: {settings.ComDevice.ProgId}");
                                m_ObservingConditions = new ObservingConditions(settings.ComDevice.ProgId);
                                break;

                            default:
                                throw new ASCOM.InvalidValueException($"CreateDevice - Unknown COM access mechanic: {settings.ComConfiguration.ComACcessMechanic}");
                        }
                        break;

                    default:
                        throw new ASCOM.InvalidValueException($"CreateDevice - Unknown technology type: {settings.DeviceTechnology}");
                }

                LogMsg("CreateDevice", MessageLevel.msgDebug, "Successfully created driver");

                WaitForAbsolute(DEVICE_DESTROY_WAIT, "Waiting for driver to initialise");
                g_Stop = false;
            }
            catch (Exception ex)
            {
                LogMsg("CreateDevice", MessageLevel.msgDebug, "Exception thrown: " + ex.Message);
                throw; // Re throw exception 
            }

            if (g_Stop) WaitFor(200);
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
                    LogMsgIssue("AveragePeriod Write", "No error generated on setting the average period < -1.0");
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
                    LogMsgOK("AveragePeriod Write", "Successfully set average period to 0.0");
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
                    LogMsgOK("AveragePeriod Write", "Successfully set average period to 5.0");
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
                    LogMsgOK("AveragePeriod Write", "Successfully restored original average period: " + averageperiod);
                }
                catch (Exception ex)
                {
                    LogMsgInfo("AveragePeriod Write", "Unable to restore original average period");
                    HandleException("AveragePeriod Write", MemberType.Property, Required.Mandatory, ex, "");
                }
            }
            else
                LogMsgInfo("AveragePeriod Write", "Test skipped because AveragePerid cold not be read");

            cloudCover = TestDouble(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.CloudCover, 0.0, 100.0, Required.Optional);
            dewPoint = TestDouble(PROPERTY_DEWPOINT, ObservingConditionsProperty.DewPoint, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional);
            humidity = TestDouble(PROPERTY_HUMIDITY, ObservingConditionsProperty.Humidity, 0.0, 100.0, Required.Optional);

            if ((IsGoodValue(dewPoint) & IsGoodValue(humidity)))
                LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both implemented per the interface specification");
            else if ((!IsGoodValue(dewPoint) & !IsGoodValue(humidity)))
                LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both not implemented per the interface specification");
            else
                LogMsg("DewPoint & Humidity", MessageLevel.msgIssue, "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            pressure = TestDouble(PROPERTY_PRESSURE, ObservingConditionsProperty.Pressure, 0.0, 1100.0, Required.Optional);
            rainRate = TestDouble(PROPERTY_RAINRATE, ObservingConditionsProperty.RainRate, 0.0, 20000.0, Required.Optional);
            skyBrightness = TestDouble(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.SkyBrightness, 0.0, 1000000.0, Required.Optional);
            skyQuality = TestDouble(PROPERTY_SKYQUALITY, ObservingConditionsProperty.SkyQuality, -20.0, 30.0, Required.Optional);
            starFWHM = TestDouble(PROPERTY_STARFWHM, ObservingConditionsProperty.StarFWHM, 0.0, 1000.0, Required.Optional);
            skyTemperature = TestDouble(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.SkyTemperature, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional);
            temperature = TestDouble(PROPERTY_TEMPERATURE, ObservingConditionsProperty.Temperature, ABSOLUTE_ZERO, WATER_BOILING_POINT, Required.Optional);
            windDirection = TestDouble(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.WindDirection, 0.0, 360.0, Required.Optional);
            windGust = TestDouble(PROPERTY_WINDGUST, ObservingConditionsProperty.WindGust, 0.0, 1000.0, Required.Optional);
            windSpeed = TestDouble(PROPERTY_WINDSPEED, ObservingConditionsProperty.WindSpeed, 0.0, 1000.0, Required.Optional);

            // Additional test to confirm that the reported direction is 0.0 if the wind speed is reported as 0.0
            if ((windSpeed == 0.0))
            {
                if ((windDirection == 0.0))
                    LogMsg(PROPERTY_WINDSPEED, MessageLevel.msgOK, "Wind direction is reported as 0.0 when wind speed is 0.0");
                else
                    LogMsg(PROPERTY_WINDSPEED, MessageLevel.msgError, string.Format("When wind speed is reported as 0.0, wind direction should also be reported as 0.0, it is actually reported as {0}", windDirection));
            }
        }

        public override void CheckMethods()
        {
            double LastUpdateTimeLatest, LastUpdateTimeCloudCover, LastUpdateTimeDewPoint, LastUpdateTimeHumidity, LastUpdateTimePressure, LastUpdateTimeRainRate, LastUpdateTimeSkyBrightness, LastUpdateTimeSkyQuality;
            double LastUpdateTimeStarFWHM, LastUpdateTimeSkyTemperature, LastUpdateTimeTemperature, LastUpdateTimeWindDirection, LastUpdateTimeWindGust, LastUpdateTimeWindSpeed;

            string SensorDescriptionCloudCover, SensorDescriptionDewPoint, SensorDescriptionHumidity, SensorDescriptionPressure, SensorDescriptionRainRate, SensorDescriptionSkyBrightness, SensorDescriptionSkyQuality;
            string SensorDescriptionStarFWHM, SensorDescriptionSkyTemperature, SensorDescriptionTemperature, SensorDescriptionWindDirection, SensorDescriptionWindGust, SensorDescriptionWindSpeed;

            // TimeSinceLastUpdate
            LastUpdateTimeLatest = TestDouble(PROPERTY_LATESTUPDATETIME, ObservingConditionsProperty.TimeSinceLastUpdateLatest, -1.0, double.MaxValue, Required.Mandatory);

            LastUpdateTimeCloudCover = TestDouble(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.TimeSinceLastUpdateCloudCover, -1.0, double.MaxValue, Required.Optional);
            LastUpdateTimeDewPoint = TestDouble(PROPERTY_DEWPOINT, ObservingConditionsProperty.TimeSinceLastUpdateDewPoint, -1.0, double.MaxValue, Required.Optional);
            LastUpdateTimeHumidity = TestDouble(PROPERTY_HUMIDITY, ObservingConditionsProperty.TimeSinceLastUpdateHumidity, -1.0, double.MaxValue, Required.Optional);

            if ((IsGoodValue(LastUpdateTimeDewPoint) & IsGoodValue(LastUpdateTimeHumidity)))
                LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both implemented per the interface specification");
            else if ((!IsGoodValue(LastUpdateTimeDewPoint) & !IsGoodValue(LastUpdateTimeHumidity)))
                LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both not implemented per the interface specification");
            else
                LogMsg("DewPoint & Humidity", MessageLevel.msgIssue, "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            LastUpdateTimePressure = TestDouble(PROPERTY_PRESSURE, ObservingConditionsProperty.TimeSinceLastUpdatePressure, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeRainRate = TestDouble(PROPERTY_RAINRATE, ObservingConditionsProperty.TimeSinceLastUpdateRainRate, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeSkyBrightness = TestDouble(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.TimeSinceLastUpdateSkyBrightness, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeSkyQuality = TestDouble(PROPERTY_SKYQUALITY, ObservingConditionsProperty.TimeSinceLastUpdateSkyQuality, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeStarFWHM = TestDouble(PROPERTY_STARFWHM, ObservingConditionsProperty.TimeSinceLastUpdateStarFWHM, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeSkyTemperature = TestDouble(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateSkyTemperature, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeTemperature = TestDouble(PROPERTY_TEMPERATURE, ObservingConditionsProperty.TimeSinceLastUpdateTemperature, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeWindDirection = TestDouble(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.TimeSinceLastUpdateWindDirection, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeWindGust = TestDouble(PROPERTY_WINDGUST, ObservingConditionsProperty.TimeSinceLastUpdateWindGust, double.MinValue, double.MaxValue, Required.Optional);
            LastUpdateTimeWindSpeed = TestDouble(PROPERTY_WINDSPEED, ObservingConditionsProperty.TimeSinceLastUpdateWindSpeed, double.MinValue, double.MaxValue, Required.Optional);

            // Refresh
            try
            {
                LogCallToDriver("AveragePeriod Write", "About to call Refresh method");
                m_ObservingConditions.Refresh();
                LogMsg("Refresh", MessageLevel.msgOK, "Refreshed OK");
            }
            catch (Exception ex)
            {
                HandleException("Refresh", MemberType.Method, Required.Optional, ex, "");
            }

            // SensorDescrtiption
            SensorDescriptionCloudCover = TestSensorDescription(PROPERTY_CLOUDCOVER, ObservingConditionsProperty.SensorDescriptionCloudCover, int.MaxValue, Required.Optional);
            SensorDescriptionDewPoint = TestSensorDescription(PROPERTY_DEWPOINT, ObservingConditionsProperty.SensorDescriptionDewPoint, int.MaxValue, Required.Optional);
            SensorDescriptionHumidity = TestSensorDescription(PROPERTY_HUMIDITY, ObservingConditionsProperty.SensorDescriptionHumidity, int.MaxValue, Required.Optional);

            if (((SensorDescriptionDewPoint == null) & (SensorDescriptionHumidity == null)))
                LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both not implemented per the interface specification");
            else if (((!(SensorDescriptionDewPoint == null)) & (!(SensorDescriptionHumidity == null))))
                LogMsg("DewPoint & Humidity", MessageLevel.msgOK, "Dew point and humidity are both implemented per the interface specification");
            else
                LogMsg("DewPoint & Humidity", MessageLevel.msgIssue, "One of Dew point or humidity is implemented and the other is not. Both must be implemented or both must not be implemented per the interface specification");

            SensorDescriptionPressure = TestSensorDescription(PROPERTY_PRESSURE, ObservingConditionsProperty.SensorDescriptionPressure, int.MaxValue, Required.Optional);
            SensorDescriptionRainRate = TestSensorDescription(PROPERTY_RAINRATE, ObservingConditionsProperty.SensorDescriptionRainRate, int.MaxValue, Required.Optional);
            SensorDescriptionSkyBrightness = TestSensorDescription(PROPERTY_SKYBRIGHTNESS, ObservingConditionsProperty.SensorDescriptionSkyBrightness, int.MaxValue, Required.Optional);
            SensorDescriptionSkyQuality = TestSensorDescription(PROPERTY_SKYQUALITY, ObservingConditionsProperty.SensorDescriptionSkyQuality, int.MaxValue, Required.Optional);
            SensorDescriptionStarFWHM = TestSensorDescription(PROPERTY_STARFWHM, ObservingConditionsProperty.SensorDescriptionStarFWHM, int.MaxValue, Required.Optional);
            SensorDescriptionSkyTemperature = TestSensorDescription(PROPERTY_SKYTEMPERATURE, ObservingConditionsProperty.SensorDescriptionSkyTemperature, int.MaxValue, Required.Optional);
            SensorDescriptionTemperature = TestSensorDescription(PROPERTY_TEMPERATURE, ObservingConditionsProperty.SensorDescriptionTemperature, int.MaxValue, Required.Optional);
            SensorDescriptionWindDirection = TestSensorDescription(PROPERTY_WINDDIRECTION, ObservingConditionsProperty.SensorDescriptionWindDirection, int.MaxValue, Required.Optional);
            SensorDescriptionWindGust = TestSensorDescription(PROPERTY_WINDGUST, ObservingConditionsProperty.SensorDescriptionWindGust, int.MaxValue, Required.Optional);
            SensorDescriptionWindSpeed = TestSensorDescription(PROPERTY_WINDSPEED, ObservingConditionsProperty.SensorDescriptionWindSpeed, int.MaxValue, Required.Optional);

            // Now check that the sensor value, description and last updated time are all either implemented or not implemented
            foreach (string sensorName in ValidSensors)
            {
                LogMsg("Consistency", MessageLevel.msgDebug, "Sensor name: " + sensorName);
                if ((sensorisImplemented[sensorName] & sensorHasDescription[sensorName] & sensorHasTimeOfLastUpdate[sensorName]))
                    LogMsg("Consistency - " + sensorName, MessageLevel.msgOK, "Sensor value, description and time since last update are all implemented as required by the specification");
                else if (((!sensorisImplemented[sensorName]) & (!sensorHasDescription[sensorName]) & (!sensorHasTimeOfLastUpdate[sensorName])))
                    LogMsg("Consistency - " + sensorName, MessageLevel.msgOK, "Sensor value, description and time since last update are all not implemented as required by the specification");
                else
                {
                    LogMsg("Consistency - " + sensorName, MessageLevel.msgIssue, "Sensor value is implemented: " + sensorisImplemented[sensorName] + ", Sensor description is implemented: " + sensorHasDescription[sensorName] + ", Sensor time since last update is implemented: " + sensorHasTimeOfLastUpdate[sensorName]);
                    LogMsg("Consistency - " + sensorName, MessageLevel.msgInfo, "The ASCOM specification requires that sensor value, description and time since last update must either all be implemented or all not be implemented.");
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
        private bool IsGoodValue(double value)
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
                SensorName = MethodName.Substring(PROPERTY_TIMESINCELASTUPDATE.Length);
                LogCallToDriver(MethodName, $"About to call TimeSinceLastUpdate({SensorName}) method");
            }
            else
            {
                SensorName = MethodName;
                LogCallToDriver(MethodName, $"About to get {SensorName} property");
            }
            LogMsg("returnValue", MessageLevel.msgDebug, "methodName: " + MethodName + ", SensorName: " + SensorName);
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
                                LogMsg(MethodName, MessageLevel.msgError, "returnValue: Unknown test type - " + p_Type.ToString());
                                break;
                            }
                    }

                    readOK = true;

                    // Successfully retrieved a value so check validity
                    switch (returnValue)
                    {
                        case object _ when returnValue < p_Min:
                            {
                                LogMsg(MethodName, MessageLevel.msgError, "Invalid value (below minimum expected - " + p_Min.ToString() + "): " + returnValue.ToString());
                                break;
                            }

                        case object _ when returnValue > p_Max:
                            {
                                LogMsg(MethodName, MessageLevel.msgError, "Invalid value (above maximum expected - " + p_Max.ToString() + "): " + returnValue.ToString());
                                break;
                            }

                        default:
                            {
                                LogMsg(MethodName, MessageLevel.msgOK, returnValue.ToString());
                                break;
                            }
                    }

                    if (MethodName.StartsWith(PROPERTY_TIMESINCELASTUPDATE))
                        sensorHasTimeOfLastUpdate[SensorName] = true;
                    else
                        sensorisImplemented[SensorName] = true;
                }
                catch (Exception ex)
                {
                    if (IsInvalidOperationException(p_Nmae, ex))
                    {
                        returnValue = BAD_VALUE;
                        retryCount = retryCount + 1;
                        LogMsg(MethodName, MessageLevel.msgInfo, "Sensor not ready, received InvalidOperationException, waiting " + settings.ObservingConditionsRetryTime + " second to retry. Attempt " + retryCount + " out of " + settings.ObservingConditionsMaxRetries);
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
            while (!readOK | (retryCount == settings.ObservingConditionsMaxRetries) | unexpectedError)// Lower than minimum value// Higher than maximum value
    ;

            if ((!readOK) & (!unexpectedError))
                LogMsg(MethodName, MessageLevel.msgInfo, "InvalidOperationException persisted for longer than " + settings.ObservingConditionsMaxRetries * settings.ObservingConditionsRetryTime + " seconds.");
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
                            LogMsg(MethodName, MessageLevel.msgError, "TestString: Unknown test type - " + p_Type.ToString());
                            break;
                        }
                }

                // Successfully retrieved a value
                switch (returnValue)
                {
                    case object _ when returnValue == null:
                        {
                            LogMsg(MethodName, MessageLevel.msgError, "The driver did not return any string at all: Nothing (VB), null (C#)");
                            break;
                        }

                    case object _ when returnValue == "":
                        {
                            LogMsg(MethodName, MessageLevel.msgOK, "The driver returned an empty string: \"\"");
                            break;
                        }

                    default:
                        {
                            if (Strings.Len(returnValue) <= p_MaxLength)
                                LogMsg(MethodName, MessageLevel.msgOK, returnValue);
                            else
                                LogMsg(MethodName, MessageLevel.msgError, "String exceeds " + p_MaxLength + " characters maximum length - " + returnValue);
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
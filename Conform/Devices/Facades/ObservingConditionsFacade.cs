using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class ObservingConditionsFacade : FacadeBaseClass, IObservingConditions
    {
        // Create the test device in the facade base class
        public ObservingConditionsFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public double AveragePeriod { get => driver.AveragePeriod; set => driver.AveragePeriod = value; }

        public double CloudCover => driver.CloudCover;

        public double DewPoint => driver.DewPoint;

        public double Humidity => driver.Humidity;

        public double Pressure => driver.Pressure;

        public double RainRate => driver.RainRate;

        public double SkyBrightness => driver.SkyBrightness;

        public double SkyQuality => driver.SkyQuality;

        public double StarFWHM => driver.StarFWHM;

        public double SkyTemperature => driver.SkyTemperature;

        public double Temperature => driver.Temperature;

        public double WindDirection => driver.WindDirection;

        public double WindGust => driver.WindGust;

        public double WindSpeed => driver.WindSpeed;

        public void Refresh()
        {
            driver.Refresh();
        }

        public string SensorDescription(string PropertyName)
        {
            return driver.SensorDescription(PropertyName);
        }

        public double TimeSinceLastUpdate(string PropertyName)
        {
            return driver.TimeSinceLastUpdate(PropertyName);
        }

        #endregion

    }
}

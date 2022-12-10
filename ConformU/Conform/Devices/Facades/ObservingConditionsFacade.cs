using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class ObservingConditionsFacade : FacadeBaseClass, IObservingConditions
    {
        // Create the test device in the facade base class
        public ObservingConditionsFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public double AveragePeriod { get => (double)FunctionNoParameters(() => driver.AveragePeriod); set => Method1Parameter((i) => driver.AveragePeriod = i, value); }

        public double CloudCover => (double)FunctionNoParameters(() => driver.CloudCover);

        public double DewPoint => (double)FunctionNoParameters(() => driver.DewPoint);

        public double Humidity => (double)FunctionNoParameters(() => driver.Humidity);

        public double Pressure => (double)FunctionNoParameters(() => driver.Pressure);

        public double RainRate => (double)FunctionNoParameters(() => driver.RainRate);

        public double SkyBrightness => (double)FunctionNoParameters(() => driver.SkyBrightness);

        public double SkyQuality => (double)FunctionNoParameters(() => driver.SkyQuality);

        public double StarFWHM => (double)FunctionNoParameters(() => driver.StarFWHM);

        public double SkyTemperature => (double)FunctionNoParameters(() => driver.SkyTemperature);

        public double Temperature => (double)FunctionNoParameters(() => driver.Temperature);

        public double WindDirection => (double)FunctionNoParameters(() => driver.WindDirection);

        public double WindGust => (double)FunctionNoParameters(() => driver.WindGust);

        public double WindSpeed => (double)FunctionNoParameters(() => driver.WindSpeed);

        public void Refresh()
        {
            MethodNoParameters(() => driver.Refresh());
        }

        public string SensorDescription(string PropertyName)
        {
            return (string)Function1Parameter((i) => driver.SensorDescription(i), PropertyName);
        }

        public double TimeSinceLastUpdate(string PropertyName)
        {
            return (double)Function1Parameter((i) => driver.TimeSinceLastUpdate(i), PropertyName);
        }

        #endregion

    }
}

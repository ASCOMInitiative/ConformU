using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class ObservingConditionsFacade : FacadeBaseClass, IObservingConditions
    {
        // Create the test device in the facade base class
        public ObservingConditionsFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public double AveragePeriod { get => FunctionNoParameters<double>(() => driver.AveragePeriod); set => Method1Parameter((i) => driver.AveragePeriod = i, value); }

        public double CloudCover
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.CloudCover);
            }
        }

        public double DewPoint
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.DewPoint);
            }
        }

        public double Humidity
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Humidity);
            }
        }

        public double Pressure
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Pressure);
            }
        }

        public double RainRate
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.RainRate);
            }
        }

        public double SkyBrightness
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.SkyBrightness);
            }
        }

        public double SkyQuality
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.SkyQuality);
            }
        }

        public double StarFWHM
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.StarFWHM);
            }
        }

        public double SkyTemperature
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.SkyTemperature);
            }
        }

        public double Temperature
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.Temperature);
            }
        }

        public double WindDirection
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.WindDirection);
            }
        }

        public double WindGust
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.WindGust);
            }
        }

        public double WindSpeed
        {
            get
            {
                return FunctionNoParameters<double>(() => driver.WindSpeed);
            }
        }

        public void Refresh()
        {
            MethodNoParameters(() => driver.Refresh());
        }

        public string SensorDescription(string PropertyName)
        {
            return Function1Parameter<string>((i) => driver.SensorDescription(i), PropertyName);
        }

        public double TimeSinceLastUpdate(string PropertyName)
        {
            return Function1Parameter<double>((i) => driver.TimeSinceLastUpdate(i), PropertyName);
        }

        #endregion

    }
}

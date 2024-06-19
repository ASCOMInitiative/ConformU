using ASCOM.Common.DeviceInterfaces;

namespace ConformU
{
    public class ObservingConditionsFacade : FacadeBaseClass, IObservingConditionsV2
    {
        // Create the test device in the facade base class
        public ObservingConditionsFacade(Settings conformSettings, ConformLogger logger) : base(conformSettings, logger) { }

        #region Interface implementation

        public double AveragePeriod { get => FunctionNoParameters<double>(() => Driver.AveragePeriod); set => Method1Parameter((i) => Driver.AveragePeriod = i, value); }

        public double CloudCover
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.CloudCover);
            }
        }

        public double DewPoint
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.DewPoint);
            }
        }

        public double Humidity
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.Humidity);
            }
        }

        public double Pressure
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.Pressure);
            }
        }

        public double RainRate
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.RainRate);
            }
        }

        public double SkyBrightness
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.SkyBrightness);
            }
        }

        public double SkyQuality
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.SkyQuality);
            }
        }

        public double StarFWHM
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.StarFWHM);
            }
        }

        public double SkyTemperature
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.SkyTemperature);
            }
        }

        public double Temperature
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.Temperature);
            }
        }

        public double WindDirection
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.WindDirection);
            }
        }

        public double WindGust
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.WindGust);
            }
        }

        public double WindSpeed
        {
            get
            {
                return FunctionNoParameters<double>(() => Driver.WindSpeed);
            }
        }

        public void Refresh()
        {
            MethodNoParameters(() => Driver.Refresh());
        }

        public string SensorDescription(string propertyName)
        {
            return Function1Parameter<string>((i) => Driver.SensorDescription(i), propertyName);
        }

        public double TimeSinceLastUpdate(string propertyName)
        {
            return Function1Parameter<double>((i) => Driver.TimeSinceLastUpdate(i), propertyName);
        }

        #endregion

    }
}

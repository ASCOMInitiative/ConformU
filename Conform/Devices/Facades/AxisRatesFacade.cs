using ASCOM;
using ASCOM.Standard.Interfaces;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public class AxisRatesFacade : IAxisRates, IEnumerable, IEnumerator, IDisposable
    {
        private TelescopeAxis m_axis;
        private RateFacade[] m_Rates;

        private int pos;
        ILogger logger;

        // Default constructor - Internal prevents public creation instances.
        internal AxisRatesFacade(TelescopeAxis Axis, dynamic driver, TelescopeFacade telescopeFacade, ILogger logger)
        {
            m_Rates = Array.Empty<RateFacade>(); // Initialise to an empty array

            this.logger = logger;
            m_axis = Axis;

            // Assign the AxisRates response to an IEnumerable variable
            IEnumerable driverAxisRates = (IEnumerable)telescopeFacade.Function1Parameter((i) => driver.AxisRates(i), Axis);

            // Copy the response values to a local array so that the driver is not continually polled for values
            int nextArrayPosition = 0;
            foreach (dynamic rate in driverAxisRates)
            {
                Array.Resize<RateFacade>(ref m_Rates, nextArrayPosition + 1); // Resize the array to add one more entry (always 1 more than the array element position, which is 0 based).
                m_Rates[nextArrayPosition] = new RateFacade(rate.Minimum, rate.Maximum); // Store the rate in the new array entry
                nextArrayPosition++; // Increment the rate counter
            }

            logger?.LogInformation($"AxisRateFacade.Init - Got {nextArrayPosition - 1} rates");
            pos = -1;
        }

        #region IAxisRates Members

        public int Count
        {
            get { return m_Rates.Length; }
        }

        public IEnumerator GetEnumerator()
        {
            logger?.LogInformation($"AxisRateFacade.GetEnumerator - Returning enumerator");

            pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }

        public IRate this[int index]
        {
            get
            {
                if (index < 1 || index > this.Count) throw new InvalidValueException("AxisRates.index", index.ToString(CultureInfo.CurrentCulture), string.Format(CultureInfo.CurrentCulture, "1 to {0}", this.Count));
                logger?.LogInformation($"AxisRateFacade.this[index] - Retuning index item: {index}");
                return (IRate)m_Rates[index - 1]; 	// 1-based
            }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        // The bulk of the clean-up code is implemented in Dispose(bool)
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                logger?.LogInformation($"AxisRateFacade.Dispose(bool) - SETTING M_RATES TO NULL!!");
                // free managed resources
                m_Rates = null;
            }
        }

        #endregion

        #region IEnumerator implementation

        public bool MoveNext()
        {
            logger?.LogInformation($"AxisRateFacade.MoveNext - SETTING Moving to next item");
            if (++pos >= m_Rates.Length) return false;
            return true;
        }

        public void Reset()
        {
            logger?.LogInformation($"AxisRateFacade.Reset - Resetting index");
            pos = -1;
        }

        System.Collections.IEnumerator IAxisRates.GetEnumerator()
        {
            logger?.LogInformation($"AxisRateFacade.IEnumerator - Returning enumerator");
            pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }

        public object Current
        {
            get
            {
                if (pos < 0 || pos >= m_Rates.Length) throw new ASCOM.InvalidOperationException();
                logger?.LogInformation($"AxisRateFacade.Current - Returning Current value");
                return m_Rates[pos];
            }
        }

        #endregion
    }
}

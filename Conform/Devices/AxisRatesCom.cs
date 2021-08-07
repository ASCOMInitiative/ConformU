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
    public class AxisRatesCom : IAxisRates, IEnumerable, IEnumerator, IDisposable
    {
        private TelescopeAxis m_axis;
        private RateFacade[] m_Rates;
        private int pos;

        //
        // Constructor - Internal prevents public creation
        // of instances. Returned by Telescope.AxisRates.
        //
        internal AxisRatesCom(TelescopeAxis Axis, dynamic driver)
        {
            m_axis = Axis;
            IEnumerable driverAxisRates = driver.AxisRates(Axis);

            int ct = 0;
            foreach (dynamic rate in driverAxisRates)
            {
                ct++;
            }

            m_Rates = new RateFacade[ct];

            ct = 0;
            foreach (dynamic rate in driverAxisRates)
            {
                m_Rates[ct] = new RateFacade(rate.Minimum, rate.Maximum);
                ct++;
            }

            pos = -1;
        }

        #region IAxisRates Members

        public int Count
        {
            get { return m_Rates.Length; }
        }

        public IEnumerator GetEnumerator()
        {
            pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }

        public IRate this[int index]
        {
            get
            {
                if (index < 1 || index > this.Count)
                    throw new InvalidValueException("AxisRates.index", index.ToString(CultureInfo.CurrentCulture), string.Format(CultureInfo.CurrentCulture, "1 to {0}", this.Count));
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
                // free managed resources
                m_Rates = null;
            }
        }

        #endregion

        #region IEnumerator implementation

        public bool MoveNext()
        {
            if (++pos >= m_Rates.Length) return false;
            return true;
        }

        public void Reset()
        {
            pos = -1;
        }

        System.Collections.IEnumerator IAxisRates.GetEnumerator()
        {
            pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }

        public object Current
        {
            get
            {
                if (pos < 0 || pos >= m_Rates.Length) throw new ASCOM.InvalidOperationException();
                return m_Rates[pos];
            }
        }

        #endregion
    }
}

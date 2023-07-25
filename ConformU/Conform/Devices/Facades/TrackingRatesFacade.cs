using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections;
using System.Globalization;

namespace ConformU
{
    public class TrackingRatesFacade : ITrackingRates, IEnumerable, IEnumerator, IDisposable
    {
        private readonly DriveRate[] m_TrackingRates;
        private int _pos = -1;

        // Default constructor - Internal prevents public creation instances.
        internal TrackingRatesFacade(dynamic driver, TelescopeFacade telescopeFacade)
        {
            // Initialise to an empty array
            m_TrackingRates = Array.Empty<DriveRate>();

            // Assign the TrackinRates response to an IEnumerable variable
            IEnumerable trackingRates = telescopeFacade.FunctionNoParameters<IEnumerable>(() => driver.TrackingRates);

            // Copy the response values to a local array so that the driver is not continually polled for values
            int nextArrayPosition = 0;
            foreach (DriveRate rate in trackingRates)
            {
                Array.Resize<DriveRate>(ref m_TrackingRates, nextArrayPosition + 1); // Resize the array to add one more entry (always 1 more than the array element position, which is 0 based).
                m_TrackingRates[nextArrayPosition] = rate; // Store the rate in the new array entry
                nextArrayPosition++; // Increment the rate counter
            }
        }

        #region ITrackingRates Members

        public int Count
        {
            get { return m_TrackingRates.Length; }
        }

        public IEnumerator GetEnumerator()
        {
            _pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }


        public DriveRate this[int index]
        {
            get
            {
                if (index < 1 || index > this.Count) throw new ASCOM.InvalidValueException("TrackingRates.this", index.ToString(CultureInfo.CurrentCulture), string.Format(CultureInfo.CurrentCulture, "1 to {0}", this.Count));
                return m_TrackingRates[index - 1];// 1-based
            }
        }
        #endregion

        #region IEnumerator implementation

        public bool MoveNext()
        {
            if (++_pos >= m_TrackingRates.Length) return false;
            return true;
        }

        public void Reset()
        {
            _pos = -1;
        }

        public object Current
        {
            get
            {
                if (_pos < 0 || _pos >= m_TrackingRates.Length) throw new System.InvalidOperationException();
                return m_TrackingRates[_pos];
            }
        }

        DriveRate ITrackingRates.this[int index]
        {
            get
            {
                if (index < 1 || index > this.Count) throw new ASCOM.InvalidValueException("TrackingRates.this", index.ToString(CultureInfo.CurrentCulture), string.Format(CultureInfo.CurrentCulture, "1 to {0}", this.Count));
                return m_TrackingRates[index - 1]; // 1-based
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
                /* Following code commented out in Platform 6.4 because m_TrackingRates is a global variable for the whole driver and there could be more than one 
                 * instance of the TrackingRates class (created by the calling application). One instance should not invalidate the variable that could be in use
                 * by other instances of which this one is unaware.

                m_TrackingRates = null;

                */
            }
        }

        IEnumerator ITrackingRates.GetEnumerator()
        {
            _pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }
        #endregion

    }
}

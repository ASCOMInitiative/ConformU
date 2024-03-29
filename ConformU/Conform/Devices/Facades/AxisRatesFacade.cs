﻿using ASCOM;
using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections;
using System.Globalization;

namespace ConformU
{
    public class AxisRatesFacade : IAxisRates, IEnumerable, IEnumerator, IDisposable
    {
        private readonly TelescopeAxis mAxis;
        private RateFacade[] mRates;

        private int pos;
        private readonly ConformLogger logger;

        // Default constructor - Internal prevents public creation instances.
        internal AxisRatesFacade(TelescopeAxis axis, dynamic driver, TelescopeFacade telescopeFacade, ConformLogger logger)
        {
            mRates = Array.Empty<RateFacade>(); // Initialise to an empty array

            this.logger = logger;
            mAxis = axis;

            // Assign the AxisRates response to an IEnumerable variable
            IEnumerable driverAxisRates = telescopeFacade.Function1Parameter<IEnumerable>((i) => driver.AxisRates(i), axis);

            // Copy the response values to a local array so that the driver is not continually polled for values
            int nextArrayPosition = 0;
            foreach (dynamic rate in driverAxisRates)
            {
                Array.Resize<RateFacade>(ref mRates, nextArrayPosition + 1); // Resize the array to add one more entry (always 1 more than the array element position, which is 0 based).
                mRates[nextArrayPosition] = new RateFacade(rate.Minimum, rate.Maximum); // Store the rate in the new array entry
                nextArrayPosition++; // Increment the rate counter
            }

            logger?.LogMessage($"AxisRateFacade.Init", MessageLevel.Debug, $"Got {nextArrayPosition - 1} rates");
            pos = -1;
        }

        #region IAxisRates Members

        public int Count
        {
            get { return mRates.Length; }
        }

        public IEnumerator GetEnumerator()
        {
            logger?.LogMessage("AxisRateFacade.GetEnumerator", MessageLevel.Debug, $"Returning enumerator");

            pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }

        public IRate this[int index]
        {
            get
            {
                if (index < 1 || index > this.Count) throw new InvalidValueException("AxisRates.index", index.ToString(CultureInfo.CurrentCulture), string.Format(CultureInfo.CurrentCulture, "1 to {0}", this.Count));
                logger?.LogMessage($"AxisRateFacade.this[index]", MessageLevel.Debug, $"Retuning index item: {index}");
                return (IRate)mRates[index - 1]; 	// 1-based
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
                logger?.LogMessage($"AxisRateFacade.Dispose(bool)", MessageLevel.Debug, $"SETTING M_RATES TO NULL!!");
                // free managed resources
                mRates = null;
            }
        }

        #endregion

        #region IEnumerator implementation

        public bool MoveNext()
        {
            logger?.LogMessage($"AxisRateFacade.MoveNext", MessageLevel.Debug, $"SETTING Moving to next item");
            if (++pos >= mRates.Length) return false;
            return true;
        }

        public void Reset()
        {
            logger?.LogMessage($"AxisRateFacade.Reset", MessageLevel.Debug, "Resetting index");
            pos = -1;
        }

        System.Collections.IEnumerator IAxisRates.GetEnumerator()
        {
            logger?.LogMessage($"AxisRateFacade.IEnumerator", MessageLevel.Debug, $"Returning enumerator");
            pos = -1; //Reset pointer as this is assumed by .NET enumeration
            return this as IEnumerator;
        }

        public object Current
        {
            get
            {
                if (pos < 0 || pos >= mRates.Length) throw new ASCOM.InvalidOperationException();
                logger?.LogMessage($"AxisRateFacade.Current", MessageLevel.Debug, $"Returning Current value");
                return mRates[pos];
            }
        }

        #endregion
    }
}

using ASCOM.Common.DeviceInterfaces;
using System;
using System.Collections;

namespace ConformU
{
    public class RateFacade : IRate, IDisposable
    {
        private double mDMaximum = 0;
        private double mDMinimum = 0;

        //
        // Default constructor - Internal prevents public creation
        // of instances. These are values for AxisRates.
        //
        internal RateFacade(double minimum, double maximum)
        {
            mDMaximum = maximum;
            mDMinimum = minimum;
        }

        #region IRate Members

        public IEnumerator GetEnumerator()
        {
            return null;
        }

        public double Maximum
        {
            get { return mDMaximum; }
            set { mDMaximum = value; }
        }

        public double Minimum
        {
            get { return mDMinimum; }
            set { mDMinimum = value; }
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
            // nothing to do?
        }

        #endregion
    }
}

using ASCOM.Alpaca.Discovery;
using System.Collections.Generic;

namespace ConformU
{
    /// Custom comparer for the Product class
    /// <summary>
    /// </summary>
    public class AscomDeviceComparer : IEqualityComparer<AscomDevice>
    {
        // Products are equal if their names and product numbers are equal.
        public bool Equals(AscomDevice x, AscomDevice y)
        {
            //Check whether the compared objects reference the same data.
            if (ReferenceEquals(x, y))
                return true;

            //Check whether any of the compared objects is null.
            if (x is null || y is null)
                return false;

            //Check whether the products' properties are equal.
            return x.UniqueId == y.UniqueId;
        }

        /// <summary>
        /// Return the hash code for an AscomDevice object
        /// </summary>
        /// <param name="ascomDevice">The AscomDevice whose hash code is required.</param>
        /// <returns>The AscomDevice's hash code.</returns>
        /// <remarks>If Equals() returns true for a pair of objects then GetHashCode() must return the same value for these objects.</remarks>
        public int GetHashCode(AscomDevice ascomDevice)
        {
            //Check whether the object is null
            if (ascomDevice is null) return 0;

            //Get hash code for the Name field if it is not null.
            return ascomDevice.UniqueId == null ? 0 : ascomDevice.UniqueId.GetHashCode();

        }
    }

}

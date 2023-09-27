using System;

namespace ConformU
{
   /// <summary>
   /// Exception raised when the user cancels a test
   /// </summary>
    public class TestCancelledException : Exception
    {
        /// <summary>
        /// Initialises with a canned exception message: Test cancelled by user.
        /// </summary>
        public TestCancelledException() : base("Test cancelled by user.") { }
    }
}

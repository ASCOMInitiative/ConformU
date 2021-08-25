using System;

namespace ConformU
{
    static class Globals
    {

        #region Global Variables

        // Variables shared between the test manager and device testers        
        internal static int g_CountError, g_CountWarning, g_CountIssue;

        // Filter wheel constants
        internal const int FWTEST_IS_MOVING = -1;
        internal const int FWTEST_TIMEOUT = 30;

        #endregion


        public static string SpaceDup(this int n)
        {
            return new String(' ', n);
        }

    }
}
// Option Strict On
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using ASCOM.Standard.Utilities;
using Microsoft.VisualBasic;
using static ConformU.ConformConstants;

namespace Conform
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

    }
}
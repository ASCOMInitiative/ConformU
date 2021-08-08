using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public static class ConformConstants
    {
        public const string TECHNOLOGY_ALPACA = "Alpaca";
        public const string TECHNOLOGY_COM = "COM";

        public const string ASCOM_PROFILE_KEY = @"SOFTWARE\ASCOM";

        public enum DeviceTechnology
        {
            Alpaca,
            COM
        }
        
        public enum MessageLevel
        {
            None = 0,
            Debug = 1,
            Comment = 2,
            Info = 3,
            OK = 4,
            Warning = 5,
            Issue = 6,
            Error = 7,
            Always = 8
        }

        public enum DeviceType
        {
            Telescope = 0,
            Camera = 1,
            Dome = 2,
            FilterWheel = 3,
            Focuser = 4,
            ObservingConditions = 5,
            Rotator = 6,
            Switch = 7,
            SafetyMonitor = 8,
            Video = 9,
            CoverCalibrator = 10
        }

        public const string NO_DEVICE_SELECTED = "No device selected"; // Text indicating that no device has been sleected

        //
        // Summary:
        //     Specifies the state of a control, such as a check box, that can be checked, unchecked,
        //     or set to an indeterminate state.
        public enum CheckState
        {
            //
            // Summary:
            //     The control is unchecked.
            Unchecked = 0,
            //
            // Summary:
            //     The control is checked.
            Checked = 1,
            //
            // Summary:
            //     The control is indeterminate. An indeterminate control generally has a shaded
            //     appearance.
            Indeterminate = 2
        }


    }


}
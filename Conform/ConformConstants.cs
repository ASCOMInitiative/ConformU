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




    }


}
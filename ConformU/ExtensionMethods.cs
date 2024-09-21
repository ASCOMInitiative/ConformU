using ASCOM.Tools;
using System;

namespace ConformU
{
    public static class ExtensionMethods
    {
        public static bool IsNumeric(this string text) => !string.IsNullOrWhiteSpace(text) && double.TryParse(text, out _);

        public static string ToHMS(this double rightAscension)
        {
            return $"{(rightAscension >= 0.0 ? "+" : "")}{Utilities.HoursToHMS(rightAscension, ":", ":", "", 2)}";
        }

        public static string ToDMS(this double declination)
        {
            return $"{(declination >= 0.0 ? "+" : "")}{Utilities.DegreesToDMS(declination, ":", ":", "", 1)}";
        }


    }
}

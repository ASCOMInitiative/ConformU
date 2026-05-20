using ASCOM.Alpaca.Discovery;
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

        /// <summary>
        /// Convert boolean value to JsonNameCaseSensitivity enum. True for CorrectCasingOnly, False for AnyCasing.
        /// </summary>
        /// <param name="value">Strict casing state</param>
        /// <returns>JsonNameCaseSensitivity.CorrectCasingOnly when value is true, JsonNameCaseSensitivity.AnyCasing when value is false</returns>
        public static JsonNameCaseSensitivity ToJsonNameCaseSensitivity(this bool value)
        {
            if(value)
                return JsonNameCaseSensitivity.CorrectCasingOnly;

            return JsonNameCaseSensitivity.AnyCasing;
        }
    }
}

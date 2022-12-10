using System;

namespace ConformU
{
    public static class ExtensionMethods
    {
        public static bool IsNumeric(this string text) => !string.IsNullOrWhiteSpace(text) && double.TryParse(text, out _);
        public static string SpaceDup(this int n)
        {
            return new String(' ', n);
        }

    }
}

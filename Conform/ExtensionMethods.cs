using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ConformU
{
    public static class ExtensionMethods
    {
        public static bool IsNumeric(this string text) => !string.IsNullOrWhiteSpace(text) && double.TryParse(text, out _);

    }
}

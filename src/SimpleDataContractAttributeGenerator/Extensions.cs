using System;
using System.Collections.Generic;
using System.Linq;

namespace SimpleDataContractAttributeGenerator
{
    internal static class Extensions
    {
        internal static bool HasContext(this string sValue) => !string.IsNullOrEmpty(sValue);

        internal static List<string> SplitWithString(this string sValue, string stringSeparator)
        {
            if (!sValue.HasContext()) throw new ArgumentNullException(nameof(sValue));
            return sValue.Split(new[] {stringSeparator}, StringSplitOptions.None).ToList();
        }

        internal static string ToLowerFirstChar(this string sValue)
        {
            if (!sValue.HasContext()) throw new ArgumentNullException(nameof(sValue));
            if (char.IsLower(sValue, 0)) return sValue;

            return char.ToLowerInvariant(sValue[0]) + sValue.Substring(1);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelMapping.Extenstions
{
    public static class StringExtenstion
    {
        public static bool IsNotEmpty(this string value)
        {
            return (!string.IsNullOrEmpty(value) && !string.IsNullOrWhiteSpace(value));
        }

        public static bool IsEmpty(this string value)
        {
            return (string.IsNullOrEmpty(value) || string.IsNullOrWhiteSpace(value));
        }
    }
}

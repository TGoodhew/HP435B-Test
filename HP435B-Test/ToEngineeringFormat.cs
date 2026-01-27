using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HP435B_Test
{
    /// <summary>
    /// Utility class for converting numeric values to engineering notation format.
    /// Credit: Steve Hageman - http://analoghome.blogspot.com/2012/01/how-to-format-numbers-in-engineering.html
    /// </summary>
    public static class ToEngineeringFormat
    {
        /// <summary>
        /// SI prefix constants for engineering notation (yocto to tera).
        /// </summary>
        private static string[] PrefixConstants = { " y", " z", " a", " f", " p", " n", " µ", " m", " ", " k", " M", " G", " T" };

        /// <summary>
        /// Converts a numeric value to engineering notation format with optional units.
        /// </summary>
        /// <param name="number">The number to convert.</param>
        /// <param name="significantDigits">The number of significant digits to return (1-15, recommended minimum of 3).</param>
        /// <param name="units">The unit of measurement (e.g., "Hz", "Farads", "Tesla").</param>
        /// <returns>A string representation of the number in engineering notation with SI prefix and units.</returns>
        public static string Convert(double number, Int16 significantDigits = 3, string units = "")
        {
            double scale = Math.Log10(Math.Abs(number));
            if (scale < 0.0)
                scale += -3.0;

            Int16 power = (Int16)((scale / 3) + 0.001);

            if (power.CompareTo(-8) < 0)
                power = -8;
            if (power.CompareTo(4) > 0)
                power = 4;

            string prefixStr = PrefixConstants[power + 8];
            double scaleFactor = Math.Pow(10.0, (double)power * 3.0);
            double baseNum = number / scaleFactor;

            if (significantDigits < 1)
                significantDigits = 1;
            if (significantDigits > 15)
                significantDigits = 15;

            string formatStr = "G" + significantDigits.ToString();
            string convertedStr = baseNum.ToString(formatStr) + prefixStr + units;

            return (convertedStr);
        }
    }
}
using System;

namespace ScipBe.Common.Office.Utils
{
    public static class StringExtensions
    {
        /// <summary>
        /// Convert the string representation of a number to its 32-bit signed integer equivalent.
        /// </summary>
        /// <param name="value">String containing a number to convert.</param>
        /// <returns>System.Int32.</returns>
        /// <remarks>
        /// The conversion fails if the string parameter is null, is not of the correct format, or represents a number
        /// less than System.Int32.MinValue or greater than System.Int32.MaxValue.
        /// </remarks>
        public static int ToInt32(this string value)
        {
            return int.TryParse(value, out int number) ? number : default;
        }

        /// <summary>
        /// Convert the string representation of a date time.
        /// </summary>
        /// <param name="value">String containing a date time to convert.</param>
        /// <returns>System.DateTime.</returns>
        public static DateTime ToDateTime(this string value)
        {
            return DateTime.TryParse(value, out DateTime dateTime)
                ? dateTime
                : DateTime.MinValue;
        }
    }
}

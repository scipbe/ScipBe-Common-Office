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
            int number = 0;
            Int32.TryParse(value, out number);
            return number;
        }

        /// <summary>
        /// Convert the string representation of a date time.
        /// </summary>
        /// <param name="value">String containing a date time to convert.</param>
        /// <returns>System.DateTime.</returns>
        public static DateTime ToDateTime(this string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return DateTime.MinValue;
            }

            DateTime dateTime;
            var success = DateTime.TryParse(value, out dateTime);

            if (!success)
            {
                return DateTime.MinValue;
            }

            return dateTime;
        }
    }
}

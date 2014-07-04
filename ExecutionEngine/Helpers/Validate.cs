using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros.ExecutionEngine
{
    public static class Validate
    {
        /// <summary>
        /// Throws an ArgumentNullException if the given object is null.
        /// </summary>
        /// <param name="o">Object to test</param>
        /// <param name="paramName">
        /// <paramref name="paramName"/> parameter used if an exception is raised
        /// </param>
        public static void IsNotNull(object o, string paramName)
        {
            if (o == null)
                throw new ArgumentNullException(paramName);
        }

        /// <summary>
        /// Throws an InvalidOperationException if the given object is not null.
        /// </summary>
        /// <param name="o">Object to test</param>
        /// <param name="paramName">
        /// <paramref name="paramName"/> parameter used if an exception is raised
        /// </param>
        public static void IsNull(object o, string paramName)
        {
            if (o != null)
            {
                string message = String.Format(CultureInfo.CurrentUICulture, Resources.ValidateError_InvalidOperation, paramName);
                throw new InvalidOperationException(message);
            }
        }

        /// <summary>
        /// Throws an ArgumentException if the given string is empty.
        /// </summary>
        /// <param name="s">String to test</param>
        /// <param name="paramName">
        /// <paramref name="paramName"/> parameter used if an exception is raised
        /// </param>
        public static void IsNotEmpty(string s, string paramName)
        {
            if (s == String.Empty)
                throw new ArgumentException(Resources.ValidateError_StringEmpty, paramName);
        }

        /// <summary>
        /// Throws an ArgumentException if the given guid is empty.
        /// </summary>
        /// <param name="s">Guid to test</param>
        /// <param name="paramName">
        /// <paramref name="paramName"/> parameter used if an exception is raised
        /// </param>
        public static void IsNotEmpty(Guid g, string paramName)
        {
            if (g == Guid.Empty)
                throw new ArgumentException(Resources.ValidateError_GuidEmpty, paramName);
        }

        /// <summary>
        /// Throws an ArgumentException if the given string contains only whitespaces.
        /// </summary>
        /// <param name="s">String to test</param>
        /// <param name="paramName">
        /// <paramref name="paramName"/> parameter used if an exception is raised
        /// </param>
        public static void IsNotWhiteSpace(string s, string paramName)
        {
            if (s != null && string.IsNullOrWhiteSpace(s))
                throw new ArgumentException(Resources.ValidateError_StringWhiteSpace, paramName);
        }

        /// <summary>
        /// Validates that the given string is both non-null and non-empty.
        /// </summary>
        /// <param name="s">String to test</param>
        /// <param name="paramName">
        /// <paramref name="paramName"/> parameter used if an exception is raised
        /// </param>
        public static void IsNotNullAndNotEmpty(string s, string paramName)
        {
            Validate.IsNotNull(s, paramName);
            Validate.IsNotEmpty(s, paramName);
        }
    }
}

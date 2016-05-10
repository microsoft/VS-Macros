//-----------------------------------------------------------------------
// <copyright file="Validate.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Globalization;

namespace ExecutionEngine
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
                string message = string.Format(CultureInfo.CurrentUICulture, Resources.ValidateErrorInvalidOperation, paramName);
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
            if (s == string.Empty)
                throw new ArgumentException(Resources.ValidateErrorStringEmpty, paramName);
        }

        /// <summary>
        /// Throws an ArgumentException if the given Guid is empty.
        /// </summary>
        /// <param name="s">Guid to test</param>
        /// <param name="paramName">
        /// <paramref name="paramName"/> parameter used if an exception is raised
        /// </param>
        public static void IsNotEmpty(Guid g, string paramName)
        {
            if (g == Guid.Empty)
                throw new ArgumentException(Resources.ValidateErrorGuidEmpty, paramName);
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
                throw new ArgumentException(Resources.ValidateErrorStringWhiteSpace, paramName);
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

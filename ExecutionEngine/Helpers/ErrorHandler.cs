//-----------------------------------------------------------------------
// <copyright file="ErrorHandler.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;

namespace ExecutionEngine
{
    internal static class ErrorHandler
    {
        public static bool Failed(int hr)
        {
            return hr < 0;
        }

        public static bool Succeeded(int hr)
        {
            return hr >= 0;
        }

        public static int ThrowOnFailure(int hr)
        {
            return ErrorHandler.ThrowOnFailure(hr, null);
        }

        public static int ThrowOnFailure(int hr, params int[] expectedHRFailure)
        {
            if (ErrorHandler.Failed(hr))
            {
                if ((expectedHRFailure == null) || (Array.IndexOf(expectedHRFailure, hr) < 0))
                {
                    Marshal.ThrowExceptionForHR(hr);
                }
            }

            return hr;
        }
    }
}

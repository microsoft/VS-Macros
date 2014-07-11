using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros.Helpers
{
    public static class ErrorHandler
    {
        public static bool Failed(int hr)
        {
            return (hr < 0);
        }

        public static bool Succeeded(int hr)
        {
            return (hr >= 0);
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

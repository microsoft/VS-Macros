//-----------------------------------------------------------------------
// <copyright file="NativeMethods.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ExecutionEngine.Helpers
{
    internal static class NativeMethods
    {
        public const int S_OK = 0x00000000;
        public const int ROTFLAGS_REGISTRATIONKEEPSALIVE = 1;
        public const int ROTFLAGS_ALLOWANYCLIENT = 2;

        [DllImport("ole32.dll")]
        internal static extern int CreateItemMoniker([MarshalAs(UnmanagedType.LPWStr)] string
           lpszDelim, [MarshalAs(UnmanagedType.LPWStr)] string lpszItem,
           out System.Runtime.InteropServices.ComTypes.IMoniker ppmk);

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved,
                                  out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved,
                                  out IBindCtx bc);
    }
}

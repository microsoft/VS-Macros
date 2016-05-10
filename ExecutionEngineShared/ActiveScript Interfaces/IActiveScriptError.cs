//-----------------------------------------------------------------------
// <copyright file="IActiveScriptError.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;

namespace ExecutionEngine.Interfaces
{
    [Guid("EAE1BA61-A4ED-11cf-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IActiveScriptError
    {
        void GetExceptionInfo(out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        void GetSourcePosition(out uint sourceContext, out uint lineNumber, out int characterPosition);
        void GetSourceLineText(out string sourceLine);
    }
}

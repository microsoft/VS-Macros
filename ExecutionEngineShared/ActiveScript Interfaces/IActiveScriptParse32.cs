//-----------------------------------------------------------------------
// <copyright file="IActiveScriptParse32.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using ExecutionEngine.Enums;

namespace ExecutionEngine.Interfaces
{
    [Guid("BB1A2AE2-A4F9-11cf-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IActiveScriptParse32
    {
        void InitNew();
        void AddScriptlet([MarshalAs(UnmanagedType.LPWStr)]string defaultName, 
            [MarshalAs(UnmanagedType.LPWStr)]string code, 
            [MarshalAs(UnmanagedType.LPWStr)]string itemName, 
            [MarshalAs(UnmanagedType.LPWStr)]string subItemName, 
            [MarshalAs(UnmanagedType.LPWStr)]string eventName, 
            [MarshalAs(UnmanagedType.LPWStr)]string delimiter, 
            uint sourceContextCookie, 
            uint startingLineNumber, 
            ScriptText flags, 
            out string name, 
            out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);

        void ParseScriptText([MarshalAs(UnmanagedType.LPWStr)]string code, 
            [MarshalAs(UnmanagedType.LPWStr)]string itemName, 
            object context, [MarshalAs(UnmanagedType.LPWStr)]string delimiter, 
            IntPtr sourceContextCookie, 
            uint startingLineNumber, 
            ScriptText flags, 
            out object result, 
            out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
    }
}

//-----------------------------------------------------------------------
// <copyright file="IActiveScriptParse64.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using ExecutionEngine.Enums;

namespace ExecutionEngine.Interfaces
{
    [Guid("C7EF7658-E1EE-480E-97EA-D52CB4D76D17")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IActiveScriptParse64
    {
        void InitNew();
        void AddScriptlet([MarshalAs(UnmanagedType.LPWStr)]string defaultName, 
            [MarshalAs(UnmanagedType.LPWStr)]string code, 
            [MarshalAs(UnmanagedType.LPWStr)]string itemName, 
            [MarshalAs(UnmanagedType.LPWStr)]string subItemName, 
            [MarshalAs(UnmanagedType.LPWStr)]string eventName, 
            [MarshalAs(UnmanagedType.LPWStr)]string delimiter, 
            IntPtr sourceContextCookie, 
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

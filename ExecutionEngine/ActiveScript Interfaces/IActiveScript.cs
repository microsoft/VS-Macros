//-----------------------------------------------------------------------
// <copyright file="IActiveScript.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using ExecutionEngine.Enums;

namespace ExecutionEngine.Interfaces
{
    [Guid("BB1A2AE1-A4F9-11cf-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IActiveScript
    {
        void SetScriptSite(IActiveScriptSite pass);
        void GetScriptSite(Guid riid, out IntPtr site);
        void SetScriptState(ScriptState state);
        void GetScriptState(out ScriptState scriptState);
        void Close();
        void AddNamedItem([MarshalAs(UnmanagedType.LPWStr)]string name, Enums.ScriptItem flags);
        void AddTypeLib(ref Guid typeLib, uint major, uint minor, uint flags);
        void GetScriptDispatch([MarshalAs(UnmanagedType.LPWStr)]string itemName, out IntPtr dispatch);
        void GetCurrentScriptThreadID(out uint thread);
        void GetScriptThreadID(uint win32ThreadId, out uint thread);
        void GetScriptThreadState(uint thread, out ScriptThreadState state);
        void InterruptScriptThread(uint thread, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo, uint flags);
        void Clone(out IActiveScript script);
    }
}

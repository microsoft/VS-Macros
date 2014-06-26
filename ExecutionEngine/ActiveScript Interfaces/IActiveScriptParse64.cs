using ExecutionEngine.Enums;
using System;
using System.Runtime.InteropServices;

namespace ExecutionEngine.Interfaces
{
    [Guid("C7EF7658-E1EE-480E-97EA-D52CB4D76D17"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IActiveScriptParse64
    {
        void InitNew();
        void AddScriptlet(string defaultName, string code, string itemName, string subItemName, string eventName, string delimiter, IntPtr sourceContextCookie, uint startingLineNumber, ScriptText flags, out string name, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        void ParseScriptText(string code, string itemName, object context, string delimiter, IntPtr sourceContextCookie, uint startingLineNumber, ScriptText flags, out object result, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
    }
}

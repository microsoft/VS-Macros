//-----------------------------------------------------------------------
// <copyright file="Site.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using ExecutionEngine.Enums;
using ExecutionEngine.Interfaces;

namespace ExecutionEngine
{
    internal class Site : IActiveScriptSite
    {
        private const int TypeEElementNotFound = unchecked((int)(0x8002802B));

        public void GetLCID(out int lcid)
        {
            lcid = Thread.CurrentThread.CurrentCulture.LCID;
        }

        public void GetItemInfo(string name, Enums.ScriptInfo returnMask, out IntPtr item, IntPtr typeInfo)
        {
            if ((returnMask & ScriptInfo.ITypeInfo) == ScriptInfo.ITypeInfo)
            {
                throw new NotImplementedException();
            }

            if (Engine.DteObject == null)
            {
                throw new COMException(null, TypeEElementNotFound);
            }

            item = Marshal.GetIUnknownForObject(Engine.DteObject);
        }

        public void GetDocVersionString(out string version)
        {
            // Debug.WriteLine("Site:IActiveScriptSite.GetDocVersionString");
            version = null;
        }

        public void OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo)
        {
            // Debug.WriteLine("Site:IActiveScriptSite.OnScriptTerminate");
        }

        public void OnStateChange(Enums.ScriptState scriptState)
        {
            // Debug.WriteLine("Site:IActiveScriptSite.OnStateChange");
        }

        public void OnScriptError(IActiveScriptError scriptError)
        {
            // Debug.WriteLine("Site:IActiveScriptSite.OnScriptError");
        }

        public void OnEnterScript()
        {
            // Debug.WriteLine("Site:IActiveScriptSite.OnEnterScript");
        }

        public void OnLeaveScript()
        {
            // Debug.WriteLine("Site:IActiveScriptSite.OnLeaveScript");
        }
    }
}

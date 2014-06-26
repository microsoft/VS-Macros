﻿using ExecutionEngine.Enums;
using ExecutionEngine.Interfaces;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExecutionEngine
{
    internal class Site : IActiveScriptSite
    {
        private const int TYPE_E_ELEMENTNOTFOUND = unchecked((int)(0x8002802B));
        internal static object DteObject;

        public void GetLCID(out int lcid)
        {
            lcid = Thread.CurrentThread.CurrentCulture.LCID;
        }

        public void GetItemInfo(string name, Enums.ScriptInfo returnMask, out IntPtr item, IntPtr typeInfo)
        {
            if ((returnMask & ScriptInfo.ITypeInfo) == ScriptInfo.ITypeInfo)
                throw new NotImplementedException();

            if (Site.DteObject == null)
                throw new COMException(null, TYPE_E_ELEMENTNOTFOUND);

            item = Marshal.GetIUnknownForObject(Site.DteObject);
        }

        public void GetDocVersionString(out string version)
        {
            //Debug.WriteLine("Site:IActiveScriptSite.GetDocVersionString");
            version = null;
        }

        public void OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo)
        {
            //Debug.WriteLine("Site:IActiveScriptSite.OnScriptTerminate");
        }

        public void OnStateChange(Enums.ScriptState scriptState)
        {
            //Debug.WriteLine("Site:IActiveScriptSite.OnStateChange");
        }

        public void OnScriptError(IActiveScriptError scriptError)
        {
            //Debug.WriteLine("Site:IActiveScriptSite.OnScriptError");
        }

        public void OnEnterScript()
        {
            //Debug.WriteLine("Site:IActiveScriptSite.OnEnterScript");
        }

        public void OnLeaveScript()
        {
            //Debug.WriteLine("Site:IActiveScriptSite.OnLeaveScript");
        }
    }
}

//-----------------------------------------------------------------------
// <copyright file="Site.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExecutionEngine.Enums;
using ExecutionEngine.Interfaces;
using VSMacros.ExecutionEngine;

namespace ExecutionEngine
{
    internal sealed class Site : IActiveScriptSite
    {
        private const int TypeEElementNotFound = unchecked((int)(0x8002802B));
        internal static bool error;
        internal static RuntimeException runtimeException;

        public void GetLCID(out int lcid)
        {
            lcid = Thread.CurrentThread.CurrentCulture.LCID;
        }

        public void GetItemInfo(string name, ScriptInfo returnMask, out IntPtr item, IntPtr typeInfo)
        {
            if ((returnMask & ScriptInfo.ITypeInfo) == ScriptInfo.ITypeInfo)
            {
                throw new NotImplementedException();
            }

            if (name.Equals("dte") && !(Engine.DteObject == null))
            {
                item = Marshal.GetIUnknownForObject(Engine.DteObject);
                //typeInfo = Marshal.GetITypeInfoForType(item.GetType());
            }

            else if (name.Equals("cmdHelper") && !(Engine.CommandHelper == null))
            {
                item = Marshal.GetIUnknownForObject(Engine.CommandHelper);
                //typeInfo = Marshal.GetITypeInfoForType(item.GetType());
            }

            else
            {
                throw new COMException(null, TypeEElementNotFound);
            }
            
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
            uint sourceContext;
            uint lineNumber;
            int characterPosition;
            System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;

            scriptError.GetSourcePosition(out sourceContext, out lineNumber, out characterPosition);
            scriptError.GetExceptionInfo(out exceptionInfo);

            string description = exceptionInfo.bstrDescription;
            string source = exceptionInfo.bstrSource;

            Site.error = true;
            Site.runtimeException = new RuntimeException(description, source, lineNumber);
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

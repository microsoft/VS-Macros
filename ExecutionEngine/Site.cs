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
using VSMacros.ExecutionEngine;

namespace ExecutionEngine
{
    internal sealed class Site : IActiveScriptSite
    {
        private const int TypeEElementNotFound = unchecked((int)(0x8002802B));
        internal static bool RuntimeError;
        internal static RuntimeException RuntimeException;
        internal static bool CriticalError;
        internal static Exception VSException;

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

            if (name.Equals("dte"))
            {   
                if (Engine.DteObject != null)
                {
                    item = Marshal.GetIUnknownForObject(Engine.DteObject);
                }
                else
                {
                    Debug.WriteLine("Engine.DteObject is null");

                    // TODO: Is this the right thing to do?
                    throw new Exception();
                }
            }
            else if (name.Equals("cmdHelper")) 
            {
                if (Engine.CommandHelper != null)
                {
                    item = Marshal.GetIUnknownForObject(Engine.CommandHelper);
                }
                else 
                {
                    Debug.WriteLine("Engine.CommandHelper is null");

                    // TODO: Is this the right thing to do?
                    throw new Exception();
                }
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

        public static void ResetError()
        {
            Site.RuntimeError = false;
            Site.RuntimeException = null;

            Site.CriticalError = false;
            Site.VSException = null;
        }

        public void OnScriptError(IActiveScriptError scriptError)
        {
            uint sourceContext;
            uint lineNumber;
            int column;
            System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;

            scriptError.GetSourcePosition(out sourceContext, out lineNumber, out column);
            scriptError.GetExceptionInfo(out exceptionInfo);

            string description = exceptionInfo.bstrDescription;
            string source = exceptionInfo.bstrSource;

            Site.RuntimeError = true;
            Site.RuntimeException = new RuntimeException(description, source, lineNumber, column);
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

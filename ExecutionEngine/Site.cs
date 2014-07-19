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
using VisualStudio.Macros.ExecutionEngine;
using VSMacros.ExecutionEngine;

namespace ExecutionEngine
{
    internal sealed class Site : IActiveScriptSite
    {
        private const int TypeEElementNotFound = unchecked((int)(0x8002802B));
        internal static bool RuntimeError;
        internal static RuntimeException RuntimeException;
        internal static bool InternalError;
        internal static InternalVSException InternalVSException;

        public void GetLCID(out int lcid)
        {
            lcid = Thread.CurrentThread.CurrentCulture.LCID;
        }

        private static void CreateInternalVSException(string message)
        {
            string source = "VSMacros.ExecutionEngine";
            string stackTrace = string.Empty;
            string targetSite = string.Empty;
            Site.InternalVSException = new InternalVSException(message, source, stackTrace, targetSite);
            Site.InternalError = true;
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
                    CreateInternalVSException(Resources.NullDte);
                    item = IntPtr.Zero;

#if DEBUG
                    Debug.WriteLine("Engine.DteObject is null.");
#endif
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
                    CreateInternalVSException(Resources.NullCommandHelper);
                    item = IntPtr.Zero;

#if DEBUG
                    Debug.WriteLine("Engine.CommandHelper is null.");
#endif
                }
            }
            else
            {
                throw new COMException(null, TypeEElementNotFound);
            }
        }

        public void GetDocVersionString(out string version)
        {
            version = null;
        }

        public void OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo)
        {
        }

        public void OnStateChange(Enums.ScriptState scriptState)
        {
        }

        public static void ResetError()
        {
            Site.RuntimeError = false;
            Site.RuntimeException = null;

            Site.InternalError = false;
            Site.InternalVSException = null;
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
        }

        public void OnLeaveScript()
        {
        }
    }
}

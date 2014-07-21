//-----------------------------------------------------------------------
// <copyright file="ParsedScript.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using VisualStudio.Macros.ExecutionEngine;
using VSMacros.ExecutionEngine;
using VSMacros.ExecutionEngine.Pipes;

namespace ExecutionEngine
{
    internal sealed class ParsedScript
    {
        private readonly Engine engine;
        private readonly object dispatch;

        internal ParsedScript(Engine engine, IntPtr dispatch)
        {
            this.engine = engine;
            this.dispatch = Marshal.GetObjectForIUnknown(dispatch);

            if (this.dispatch == null)
            {
                string message = Resources.NullDispatch;
                string source = "ExecutionEngine";
                string stackTrace = string.Empty;
                string targetSite = string.Empty;

                byte[] criticalErrorMessage = Client.PackageCriticalError(message, source, stackTrace, targetSite);
                Client.SendMessageToServer(Client.ClientStream, criticalErrorMessage);
            }
        }

        public bool CallMethod(string methodName, params object[] arguments)
        {
            try
            {
                this.dispatch.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, this.dispatch, arguments);
                return true;
            }
            catch (Exception e)
            {
                var internalException = e.InnerException;
                if (!Site.RuntimeError)
                {
                    Site.InternalError = true;
                    Site.InternalVSException = new InternalVSException(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                }
                return false;
            }
        }
    }
}

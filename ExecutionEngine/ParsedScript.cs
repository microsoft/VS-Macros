//-----------------------------------------------------------------------
// <copyright file="ParsedScript.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
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
                // TODO: Put these in resources
                // TODO: Check if these messages are correct.
                string message = "The dispatch object was null.";
                string source = "Execution engine";
                string stackTrace = "ParsedScript";
                string targetSite = "ParsedScript??";

                byte[] criticalErrorMessage = Client.PackageCriticalError(message, source, stackTrace, targetSite);
                Client.SendMessageToServer(Client.ClientStream, criticalErrorMessage);
            }
        }

        public bool CallMethod(string methodName, params object[] arguments)
        {
            try
            {
                this.dispatch.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, this.dispatch, arguments);

                byte[] successMessage = Client.PackageSuccessMessage();
                string message = Encoding.Unicode.GetString(successMessage);
                Client.SendMessageToServer(Client.ClientStream, successMessage);

                return true;
            }
            catch (Exception e)
            {
                if (!Site.RuntimeError)
                {
                    Site.CriticalError = true;
                    Site.VSException = new Exception(e.Message, e);
                }
                return false;
            }
        }
    }
}

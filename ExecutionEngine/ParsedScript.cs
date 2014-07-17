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

        public object CallMethod(string methodName, params object[] arguments)
        {
            try
            {
                object result = this.dispatch.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, this.dispatch, arguments);

                byte[] successMessage = Client.PackageSuccessMessage();
                string message = Encoding.Unicode.GetString(successMessage);
                Client.SendMessageToServer(Client.ClientStream, successMessage);

                return result;
            }
            catch (Exception e)
            {
                if (Site.Error)
                {
                    uint activeDocumentAddition = 1;
                    var ex = Site.RuntimeException;
                    uint modifiedLineNumber = ex.Line + activeDocumentAddition;

                    byte[] scriptErrorMessage = Client.PackageScriptError(modifiedLineNumber, ex.CharacterPosition, ex.Source, ex.Description);
                    string message = Encoding.Unicode.GetString(scriptErrorMessage);
                    Client.SendMessageToServer(Client.ClientStream, scriptErrorMessage);
                    return null;
                }
                else
                {
                    byte[] criticalErrorMessage = Client.PackageCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                    Client.SendMessageToServer(Client.ClientStream, criticalErrorMessage);
                    return null;
                }
            }
        }
    }
}

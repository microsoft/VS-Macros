//-----------------------------------------------------------------------
// <copyright file="ParsedScript.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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
                throw new InvalidOperationException();
            }
        }

        public object CallMethod(string methodName, params object[] arguments)
        {
            try
            {
                return this.dispatch.GetType().InvokeMember(Program.MacroName, BindingFlags.InvokeMethod, null, this.dispatch, arguments);
            }
            catch (Exception e)
            {
                if (Site.Error)
                {
                    var exception = Site.RuntimeException;
                    var exceptionMessage = string.Format("{0}: {1} at line {2}", exception.Source, exception.Description, exception.Line);
                    MessageBox.Show("Error in the script: " + exceptionMessage);

                    // Error messaging for IPC is acting weird for now
                    // byte[] scriptErrorMessage = Client.PackageScriptError(exception.Line, exception.Source, exception.Description);
                    // string message = System.Text.UnicodeEncoding.Unicode.GetString(scriptErrorMessage);
                    // Client.SendMessageToServer(Client.ClientStream, scriptErrorMessage);
                    return null;
                }
                else
                {
                    var errorMessage = string.Format("An error occurred: {0}: {1}", e.Message, e.GetBaseException());
                    MessageBox.Show(errorMessage);
                    throw;
                }
            }
        }
    }
}

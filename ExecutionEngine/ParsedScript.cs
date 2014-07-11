//-----------------------------------------------------------------------
// <copyright file="ParsedScript.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using VSMacros.ExecutionEngine;

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
                return this.dispatch.GetType().InvokeMember(Program.macroName, BindingFlags.InvokeMethod, null, this.dispatch, arguments);
            }

            catch (Exception e)
            {
                if (Site.error)
                {
                    var exception = Site.runtimeException;
                    var exceptionMessage = string.Format("{0}: {1} at line {2}", exception.Source, exception.Description, exception.Line);
                    MessageBox.Show("from CallMethod, Site.error: " + exceptionMessage);
                    return null;
                }
                else
                {
                    var errorMessage = string.Format("An error occurred: {0}: {1}", e.Message, e.GetBaseException());
                    MessageBox.Show("From CallMethod, general exception: " + errorMessage);
                    throw;
                }
            }
        }
    }
}

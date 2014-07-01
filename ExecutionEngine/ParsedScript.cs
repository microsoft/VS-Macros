﻿//-----------------------------------------------------------------------
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
            if (methodName == null)
            {
                throw new ArgumentNullException("methodName");
            }

            try
            {
                return this.dispatch.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, this.dispatch, arguments);
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                throw;
            }
        }
    }
}

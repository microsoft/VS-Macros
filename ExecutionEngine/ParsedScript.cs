using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExecutionEngine
{
    internal class ParsedScript : IDisposable
    {
        private object dispatch;
        private readonly Engine engine;

        internal ParsedScript(Engine engine, IntPtr dispatch)
        {
            this.engine = engine;
            this.dispatch = Marshal.GetObjectForIUnknown(dispatch);
        }


        public object CallMethod(string methodName, params object[] arguments)
        {
            if (this.dispatch == null)
                throw new InvalidOperationException();

            if (methodName == null)
                throw new ArgumentNullException("methodName");

            try
            {
                return this.dispatch.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, this.dispatch, arguments);
            }
            catch
            {
                throw;
            }
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}

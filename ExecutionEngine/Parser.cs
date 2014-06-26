using ExecutionEngine.Enums;
using ExecutionEngine.Interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExecutionEngine
{
    class Parser
    {
        private IActiveScriptParse32 parse32;
        private IActiveScriptParse64 parse64;
        private bool isParse32;

        public Parser(IActiveScript engine)
        {
            this.isParse32 = DeterminePointerSize();
            InitializeParsers(engine);
        }

        void InitializeParsers(IActiveScript engine)
        {
            if (this.isParse32)
            {
                this.parse32 = engine as IActiveScriptParse32;
                this.parse32.InitNew();
            }
            else
            {
                this.parse64 = engine as IActiveScriptParse64;
                this.parse64.InitNew();
            }
        }

        bool DeterminePointerSize()
        {
            if (IntPtr.Size == 4)
                return true;

            return false;
        }

        internal void Parse(string unparsed)
        {
            object result;
            System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
            ScriptText flags = ScriptText.None;

            if (this.isParse32)
            {
                this.parse32.ParseScriptText(unparsed, null, null, null, IntPtr.Zero, 0, flags, out result, out exceptionInfo);
            }
            else
            {
                this.parse64.ParseScriptText(unparsed, null, null, null, IntPtr.Zero, 0, flags, out result, out exceptionInfo);
            }
        }

        public void Dispose()
        {
            if (isParse32)
            {
                Marshal.ReleaseComObject(this.parse32);
                this.parse32 = null;
            }
            else
            {
                Marshal.ReleaseComObject(this.parse64);
                this.parse64 = null;
            }
        }
    }
}

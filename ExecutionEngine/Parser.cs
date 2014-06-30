//-----------------------------------------------------------------------
// <copyright file="Parser.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using ExecutionEngine.Enums;
using ExecutionEngine.Interfaces;

namespace ExecutionEngine
{
    internal sealed class Parser
    {
        private IActiveScriptParse32 parse32;
        private IActiveScriptParse64 parse64;
        private bool isParse32;

        public Parser(IActiveScript engine)
        {
            this.isParse32 = this.is32BitEnvironment();
            this.InitializeParsers(engine);
        }

        internal void InitializeParsers(IActiveScript engine)
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

        internal bool is32BitEnvironment()
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
            if (this.isParse32 && this.parse32 != null)
            {
                Marshal.ReleaseComObject(this.parse32);
                this.parse32 = null;
            }
            if (this.parse64 != null)
            {
                Marshal.ReleaseComObject(this.parse64);
                this.parse64 = null;
            }
        }
    }
}

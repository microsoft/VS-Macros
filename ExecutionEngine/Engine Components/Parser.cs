//-----------------------------------------------------------------------
// <copyright file="Parser.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
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
            this.isParse32 = this.Is32BitEnvironment();
            this.InitializeParsers(engine);
        }

        internal void InitializeParsers(IActiveScript engine)
        {
            if (this.isParse32)
            {
                this.parse32 = (IActiveScriptParse32)engine;
                this.parse32.InitNew();
            }
            else
            {
                this.parse64 = (IActiveScriptParse64)engine;
                this.parse64.InitNew();
            }
        }

        internal bool Is32BitEnvironment()
        {
            return IntPtr.Size == 4;
        }

        internal void Parse(string unparsed)
        {
            object result;
            System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
            ScriptText flags = ScriptText.None;

            if (this.isParse32)
            {
                this.parse32.ParseScriptText(unparsed,
                    itemName: null,
                    context: null, 
                    delimiter: null, 
                    sourceContextCookie: IntPtr.Zero, 
                    startingLineNumber: 0, 
                    flags: flags, 
                    result: out result, 
                    exceptionInfo: out exceptionInfo);
            }
            else
            {
                this.parse64.ParseScriptText(unparsed,
                    itemName: null,
                    context: null,
                    delimiter: null,
                    sourceContextCookie: IntPtr.Zero,
                    startingLineNumber: 0,
                    flags: flags,
                    result: out result,
                    exceptionInfo: out exceptionInfo);
            }
        }
    }
}

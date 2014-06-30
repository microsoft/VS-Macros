//-----------------------------------------------------------------------
// <copyright file="Engine.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ExecutionEngine.Enums;
using ExecutionEngine.Helpers;
using ExecutionEngine.Interfaces;
using System.Globalization;

namespace ExecutionEngine
{
    internal class Engine : IDisposable
    {
        private IActiveScript engine;
        private Parser parser;
        private Site scriptSite;

        private static object dteObject;
        public static object DteObject
        {
            get { return dteObject; }
            set { dteObject = value; }
        }

        private void InitializeDteObject(int pid)
        {
            IMoniker moniker;
            NativeMethods.CreateItemMoniker("!", string.Format(CultureInfo.InvariantCulture, "VisualStudio.DTE.12.0:{0}", pid), out moniker);

            IRunningObjectTable rot;
            NativeMethods.GetRunningObjectTable(0, out rot);

            rot.GetObject(moniker, out Engine.dteObject);
        }

        internal IActiveScript CreateEngine()
        {
            const string language = "jscript";

            Type engine = Type.GetTypeFromProgID(language, true);
            return Activator.CreateInstance(engine) as IActiveScript;
        }

        public Engine(int pid)
        {
            const string dte = "dte";
            this.engine = this.CreateEngine();
            this.scriptSite = new Site();
            this.parser = new Parser(this.engine);

            if (Engine.dteObject == null)
            {
                this.InitializeDteObject(pid);
                this.engine.SetScriptSite(this.scriptSite);
                this.engine.AddNamedItem(dte, ScriptItem.CodeOnly | ScriptItem.IsVisible);
            }
        }

        public void Dispose()
        {
            if (this.parser != null)
            {
                this.parser.Dispose();
                this.parser = null;
            }

            if (this.engine != null)
            {
                Marshal.ReleaseComObject(this.engine);
                this.engine = null;
            }
        }

        internal ParsedScript GenerateParsedScript()
        {
            IntPtr dispatch;
            this.engine.GetScriptDispatch(null, out dispatch);
            return new ParsedScript(this, dispatch);
        }

        internal ParsedScript Parse(string unparsed)
        {
            if (string.IsNullOrEmpty(unparsed))
            {
                throw new ArgumentNullException("unparsed");
            }

            this.engine.SetScriptState(ScriptState.Connected);
            this.parser.Parse(unparsed);
            var parsedScript = this.GenerateParsedScript();

            return parsedScript;
        }
    }
}

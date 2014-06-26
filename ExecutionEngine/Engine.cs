using ExecutionEngine.Enums;
using ExecutionEngine.Helpers;
using ExecutionEngine.Interfaces;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;


namespace ExecutionEngine
{
    internal class Engine : IDisposable
    {
        private IActiveScript engine;
        private Parser parser;
        internal Site site;

        void InitializeDteObject(int pid)
        {
            IMoniker moniker;
            NativeMethods.CreateItemMoniker("!", string.Format("VisualStudio.DTE.12.0:{0}", pid), out moniker);

            IRunningObjectTable rot;
            NativeMethods.GetRunningObjectTable(0, out rot);

            rot.GetObject(moniker, out Site.DteObject);
        }

        IActiveScript CreateEngine()
        {
            string language = "jscript";

            Type engine = Type.GetTypeFromProgID(language, true);
            return Activator.CreateInstance(engine) as IActiveScript;
        }

        public Engine(int pid)
        {
            this.engine = CreateEngine();
            this.site = new Site();
            this.parser = new Parser(this.engine);

            if (Site.DteObject == null)
            {
                InitializeDteObject(pid);
                this.engine.SetScriptSite(this.site);
                this.engine.AddNamedItem("dte", ScriptItem.CodeOnly | ScriptItem.IsVisible);
            }
        }

        public void Dispose()
        {
            if (this.parser != null)
            {
                this.parser.Dispose();
            }

            if (this.engine != null)
            {
                Marshal.ReleaseComObject(this.engine);
                this.engine = null;
            }
        }

        ParsedScript GenerateParsedScript()
        {
            IntPtr dispatch;
            this.engine.GetScriptDispatch(null, out dispatch);
            return new ParsedScript(this, dispatch);
        }

        internal ParsedScript Parse(string unparsed)
        {
            if (String.IsNullOrEmpty(unparsed))
            {
                throw new ArgumentNullException("unparsed");
            }

            this.engine.SetScriptState(ScriptState.Connected);
            this.parser.Parse(unparsed);
            var parsedScript = GenerateParsedScript();

            return parsedScript;
        }
    }
}

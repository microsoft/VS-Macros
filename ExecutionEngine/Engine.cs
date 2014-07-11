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
using VSMacros.ExecutionEngine;
using Microsoft.VisualStudio.Shell;
using System.Windows.Forms;

namespace ExecutionEngine
{
    internal sealed class Engine : IDisposable
    {
        private IActiveScript engine;
        private Parser parser;
        private Site scriptSite;

        public static object DteObject { get; private set; }
        public static object CommandHelper { get; private set; }

        IMoniker GetItemMoniker(int pid)
        {
            IMoniker moniker;
            // TODO: make it so it works with other versions of VS as well
            int hr = NativeMethods.CreateItemMoniker("!", string.Format(CultureInfo.InvariantCulture, "VisualStudio.DTE.12.0:{0}", pid), out moniker);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return moniker;
        }

        IRunningObjectTable GetRunningObjectTable()
        {
            IRunningObjectTable rot;
            int hr = NativeMethods.GetRunningObjectTable(0, out rot);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return rot;
        }

        object GetDteObject(IRunningObjectTable rot, IMoniker moniker)
        {
            object dteObject;
            int hr = rot.GetObject(moniker, out dteObject);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return dteObject;
        }

        private void InitializeDteObject(int pid)
        {
            IMoniker moniker = GetItemMoniker(pid);
            IRunningObjectTable rot = GetRunningObjectTable();
            Engine.DteObject = GetDteObject(rot, moniker);

            Validate.IsNotNull(Engine.DteObject, "Engine.DteObject");
        }

        private void InitializeCommandHelper()
        {
            var globalProvider = ServiceProvider.GlobalProvider;
            if (globalProvider == null)
            {
                //MessageBox.Show("global provider is null");
            }
            //Engine.CommandHelper = new CommandHelper(ServiceProvider.GlobalProvider);
            Engine.CommandHelper = new CommandHelper(globalProvider);
            Validate.IsNotNull(Engine.CommandHelper, "Engine.CommandHelper");
        }

        internal IActiveScript CreateEngine()
        {
            const string language = "jscript";

            Type engine = Type.GetTypeFromProgID(language, true);
            return Activator.CreateInstance(engine) as IActiveScript;
        }

        public Engine(int pid)
        {
            MessageBox.Show("hi :)");
            const string dte = "dte";
            const string cmdHelper = "cmdHelper";

            this.engine = this.CreateEngine();
            this.scriptSite = new Site();
            this.parser = new Parser(this.engine);

            this.InitializeCommandHelper();
            this.InitializeDteObject(pid);
            this.engine.SetScriptSite(this.scriptSite);
            this.engine.AddNamedItem(dte, ScriptItem.CodeOnly | ScriptItem.IsVisible);
            this.engine.AddNamedItem(cmdHelper, ScriptItem.CodeOnly | ScriptItem.IsVisible);
        }

        public void Dispose()
        {
            this.parser.Dispose();

            if (this.engine != null)
            {
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
            Validate.IsNotNullAndNotEmpty(unparsed, "unparsed");

            this.engine.SetScriptState(ScriptState.Connected);
            this.parser.Parse(unparsed);
            var parsedScript = this.GenerateParsedScript();

            return parsedScript;
        }
    }
}

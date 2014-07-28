//-----------------------------------------------------------------------
// <copyright file="Engine.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Globalization;
using System.Runtime.InteropServices.ComTypes;
using ExecutionEngine.Enums;
using ExecutionEngine.Helpers;
using ExecutionEngine.Interfaces;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;

namespace ExecutionEngine
{
    public sealed class Engine : IDisposable
    {
        private IActiveScript engine;
        private Parser parser;
        private Site scriptSite;

        public static object DteObject { get; private set; }
        public static object CommandHelper { get; private set; }

        private IMoniker GetItemMoniker(int pid, string version)
        {
            IMoniker moniker;
            ErrorHandler.ThrowOnFailure(NativeMethods.CreateItemMoniker("!", string.Format(CultureInfo.InvariantCulture, "VisualStudio.DTE.{0}:{1}", version, pid), out moniker));

            return moniker;
        }

        private IRunningObjectTable GetRunningObjectTable()
        {
            IRunningObjectTable rot;
            int hr = NativeMethods.GetRunningObjectTable(0, out rot);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return rot;
        }

        private object GetDteObject(IRunningObjectTable rot, IMoniker moniker)
        {
            object dteObject;
            int hr = rot.GetObject(moniker, out dteObject);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return dteObject;
        }

        private void InitializeDteObject(int pid, string version)
        {
            IMoniker moniker = this.GetItemMoniker(pid, version);
            IRunningObjectTable rot = this.GetRunningObjectTable();
            Engine.DteObject = this.GetDteObject(rot, moniker);

            if (Engine.DteObject == null)
            {
                throw new InvalidOperationException();
            }

            //var serviceprovider = (microsoft.visualstudio.ole.interop.iserviceprovider)engine.dteobject;
            
            //guid cmdnameguid = marshal.generateguidfortype(typeof(svscmdnamemapping));
            //guid suihostguid = marshal.generateguidfortype(typeof(suihostcommanddispatcher));

            //intptr cmdnameptr, suihostptr;
            //guid dummy = guid.empty;
            //serviceprovider.queryservice(ref cmdnameguid, ref dummy, out cmdnameptr);
            //serviceprovider.queryservice(ref suihostguid, ref dummy, out suihostptr);

            //console.writeline(cmdnameptr.tostring());
            //console.writeline(suihostptr.tostring());
        }

        private void InitializeCommandHelper()
        {
            var globalProvider = ServiceProvider.GlobalProvider;
            Validate.IsNotNull(globalProvider, "globalProvider");

            Engine.CommandHelper = new CommandHelper(globalProvider);
            Validate.IsNotNull(Engine.CommandHelper, "Engine.CommandHelper");
        }

        internal IActiveScript CreateEngine()
        {
            Type engine = Type.GetTypeFromProgID("jscript", true);
            return Activator.CreateInstance(engine) as IActiveScript;
        }

        public Engine(int pid, string version)
        {
            const string dte = "dte";
            const string cmdHelper = "cmdHelper";

            this.engine = this.CreateEngine();
            this.scriptSite = new Site();
            this.parser = new Parser(this.engine);

            this.InitializeCommandHelper();
            this.InitializeDteObject(pid, version);
            this.engine.SetScriptSite(this.scriptSite);
            this.engine.AddNamedItem(dte, ScriptItem.CodeOnly | ScriptItem.IsVisible);
            this.engine.AddNamedItem(cmdHelper, ScriptItem.CodeOnly | ScriptItem.IsVisible);
        }

        public void InterruptScript(int threadId)
        {
            System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
            uint flag = 2;

            this.engine.InterruptScriptThread((uint)threadId, ref exceptionInfo, flag);
        }

        public void Dispose()
        {
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
            try
            {
                this.engine.SetScriptState(ScriptState.Connected);
                this.parser.Parse(unparsed);
                var parsedScript = this.GenerateParsedScript();
                return parsedScript;
            }
            catch (Exception e)
            {
                // TODO: Package into internal error
                Console.WriteLine("oops:" + e.Message + e.Source + e.StackTrace + e.TargetSite.ToString());
                return null;
            }
        }
    }
}

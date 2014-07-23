//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.ExecutionEngine.Pipes.Shared;
using VSMacros.Helpers;
using VSMacros.Interfaces;
using VSMacros.Pipes;

namespace VSMacros.Engines
{
    /// <summary>
    /// Implements the execution engine.
    /// </summary>   
    internal sealed class Executor : IExecutor
    {
        /// <summary>
        /// The execution engine.
        /// </summary>
        internal static Process executionEngine;
        internal static bool IsEngineInitialized;
        internal static JobHandle Job;

        /// <summary>
        /// Informs subscribers of success or error during execution.
        /// </summary>
        public event EventHandler<CompletionReachedEventArgs> Complete;

        internal void SendCompletionMessage(bool isError, string errorMessage)
        {
            if (this.Complete != null)
            {
                var eventArgs = new CompletionReachedEventArgs(isError, errorMessage);
                this.Complete(this, eventArgs);
            }
        }

        internal void ResetMessages()
        {
            this.Complete = null;
        }

        #region Helpers
        private string ProvidePipeArguments(Guid guid, string version)
        {
            int pid = Process.GetCurrentProcess().Id;
            return string.Format("{0}{1}{2}{1}{3}", guid, SharedVariables.Delimiter, pid, version);
        }

        private System.Runtime.InteropServices.ComTypes.IRunningObjectTable GetRunningObjectTable()
        {
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            ErrorHandler.ThrowOnFailure(NativeMethods.GetRunningObjectTable(0, out rot));

            return rot;
        }

        private static string GetProgID(Type type)
        {
            Guid guid = Marshal.GenerateGuidForType(type);
            string progID = String.Format(CultureInfo.InvariantCulture, "{0:B}", guid);
            return progID;
        }

        private void RegisterCommandDispatcherinROT()
        {
            int hResult;
            System.Runtime.InteropServices.ComTypes.IBindCtx bc;
            hResult = NativeMethods.CreateBindCtx(0, out bc);

            if (hResult == NativeMethods.S_OK && bc != null)
            {
                System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot = GetRunningObjectTable();
                Validate.IsNotNull(rot, "rot");

                string delim = "!";
                string progID = GetProgID(typeof(SUIHostCommandDispatcher));

                System.Runtime.InteropServices.ComTypes.IMoniker moniker;
                hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);
                Validate.IsNotNull(moniker, "moniker");

                if (hResult == NativeMethods.S_OK)
                {
                    var serviceProvider = ServiceProvider.GlobalProvider;

                    var suiHost = serviceProvider.GetService(typeof(SUIHostCommandDispatcher));
                    hResult = rot.Register(NativeMethods.ROTFLAGS_REGISTRATIONKEEPSALIVE, suiHost, moniker);
                    if (hResult != NativeMethods.S_OK)
                    {
                        // Todo: logging an error
                    }
                }
            }
        }

        private void RegisterCmdNameMappinginROT()
        {
            int hResult;
            System.Runtime.InteropServices.ComTypes.IBindCtx bc;
            hResult = NativeMethods.CreateBindCtx(0, out bc);

            if (hResult == NativeMethods.S_OK && bc != null)
            {
                System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot = GetRunningObjectTable();
                Validate.IsNotNull(rot, "rot");

                string delim = "!";
                string progID = GetProgID(typeof(SVsCmdNameMapping));

                System.Runtime.InteropServices.ComTypes.IMoniker moniker;
                hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);
                Validate.IsNotNull(moniker, "moniker");

                if (hResult == NativeMethods.S_OK)
                {
                    var serviceProvider = ServiceProvider.GlobalProvider;

                    var svsCmdName = serviceProvider.GetService(typeof(SVsCmdNameMapping));
                    hResult = rot.Register(NativeMethods.ROTFLAGS_REGISTRATIONKEEPSALIVE, svsCmdName, moniker);
                    if (hResult != NativeMethods.S_OK)
                    {
                        // Todo: logging an error
                    }
                }
            }
        }
        #endregion

        /// <summary>
        /// Initializes the engine.
        /// </summary>
        /// 
        public void InitializeEngine()
        {
            this.RegisterCmdNameMappinginROT();
            this.RegisterCommandDispatcherinROT();

            Server.InitializeServer();
            Server.serverWait = new Thread(new ThreadStart(Server.WaitForMessage));
            Server.serverWait.Start();

            EnvDTE.DTE dte = ((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SDTE)) as EnvDTE.DTE;
            string version = dte.Version;

            Executor.executionEngine = new Process();
            string processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualStudio.Macros.ExecutionEngine.exe");
            Executor.executionEngine.StartInfo.FileName = processName;
            Executor.executionEngine.StartInfo.UseShellExecute = false;
            Executor.executionEngine.StartInfo.Arguments = ProvidePipeArguments(Server.Guid, version);
            Executor.executionEngine.Start();

            Executor.IsEngineInitialized = true;
            Executor.Job = JobHandle.CreateNewJob();
            Executor.Job.AddProcess(Executor.executionEngine);
        }

        /// <summary>
        /// If engine is initialized, runs the engine.  Otherwise, initializes and runs the engine.
        /// </summary>
        /// 
        public void RunEngine(int iterations, string path)
        {
            if (Executor.IsEngineInitialized && Server.ServerStream.IsConnected)
            {
                Server.SendFilePath(iterations, path);
            }
            else
            {
                this.InitializeEngine();
                Thread waitsUntilClientIsConnected = new Thread(() =>
                    {
                        while (true) 
                        {
                            if (Server.ServerStream.IsConnected)
                            {
                                Server.SendFilePath(iterations, path);
                                break;
                            }
                        }
                    }
                );
                waitsUntilClientIsConnected.Start();
            }
        }
    }
}

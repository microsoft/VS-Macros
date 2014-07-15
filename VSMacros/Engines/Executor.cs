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
using System.Windows;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
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
        public static Process executionEngine;
        public static bool IsEngineInitialized = false;
        public static Job Job;

        /// <summary>
        /// Informs subscribers of an error during execution.
        /// </summary>
        public event EventHandler<CompletionReachedEventArgs> Complete;

        public void SendCompletionMessage(bool isError, string errorMessage)
        {
            if (this.Complete != null)
            {
                var eventArgs = new CompletionReachedEventArgs(isError, errorMessage);
                this.Complete(this, eventArgs);
            }
        }

        #region Helpers
        private string ProvideArguments(int iterations, string path)
        {
            string pid = Process.GetCurrentProcess().Id.ToString();
            var delimiter = "[delimiter]";
            return pid + delimiter + iterations.ToString() + delimiter + path;
        }

        private string ProvidePipeArguments(Guid guid)
        {
            var pipeToken = "@";
            var delimiter = "[delimiter]";
            string pid = Process.GetCurrentProcess().Id.ToString();

            return pipeToken + delimiter + guid.ToString() + delimiter + pid;
        }

        private System.Runtime.InteropServices.ComTypes.IRunningObjectTable GetRunningObjectTable()
        {
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            ErrorHandler.ThrowOnFailure(NativeMethods.GetRunningObjectTable(0, out rot));

            return rot;
        }

        private static string GetProgID(Type type)
        {
            Guid guid = Marshal.GenerateGuidForType(type); // what about IOleCommandTarget?
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

            Executor.executionEngine = new Process();
            string processName = @"C:\Users\t-grawa\Source\Repos\Macro Extension\ExecutionEngine\bin\Debug\VisualStudio.Macros.ExecutionEngine.exe";
            //string processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualStudio.Macros.ExecutionEngine.exe");
            Executor.executionEngine.StartInfo.FileName = processName;
            Executor.executionEngine.StartInfo.Arguments = ProvidePipeArguments(Server.Guid);
            Executor.executionEngine.Start();

            Server.ServerStream.WaitForConnection();
            Server.serverWait = new Thread(new ThreadStart(Server.WaitForMessage));
            Server.serverWait.Start();

            Executor.IsEngineInitialized = true;
            Executor.Job = new Job();
            Executor.Job.AddProcess(Executor.executionEngine.Handle);
        }

        /// <summary>
        /// Initializes the engine.
        /// </summary>
        /// 
        public void RunEngine(int iterations, string path)
        {
            if (Server.ServerStream.IsConnected)
            {
                Server.SendFilePath(iterations, path);
            }
            else
            {
                InitializeEngine();
            }
        }
    }
}

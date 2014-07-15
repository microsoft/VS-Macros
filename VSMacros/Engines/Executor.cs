//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Helpers;
using VSMacros.Interfaces;
using VSMacros.Pipes;



using System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Reflection;
//using System.Windows.Forms;

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
        public static bool IsServerInitialized = false;
        public static Job job;

        /// <summary>
        /// Informs subscribers of an error during execution.
        /// </summary>
        public event EventHandler<CompletionReachedEventArgs> Complete;

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

        System.Runtime.InteropServices.ComTypes.IRunningObjectTable GetRunningObjectTable()
        {
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            int hr = NativeMethods.GetRunningObjectTable(0, out rot);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return rot;
        }

        public void RegisterCommandDispatcherinROT()
        {
            int hResult;
            System.Runtime.InteropServices.ComTypes.IBindCtx bc;
            hResult = NativeMethods.CreateBindCtx(0, out bc);
            if (hResult == NativeMethods.S_OK && bc != null)
            {
                System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot = GetRunningObjectTable();
                if (rot != null)
                {
                    string delim = "!";
                    Guid guid = Marshal.GenerateGuidForType(typeof(SUIHostCommandDispatcher)); // what about IOleCommandTarget?

                    string progID = String.Format(CultureInfo.InvariantCulture, "{0:B}", guid);

                    System.Runtime.InteropServices.ComTypes.IMoniker moniker;
                    hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);

                    if (hResult == NativeMethods.S_OK && moniker != null)
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
        }

        public void RegisterCmdNameMappinginROT()
        {
            // thank you http://index/#HlpViewer/Application/HelpViewer.cs,452 you were very helpful

            int hResult;
            System.Runtime.InteropServices.ComTypes.IBindCtx bc;
            hResult = NativeMethods.CreateBindCtx(0, out bc);
            if (hResult == NativeMethods.S_OK && bc != null)
            {
                System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot = GetRunningObjectTable();
                if (rot != null)
                {
                    string delim = "!";
                    Guid guid = Marshal.GenerateGuidForType(typeof(SVsCmdNameMapping)); // what about IVsCmdNameMapping?  // think i can just cast later

                    string progID = String.Format(CultureInfo.InvariantCulture, "{0:B}", guid);
                    System.Runtime.InteropServices.ComTypes.IMoniker moniker;
                    hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);

                    if (hResult == NativeMethods.S_OK && moniker != null)
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
        }

        /// <summary>
        /// Initializes the engine and then runs the macro script.
        /// This method will be removed after IPC is implemented.
        /// </summary>
        /// 
        public void InitializeEngine()
        {
            this.RegisterCmdNameMappinginROT();
            this.RegisterCommandDispatcherinROT();
  
            Server.InitializeServer();
            Executor.executionEngine = new Process();

            // Debug.WriteLine("for some reason it's not finding the executable, so the path is hardcoded for now");
             string processName = @"C:\Users\t-grawa\Source\Repos\Macro Extension\ExecutionEngine\bin\Debug\VisualStudio.Macros.ExecutionEngine.exe";
            //string processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualSt .Macros.ExecutionEngine.exe");

            Executor.executionEngine.StartInfo.FileName = processName;
            Executor.executionEngine.StartInfo.Arguments = ProvidePipeArguments(Server.Guid);

            Debug.WriteLine(string.Format("guid of server is: {0}", Server.Guid));

            Executor.executionEngine.Start();

            Server.ServerStream.WaitForConnection();
            Server.serverWait = new Thread(new ThreadStart(Server.WaitForMessage));
            Server.serverWait.Start();

            Executor.IsEngineInitialized = true;

            job = new Job();
            job.AddProcess(Executor.executionEngine.Handle);
        }

        internal void RunEngine(int iterations, string path)
        {
            if (Server.ServerStream.IsConnected)
            {
                Server.SendFilePath(iterations, path);
            }
            else
            {
                MessageBox.Show("The server is not connected.");
                InitializeEngine();
            }
        }

        /// <summary>
        /// Will run the macro file.
        /// </summary>
        public void StartExecution(int iterations, string path)
        {
            Debug.WriteLine("path is: " + path);

            string processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualStudio.Macros.ExecutionEngine.exe");
            string encodedPath = path.Replace(" ", "%20");
            Executor.executionEngine = new Process();
            Executor.executionEngine.StartInfo.FileName = processName;
            Executor.executionEngine.StartInfo.UseShellExecute = false;
            Executor.executionEngine.StartInfo.Arguments = ProvideArguments(iterations, encodedPath);
            Debug.WriteLine("arguments are: " + Executor.executionEngine.StartInfo.Arguments);
            Executor.executionEngine.Start();
        }

        /// <summary>
        /// Will stop the currently executing macro file.
        /// We are considering removing this.
        /// </summary>
        public void StopExecution()
        {
            // throw new NotImplementedException();
        }
    }
}

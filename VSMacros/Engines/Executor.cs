//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Reflection;
using System.Threading;
using System.Windows;
using MicrosoftCorporation.VSMacros.Pipes;
//using EnvDTE;
using VSMacros.Engines;
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
        private Process executionEngine;
        public static bool IsEngineInitialized = false;
        public static bool IsServerInitialized = false;
        private static bool enablePipes = false;

        /// <summary>
        /// Informs subscribers of an error during execution.
        /// </summary>
        public event EventHandler<CompletionReachedEventArgs> Complete;

        private string ProvideArguments(int iterations, string path)
        {
            string pid = Process.GetCurrentProcess().Id.ToString();
            var delimiter = "[delimiter]";
            //path = string.Replac
            return pid + delimiter + iterations.ToString() + delimiter + path;
        }

        private string ProvideClientArguments(Guid guid)
        {
            var pipeToken = "@";
            var delimiter = "[delimiter]";
            string pid = Process.GetCurrentProcess().Id.ToString();

            return pipeToken + delimiter + guid.ToString() + delimiter + pid;
        }

        /// <summary>
        /// Initializes the engine and then runs the macro script.
        /// This method will be removed after IPC is implemented.
        /// </summary>
        /// 

        public void InitializeEngine()
        {
            Debug.WriteLine("Initializing the engine");
            string processName;
  
            Server.InitializeServer();
            var client = new System.Diagnostics.Process();

            Debug.WriteLine("for some reason it's not finding the executable, so the path is hardcoded for now");
            processName = @"C:\Users\t-grawa\Source\Repos\Macro Extension\ExecutionEngine\bin\Debug\VisualStudio.Macros.ExecutionEngine.exe";
            //var processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualStudio.Macros.ExecutionEngine.exe");

            client.StartInfo.FileName = processName;
            client.StartInfo.Arguments = ProvideClientArguments(Server.Guid);

            Debug.WriteLine(string.Format("guid of server is: {0}", Server.Guid));

            client.Start();

            Server.ServerStream.WaitForConnection();
            Server.serverWait = new Thread(new ThreadStart(Server.WaitForMessage));
            Server.serverWait.Start();


            Executor.IsEngineInitialized = true;
        }

        public void RunEngine(string path)
        {
            if (Server.ServerStream.IsConnected)
            {
                //MessageBox.Show("hello");
                //Server.SendMessage();
                Server.SendFilePath(path);
            }
            else
            {
                MessageBox.Show("The server is not connected.");
            }
        }

        /// <summary>
        /// Will run the macro file.
        /// </summary>
        public void StartExecution(string path, int iterations)
        {
            Debug.WriteLine("path is: " + path);

            Debug.WriteLine("for some reason it's not finding the executable, so the path is hardcoded for now");
            var processName = @"C:\Users\t-grawa\Source\Repos\Macro Extension\ExecutionEngine\bin\Debug\VisualStudio.Macros.ExecutionEngine.exe";
            //var processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualStudio.Macros.ExecutionEngine.exe");
            var encodedPath = path.Replace(" ", "%20");
            this.executionEngine = new Process();
            this.executionEngine.StartInfo.FileName = processName;
            this.executionEngine.StartInfo.UseShellExecute = false;
            this.executionEngine.StartInfo.Arguments = ProvideArguments(iterations, encodedPath);
            Debug.WriteLine("arguments are: " + this.executionEngine.StartInfo.Arguments);
            this.executionEngine.Start();
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

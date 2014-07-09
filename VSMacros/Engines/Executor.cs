//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
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
        public static bool IsServerInitialized = false;

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

        private string ProvidePipeArguments(Guid guid)
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
            Executor.executionEngine = new Process();

            Debug.WriteLine("for some reason it's not finding the executable, so the path is hardcoded for now");
            processName = @"C:\Users\t-grawa\Source\Repos\Macro Extension\ExecutionEngine\bin\Debug\VisualStudio.Macros.ExecutionEngine.exe";
            //var processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualStudio.Macros.ExecutionEngine.exe");

            Executor.executionEngine.StartInfo.FileName = processName;
            Executor.executionEngine.StartInfo.Arguments = ProvidePipeArguments(Server.Guid);

            Debug.WriteLine(string.Format("guid of server is: {0}", Server.Guid));

            Executor.executionEngine.Start();

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
                InitializeEngine();
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

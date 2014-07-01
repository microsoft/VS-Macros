//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using MicrosoftCorporation.VSMacros.Engines;
using VSMacros.Interfaces;

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

        /// <summary>
        /// Initializes the engine and then runs the macro script.
        /// This method will be removed after IPC is implemented.
        /// </summary>
        public void InitializeEngine()
        {
        }


        /// <summary>
        /// Will run the macro file.
        /// </summary>
        public void StartExecution(string path, int iterations)
        {
            var processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "VisualStudio.Macros.ExecutionEngine.exe");
            var encodedPath = path.Replace(" ", "%20");
            this.executionEngine = new Process();
            this.executionEngine.StartInfo.FileName = processName;
            this.executionEngine.StartInfo.UseShellExecute = false;
            this.executionEngine.StartInfo.Arguments = ProvideArguments(iterations, encodedPath);
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

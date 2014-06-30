//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using VSMacros.Interfaces;
using VSMacros.Engines;

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

        private string ProvideArguments(int iterations, string script)
        {
            string pid = Process.GetCurrentProcess().Id.ToString();
            var delimiter = "[delimiter]";
            return pid + delimiter + iterations.ToString() + delimiter + script;
        }

        /// <summary>
        /// Initializes the engine and then runs the macro script.
        /// This method will be removed after IPC is implemented.
        /// </summary>
        public void InitializeEngine()
        {
        }

        private static string CreateScriptFromReader(StreamReader reader)
        {
            var script = string.Empty;
            using (reader)
            {
                script = reader.ReadToEnd();
            }
            return script;
        }

        /// <summary>
        /// Will run the macro file.
        /// </summary>
        public void StartExecution(StreamReader reader, int iterations)
        {
            var processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ExecutionEngine.exe");
            var script = CreateScriptFromReader(reader);
            this.executionEngine = new Process();

            this.executionEngine.StartInfo.FileName = processName;
            this.executionEngine.StartInfo.UseShellExecute = false;
            this.executionEngine.StartInfo.Arguments = ProvideArguments(iterations, script);
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

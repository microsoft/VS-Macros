//-----------------------------------------------------------------------
// <copyright file="Executor.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace VSMacros.Engines
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;

    /// <summary>
    /// Exposes the execution engine.
    /// </summary>
    internal interface IExecutor
    {
        /// <summary>
        /// Informs subscribers of an error during execution.
        /// </summary>
        event EventHandler OnError;

        /// <summary>
        /// Informs subscribers of success after execution.
        /// </summary>
        event EventHandler OnSuccess;

        /// <summary>
        /// Initializes the engine and then runs the macro script.
        /// This method will be removed after IPC is implemented.
        /// </summary>
        void InitializeAndRunEngine();

        /// <summary>
        /// Will run the macro file.
        /// <param name="macro">Name of macro.</param>
        /// <param name="times">Times to be executed.</param>
        /// </summary>
        void StartExecution(StreamReader macro, int times);

        /// <summary>
        /// Will stop the currently executing macro file.
        /// We are considering removing this.
        /// </summary>
        void StopExecution();
    }

    /// <summary>
    /// Implements the execution engine.
    /// </summary>   
    internal class Executor : IExecutor
    {
        /// <summary>
        /// The execution engine.
        /// </summary>
        private Process executionEngine;

        /// <summary>
        /// Informs subscribers of an error during execution.
        /// </summary>
        public event EventHandler OnError;

        /// <summary>
        /// Informs subscribers of success after execution.
        /// </summary>
        public event EventHandler OnSuccess;

        /// <summary>
        /// Initializes the engine and then runs the macro script.
        /// This method will be removed after IPC is implemented.
        /// </summary>
        public void InitializeAndRunEngine()
        {
            var currentProcess = Process.GetCurrentProcess();
            var processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ExecutionEngine.exe");
            this.executionEngine = new Process();

            this.executionEngine.StartInfo.FileName = processName;
            this.executionEngine.StartInfo.Arguments = Process.GetCurrentProcess().Id.ToString();
            this.executionEngine.Start();
        }

        /// <summary>
        /// Will run the macro file.
        /// </summary>
        public void StartExecution(StreamReader macro, int times)
        {
            // throw new NotImplementedException();
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

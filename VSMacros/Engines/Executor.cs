using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace VSMacros.Engines
{
    interface IExecutor
    {
        void InitializeAndRunEngine();
        void StartExecution(Stream macro, int times);
        void StopExecution();
        event EventHandler OnError;
        event EventHandler OnSuccess;
    }
        
    internal class Executor : IExecutor
    {
        Process executionEngine;

        public void InitializeAndRunEngine()
        {
            var currentProcess = Process.GetCurrentProcess();
            var processName = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ExecutionEngine.exe");
            executionEngine = new Process();

            executionEngine.StartInfo.FileName = processName;
            executionEngine.StartInfo.Arguments = Process.GetCurrentProcess().Id.ToString();
            executionEngine.Start();
        }

        public void StartExecution(Stream macro, int times)
        {
            //throw new NotImplementedException();
        }

        public void StopExecution()
        {
            //throw new NotImplementedException();
        }

        public event EventHandler OnError;

        public event EventHandler OnSuccess;
    }
}

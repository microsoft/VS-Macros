using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell.Interop;
using System.IO;
using System.Reflection;

namespace VSMacros
{
    class MacroRecorder:IMacroRecorder, IMacroRecorderPrivate, IDisposable
    {
        private WindowActivationWatcher activationWatcher;
        private CommandExecutionWatcher commandWatcher;
        private DataModel dataModel = new DataModel();
        private IServiceProvider serviceProvider;
        private bool recording;

        public MacroRecorder(IServiceProvider serviceProvider)
        {
            this.serviceProvider = serviceProvider;
        }
        public void StartRecording()
        {
            this.ClearData();

            if (this.activationWatcher == null)
            {
                this.activationWatcher = new WindowActivationWatcher(serviceProvider: this.serviceProvider);
            }
            if (this.commandWatcher == null)
            {
                this.commandWatcher = new CommandExecutionWatcher(serviceProvider: this.serviceProvider);
            }
            this.recording = true;         
        }

        public void StopRecording()
        {
            this.recording = false;
            //string ParentPath = "c:\\users\\t-xindo\\documents\\vsmacros\\vsmacros\\vsmacros";
            //string fileName = Path.Combine(ParentPath, "test1.txt");

            //using (StreamWriter fs = new StreamWriter(fileName))
            //{
            //    foreach (var action in this.dataModel.Actions)
            //    {
            //        action.ConvertToJavascript(fs);
            //    }
            //}
        }

        public void AddCommandData(Guid commandSet, uint identifier, string commandName, char input)
        {
            this.dataModel.AddExecutedCommand(commandSet,identifier,commandName,input);
        }

        public void AddWindowActivation(Guid toolWindowID, string name)
        {
            this.dataModel.AddWindow(toolWindowID,name);
        }

        public void AddWindowActivation(string path)
        {
            this.dataModel.AddWindow(path);
        }

        public void ClearData()
        {
            this.dataModel.ClearActions();
        }

        public void Dispose()
        {
            using(this.commandWatcher)
            using(this.activationWatcher)
            {
                this.commandWatcher = null;
                this.activationWatcher = null;
            }
        }

        public bool Recording
        {
            get { return this.recording; }
        }
    }
}

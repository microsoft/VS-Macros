using System;
using Microsoft.Internal.VisualStudio.Shell;
using VSMacros.Interfaces;

namespace VSMacros.Engines
{
    class Recorder:IRecorder, IRecorderPrivate, IDisposable
    {
        private WindowActivationWatcher activationWatcher;
        private CommandExecutionWatcher commandWatcher;
        private RecorderDataModel dataModel;
        private IServiceProvider serviceProvider;
        private bool recording;

        public Recorder(IServiceProvider serviceProvider)
        {
            Validate.IsNotNull(serviceProvider, "serviceProvider");
            this.serviceProvider = serviceProvider;
            dataModel = new RecorderDataModel();
        }
        public void StartRecording()
        {
            this.ClearData();
            this.activationWatcher = this.activationWatcher ?? new WindowActivationWatcher(this.serviceProvider);
            this.commandWatcher = this.commandWatcher?? new CommandExecutionWatcher(this.serviceProvider);
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

        public bool Recording
        {
            get { return this.recording; }
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
    }
}

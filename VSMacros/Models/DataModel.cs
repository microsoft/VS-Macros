using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace VSMacros
{
    internal sealed class DataModel
    {
        private string currWindow = "Macro Explorer";
        private readonly ObservableCollection<RecordedActionBase> actions = new ObservableCollection<RecordedActionBase>();
        private bool recording = false;

        public void CurrWindow(string window)
        {
             this.currWindow = window;
        }

        public IEnumerable<RecordedActionBase> Actions
        {
            get
            {
                return this.actions;
            }
        }

        internal void AddExecutedCommand(Guid commandSet, uint identifier, string commandName, char input)
        {
            this.actions.Add(new RecordedCommand(commandSet, identifier, commandName, input));
        }

        internal void AddWindow(Guid toolWindowID,string name)
        {
            this.actions.Add(new RecordedWindowActivation(toolWindowID,name));
        }

        internal void AddWindow(string path)
        {
            this.actions.Add(new RecordedDocumentActivation(path));
        }

        public void ClearActions()
        {
            this.actions.Clear();
        }
    }
}

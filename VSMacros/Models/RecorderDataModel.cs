using System.Collections.Generic;
using System.Collections.ObjectModel;
using System;

namespace VSMacros
{
    internal sealed class RecorderDataModel
    {
        private string currWindow = "Macro Explorer";
        private readonly List<RecordedActionBase> actions = new List<RecordedActionBase>();
        private bool recording = false;

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

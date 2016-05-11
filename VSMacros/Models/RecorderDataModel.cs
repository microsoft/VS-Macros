//-----------------------------------------------------------------------
// <copyright file="RecorderDataModel.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using VSMacros.RecorderOutput;

namespace VSMacros.Model
{
    internal sealed class RecorderDataModel
    {
        private readonly List<RecordedActionBase> actions = new List<RecordedActionBase>();
        public Guid currWindow;
        public string currDoc;
        public bool isDoc = false;

        public List<RecordedActionBase> Actions
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

        internal void AddWindow(Guid toolWindowID, string name)
        {
            this.isDoc = false;
            this.actions.Add(new RecordedWindowActivation(toolWindowID, name));
            this.currWindow = toolWindowID;
        }

        internal void AddWindow(string path)
        {
            this.isDoc = true;
            this.actions.Add(new RecordedDocumentActivation(path));
            this.currDoc = path;
        }

        public void ClearActions()
        {
            this.actions.Clear();
        }
    }
}

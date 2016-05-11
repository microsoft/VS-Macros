//-----------------------------------------------------------------------
// <copyright file="Recorder.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Internal.VisualStudio.Shell;
using VSMacros.Interfaces;
using VSMacros.Model;
using VSMacros.RecorderListeners;
using VSMacros.RecorderOutput;

namespace VSMacros.Engines
{
    class Recorder : IRecorder, IRecorderPrivate, IDisposable
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
            this.activationWatcher = this.activationWatcher ?? new WindowActivationWatcher(serviceProvider: this.serviceProvider, dataModel: this.dataModel);
        }

        public void StartRecording()
        {
            this.ClearData();
            this.commandWatcher = this.commandWatcher ?? new CommandExecutionWatcher(this.serviceProvider);
            this.recording = true;
        }

        public void StopRecording(string path)
        {
            using (StreamWriter fs = new StreamWriter(path))
            {
                // Add reference to DTE for Intellisense
                fs.WriteLine(string.Format("/// <reference path=\"" + Manager.IntellisensePath + "\" />{0}", Environment.NewLine));

                bool inDocument = Manager.Instance.PreviousWindowIsDocument;

                for (int i = 0; i < this.dataModel.Actions.Count; i++)
               {
                    RecordedActionBase action = this.dataModel.Actions[i];

                    if (action is RecordedCommand)
                    {
                        RecordedCommand current = action as RecordedCommand;
                        RecordedCommand empty = new RecordedCommand(Guid.Empty, 0, string.Empty, '\0');

                        // If next action is a recorded command, try to merge
                        if (i < this.dataModel.Actions.Count - 1 &&
                        this.dataModel.Actions[i + 1] is RecordedCommand)
                        {
                            RecordedCommand next = this.dataModel.Actions[i + 1] as RecordedCommand;

                            if (current.IsInsert())
                            {
                                List<char> buffer = new List<char>();

                                // Setup for the loop
                                next = current;

                                // Get all the characters that forms the input string
                                do
                                {
                                    if (next.IsValidCharacter())
                                    {
                                        buffer.Add(next.Input);
                                    }

                                    next = this.dataModel.Actions[++i] as RecordedCommand ?? empty;
                                } while (next.IsInsert() && i + 1 < this.dataModel.Actions.Count);

                                // Process last character
                                if (next.IsInsert())
                                {
                                    if (next.IsValidCharacter())
                                    {
                                        buffer.Add(next.Input);
                                    }
                                }
                                else
                                {
                                    // The loop has incremented i an extra time, backtrack
                                    i--;
                                }

                                // Output the text
                                current.ConvertToJavascript(fs, buffer);

                                buffer = new List<char>();
                            }
                            else
                            {
                                // Compute the number of iterations of the same command
                                int iterations = 1;
                                while (current.CommandName == next.CommandName && (i + 2 < this.dataModel.Actions.Count || i + 1 == this.dataModel.Actions.Count))
                                {
                                    iterations++;
                                    current = next;
                                    next = this.dataModel.Actions[++i + 1] as RecordedCommand ?? empty;
                                }

                                if (current.CommandName == next.CommandName)
                                {
                                    iterations++;
                                    i++;
                                }

                                current.ConvertToJavascript(fs, iterations, inDocument);
                            }
                        }
                        else
                        {
                            if (current.CommandName == "keyboard")
                            {
                                current.ConvertToJavascript(fs, new List<char>() { current.Input });
                            }
                            else
                            {
                                current.ConvertToJavascript(fs, 1, inDocument);
                            }
                        }
                    }
                    else
                    {
                        action.ConvertToJavascript(fs);
                        inDocument = action is RecordedDocumentActivation;
                    }
                }
            }

            this.recording = false;
        }

        public bool IsRecording
        {
            get { return this.recording; }
        }

        public void AddCommandData(Guid commandSet, uint identifier, string commandName, char input)
        {
            this.dataModel.AddExecutedCommand(commandSet, identifier, commandName, input);
        }

        public void AddWindowActivation(Guid toolWindowID, string name)
        {
            this.dataModel.AddWindow(toolWindowID, name);
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
            using (this.commandWatcher)
            using (this.activationWatcher)
            {
                this.commandWatcher = null;
                this.activationWatcher = null;
            }
        }
    }
}

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
        }
        public void StartRecording()
        {
            this.ClearData();
            this.activationWatcher = this.activationWatcher ?? new WindowActivationWatcher(this.serviceProvider);
            this.commandWatcher = this.commandWatcher ?? new CommandExecutionWatcher(this.serviceProvider);
            this.recording = true;
        }

        public void StopRecording(string path)
        {
            using (StreamWriter fs = new StreamWriter(path))
            {
                // Add reference to DTE for Intellisense
                fs.WriteLine(string.Format("/// <reference path=\"" + Manager.dteIntellisensePath + "\" />{0}", Environment.NewLine));

                for (int i = 0; i < this.dataModel.Actions.Count; i++)
                {
                    RecordedActionBase action = this.dataModel.Actions[i];

                    // If both current and next commands are RecordedCommand and if the next command exists
                    if (action is RecordedCommand && i < this.dataModel.Actions.Count - 1 && this.dataModel.Actions[i + 1] is RecordedCommand)
                    {
                        RecordedCommand current = action as RecordedCommand;
                        RecordedCommand next = this.dataModel.Actions[i + 1] as RecordedCommand;                        

                        if (current.IsInsert())
                        {
                            List<char> buffer = new List<char>();

                            // Setup for the loop
                            next = current;

                            // Get all the characters that forms the input string
                            do
                            {
                                if (next.Input != '\0')
                                {
                                    buffer.Add(next.Input);
                                }

                                next = this.dataModel.Actions[++i] as RecordedCommand;
                            } while (next.IsInsert() && i + 1 < this.dataModel.Actions.Count);

                            // Process last character
                            if (next.IsInsert())
                            {
                                if (next.Input != '\0')
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
                            while (current.CommandName == next.CommandName && i + 2 < this.dataModel.Actions.Count)
                            {
                                iterations++;
                                current = next;
                                next = this.dataModel.Actions[++i + 1] as RecordedCommand;
                            }

                            current.ConvertToJavascript(fs, iterations);
                        }
                    }
                    else
                    {
                        action.ConvertToJavascript(fs);
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

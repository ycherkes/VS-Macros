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
        private readonly RecorderDataModel dataModel;
        private readonly IServiceProvider serviceProvider;
        private bool recording;

        public Recorder(IServiceProvider serviceProvider)
        {
            Validate.IsNotNull(serviceProvider, "serviceProvider");
            this.serviceProvider = serviceProvider;
            dataModel = new RecorderDataModel();
            activationWatcher = new WindowActivationWatcher(serviceProvider: serviceProvider, dataModel: dataModel);
        }

        public void StartRecording()
        {
            ClearData();
            commandWatcher = commandWatcher ?? new CommandExecutionWatcher(serviceProvider);
            recording = true;
        }

        public void StopRecording(string path)
        {
            using (StreamWriter fs = new StreamWriter(path))
            {
                // Add reference to DTE for Intellisense
                fs.WriteLine("/// <reference path=\"" + Manager.IntellisensePath + "\" />{0}", Environment.NewLine);

                bool inDocument = Manager.Instance.PreviousWindowIsDocument;

                for (int i = 0; i < dataModel.Actions.Count; i++)
                {
                    RecordedActionBase action = dataModel.Actions[i];

                    if (action is RecordedCommand)
                    {
                        RecordedCommand current = action as RecordedCommand;
                        RecordedCommand empty = new RecordedCommand(Guid.Empty, 0, string.Empty, '\0');

                        // If next action is a recorded command, try to merge
                        if (i < dataModel.Actions.Count - 1 &&
                            dataModel.Actions[i + 1] is RecordedCommand)
                        {
                            RecordedCommand next = dataModel.Actions[i + 1] as RecordedCommand;

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

                                    next = dataModel.Actions[++i] as RecordedCommand ?? empty;
                                } while (next.IsInsert() && i + 1 < dataModel.Actions.Count);

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
                                while (current.CommandName == next.CommandName && (i + 2 < dataModel.Actions.Count || i + 1 == dataModel.Actions.Count))
                                {
                                    iterations++;
                                    current = next;
                                    next = dataModel.Actions[++i + 1] as RecordedCommand ?? empty;
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

            recording = false;
        }

        public bool IsRecording
        {
            get { return recording; }
        }

        public void AddCommandData(Guid commandSet, uint identifier, string commandName, char input)
        {
            dataModel.AddExecutedCommand(commandSet, identifier, commandName, input);
        }

        public void AddWindowActivation(Guid toolWindowID, string name)
        {
            dataModel.AddWindow(toolWindowID, name);
        }

        public void AddWindowActivation(string path)
        {
            dataModel.AddWindow(path);
        }

        public void ClearData()
        {
            dataModel.ClearActions();
        }

        public void Dispose()
        {
            using (commandWatcher)
            using (activationWatcher)
            {
                commandWatcher = null;
                activationWatcher = null;
            }
        }
    }
}

﻿//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Windows.Forms;
using ExecutionEngine.Helpers;
using Microsoft.Internal.VisualStudio.Shell;
using VisualStudio.Macros.ExecutionEngine.Pipes;
using VSMacros.ExecutionEngine.Pipes;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static BinaryFormatter serializer;
        internal const string MacroName = "currentScript";

        internal static void RunMacro(string script, int iterations)
        {
            Validate.IsNotNullAndNotEmpty(script, "script");
            Program.engine.Parse(script);

            int i = 0;
            for (i = 0; i < iterations; i++)
            {
                if (!Program.engine.CallMethod(Program.MacroName))
                {
                    if (Site.RuntimeError)
                    {
                        uint activeDocumentModification = 1;
                        uint macroInsertTextModification = 0;
                        var e = Site.RuntimeException;
                        uint modifiedLineNumber = e.Line - activeDocumentModification + macroInsertTextModification;

                        Client.SendScriptError(modifiedLineNumber, e.CharacterPosition, e.Source, e.Description);
                    }
                    else
                    {
                        var e = Site.InternalVSException;
                        Client.SendCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                    }
                    Site.ResetError();
                    break;
                }
            }

            Client.SendSuccessMessage();
        }

        private static void HandleInput()
        {
            var type = (PacketType)Program.serializer.Deserialize(Client.ClientStream);

            // I know a switch statement seems useless but just preparing for the possibility of other packets.
            switch (type)
            {
                case PacketType.FilePath:
                    HandleFilePath();
                    break;
            }
        }

        private static void HandleFilePath()
        {
            // Just make one static formatter
            var filePath = (FilePath)Program.serializer.Deserialize(Client.ClientStream);

            int iterations = filePath.Iterations;
            string message = filePath.Path;
            string unwrappedScript = InputParser.ExtractScript(message);
            string wrappedScript = InputParser.WrapScript(unwrappedScript);
            Program.RunMacro(wrappedScript, iterations);
        }

        internal static Thread CreateReadingExecutingThread(int pid, string version)
        {
            Thread readAndExecuteThread = new Thread(() =>
            {
                try
                {
                    Program.serializer = new BinaryFormatter();
                    Program.serializer.Binder = new BinderHelper();
                    Program.engine = new Engine(pid, version);

                    while (true)
                    {
                        HandleInput();
                    }
                }
                catch (Exception e)
                {
                    if (Client.ClientStream.IsConnected)
                    {
                        Client.SendCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                    }
                    else
                    {
#if DEBUG
                        MessageBox.Show("Execution engine's pipe is not connected.");
#endif
                    }
                }
                finally
                {
                    if (Client.ClientStream.IsConnected)
                    {
                        Client.ShutDownServer(Client.ClientStream);
                        Client.ClientStream.Close();
                    }
                    else
                    {
#if DEBUG
                        MessageBox.Show("Execution engine's pipe is not connected.");
#endif
                    }
                }
            });

            readAndExecuteThread.SetApartmentState(ApartmentState.STA);
            return readAndExecuteThread;
        }

        private static void RunFromPipe(string[] separatedArgs)
        {
            Guid guid = InputParser.GetGuid(separatedArgs[0]);
            Client.InitializePipeClientStream(guid);

            int pid = InputParser.GetPid(separatedArgs[1]);
            string version = separatedArgs[2];

            Thread readAndExecuteThread = CreateReadingExecutingThread(pid, version);
            readAndExecuteThread.Start();
        }

        internal static void Main(string[] args)
        {
            try
            {
                MessageBox.Show("hello");
                string[] separatedArgs = InputParser.SeparateArgs(args);
                RunFromPipe(separatedArgs);
            }
            catch (Exception e)
            {
                Client.SendCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
            }
        }
    }
}

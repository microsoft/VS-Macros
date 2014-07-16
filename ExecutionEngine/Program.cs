//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using ExecutionEngine.Enums;
using ExecutionEngine.Helpers;
using VSMacros.ExecutionEngine.Pipes;
using VisualStudio.Macros.ExecutionEngine;
using Microsoft.Internal.VisualStudio.Shell;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        internal static string MacroName = "currentScript";
        private static bool exit;

        internal static void RunMacro(string script, int iterations)
        {
            Validate.IsNotNullAndNotEmpty(script, "script");
            Program.parsedScript = Program.engine.Parse(script);

            for (int i = 0; i < iterations; i++)
            {
                Program.parsedScript.CallMethod(Program.MacroName);
            }
        }

        private static void HandleInput()
        {
            int typeOfMessage = Client.GetInt(Client.ClientStream);

            switch ((Packet)typeOfMessage)
            {
                case Packet.FilePath:
                    HandleFilePath();
                    break;

                case Packet.Close:
                    HandleCloseRequest();
                    break;
            }
        }

        private static void HandleCloseRequest()
        {
            Client.ShutDownServer(Client.ClientStream);
            Client.ClientStream.Close();
            Program.exit = true;
        }

        private static void HandleFilePath()
        {
            int iterations = Client.GetIterations(Client.ClientStream);
            string message = Client.GetFilePath(Client.ClientStream);
            string unwrappedScript = InputParser.ExtractScript(message);
            string wrappedScript = InputParser.WrapScript(unwrappedScript);
            Program.RunMacro(wrappedScript, iterations);
        }

        internal static Thread CreateReadingThread(int pid)
        {
            Program.exit = false;
            Thread readThread = new Thread(() =>
            {
                try
                {
                    Program.engine = new Engine(pid);
                    while (!Program.exit)
                    {
                        HandleInput();
                    }
                }
                catch (Exception e)
                {
                    Client.ShutDownServer(Client.ClientStream);
                    Client.ClientStream.Close();
                    Program.exit = true;

                    byte[] criticalError = Client.PackageCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                    Client.SendMessageToServer(Client.ClientStream, criticalError);
                }
            });

            readThread.SetApartmentState(ApartmentState.STA);
            return readThread;
        }

        private static void RunFromPipe(string[] separatedArgs)
        {
            Guid guid = InputParser.GetGuid(separatedArgs[0]);
            int pid = InputParser.GetPid(separatedArgs[1]);

            Client.InitializePipeClientStream(guid);

            Thread readThread = CreateReadingThread(pid);
            readThread.Start();
        }

        internal static void Main(string[] args)
        {
            try
            {
                string[] separatedArgs = InputParser.SeparateArgs(args);
                RunFromPipe(separatedArgs);
            }
            catch (Exception e)
            {
                byte[] criticalError = Client.PackageCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                Client.SendMessageToServer(Client.ClientStream, criticalError);
            }
        }
    }
}

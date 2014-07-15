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
using VSMacros.ExecutionEngine;
using VSMacros.ExecutionEngine.Helpers;
using VSMacros.ExecutionEngine.Pipes;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        internal static string MacroName = "currentScript";
        private const string close = "close";
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

        internal static void RunFromExtension(string[] separatedArgs)
        {
            string unparsedPid = separatedArgs[0];
            string unparsedIter = separatedArgs[1];
            string encodedPath = separatedArgs[2];

            int pid = InputParser.GetPid(unparsedPid);
            short iterations = InputParser.GetNumberOfIterations(unparsedIter);
            string decodedPath = InputParser.DecodePath(encodedPath);
            var unwrapped = InputParser.ExtractScript(decodedPath);
            var wrapped = InputParser.WrapScript(unwrapped);

            Program.engine = new Engine(pid);
            RunMacro(wrapped, iterations);
        }

        internal static void RunAsStartupProject()
        {
            Debug.WriteLine("Warning: Hardcoded devenv pid");
            short pidOfCurrentDevenv = 9800;

            Program.engine = new Engine(pidOfCurrentDevenv);

            Debug.WriteLine("Warning: Hardcoded path");
            string unwrapped = File.ReadAllText(@"C:\Users\t-grawa\Desktop\test.js");
            string wrapped = InputParser.WrapScript(unwrapped);

            RunMacro(wrapped, 1);
        }

        private static void HandleInput()
        {
            int typeOfMessage = Client.GetIntFromStream(Client.ClientStream);

            switch ((Packet)typeOfMessage)
            {
                case (Packet.FilePath):
                    HandleFilePath();
                    break;

                case (Packet.Close):
                    HandleCloseRequest();
                    break;

                case (Packet.Success):
                    Client.HandlePacketSuccess(Client.ClientStream);
                    break;

                case (Packet.ScriptError):
                    Client.HandlePacketScriptError(Client.ClientStream);
                    break;

                case (Packet.Iterations):
                    Client.HandlePacketOtherError(Client.ClientStream);
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
            int iterations = Client.ParseIterations(Client.ClientStream);
            string message = Client.ParseFilePath(Client.ClientStream);
            string unwrappedScript = InputParser.ExtractScript(message);
            string wrappedScript = InputParser.WrapScript(unwrappedScript);
            RunMacro(wrappedScript, iterations);
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

                    var errorMessage = string.Format("From thread: An error occurred: {0}: {1}", e.Message, e.GetBaseException());
                }
            });

            readThread.SetApartmentState(ApartmentState.STA);
            return readThread;
        }

        private static void RunFromPipe(string[] separatedArgs)
        {
            string guid = InputParser.GetGuid(separatedArgs[1]);
            int pid = InputParser.GetPid(separatedArgs[2]);

            Client.InitializePipeClientStream(guid);

            Thread readThread = CreateReadingThread(pid);
            readThread.Start();

            // byte[] packet = Client.PackageFilePathMessage("hello from engine!");
            // Client.SendMessageToServer(Client.ClientStream, packet);
        }

        internal static void Main(string[] args)
        {
            try
            {
                if (args.Length > 0)
                {
                    var pipeToken = "@";
                    string[] separatedArgs = InputParser.SeparateArgs(args);

                    if (separatedArgs[0].Equals(pipeToken))
                    {
                        RunFromPipe(separatedArgs);
                    }
                    else
                    {
                        RunFromExtension(separatedArgs);
                    }
                }
                else
                {
                    RunAsStartupProject();
                }
            }
            catch (Exception e)
            {
                var error = string.Format("Error at {0} from method {1}\n\n{2}", e.Source, e.TargetSite, e.Message);
                MessageBox.Show(error);
            }
        }
    }
}

﻿//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using VSMacros.ExecutionEngine;
using VSMacros.ExecutionEngine.Pipes;
using VSMacros.ExecutionEngine.Stubs;
using VSMacros.ExecutionEngine.Helpers;
using ExecutionEngine.Enums;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        internal static string macroName = "currentScript";
        private const string close = "close";
        private static bool exit;

        internal static void RunMacro(string script, int iterations = 1)
        {
            Validate.IsNotNullAndNotEmpty(script, "script");
            Program.parsedScript = Program.engine.Parse(script);

            for (int i = 0; i < iterations; i++)
            {
                Program.parsedScript.CallMethod(Program.macroName);
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
            //string cleaned = InputParser.RemoveComments(unwrapped);
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
                    Client.HandlePacketClose(Client.ClientStream);
                    break;

                case (Packet.Success):
                    Client.HandlePacketSuccess(Client.ClientStream);
                    break;

                case (Packet.ScriptError):
                    Client.HandlePacketScriptError(Client.ClientStream);
                    break;

                case (Packet.OtherError):
                    Client.HandlePacketOtherError(Client.ClientStream);
                    break;
            }
        }

        private static void HandleFilePath()
        {
            string message = Client.ParseFilePath(Client.ClientStream);
            if (InputParser.IsDebuggerStopped(message))
            {
                Client.ClientStream.Close();
                Program.exit = true;
            }

            else
            {
                if (InputParser.IsRequestToClose(message))
                {
                    Client.ShutDownServer(Client.ClientStream);
                    Client.ClientStream.Close();
                    Program.exit = true;
                }
                else
                {
                    string unwrappedScript = InputParser.ExtractScript(message);
                    string wrappedScript = InputParser.WrapScript(unwrappedScript);
                    RunMacro(wrappedScript, iterations: 1);
                }
            }
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
                    var errorMessage = string.Format("An error occurred: {0}: {1}", e.Message, e.GetBaseException());
                    MessageBox.Show(errorMessage);

                    Client.ShutDownServer(Client.ClientStream);
                    Client.ClientStream.Close();
                    Program.exit = true;
                }
            });

            return readThread;
        }

        private static void RunFromPipe(string[] separatedArgs)
        {
            string guid = InputParser.GetGuid(separatedArgs[1]);
            int pid = InputParser.GetPid(separatedArgs[2]);

            Client.InitializePipeClientStream(guid);

            Thread readThread = CreateReadingThread(pid);
            readThread.Start();

            //while (!Program.exit)
            //{
            //    Client.SendMessageToServer();
            //}
        }

        internal static void Main(string[] args)
        {
            // TODO: Close pipes when Visual Studio closes

            Console.WriteLine("Hello there!  Welcome to our macro extension!");

            try
            {
                if (args.Length > 0)
                {
                    var pipeToken = "@";
                    string[] separatedArgs = InputParser.SeparateArgs(args);

                    if (separatedArgs[0].Equals(pipeToken))
                    {
                        Console.WriteLine("running macro from pipe");
                        RunFromPipe(separatedArgs);
                    }
                    else
                    {
                        Console.WriteLine("running macro from extension");
                        RunFromExtension(separatedArgs);
                    }
                }
                else
                {
                    Console.WriteLine("running as startup");
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

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
using VisualStudio.Macros.ExecutionEngine.Pipes;
using VisualStudio.Macros.ExecutionEngine.Stubs;
using VSMacros.ExecutionEngine;
using VSMacros.ExecutionEngine.Helpers;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        private const string macroName = "currentScript";
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
            string unwrappedScript = InputParser.ExtractScript(decodedPath);
            string wrappedScript = InputParser.WrapScript(unwrappedScript);

            Program.engine = new Engine(pid);
            var unwrapped = InputParser.ExtractScript(decodedPath);
            var wrapped = InputParser.WrapScript(unwrapped);

            RunMacro(wrapped, iterations);
        }

        internal static void RunAsStartupProject()
        {
            Debug.WriteLine("Warning: Hardcoded devenv pid");
            short pidOfCurrentDevenv = 8516;

            Program.engine = new Engine(pidOfCurrentDevenv);

            Debug.WriteLine("Warning: Hardcoded path");
            string unwrapped = File.ReadAllText(@"C:\Users\t-grawa\Desktop\test.js");
            string wrapped = InputParser.WrapScript(unwrapped);

            RunMacro(wrapped, 1);
        }

        internal static string GetMacroFromStream()
        {
            int sizeOfFilePath = Client.GetSizeOfMessageFromStream(Client.ClientStream);
            string filePath = Client.GetMessageFromStream(Client.ClientStream, sizeOfFilePath);
            string unwrappedScript = InputParser.ExtractScript(filePath);
            string wrappedScript = InputParser.WrapScript(unwrappedScript);

            return wrappedScript;
        }

        internal static void RunMacroAsThread(out Thread readThread, int pid)
        {
            Program.exit = false;
            readThread = new Thread(() =>
            {
                try
                {
                    Program.engine = new Engine(pid);
                    while (!Program.exit)
                    {
                        string script = GetMacroFromStream();
                        RunMacro(script, iterations: 1);
                    }
                }
                catch (Exception e)
                {
                    var errorMessage = string.Format("An error occurred: {0}", e.Message);
                    MessageBox.Show(errorMessage);
                    Client.ShutDownServer(Client.ClientStream, close);
                    Client.ClientStream.Close();
                    Program.exit = true;
                }
            });
        }

        private static void SendMessageToServer()
        {
            Console.Write("\n>> ");
            string line = Console.ReadLine();
            if (line.ToLower() == "close")
            {
                Client.ShutDownServer(Client.ClientStream, close);
                Program.exit = true;
            }

            byte[] message = Client.PackageMessage(line);
            Client.SendMessageToServer(Client.ClientStream, message);
        } 

        private static void RunFromPipe(string[] separatedArgs)
        {
            Console.WriteLine("Initializing the engine and the pipes");
            string guid = InputParser.GetGuid(separatedArgs[1]);
            int pid = InputParser.GetPid(separatedArgs[2]);

            Client.InitializePipeClientStream(guid);

            Thread readMacroFromStream;
            RunMacroAsThread(out readMacroFromStream, pid);
            readMacroFromStream.Start();

            //while (!Program.exit)
            //{
                //SendMessageToServer();
            //}
        }

        internal static void Main(string[] args)
        {
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
                var errorMessage = string.Format("An error occurred: {0}", e.Message);
                MessageBox.Show(errorMessage);
            }
        }
    }
}

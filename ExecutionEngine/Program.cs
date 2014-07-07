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
            var unwrapped = Script.CreateScriptStub();
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

        internal static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Hello there!  Welcome to our macro extension!");

                if (args.Length > 0)
                {
                    string[] separatedArgs = InputParser.SeparateArgs(args);

                    if (separatedArgs[0] == "@")
                    {
                        Console.WriteLine("Initializing the engine and the pipes");
                        string guid = separatedArgs[1];
                        Console.WriteLine("guid is: " + guid);
                        int pid = InputParser.GetPid(separatedArgs[2]);
                        Console.WriteLine("pid is: " + pid);
                        // TODO: Check if guid is null

                        //MessageBox.Show("set breakpoint here");
                        Client.InitializePipeClientStream(guid);

                        bool exit = false;
                        Thread readThread = new Thread(() =>
                        {
                            try
                            {
                                Program.engine = new Engine(pid);
                                while (!exit)
                                {
                                    int sizeOfFilePath = Client.GetSizeOfMessageFromStream(Client.ClientStream);
                                    string filePath = Client.GetMessageFromStream(Client.ClientStream, sizeOfFilePath);
                                    string unwrappedScript = InputParser.ExtractScript(filePath);
                                    string wrappedScript = InputParser.WrapScript(unwrappedScript);
                                    RunMacro(wrappedScript, iterations: 1);
                                }
                            }
                            catch (Exception e)
                            {
                                var errorMessage = string.Format("An error occurred: {0}", e.Message);
                                MessageBox.Show(errorMessage);
                            }
                        });

                        readThread.Start();

                        while (true)
                        {
                            Console.Write("\n>> ");
                            string line = Console.ReadLine();
                            if (line.ToLower() == "close")
                            {
                                Client.ShutDownServer(Client.ClientStream, line);
                                break;
                            }

                            byte[] message = Client.PackageMessage(line);
                            Client.SendMessageToServer(Client.ClientStream, message);
                        }
                        Client.ClientStream.Close();
                        exit = true;
                    }
                    else
                    {
                        Console.WriteLine("Running macro as normal");
                        //MessageBox.Show("Attach debugger here");
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

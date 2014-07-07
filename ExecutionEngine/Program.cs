//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using VisualStudio.Macros.ExecutionEngine.Pipes;
using VisualStudio.Macros.ExecutionEngine.Stubs;
using VSMacros.ExecutionEngine;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        private static string macroName = "currentScript";
        //public static NamedPipeClientStream clientStream;

        private static string[] SeparateArgs(string[] args)
        {
            string[] stringSeparator = new string[] {"[delimiter]"};
            string[] separatedArgs = args[0].Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);
            return separatedArgs;
        }

        private static short GetNumberOfIterations(string iter)
        {
            short iterations;

            Validate.IsNotNullAndNotEmpty(iter, "iter");

            if (!short.TryParse(iter, out iterations))
            {
                throw new ArgumentException(string.Format(CultureInfo.CurrentUICulture, Resources.InvalidIterationsArgument, iter, "iter"));
            }

            return iterations;
        }

        private static string ExtractScript(string path)
        {
            Validate.IsNotNullAndNotEmpty(path, "path");
            return File.ReadAllText(path);
        }

        private static string WrapScript(string unwrapped)
        {
            string wrapped = "function currentScript() {";
            wrapped += unwrapped;
            wrapped += "}";

            return wrapped;
        }

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

            int pid = GetPid(unparsedPid);
            short iterations = GetNumberOfIterations(unparsedIter);
            string decodedPath = DecodePath(encodedPath);
            string unwrappedScript = ExtractScript(decodedPath);
            string wrappedScript = WrapScript(unwrappedScript);

            Program.engine = new Engine(pid);
            var unwrapped = Script.CreateScriptStub();
            var wrapped = WrapScript(unwrapped);

            //Console.WriteLine("running the wrapped scdript");
            RunMacro(wrapped, iterations);
        }

        private static int GetPid(string unparsedPid)
        {
            int pid;

            Validate.IsNotNullAndNotEmpty(unparsedPid, "unparsedPid");

            if (!int.TryParse(unparsedPid, out pid))
            {
                MessageBox.Show("The pid is invalid.");
                throw new ArgumentException(string.Format(CultureInfo.CurrentUICulture, Resources.InvalidPIDArgument, unparsedPid, "unparsedPid"));
            }

            return pid;
        }

        private static string DecodePath(string encodedPath)
        {
            return encodedPath.Replace("%20", " ");
        }

        internal static void RunAsStartupProject()
        {
            Debug.WriteLine("Warning: Hardcoded devenv pid");
            short pidOfCurrentDevenv = 8516;

            Program.engine = new Engine(pidOfCurrentDevenv);

            Debug.WriteLine("Warning: Hardcoded path");
            string unwrapped = File.ReadAllText(@"C:\Users\t-grawa\Desktop\test.js");
            string wrapped = WrapScript(unwrapped);
            //Console.WriteLine("wrapped is: " + wrapped);

            RunMacro(wrapped, 1);
        }

        internal static void Main(string[] args)
        {
            Console.WriteLine("Hello there!  Welcome to our macro extension!");

            if (args.Length > 0)
            {
                string[] separatedArgs = SeparateArgs(args);

                if (separatedArgs[0] == "@")
                {
                    Console.WriteLine("Initializing the engine and the pipes");
                    string guid = separatedArgs[1];
                    Console.WriteLine("guid is: " + guid);
                    int pid = GetPid(separatedArgs[2]);
                    Console.WriteLine("pid is: " + pid);
                    // TODO: Check if guid is null

                    //MessageBox.Show("set breakpoint here");
                    Client.InitializePipeClientStream(guid);

                    bool exit = false;
                    Thread readThread = new Thread(() => 
                    {
                        Program.engine = new Engine(pid);
                        while(!exit)
                        {
                            int sizeOfFilePath = Client.GetSizeOfMessageFromStream(Client.ClientStream);
                            string filePath = Client.GetMessageFromStream(Client.ClientStream, sizeOfFilePath);
                            string unwrappedScript = ExtractScript(filePath);
                            string wrappedScript = WrapScript(unwrappedScript);
                            RunMacro(wrappedScript, iterations: 1);
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
    }
}

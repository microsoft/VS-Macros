//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        private static short pid;

        internal static void CreateMacroFile(string path)
        {
            using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine("function dteTest()");
                sw.WriteLine("{");
                sw.WriteLine("dte.ExecuteCommand('File.NewFile')");
                sw.Write("}");
            }
        }

        internal static string ReadFromMacroFile(string path)
        {
            string script = string.Empty;

            if (!File.Exists(path))
            {
                throw new FileNotFoundException(path);
            }
            
            using (StreamReader sr = new StreamReader(path))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    script += line;
                }
            }

            return script;
        }

        internal static void RunMacro(string macroName, int pid)
        {
            var macroPath = AppendExtension(macroName);

            CreateMacroFile(macroPath);
            var script = ReadFromMacroFile(macroPath);

            if (!string.IsNullOrEmpty(script))
            {
                parsedScript = engine.Parse(script);
            }
            else
            {
                throw new NullReferenceException(script);
            }
            
            var output = parsedScript.CallMethod(macroName);
        }

        internal static bool IsVSShuttingDown()
        {
            // TODO: Don't know how to implement this
            // Is there an event I can subscribe to here?
            return false;
        }
        internal static bool ListenForInput(string macroName)
        {
            var willKeepListening = true;

            var input = Console.ReadLine();
            if (!string.IsNullOrEmpty(input))
            {
                if (input.Equals("r"))
                {
                    Console.WriteLine(">> Running your macro");
                    RunMacro(macroName, pid);
                }
                else if (input.Equals("q"))
                {
                    Console.WriteLine(">> Quitting");
                    willKeepListening = false;
                }
            }

            if (IsVSShuttingDown())
            {
                willKeepListening = false;
            }

            return willKeepListening;
        }

        internal static void RunFromExtension(string macroName, string[] args)
        {
            if (short.TryParse(args[0], out pid))
            {
                engine = new Engine(pid);
                RunMacro(macroName, pid);
            }
            else
            {
                // TODO: Is throwing an exception the right thing to do here?
                // And if it is, is this the right exception to throw?
                throw new ArgumentException(args[0]);
            }
        }

        internal static void RunAsStartupProject(string macroName, short temp_pid)
        {
            pid = temp_pid;
            engine = new Engine(pid);
            RunMacro(macroName, pid);
        }

        internal static string AppendExtension(string name) 
        {
            return name + ".txt";
        }

        internal static void Main(string[] args)
        {
            var macroName = "dteTest";

            Console.WriteLine("Hello there!  Welcome to our macro extension!");

            if (args.Length > 0)
            {
                RunFromExtension(macroName, args);
            }
            else
            {
                short pidOfCurrentDevenv = 9344;
                RunAsStartupProject(macroName, pidOfCurrentDevenv);
            }

            // TODO: this while loop is a temp fix for now until I figure out named pipes
            // My hacky way of IPC for now
            // while (true)
            // {
            //    if (!ListenForInput(macroName)) 
            //        break;
            // }
        }
    }
}

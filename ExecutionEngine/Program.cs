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

        // Helper methods

        internal static string AppendExtension(string name)
        {
            return name + ".txt";
        }

        private static string[] SeparateArgs(string[] args)
        {
            return args[0].Split(',');
        }
        private static short DetermineNumberOfTimes(string[] args)
        {
            short times;

            if (args.Length < 2 || !short.TryParse(args[1], out times))
            {
                return 1;
            }

            return times;
        }
        internal static bool IsVSShuttingDown()
        {
            // TODO: Don't know how to implement this
            // Is there an event I can subscribe to here?
            return false;
        }

        internal static StreamReader CreateMacroStreamReader(string path)
        {
            using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine("function dteTest()");
                sw.WriteLine("{");
                sw.WriteLine("dte.ExecuteCommand('File.NewFile');");
                sw.Write("}");
            }

            return new StreamReader(path);
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

        private static string CreateScriptFromReader(StreamReader reader)
        {
            var script = string.Empty;
            using (reader)
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    script += line;
                }
            }
            return script;
        }

        internal static void RunMacro(string name, string script, int times = 1)
        {
            if (!string.IsNullOrEmpty(script))
            {
                parsedScript = engine.Parse(script);
            }
            else
            {
                throw new NullReferenceException(script);
            }

            for (int i = 0; i < times; i++)
            {
                parsedScript.CallMethod(name);
            }
        }

        internal static bool ListenForInput(string name)
        {
            var willKeepListening = true;

            var input = Console.ReadLine();
            if (!string.IsNullOrEmpty(input))
            {
                if (input.Equals("r"))
                {
                    Console.WriteLine(">> Running your macro");
                    var path = AppendExtension(name);
                    var reader = CreateMacroStreamReader(path);
                    var script = CreateScriptFromReader(reader);
                    RunMacro(name, script);
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

        internal static void RunFromExtension(string name, StreamReader reader, string[] args)
        {
            if (short.TryParse(args[0], out pid))
            {
                engine = new Engine(pid);

                var times = DetermineNumberOfTimes(args);
                var macroScript = CreateScriptFromReader(reader);
                RunMacro(name, macroScript, times);
            }
            else
            {
                throw new ArgumentException(args[0]);
            }
        }

        internal static void RunAsStartupProject(string name, StreamReader reader, short tempPid)
        {
            pid = tempPid;
            engine = new Engine(pid);
            var script = CreateScriptFromReader(reader);
            RunMacro(name, script);
        }

        internal static void Main(string[] args)
        {
            var name = "dteTest";
            var path = AppendExtension(name);
            var reader = CreateMacroStreamReader(path);

            Console.WriteLine("Hello there!  Welcome to our macro extension!");

            if (args.Length > 0)
            {
                var separatedArgs = SeparateArgs(args);
                RunFromExtension(name, reader, separatedArgs);
            }
            else
            {
                short pidOfCurrentDevenv = 1860;
                RunAsStartupProject(name, reader, pidOfCurrentDevenv);
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

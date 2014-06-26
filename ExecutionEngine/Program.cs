using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExecutionEngine
{
    class Program
    {
        static Engine engine;
        static ParsedScript parsedScript;
        private static short pid;

        static void CreateMacroFile(string path)
        {
            using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine("function dteTest()");
                sw.WriteLine("{");
                sw.WriteLine("dte.ExecuteCommand('File.NewFile')");
                sw.Write("}");
            }
        }

        static string ReadFromMacroFile(string path)
        {
            string script = "";

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

        static void RunMacro(string macroName, int pid)
        {
            var macroPath = AppendExtension(macroName);

            CreateMacroFile(macroPath);
            var script = ReadFromMacroFile(macroPath);

            if (!String.IsNullOrEmpty(script))
            {
                parsedScript = engine.Parse(script);
            }
            else
            {
                throw new NullReferenceException(script);
            }
            
            var output = parsedScript.CallMethod(macroName);
        }

        static bool IsVSShuttingDown()
        {
            // TODO: Don't know how to implement this
            // Is there an event I can subscribe to here?
            return false;
        }
        static bool ListenForInput(string macroName)
        {
            var willKeepListening = true;

            var input = Console.ReadLine();
            if (!String.IsNullOrEmpty(input))
            {
                Console.WriteLine(">> You typed in: " + input);
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

            input = null;
            return willKeepListening;
        }

        static void RunAsStartupProject(string macroName, short temp_pid)
        {
            pid = temp_pid;
            engine = new Engine(pid);
            RunMacro(macroName, pid);
        }

        static void RunFromExtension(string macroName, string[] args)
        {
            if (Int16.TryParse(args[0], out pid))
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

        static string AppendExtension(string name) 
        {
            return name + ".txt";
        }

        static void Main(string[] args)
        {
            var macroName = "dteTest";

            Console.WriteLine("hello there!");

            if (args.Length > 0)
            {
                RunFromExtension(macroName, args);
            }
            else
            {
                RunAsStartupProject(macroName, 96);
            }

            // TODO: this while loop is a temp fix for now until I figure out named pipes
            while (true)
            {
                if (!ListenForInput(macroName)) 
                    break;
            }
        }
    }
}

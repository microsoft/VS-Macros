using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExecutionEngine
{
    class Program
    {
        static Engine engine;
        static ParsedScript parsedScript;
        private static short pid;

        static void LaunchVS()
        {
            var processName = @"devenv.exe";
            var process = new Process();

            process.StartInfo.FileName = processName;
            process.Start();
        }

        static void CreateHardcodedStream(string fileName)
        {
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            using (StreamWriter sw = new StreamWriter(fileName))
            {
                sw.WriteLine("function dteTest()");
                sw.WriteLine("{");
                sw.WriteLine("dte.ExecuteCommand('File.NewFile')");
                sw.Write("}");
            }
        }

        static string ReadFromStream(string fileName)
        {
            string script = "";
            using (StreamReader sr = new StreamReader(fileName))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    script += line;
                }
            }

            Console.WriteLine(script);
            return script;
        }

        static void CreateEngine(int pid)
        {
            engine = new Engine(pid);
        }

        static void RunMacro(string nameWithoutExtension, int pid)
        {
            var fileName = AppendExtension(nameWithoutExtension);

            CreateHardcodedStream(fileName);
            var script = ReadFromStream(fileName);

            parsedScript = engine.Parse(script);
            var output = parsedScript.CallMethod(nameWithoutExtension);
        }

        static bool IsRunRequested(string s)
        {
            return (s == "r");
        }

        static bool IsInitializeRequested(string s)
        {
            return (s == "i");
        }

        static bool IsVSShuttingDown()
        {
            // TODO: Don't know how to implement this
            // Is there an event I can subscribe to here?
            return false;
        }
        static bool ListenForInput(string nameWithoutExtension)
        {
            var willKeepListening = true;

            var input = Console.ReadLine();
            if (!String.IsNullOrEmpty(input))
            {
                Console.WriteLine(">> You typed in: " + input);
                if (input.Equals("r"))
                {
                    Console.WriteLine(">> Running your macro");
                    RunMacro(nameWithoutExtension, pid);
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

        static void RunningAsStartupProject(string nameWithoutExtension, short temp_pid)
        {
            pid = temp_pid;
            CreateEngine(pid);
            RunMacro(nameWithoutExtension, pid);
        }

        static void RunningFromExtension(string nameWithoutExtension, string[] args)
        {
            var command = args[0].Split(',');

            if (IsInitializeRequested(command[0]))
            {
                if (Int16.TryParse(command[1], out pid))
                {
                    CreateEngine(pid);
                    RunMacro(nameWithoutExtension, pid);
                }
            }
            if (IsRunRequested(command[0]))
            {
                RunMacro(nameWithoutExtension, pid);
            }
        }

        static string AppendExtension(string name) 
        {
            return name + ".txt";
        }

        static void Main(string[] args)
        {
            var nameWithoutExtension = "dteTest";
            var nameWithExtension = AppendExtension(nameWithoutExtension);

            Console.WriteLine("hello there!");

            if (args.Length > 0)
            {
                RunningFromExtension(nameWithoutExtension, args);
            }
            else
            {
                RunningAsStartupProject(nameWithoutExtension, 96);
            }

            //LaunchVS();
            //Debug.WriteLine("Current directory: " + Directory.GetCurrentDirectory());
            //Debug.WriteLine("Using the assembly thing: " + Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)));

            // TODO: this while loop is a temp fix for now until I can figure out how to specify Initialize versus Run from VSMacros
            while (true)
            {
                if (!ListenForInput(nameWithoutExtension)) 
                    break;
            }
        }
    }
}

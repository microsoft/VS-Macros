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
        private static int pid;
        private static string macroName = "currentScript";

        // Helper methods
        private static string[] SeparateArgs(string[] args)
        {
            string[] stringSeparator = new string[] {"[delimiter]"};
            string[] separatedArgs = args[0].Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine(args[i]);
            }

            return separatedArgs;
        }
        private static short DetermineNumberOfIterations(string[] args)
        {
            short iterations;

            if (args.Length < 2 || !short.TryParse(args[1], out iterations))
            {
                return 1;
            }

            return iterations;
        }

        private static string ExtractScript(string[] args)
        {
            if (args.Length > 2)
            {
                return args[2];
            }
            else
            {
                MessageBox.Show("You did not provide a script");
                return string.Empty;
            }
        }

        private static string WrapScriptInFunction(string unwrapped)
        {
            string wrapped = "function currentScript() {";
            wrapped += unwrapped;
            wrapped += "}";

            return wrapped;
        }

        internal static void RunMacro(string script, int iterations = 1)
        {
            if (!string.IsNullOrEmpty(script))
            {
                Program.parsedScript = Program.engine.Parse(script);
            }
            else
            {
                throw new Exception(script);
            }

            for (int i = 0; i < iterations; i++)
            {
                Program.parsedScript.CallMethod(Program.macroName);
            }
        }

        internal static void RunFromExtension(string[] args)
        {
            if (int.TryParse(args[0], out Program.pid))
            {
                Program.engine = new Engine(Program.pid);

                short iterations = DetermineNumberOfIterations(args);
                string unwrappedScript = ExtractScript(args);
                string wrappedScript = WrapScriptInFunction(unwrappedScript);

                RunMacro(wrappedScript, iterations);
            }
            else
            {
                MessageBox.Show("You did not provide any arguments to the Execution Engine");
                throw new ArgumentException(args[0]);
            }
        }

        internal static void RunAsStartupProject(int tempPid)
        {
            Program.pid = tempPid;
            Program.engine = new Engine(Program.pid);
            string script = CreateScriptStub();
            RunMacro(script, 2);
        }

        private static string CreateScriptStub()
        {
            return "function currentScript() { dte.ExecuteCommand('File.NewFile'); } ";
        }

        internal static void Main(string[] args)
        {
            Console.WriteLine("Hello there!  Welcome to our macro extension!");

            if (args.Length > 0)
            {
                string[] separatedArgs = SeparateArgs(args);
                RunFromExtension(separatedArgs);
            }
            else
            {
                var path = @"C:\Users\t-grawa\Source\Repos\Macro Extension\ExecutionEngine\bin\Debug\dteTest.js";
                var reader = new StreamReader(path);
                short pidOfCurrentDevenv = 4300;
                RunAsStartupProject(pidOfCurrentDevenv);
            }
        }

        private static string GetPathFromArgs(string[] args)
        {
            if (args[2] == null)
            {
                MessageBox.Show("A path to the macro file was not provided");
                return string.Empty;
            }

            return args[2];
        }
    }
}

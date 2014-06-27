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
        private static string macroName = "currentScript";

        // Helper methods
        private static string[] SeparateArgs(string[] args)
        {
            string[] stringSeparator = new string[] {"[delimiter]"};
            var separatedArgs = args[0].Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < args.Length; i++)
            {
                Console.WriteLine(args[i]);
            }

            return separatedArgs;
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
            var wrapped = "function currentScript() {";
            wrapped += unwrapped;
            wrapped += "}";

            return wrapped;
        }

        internal static void RunMacro(string script, int times = 1)
        {
            if (!string.IsNullOrEmpty(script))
            {
                parsedScript = engine.Parse(script);
            }
            else
            {
                MessageBox.Show("The script is empty");
                throw new NullReferenceException(script);
            }

            for (int i = 0; i < times; i++)
            {
                parsedScript.CallMethod(Program.macroName);
            }
        }

        internal static void RunFromExtension(string[] args)
        {
            if (short.TryParse(args[0], out pid))
            {
                engine = new Engine(pid);

                var times = DetermineNumberOfTimes(args);
                var unwrappedScript = ExtractScript(args);
                var wrappedScript = WrapScriptInFunction(unwrappedScript);
                RunMacro(wrappedScript, times);
            }
            else
            {
                MessageBox.Show("You did not provide any arguments to the Execution Engine");
                throw new ArgumentException(args[0]);
            }
        }

        internal static void RunAsStartupProject(short tempPid)
        {
            pid = tempPid;
            engine = new Engine(pid);
            var script = CreateScriptStub();
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
                var separatedArgs = SeparateArgs(args);
                RunFromExtension(separatedArgs);
            }
            else
            {
                var path = @"C:\Users\t-grawa\Source\Repos\Macro Extension\ExecutionEngine\bin\Debug\dteTest.txt";
                var reader = new StreamReader(path);
                short pidOfCurrentDevenv = 1860;
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
